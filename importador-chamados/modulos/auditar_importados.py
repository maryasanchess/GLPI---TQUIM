#!/usr/bin/env python3
"""Audita, chamado por chamado, o que já foi importado para o GLPI.

Diferente de verificar_pre_importacao.py (que só olha a planilha ANTES de
importar), este módulo consulta o GLPI de verdade e confere se cada
referência EXCEL-N já registrada no controle está com:
  - a mesma data de abertura da planilha;
  - o mesmo status da planilha;
  - o requerente (solicitante) vinculado como usuário requerente;
  - o técnico (Responsável) vinculado como usuário atribuído;
  - o grupo/setor do Departamento vinculado como grupo requerente;
  - algum grupo de TI vinculado como grupo atribuído.

Este módulo só lê o GLPI. Ele nunca cria, atualiza ou apaga nada.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.error
from collections import Counter
from pathlib import Path

import pandas as pd

# Reaproveita as funções já testadas do importador (mesma pasta modulos/).
import importar_chamados as imp

TIPO_TICKET_USER_REQUERENTE = 1
TIPO_TICKET_USER_ATRIBUIDO = 2
TIPO_GROUP_TICKET_REQUERENTE = 1
TIPO_GROUP_TICKET_ATRIBUIDO = 2


class Relatorio:
    def __init__(self):
        self.linhas = []
        self.contadores = Counter()

    def secao(self, titulo):
        self.linhas.append("")
        self.linhas.append(f"== {titulo} ==")

    def linha(self, msg):
        self.linhas.append(msg)

    def texto_final(self, total_verificados):
        cabecalho = [
            "AUDITORIA DOS CHAMADOS JÁ IMPORTADOS",
            f"Chamados verificados: {total_verificados}",
            f"Divergências: {dict(self.contadores)}",
        ]
        return "\n".join(cabecalho + self.linhas) + "\n"


def limitar(itens, maximo=15):
    itens = list(itens)
    if len(itens) <= maximo:
        return ", ".join(str(i) for i in itens)
    resto = len(itens) - maximo
    return ", ".join(str(i) for i in itens[:maximo]) + f" (+{resto} outros)"


def montar_mapas_normalizados(config):
    return {
        "tecnicos": {imp.chave(k): v for k, v in config.get("technician_map", {}).items()},
        "requerentes": {imp.chave(k): v for k, v in config.get("requester_user_map", {}).items()},
        "departamentos": {imp.chave(k): v for k, v in config.get("department_group_map", {}).items()},
    }


def auditar_ticket(glpi, numero, ticket_id, linha, mapas, rel):
    """Confere um único chamado. Retorna True se está tudo certo."""
    ticket = glpi.obter_ou_none("Ticket", ticket_id)
    if ticket is None:
        rel.contadores["chamado_ausente"] += 1
        rel.linha(f"[AUSENTE] {numero} -> GLPI #{ticket_id} não existe mais (foi apagado?).")
        return False

    problemas = []

    # --- Status ---
    status_planilha = imp.chave(linha.get("Status"))
    status_esperado = imp.STATUS_GLPI.get(status_planilha, 1)
    status_real = int(ticket.get("status") or 0)
    if status_real != status_esperado:
        rel.contadores["status_divergente"] += 1
        problemas.append(f"status esperado={status_esperado} real={status_real}")

    # --- Data de abertura ---
    data_esperada = imp.data_glpi(linha.get("Carimbo de data/hora"))
    if data_esperada:
        data_real = imp.texto(ticket.get("date"))
        if data_real[:10] != data_esperada[:10]:
            rel.contadores["data_divergente"] += 1
            problemas.append(f"data esperada={data_esperada[:10]} real={data_real[:10] or 'vazia'}")

    # --- Vínculos (Ticket_User: requerente/técnico) ---
    try:
        vinculos_usuario = glpi.requisicao("GET", f"Ticket/{ticket_id}/Ticket_User?range=0-99") or []
    except Exception:
        vinculos_usuario = []
    ids_por_tipo_usuario = {}
    for v in vinculos_usuario if isinstance(vinculos_usuario, list) else []:
        tipo = int(v.get("type") or 0)
        ids_por_tipo_usuario.setdefault(tipo, set()).add(int(v.get("users_id") or 0))

    nome_completo = " ".join(filter(None, [
        imp.texto(linha.get("Nome do Solicitante")),
        imp.texto(linha.get("Sobrenome do Solicitante")),
    ]))
    id_requerente_esperado = mapas["requerentes"].get(imp.chave(nome_completo))
    if id_requerente_esperado:
        if int(id_requerente_esperado) not in ids_por_tipo_usuario.get(TIPO_TICKET_USER_REQUERENTE, set()):
            rel.contadores["requerente_nao_vinculado"] += 1
            problemas.append(f"requerente esperado (usuário {id_requerente_esperado}) não está vinculado")
    else:
        rel.contadores["requerente_sem_mapeamento"] += 1

    responsavel = imp.texto(linha.get("Responsável"))
    id_tecnico_esperado = mapas["tecnicos"].get(imp.chave(responsavel))
    if id_tecnico_esperado:
        if int(id_tecnico_esperado) not in ids_por_tipo_usuario.get(TIPO_TICKET_USER_ATRIBUIDO, set()):
            rel.contadores["tecnico_nao_vinculado"] += 1
            problemas.append(f"técnico esperado (usuário {id_tecnico_esperado}) não está atribuído")
    elif responsavel:
        rel.contadores["tecnico_sem_mapeamento"] += 1

    # --- Vínculos (Group_Ticket: setor requerente/TI atribuído) ---
    try:
        vinculos_grupo = glpi.requisicao("GET", f"Ticket/{ticket_id}/Group_Ticket?range=0-99") or []
    except Exception:
        vinculos_grupo = []
    ids_por_tipo_grupo = {}
    for v in vinculos_grupo if isinstance(vinculos_grupo, list) else []:
        tipo = int(v.get("type") or 0)
        ids_por_tipo_grupo.setdefault(tipo, set()).add(int(v.get("groups_id") or 0))

    departamento = imp.texto(linha.get("Departamento"))
    id_grupo_esperado = mapas["departamentos"].get(imp.chave(departamento))
    if id_grupo_esperado:
        if int(id_grupo_esperado) not in ids_por_tipo_grupo.get(TIPO_GROUP_TICKET_REQUERENTE, set()):
            rel.contadores["grupo_setor_nao_vinculado"] += 1
            problemas.append(f"grupo do setor esperado (grupo {id_grupo_esperado}) não está vinculado")
    elif departamento:
        rel.contadores["grupo_setor_sem_mapeamento"] += 1

    if not ids_por_tipo_grupo.get(TIPO_GROUP_TICKET_ATRIBUIDO):
        rel.contadores["grupo_ti_ausente"] += 1
        problemas.append("nenhum grupo de TI atribuído como responsável")

    if problemas:
        rel.linha(f"[DIVERGENTE] {numero} (GLPI #{ticket_id}): " + "; ".join(problemas))
        return False
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Audita no GLPI (somente leitura) os chamados já importados: "
        "requerente, data, status, técnico e grupo/setor."
    )
    parser.add_argument("excel", help="Planilha atual (.xlsx)")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--estado", default="importacao_glpi_estado.json")
    parser.add_argument("--limit", type=int, default=0, help="Máximo de chamados a auditar; 0 = todos")
    parser.add_argument(
        "--somente-recentes",
        type=int,
        default=0,
        help="Audita só as N referências mais recentes do controle (0 = usa --limit normalmente)",
    )
    parser.add_argument("--relatorio", default="auditoria_importados.txt")
    args = parser.parse_args()

    config = imp.carregar_config(Path(args.config).resolve())
    estado = imp.carregar_estado(Path(args.estado).resolve())
    df = pd.read_excel(Path(args.excel).resolve())
    df = imp.normalizar_colunas(df)
    mapas = montar_mapas_normalizados(config)

    importados = estado.get("importados", {})
    referencias = [(k, v) for k, v in importados.items() if isinstance(v, dict) and v.get("ticket_id")]
    referencias.sort(key=lambda item: imp.valor_referencia(item[0]))
    if args.somente_recentes > 0:
        referencias = referencias[-args.somente_recentes:]
    if args.limit > 0:
        referencias = referencias[:args.limit]

    rel = Relatorio()
    glpi = imp.Glpi(config)
    glpi.iniciar()
    total_verificados = 0
    ok = 0
    try:
        for numero, dados in referencias:
            indice = imp.valor_referencia(numero) - 2
            if indice not in df.index:
                rel.contadores["fora_da_planilha_atual"] += 1
                rel.linha(f"[FORA] {numero} não existe mais na planilha atual (posição {indice}).")
                continue
            linha = df.loc[indice]
            total_verificados += 1
            if auditar_ticket(glpi, numero, int(dados["ticket_id"]), linha, mapas, rel):
                ok += 1
            time.sleep(float(config.get("interval_seconds", 0.15)))
            if total_verificados % 100 == 0:
                print(f"... {total_verificados} chamados auditados até agora.")
    finally:
        glpi.finalizar()

    saida = rel.texto_final(total_verificados)
    print(saida)
    print(f"RESUMO: {ok} de {total_verificados} chamado(s) sem nenhuma divergência.")
    relatorio_path = Path(args.relatorio).resolve()
    relatorio_path.parent.mkdir(parents=True, exist_ok=True)
    relatorio_path.write_text(saida, encoding="utf-8")
    print(f"\nRelatório salvo em: {relatorio_path}")

    if total_verificados - ok > 0:
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as erro:
        corpo = erro.read().decode("utf-8", errors="replace")[:1000]
        print(f"Erro da API: HTTP {erro.code} - {erro.reason}\n{corpo}", file=sys.stderr)
        sys.exit(2)
    except Exception as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        sys.exit(1)
