# -*- coding: utf-8 -*-
"""Lê a planilha de validação (devolvida pela Qualidade, já com as respostas
dos setores) e atualiza no GLPI o campo "Inserir na avaliação motorista?" de
cada chamado.

O que este script NÃO faz (de propósito, por decisão da Qualidade): não
calcula desconto de pontos, não decide o que é maioria/empate entre setores.
Isso é feito por fora, pela Qualidade — o sistema só reflete a decisão final
que já veio pronta na planilha.

Uso:
    python importar_validacao_qualidade.py caminho_planilha_preenchida.xlsx

Requer config.json na mesma pasta. Não dispara notificação nenhuma (mesmo
padrão usado no importador geral: _disablenotifications + _disablenotif).
"""
import json
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

CONTAINER_CLASSIFICACAO = 6
PASTA = Path(__file__).resolve().parent
ARQ_ESTADO = PASTA / "estado_validacao_qualidade.json"

RESPOSTAS_SIM = {"sim", "s", "yes", "1"}
RESPOSTAS_NAO = {"não", "nao", "n", "no", "0"}


def api_request(base, app_token, method, path, headers=None, data=None):
    url = base + path
    h = {"App-Token": app_token, "Content-Type": "application/json"}
    if headers:
        h.update(headers)
    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, method=method, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Erro na API ({e.code}): {e.read().decode()}") from e


def carregar_estado():
    if ARQ_ESTADO.exists():
        return json.loads(ARQ_ESTADO.read_text(encoding="utf-8"))
    return {"exportados": {}, "importados": {}}


def salvar_estado(estado):
    ARQ_ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def ler_planilha(caminho):
    wb = load_workbook(caminho, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, 8):
        valores = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if "ID" in valores:
            header_row = r
            break
    if header_row is None:
        raise RuntimeError("Não encontrei a linha de cabeçalho (coluna \"ID\") na planilha.")

    col_por_nome = {}
    for c in range(1, ws.max_column + 1):
        nome = ws.cell(row=header_row, column=c).value
        if nome:
            col_por_nome[str(nome).strip()] = c

    col_id = col_por_nome.get("ID")
    col_decisao = col_por_nome.get("Confirmado como culpa do motorista? (Sim/Não)")
    if not col_id or not col_decisao:
        raise RuntimeError("Planilha não tem as colunas esperadas (\"ID\" e a coluna de decisão).")

    decisoes = []
    for r in range(header_row + 1, ws.max_row + 1):
        tid = ws.cell(row=r, column=col_id).value
        resposta = ws.cell(row=r, column=col_decisao).value
        if not tid or not resposta:
            continue
        resposta_norm = str(resposta).strip().lower()
        if resposta_norm in RESPOSTAS_SIM:
            decisoes.append((int(tid), 1, str(resposta).strip()))
        elif resposta_norm in RESPOSTAS_NAO:
            decisoes.append((int(tid), 0, str(resposta).strip()))
        else:
            print(f"  AVISO: linha {r} (chamado {tid}) tem resposta \"{resposta}\" não reconhecida (esperado Sim/Não) — ignorada.")
    return decisoes


def buscar_registro_container6(sreq, ticket_id):
    linhas = sreq("GET", f"/PluginFieldsTicketdadosdaocorrncia?range=0-3000")
    for l in linhas:
        if l.get("items_id") == ticket_id and l.get("plugin_fields_containers_id") == CONTAINER_CLASSIFICACAO:
            return l
    return None


def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_validacao_qualidade.py caminho_planilha_preenchida.xlsx")
        sys.exit(1)

    caminho = Path(sys.argv[1])
    if not caminho.exists():
        print(f"Arquivo não encontrado: {caminho}")
        sys.exit(1)

    cfg = json.loads((PASTA / "config.json").read_text(encoding="utf-8-sig"))
    base_url = cfg["api_url"]
    app_token = cfg["app_token"]
    user_token = cfg["user_token"]

    print(f"Lendo {caminho.name}...")
    decisoes = ler_planilha(caminho)
    print(f"{len(decisoes)} decisão(ões) reconhecida(s) na planilha.")
    if not decisoes:
        print("Nada pra importar.")
        return

    init = api_request(base_url, app_token, "GET", "/initSession", {"Authorization": "user_token " + user_token})
    session_token = init["session_token"]

    def sreq(method, path, data=None):
        return api_request(base_url, app_token, method, path, {"Session-Token": session_token}, data)

    try:
        print("Buscando registros de Classificação de todos os chamados (uma chamada só)...")
        todas_linhas_container6 = sreq("GET", "/PluginFieldsTicketdadosdaocorrncia?range=0-3000")
        por_ticket = {
            l["items_id"]: l for l in todas_linhas_container6
            if l.get("plugin_fields_containers_id") == CONTAINER_CLASSIFICACAO
        }

        estado = carregar_estado()
        agora = f"{datetime.now():%d/%m/%Y}"
        aplicados = 0
        for ticket_id, valor, resposta_original in decisoes:
            registro = por_ticket.get(ticket_id)
            if not registro:
                print(f"  chamado {ticket_id}: sem registro de Classificação no GLPI — pulado.")
                continue

            sreq(
                "PUT",
                f"/PluginFieldsTicketdadosdaocorrncia/{registro['id']}",
                {"input": {
                    "id": registro["id"],
                    "avaliacaomotoristafield": valor,
                    "_disablenotifications": 1,
                    "_disablenotif": 1,
                }},
            )
            estado["importados"][str(ticket_id)] = {
                "data_importado": agora,
                "decisao": resposta_original,
                "arquivo": caminho.name,
            }
            aplicados += 1
            print(f"  chamado {ticket_id}: \"Inserir na avaliação motorista?\" = {resposta_original}")

        salvar_estado(estado)
        print(f"\n{aplicados} chamado(s) atualizado(s) no GLPI. Estado salvo em {ARQ_ESTADO.name}.")
    finally:
        sreq("GET", "/killSession")


if __name__ == "__main__":
    main()
