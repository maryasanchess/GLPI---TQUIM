#!/usr/bin/env python3
"""Move os chamados já importados para a entidade de TI (ex.: TQUIM > TI).

Usa os vínculos de ``importacao_glpi_estado.json`` (mesma referência EXCEL-N
-> chamado GLPI dos outros complementos). Não cria, não apaga e não altera
requerente, técnico, categoria, grupo, status, solução ou anexos - só o
campo entities_id do chamado. Chamados que já estão na entidade certa são
pulados (sem gravação).
"""

import argparse
import csv
import time
from datetime import datetime
from pathlib import Path

from atribuir_grupo_ti_chamados_glpi import chave_referencia, montar_plano
from atribuir_setores_chamados_glpi import (
    Glpi,
    carregar_json,
    normalizar,
    salvar_json,
    texto,
)


ENTIDADE_PADRAO = "TQUIM > TI"


def registrar_erro(caminho, referencia, ticket_id, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(["Referência", "Chamado GLPI", "Resultado", "Detalhe"])
        escritor.writerow([referencia, ticket_id, "ERRO", detalhe])


def indice_entidades(entidades):
    # A entidade raiz normalmente tem ID 0, que é "falsy" em Python -
    # "if not entidade.get('id')" descartaria ela por engano.
    indice = {}
    for entidade in entidades:
        if entidade.get("id") is None:
            continue
        completo = texto(entidade.get("completename")) or texto(entidade.get("name"))
        indice[normalizar(completo)] = int(entidade["id"])
    return indice


def localizar_entidade(entidades, config):
    """Resolve o ID da entidade de destino.

    Prioridade: "ti_entity_id" no config.json (explícito). Sem isso, procura
    pelo caminho completo em "ti_entity_completename" (padrão "TQUIM > TI").
    A entidade PRECISA já existir no GLPI (Administração > Entidades); este
    módulo nunca cria entidade.
    """
    entidade_id_config = config.get("ti_entity_id")
    if entidade_id_config:
        entidade_id_config = int(entidade_id_config)
        if any(int(item.get("id") or 0) == entidade_id_config for item in entidades):
            return entidade_id_config
        raise RuntimeError(
            f"A entidade configurada em ti_entity_id={entidade_id_config} "
            "não foi encontrada no GLPI."
        )

    nome_completo = texto(config.get("ti_entity_completename")) or ENTIDADE_PADRAO
    indice = indice_entidades(entidades)
    chave = normalizar(nome_completo)
    if chave not in indice:
        raise RuntimeError(
            f"A entidade '{nome_completo}' não foi encontrada no GLPI. Crie-a em "
            "Administração > Entidades (filha da entidade atual) ou ajuste "
            '"ti_entity_completename" (ou "ti_entity_id") no config.json.'
        )
    return indice[chave]


def main():
    parser = argparse.ArgumentParser(
        description="Move os chamados já importados para a entidade de TI"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--estado", default="importacao_glpi_estado.json")
    parser.add_argument(
        "--limit",
        type=int,
        default=10,
        help="Máximo de chamados; use 0 para todos (padrão: 10)",
    )
    parser.add_argument(
        "--executar",
        action="store_true",
        help="Aplica no GLPI; sem esta opção, apenas simula",
    )
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    estado_path = Path(args.estado).resolve()
    if not estado_path.exists():
        raise FileNotFoundError(f"Controle da importação não encontrado: {estado_path}")

    estado = carregar_json(estado_path)
    config = carregar_json(config_path) if config_path.exists() else {}
    if args.executar and not config:
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")

    plano = montar_plano(estado)
    selecionados = plano[: args.limit] if args.limit > 0 else plano
    print(
        f"PLANO ENTIDADE: {len(plano)} chamados vinculados no estado; "
        f"{len(selecionados)} chamados selecionados."
    )

    if not args.executar:
        alvo = texto(config.get("ti_entity_completename")) or ENTIDADE_PADRAO
        for item in selecionados[:50]:
            print(
                f"SIMULAÇÃO: {item['referencia']} -> GLPI #{item['ticket_id']} "
                f"iria para a entidade '{alvo}'."
            )
        if len(selecionados) > 50:
            print(f"... mais {len(selecionados) - 50} chamados na simulação.")
        print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
        return

    glpi = Glpi(config)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    erros_path = Path(f"atribuicao_entidade_erros_{carimbo}.csv").resolve()
    sucessos = 0
    existentes = 0
    falhas = 0
    try:
        glpi.iniciar()
        entidades = glpi.listar_todos("Entity")
        entidade_id = localizar_entidade(entidades, config)
        print(f"ENTIDADE TI: ID {entidade_id}.")

        limite_varredura = int(config.get("duplicate_scan_range", 30000))
        tickets = glpi.listar_todos("Ticket", limite=limite_varredura)
        entidade_atual_por_ticket = {
            int(ticket["id"]): int(ticket.get("entities_id") or 0)
            for ticket in tickets
            if ticket.get("id")
        }

        for item in selecionados:
            entidade_atual = entidade_atual_por_ticket.get(item["ticket_id"])
            if entidade_atual == entidade_id:
                existentes += 1
                print(
                    f"JÁ OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                    "já está na entidade de TI."
                )
                continue
            try:
                glpi.atualizar("Ticket", item["ticket_id"], {"entities_id": entidade_id})
                entidade_atual_por_ticket[item["ticket_id"]] = entidade_id
                sucessos += 1
                print(
                    f"OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                    "movido para a entidade de TI."
                )
            except Exception as erro:
                falhas += 1
                registrar_erro(erros_path, item["referencia"], item["ticket_id"], str(erro))
                print(f"ERRO: {item['referencia']} / GLPI #{item['ticket_id']} | {erro}")
            time.sleep(float(config.get("interval_seconds", 0.15)))
    finally:
        glpi.finalizar()

    estado.setdefault("sincronizacao_entidade_ti", {})["ultima_execucao"] = {
        "data": datetime.now().isoformat(timespec="seconds"),
        "chamados_selecionados": len(selecionados),
        "movidos": sucessos,
        "ja_na_entidade": existentes,
        "falhas": falhas,
    }
    salvar_json(estado_path, estado)

    print(
        f"CONCLUÍDO ENTIDADE: {sucessos} chamado(s) movido(s); "
        f"{existentes} já estavam corretos; {falhas} falhas."
    )
    if falhas:
        print(f"ERROS: {erros_path}")


if __name__ == "__main__":
    main()
