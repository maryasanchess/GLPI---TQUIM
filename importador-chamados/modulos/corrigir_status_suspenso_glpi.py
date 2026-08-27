#!/usr/bin/env python3
"""Corrige no GLPI os chamados já importados como "Suspenso".

Até 17/08/2026, o importador mapeava "Suspenso" para o status 4 (Pendente)
no GLPI. A regra mudou: "Suspenso" agora deve ficar como Fechado (6), igual
a "Finalizado". Este módulo corrige apenas os chamados que já foram
importados com o mapeamento antigo e continuam com o status errado; não
reimporta, não cria chamados e não duplica solução já existente.
"""

import argparse
import csv
import html
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

from importar_chamados import (  # noqa: E402
    Glpi,
    carregar_config,
    carregar_estado,
    chave,
    data_glpi,
    normalizar_colunas,
    salvar_estado,
    texto,
)

STATUS_FECHADO = 6
STATUS_PENDENTE = 4


def montar_plano(df, estado):
    plano = []
    for indice, linha in df.iterrows():
        if chave(linha.get("Status")) != "suspenso":
            continue
        referencia = f"EXCEL-{indice + 2}"
        controle = estado.get("importados", {}).get(referencia)
        if not controle or not controle.get("ticket_id"):
            continue
        plano.append(
            {
                "referencia": referencia,
                "ticket_id": int(controle["ticket_id"]),
                "solucao_excel": texto(linha.get("Desdobramento")),
                "data_final": data_glpi(linha.get("Data da Finalização")),
            }
        )
    return plano


def registrar_csv(caminho, item, resultado, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(["Referência", "Chamado GLPI", "Resultado", "Detalhe"])
        escritor.writerow([item["referencia"], item["ticket_id"], resultado, detalhe])


def main():
    parser = argparse.ArgumentParser(
        description="Corrige o status dos chamados 'Suspenso' já importados para Fechado"
    )
    parser.add_argument("excel", help="Última planilha .xlsx")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--estado", default="importacao_glpi_estado.json")
    parser.add_argument(
        "--executar",
        action="store_true",
        help="Aplica no GLPI; sem esta opção, apenas simula",
    )
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    config_path = Path(args.config).resolve()
    estado_path = Path(args.estado).resolve()
    if not excel_path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")

    df = pd.read_excel(excel_path)
    df = normalizar_colunas(df)
    if "Status" not in df.columns:
        raise ValueError("A planilha não possui a coluna Status.")

    estado = carregar_estado(estado_path)
    config = carregar_config(config_path) if args.executar else {}

    plano = montar_plano(df, estado)
    print(f"PLANO: {len(plano)} chamado(s) 'Suspenso' já importado(s) para revisar.")

    if not args.executar:
        for item in plano[:50]:
            print(f"SIMULAÇÃO: {item['referencia']} -> GLPI #{item['ticket_id']}")
        if len(plano) > 50:
            print(f"... mais {len(plano) - 50} chamados na simulação.")
        print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
        return

    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    pendencias_path = Path(f"correcao_suspenso_pendencias_{carimbo}.csv").resolve()
    ja_fechados = 0
    corrigidos = 0
    falhas = 0

    glpi = Glpi(config)
    try:
        glpi.iniciar()
        for item in plano:
            ticket_id = item["ticket_id"]
            try:
                ticket_atual = glpi.requisicao("GET", f"Ticket/{ticket_id}")
            except Exception as erro:
                falhas += 1
                registrar_csv(pendencias_path, item, "ERRO API", f"Falha ao consultar: {erro}")
                print(f"ERRO: {item['referencia']} / GLPI #{ticket_id} | {erro}")
                continue

            status_atual = int(ticket_atual.get("status") or 0)
            if status_atual == STATUS_FECHADO:
                ja_fechados += 1
                print(f"JÁ OK: {item['referencia']} / GLPI #{ticket_id} já está Fechado.")
                continue

            try:
                solucoes = glpi.requisicao("GET", f"Ticket/{ticket_id}/ITILSolution?range=0-99")
                solucoes = solucoes if isinstance(solucoes, list) else []
            except Exception:
                solucoes = []

            if not solucoes:
                solucao = item["solucao_excel"] or (
                    "Chamado concluído no controle antigo; a solução não foi registrada no Excel."
                )
                glpi.criar(
                    "ITILSolution",
                    {
                        "itemtype": "Ticket",
                        "items_id": ticket_id,
                        "content": html.escape(solucao).replace("\n", "<br>"),
                    },
                )

            if item["data_final"]:
                glpi.atualizar("Ticket", ticket_id, {"solvedate": item["data_final"]})
            glpi.atualizar("Ticket", ticket_id, {"status": STATUS_FECHADO})

            estado["importados"][item["referencia"]]["status_sincronizado_excel"] = STATUS_FECHADO
            estado["importados"][item["referencia"]]["status_corrigido_em"] = datetime.now().isoformat(
                timespec="seconds"
            )
            salvar_estado(estado_path, estado)
            corrigidos += 1
            print(f"OK: {item['referencia']} / GLPI #{ticket_id} corrigido para Fechado (era {status_atual}).")
    finally:
        glpi.finalizar()

    print(
        f"CONCLUÍDO: {corrigidos} corrigido(s); {ja_fechados} já estavam Fechado(s); "
        f"{falhas} falha(s)."
    )
    if falhas:
        print(f"PENDÊNCIAS: {pendencias_path}")


if __name__ == "__main__":
    main()
