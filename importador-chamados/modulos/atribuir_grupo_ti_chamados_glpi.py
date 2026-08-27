#!/usr/bin/env python3
"""Adiciona o grupo TI como grupo atribuído aos chamados importados.

O programa usa os vínculos de ``importacao_glpi_estado.json``, não remove
atores existentes, evita vínculos duplicados e também processa chamados
solucionados ou fechados.
"""

import argparse
import csv
import time
from datetime import datetime
from pathlib import Path

from atribuir_setores_chamados_glpi import (
    Glpi,
    carregar_json,
    indice_grupos,
    normalizar,
    salvar_json,
    texto,
)


TIPO_GRUPO_ATRIBUIDO = 2


def chave_referencia(referencia):
    try:
        return int(str(referencia).rsplit("-", 1)[-1])
    except (TypeError, ValueError):
        return 10**12


def montar_plano(estado):
    plano = []
    vistos = set()
    importados = estado.get("importados", {})
    for referencia in sorted(importados, key=chave_referencia):
        controle = importados.get(referencia) or {}
        ticket_ids = controle.get("ticket_ids") or [controle.get("ticket_id")]
        for ticket_id in ticket_ids:
            if not ticket_id:
                continue
            ticket_id = int(ticket_id)
            if ticket_id in vistos:
                continue
            vistos.add(ticket_id)
            plano.append({"referencia": referencia, "ticket_id": ticket_id})
    return plano


def registrar_erro(caminho, referencia, ticket_id, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(["Referência", "Chamado GLPI", "Resultado", "Detalhe"])
        escritor.writerow([referencia, ticket_id, "ERRO", detalhe])


def localizar_grupo_ti(grupos, config):
    grupo_id_config = config.get("ti_assigned_group_id")
    if grupo_id_config:
        grupo_id_config = int(grupo_id_config)
        if any(int(grupo.get("id") or 0) == grupo_id_config for grupo in grupos):
            return grupo_id_config
        raise RuntimeError(
            f"O grupo configurado em ti_assigned_group_id={grupo_id_config} "
            "não foi encontrado no GLPI."
        )

    grupos_por_nome, ambiguos = indice_grupos(grupos)
    chave = normalizar("TI")
    if chave in ambiguos:
        raise RuntimeError(
            "Há mais de um grupo chamado TI no GLPI. Acrescente ao config.json "
            'o campo "ti_assigned_group_id" com o ID correto.'
        )
    grupo_id = grupos_por_nome.get(chave)
    if not grupo_id:
        raise RuntimeError("O grupo TI não foi encontrado no GLPI.")
    return grupo_id


def main():
    parser = argparse.ArgumentParser(
        description="Adiciona o grupo TI como grupo atribuído aos chamados importados"
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
        f"PLANO TI: {len(plano)} chamados vinculados no estado; "
        f"{len(selecionados)} chamados selecionados."
    )

    if not args.executar:
        for item in selecionados[:50]:
            print(
                f"SIMULAÇÃO: {item['referencia']} -> GLPI #{item['ticket_id']} "
                "receberá o grupo atribuído TI."
            )
        if len(selecionados) > 50:
            print(f"... mais {len(selecionados) - 50} chamados na simulação.")
        print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
        return

    glpi = Glpi(config)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    erros_path = Path(f"atribuicao_grupo_ti_erros_{carimbo}.csv").resolve()
    sucessos = 0
    existentes = 0
    falhas = 0
    try:
        glpi.iniciar()
        grupos = glpi.listar_todos("Group")
        grupo_ti_id = localizar_grupo_ti(grupos, config)
        print(f"GRUPO TI: ID {grupo_ti_id}.")

        vinculos = glpi.listar_todos("Group_Ticket")
        existentes_glpi = {
            (
                int(vinculo.get("tickets_id") or 0),
                int(vinculo.get("groups_id") or 0),
                int(vinculo.get("type") or 0),
            )
            for vinculo in vinculos
        }

        for item in selecionados:
            chave_vinculo = (
                item["ticket_id"],
                grupo_ti_id,
                TIPO_GRUPO_ATRIBUIDO,
            )
            if chave_vinculo in existentes_glpi:
                existentes += 1
                print(
                    f"JÁ OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                    "já está atribuído ao grupo TI."
                )
                continue
            try:
                glpi.criar(
                    "Group_Ticket",
                    {
                        "tickets_id": item["ticket_id"],
                        "groups_id": grupo_ti_id,
                        "type": TIPO_GRUPO_ATRIBUIDO,
                        "_disablenotif": 1,
                    },
                )
                existentes_glpi.add(chave_vinculo)
                sucessos += 1
                print(
                    f"OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                    "foi atribuído ao grupo TI."
                )
            except Exception as erro:
                falhas += 1
                registrar_erro(
                    erros_path,
                    item["referencia"],
                    item["ticket_id"],
                    str(erro),
                )
                print(
                    f"ERRO: {item['referencia']} / GLPI #{item['ticket_id']} | {erro}"
                )
            time.sleep(float(config.get("interval_seconds", 0.15)))
    finally:
        glpi.finalizar()

    estado.setdefault("sincronizacao_grupo_ti", {})["ultima_execucao"] = {
        "data": datetime.now().isoformat(timespec="seconds"),
        "chamados_selecionados": len(selecionados),
        "vinculos_criados": sucessos,
        "vinculos_ja_existentes": existentes,
        "falhas": falhas,
    }
    salvar_json(estado_path, estado)

    print(
        f"CONCLUÍDO TI: {sucessos} vínculos criados; "
        f"{existentes} já existentes; {falhas} falhas."
    )
    if falhas:
        print(f"ERROS: {erros_path}")


if __name__ == "__main__":
    main()
