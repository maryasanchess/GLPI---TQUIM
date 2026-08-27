#!/usr/bin/env python3
"""Diagnostica nomes e contagens dos grupos requerentes usados no dashboard.

O programa faz apenas consultas à API. Nenhum chamado, grupo, usuário ou
configuração do GLPI é alterado.
"""

import argparse
import csv
from collections import Counter
from datetime import datetime
from pathlib import Path

from atribuir_setores_chamados_glpi import Glpi, carregar_json, texto


TIPO_GRUPO_REQUERENTE = 1


def nome_grupo(grupo):
    return texto(grupo.get("name")) or texto(grupo.get("completename"))


def gravar_csv(caminho, linhas):
    with open(caminho, "w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        escritor.writerow(
            [
                "Posição",
                "ID do grupo",
                "Nome",
                "Nome completo",
                "Vínculos como grupo requerente",
                "Diagnóstico",
            ]
        )
        for linha in linhas:
            escritor.writerow(
                [
                    linha["posicao"],
                    linha["grupo_id"],
                    linha["nome"],
                    linha["nome_completo"],
                    linha["quantidade"],
                    linha["diagnostico"],
                ]
            )


def montar_resultado(grupos, vinculos):
    grupos_por_id = {
        int(grupo.get("id")): grupo
        for grupo in grupos
        if grupo.get("id")
    }
    contagem = Counter(
        int(vinculo.get("groups_id") or 0)
        for vinculo in vinculos
        if int(vinculo.get("type") or 0) == TIPO_GRUPO_REQUERENTE
        and int(vinculo.get("groups_id") or 0) > 0
    )

    linhas = []
    for posicao, (grupo_id, quantidade) in enumerate(
        sorted(contagem.items(), key=lambda item: (-item[1], item[0])),
        start=1,
    ):
        grupo = grupos_por_id.get(grupo_id)
        if not grupo:
            nome = ""
            nome_completo = ""
            diagnostico = "ID vinculado, mas grupo não retornou pela API"
        else:
            nome = texto(grupo.get("name"))
            nome_completo = texto(grupo.get("completename"))
            if nome or nome_completo:
                diagnostico = "OK - nome disponível"
            else:
                diagnostico = "Grupo localizado, mas sem nome"
        linhas.append(
            {
                "posicao": posicao,
                "grupo_id": grupo_id,
                "nome": nome,
                "nome_completo": nome_completo,
                "quantidade": quantidade,
                "diagnostico": diagnostico,
            }
        )
    return linhas


def main():
    parser = argparse.ArgumentParser(
        description="Diagnostica a legenda do dashboard de grupos requerentes"
    )
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    if not config_path.exists():
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")
    config = carregar_json(config_path)

    glpi = Glpi(config)
    try:
        print("Conectando ao GLPI somente para leitura...")
        glpi.iniciar()
        grupos = glpi.listar_todos("Group")
        vinculos = glpi.listar_todos("Group_Ticket")
    finally:
        glpi.finalizar()

    linhas = montar_resultado(grupos, vinculos)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    resultado_path = Path(
        f"diagnostico_dashboard_grupos_requerentes_{carimbo}.csv"
    ).resolve()
    gravar_csv(resultado_path, linhas)

    total_vinculos = sum(linha["quantidade"] for linha in linhas)
    com_nome = sum(
        1 for linha in linhas if linha["diagnostico"] == "OK - nome disponível"
    )
    sem_nome = len(linhas) - com_nome

    print()
    print(f"GRUPOS CONSULTADOS NA API: {len(grupos)}")
    print(f"GRUPOS REQUERENTES COM VÍNCULOS: {len(linhas)}")
    print(f"TOTAL DE VÍNCULOS REQUERENTES: {total_vinculos}")
    print(f"GRUPOS COM NOME DISPONÍVEL: {com_nome}")
    print(f"GRUPOS COM PROBLEMA DE NOME/ACESSO: {sem_nome}")
    print()
    print("TOP GRUPOS REQUERENTES:")
    for linha in linhas[:20]:
        nome = linha["nome"] or linha["nome_completo"] or "SEM NOME"
        print(
            f"- {linha['quantidade']:>5} chamados | grupo ID "
            f"{linha['grupo_id']}: {nome} | {linha['diagnostico']}"
        )

    print()
    if linhas and sem_nome == 0:
        print(
            "RESULTADO: os IDs, nomes e contagens estão corretos na API. "
            "A falha está na exibição do dashboard, não na importação."
        )
    elif linhas:
        print(
            "RESULTADO: existem grupos cujo nome não foi retornado pela API. "
            "Verifique o CSV para identificar os IDs e as permissões."
        )
    else:
        print(
            "RESULTADO: nenhum vínculo de grupo requerente foi localizado. "
            "Verifique a entidade e o perfil usado pelo token."
        )
    print(f"ARQUIVO GERADO: {resultado_path}")


if __name__ == "__main__":
    main()
