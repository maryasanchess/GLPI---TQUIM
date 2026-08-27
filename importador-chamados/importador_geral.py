#!/usr/bin/env python3
"""Menu único e seguro para importação e manutenção dos chamados no GLPI."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


RAIZ = Path(__file__).resolve().parent
MODULOS = RAIZ / "modulos"
PLANILHA = RAIZ / "planilha_atualizada.xlsx"
PLANILHA_ANTERIOR = RAIZ / "planilha_anterior.xlsx"
CONFIG = RAIZ / "config.json"
ESTADO = RAIZ / "importacao_glpi_estado.json"
CATEGORIAS = RAIZ / "dados_categorias" / "chamados_sem_categoria.xlsx"
RESULTADOS = RAIZ / "resultados"
BACKUPS = RAIZ / "backups"
LOGS = RAIZ / "logs"
LOG_ATUAL: Path | None = None
MODULOS_INTERATIVOS = {
    "preparar_config_e_reconstruir_estado.py",
    "validar_usuarios_requerentes.py",
}


ACOES_MUTAVEIS = {
    "testar1",
    "importar",
    "complementar",
    "status",
    "categorias_aplicar",
    "anexos",
    "tudo",
    "corrigir_suspenso",
}


def registrar(mensagem: str) -> None:
    if LOG_ATUAL:
        LOG_ATUAL.parent.mkdir(parents=True, exist_ok=True)
        with LOG_ATUAL.open("a", encoding="utf-8") as arquivo:
            arquivo.write(mensagem.rstrip() + "\n")


def iniciar_log(acao: str) -> None:
    global LOG_ATUAL
    LOGS.mkdir(exist_ok=True)
    momento = datetime.now().strftime("%Y%m%d_%H%M%S")
    LOG_ATUAL = LOGS / f"execucao_{acao}_{momento}.log"
    registrar(f"Ação: {acao}")
    registrar(f"Início: {datetime.now().isoformat(timespec='seconds')}")


def executar(modulo: str, *argumentos: object) -> None:
    comando = [sys.executable, str(MODULOS / modulo), *(str(a) for a in argumentos)]
    print(f"\n>> Executando: {modulo}")
    registrar(f"\n>> Executando: {modulo}")
    ambiente = os.environ.copy()
    ambiente["PYTHONIOENCODING"] = "utf-8"
    if modulo in MODULOS_INTERATIVOS:
        registrar("Módulo interativo: as respostas digitadas não são gravadas no log.")
        resultado = subprocess.run(comando, cwd=RAIZ, env=ambiente)
        if resultado.returncode != 0:
            raise RuntimeError(f"O módulo {modulo} terminou com erro {resultado.returncode}.")
        return
    processo = subprocess.Popen(
        comando,
        cwd=RAIZ,
        env=ambiente,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert processo.stdout is not None
    for linha in processo.stdout:
        print(linha, end="")
        registrar(linha)
    codigo = processo.wait()
    if codigo != 0:
        raise RuntimeError(f"O módulo {modulo} terminou com erro {codigo}.")


def confirmar(acao: str) -> None:
    if acao not in ACOES_MUTAVEIS:
        return
    print("\nATENÇÃO: esta opção altera dados no GLPI.")
    resposta = input("Digite APLICAR para continuar: ").strip()
    if resposta != "APLICAR":
        raise RuntimeError("Operação cancelada. Nenhuma alteração foi solicitada.")


def backup_estado(acao: str) -> Path:
    BACKUPS.mkdir(exist_ok=True)
    momento = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUPS / f"estado_antes_{acao}_{momento}.json"
    shutil.copy2(ESTADO, destino)
    print(f"Backup automático do controle: {destino.relative_to(RAIZ)}")
    registrar(f"Backup do estado: {destino}")
    return destino


def argumentos_importador(limite: int) -> list[object]:
    return [
        PLANILHA,
        "--config", CONFIG,
        "--excel-anterior", PLANILHA_ANTERIOR,
        "--depois-de", 0,
        "--limit", limite,
    ]


def complementos(limite: int, executar_agora: bool) -> None:
    extra = ["--executar"] if executar_agora else []
    executar(
        "atribuir_setores_chamados_glpi.py", PLANILHA,
        "--config", CONFIG, "--estado", ESTADO, "--limit", limite, *extra,
    )
    executar(
        "atribuir_grupo_ti_chamados_glpi.py",
        "--config", CONFIG, "--estado", ESTADO, "--limit", limite, *extra,
    )
    executar(
        "atribuir_usuarios_requerentes_glpi.py", PLANILHA,
        "--config", CONFIG, "--estado", ESTADO, "--limit", limite, *extra,
    )


def relatorio_categoria() -> Path:
    RESULTADOS.mkdir(exist_ok=True)
    momento = datetime.now().strftime("%Y%m%d_%H%M%S")
    return RESULTADOS / f"resultado_categorias_{momento}.csv"


def relatorio_verificacao() -> Path:
    RESULTADOS.mkdir(exist_ok=True)
    momento = datetime.now().strftime("%Y%m%d_%H%M%S")
    return RESULTADOS / f"verificacao_pre_importacao_{momento}.txt"


def relatorio_auditoria() -> Path:
    RESULTADOS.mkdir(exist_ok=True)
    momento = datetime.now().strftime("%Y%m%d_%H%M%S")
    return RESULTADOS / f"auditoria_importados_{momento}.txt"


def verificar_pre_importacao(testar_api: bool = False) -> None:
    argumentos: list[object] = [
        "--config", CONFIG,
        "--estado", ESTADO,
        "--planilha", PLANILHA,
        "--relatorio", relatorio_verificacao(),
    ]
    if testar_api:
        argumentos.append("--api")
    executar("verificar_pre_importacao.py", *argumentos)


def validar_pacote() -> None:
    obrigatorios = [
        CONFIG, ESTADO, PLANILHA, PLANILHA_ANTERIOR,
        MODULOS / "importar_chamados.py",
        MODULOS / "atribuir_setores_chamados_glpi.py",
        MODULOS / "atribuir_grupo_ti_chamados_glpi.py",
        MODULOS / "atribuir_usuarios_requerentes_glpi.py",
        MODULOS / "corrigir_datas_glpi.py",
        MODULOS / "verificar_pre_importacao.py",
    ]
    ausentes = [str(p.relative_to(RAIZ)) for p in obrigatorios if not p.exists()]
    if ausentes:
        raise RuntimeError("Arquivos ausentes: " + ", ".join(ausentes))

    config = json.loads(CONFIG.read_text(encoding="utf-8-sig"))
    for chave in ("api_url", "app_token", "user_token"):
        if not str(config.get(chave, "")).strip():
            raise RuntimeError(f"Preencha o campo {chave!r} no config.json.")
    estado = json.loads(ESTADO.read_text(encoding="utf-8-sig"))
    if not isinstance(estado, dict):
        raise RuntimeError("O arquivo de estado não possui uma estrutura válida.")

    verificar_pre_importacao(False)
    print("Pacote validado: arquivos, configuração, estado e planilha estão prontos.")


def realizar(acao: str) -> None:
    iniciar_log(acao)
    if acao == "validar_pacote":
        validar_pacote()
        return
    if acao == "verificar_api":
        verificar_pre_importacao(True)
        return
    if acao in ACOES_MUTAVEIS:
        verificar_pre_importacao(True)
        confirmar(acao)
        backup_estado(acao)

    if acao == "configurar":
        executar("preparar_config_e_reconstruir_estado.py")
    elif acao == "simular":
        verificar_pre_importacao(True)
        executar("importar_chamados.py", *argumentos_importador(10))
        complementos(10, False)
        if CATEGORIAS.exists():
            executar(
                "atribuir_categorias_glpi.py", CATEGORIAS,
                "--config", CONFIG, "--limit", 10, "--relatorio", relatorio_categoria(),
            )
        print("\nSimulação concluída. Nenhum dado foi alterado no GLPI.")
    elif acao == "testar1":
        executar("importar_chamados.py", *argumentos_importador(1), "--executar")
    elif acao == "importar":
        validar_pacote()
        executar("importar_chamados.py", *argumentos_importador(0), "--executar", "--reparar-preenchidos")
        complementos(0, True)
        print("\nImportação e complementos concluídos.")
    elif acao == "tudo":
        print(
            "\n>> Ação única: validar (com API), importar só o que é novo, completar "
            "setor/TI/requerente, aplicar categorias e auditar o resultado no GLPI."
        )
        validar_pacote()
        referencias_antes = set(json.loads(ESTADO.read_text(encoding="utf-8-sig")).get("importados", {}))
        executar("importar_chamados.py", *argumentos_importador(0), "--executar", "--reparar-preenchidos")
        complementos(0, True)
        if CATEGORIAS.exists():
            executar(
                "atribuir_categorias_glpi.py", CATEGORIAS,
                "--config", CONFIG, "--limit", 0, "--relatorio", relatorio_categoria(),
                "--executar",
            )
        referencias_depois = set(json.loads(ESTADO.read_text(encoding="utf-8-sig")).get("importados", {}))
        novas = len(referencias_depois - referencias_antes)
        if novas == 0:
            print("\nTUDO CONCLUÍDO: nenhum chamado novo para importar (tudo já estava no GLPI).")
            return
        try:
            executar(
                "auditar_importados.py", PLANILHA,
                "--config", CONFIG, "--estado", ESTADO,
                "--somente-recentes", novas,
                "--relatorio", relatorio_auditoria(),
            )
            print(f"\nTUDO CONCLUÍDO: {novas} chamado(s) novo(s) importado(s), sem divergência na auditoria.")
        except RuntimeError:
            print(
                f"\nIMPORTAÇÃO CONCLUÍDA ({novas} chamado(s) novo(s)), mas a auditoria encontrou "
                "divergências - veja o relatório em resultados/auditoria_importados_*.txt antes de "
                "considerar estes chamados 100% corretos."
            )
    elif acao == "complementar":
        complementos(0, True)
    elif acao == "status":
        executar(
            "importar_chamados.py", *argumentos_importador(0),
            "--sincronizar-status-antigos", "--executar",
        )
    elif acao in {"categorias_simular", "categorias_aplicar"}:
        argumentos: list[object] = [
            CATEGORIAS, "--config", CONFIG, "--limit", 0,
            "--relatorio", relatorio_categoria(),
        ]
        if acao == "categorias_aplicar":
            argumentos.append("--executar")
        executar("atribuir_categorias_glpi.py", *argumentos)
    elif acao == "datas":
        executar("corrigir_datas_glpi.py", "--config", CONFIG, "--planilha", PLANILHA)
        print("\nSomente o plano e os SQLs foram gerados; nenhuma data foi alterada.")
    elif acao == "anexos":
        executar(
            "importar_chamados.py", *argumentos_importador(0),
            "--reprocessar-anexos-pendentes", "--executar",
        )
    elif acao == "validar_requerentes":
        executar(
            "validar_usuarios_requerentes.py", PLANILHA,
            "--config", CONFIG, "--estado", ESTADO,
        )
    elif acao == "diagnosticar_dashboard":
        executar("diagnosticar_dashboard_grupos_requerentes.py", "--config", CONFIG)
    elif acao == "auditar_importados":
        executar(
            "auditar_importados.py", PLANILHA,
            "--config", CONFIG, "--estado", ESTADO,
            "--relatorio", relatorio_auditoria(),
        )
    elif acao == "corrigir_suspenso":
        executar(
            "corrigir_status_suspenso_glpi.py", PLANILHA,
            "--config", CONFIG, "--estado", ESTADO, "--executar",
        )
    else:
        raise ValueError(f"Ação desconhecida: {acao}")


MENU = """
IMPORTADOR GERAL DO GLPI
========================
16 - FAZER TUDO: validar, importar o que é novo, completar e auditar
     (uma ação só - recomendado no dia a dia)
------------------------------------------------------------------
1  - Configurar API e reconstruir o controle EXCEL-N
2  - Simular fluxo completo (não altera o GLPI)
3  - Importar 1 chamado de teste
4  - Importar todos os novos chamados e aplicar complementos
5  - Completar chamados existentes (setor, TI e requerentes)
6  - Sincronizar status antigos
7  - Simular categorias
8  - Aplicar categorias
9  - Gerar plano/SQL de correção de datas (não aplica)
10 - Reprocessar anexos pendentes
11 - Validar usuários requerentes
12 - Diagnosticar grupos do dashboard
13 - Validar arquivos do pacote (offline)
14 - Testar planilha, controle, API e permissões de leitura
15 - Auditar chamados já importados (requerente, data, status, técnico, grupo)
17 - Corrigir chamados 'Suspenso' já importados para Fechado
0  - Sair
"""

OPCOES = {
    "1": "configurar", "2": "simular", "3": "testar1", "4": "importar",
    "5": "complementar", "6": "status", "7": "categorias_simular",
    "8": "categorias_aplicar", "9": "datas", "10": "anexos",
    "11": "validar_requerentes", "12": "diagnosticar_dashboard",
    "13": "validar_pacote", "14": "verificar_api",
    "15": "auditar_importados", "16": "tudo", "17": "corrigir_suspenso",
}


def main() -> int:
    parser = argparse.ArgumentParser(description="Importador geral simplificado do GLPI")
    parser.add_argument("--acao", choices=sorted(set(OPCOES.values())))
    parser.add_argument("--listar", action="store_true", help="Mostra as opções e encerra")
    args = parser.parse_args()
    if args.listar:
        print(MENU)
        return 0
    if args.acao:
        realizar(args.acao)
        return 0

    while True:
        print(MENU)
        opcao = input("Escolha uma opção: ").strip()
        if opcao == "0":
            return 0
        acao = OPCOES.get(opcao)
        if not acao:
            print("Opção inválida.")
            continue
        try:
            realizar(acao)
        except (RuntimeError, ValueError, FileNotFoundError) as erro:
            print(f"\nERRO: {erro}")
        input("\nPressione ENTER para voltar ao menu...")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nOperação interrompida pelo usuário.")
        raise SystemExit(130)
