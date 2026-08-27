#!/usr/bin/env python3
"""Atribui o grupo requerente dos chamados importados conforme o Departamento do Excel.

O vínculo entre a linha do Excel e o chamado do GLPI vem de
``importacao_glpi_estado.json``. O programa não cria chamados, usuários ou
grupos e não altera status, solução, datas, técnico, categoria ou anexos.
"""

import argparse
import csv
import json
import re
import time
import unicodedata
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd


def texto(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


# A coluna de data de abertura já apareceu com nomes diferentes em
# exportações diferentes do Microsoft Forms; qualquer um destes é aceito
# e tratado internamente como "Carimbo de data/hora".
ALIASES_COLUNA_ABERTURA = {"carimbo de data/hora", "data de abertura"}


def normalizar_colunas(df):
    renomear = {}
    for coluna in df.columns:
        if str(coluna).strip().casefold() in ALIASES_COLUNA_ABERTURA:
            renomear[coluna] = "Carimbo de data/hora"
    return df.rename(columns=renomear) if renomear else df


def normalizar(valor):
    valor = unicodedata.normalize("NFKD", texto(valor))
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    valor = valor.casefold().replace("–", "-").replace("—", "-")
    return re.sub(r"[^a-z0-9]+", " ", valor).strip()


def linha_valida(linha):
    return bool(
        texto(linha.get("Carimbo de data/hora"))
        and texto(linha.get("Nome do Solicitante"))
        and (
            texto(linha.get("Descrição do Chamado"))
            or texto(linha.get("Motivo da Solicitação"))
        )
    )


def carregar_json(caminho):
    with open(caminho, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def salvar_json(caminho, dados):
    temporario = caminho.with_suffix(".tmp")
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)
    temporario.replace(caminho)


def registrar_csv(caminho, referencia, ticket_id, departamento, resultado, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(
                ["Referência", "Chamado GLPI", "Departamento Excel", "Resultado", "Detalhe"]
            )
        escritor.writerow([referencia, ticket_id, departamento, resultado, detalhe])


class Glpi:
    def __init__(self, config):
        obrigatorios = ["api_url", "app_token", "user_token"]
        ausentes = [campo for campo in obrigatorios if not texto(config.get(campo))]
        if ausentes:
            raise ValueError("Preencha no config.json: " + ", ".join(ausentes))
        self.config = config
        self.api_url = texto(config["api_url"]).rstrip("/")
        self.headers = {
            "Content-Type": "application/json",
            "App-Token": config["app_token"],
            "Authorization": f"user_token {config['user_token']}",
        }
        self.session_token = None

    def requisicao(self, metodo, endpoint, dados=None, timeout=60):
        corpo = json.dumps(dados).encode("utf-8") if dados is not None else None
        requisicao = urllib.request.Request(
            f"{self.api_url}/{endpoint.lstrip('/')}",
            data=corpo,
            headers=self.headers,
            method=metodo,
        )
        try:
            with urllib.request.urlopen(requisicao, timeout=timeout) as resposta:
                conteudo = resposta.read().decode("utf-8")
        except urllib.error.HTTPError as erro:
            resposta_erro = erro.read().decode("utf-8", errors="replace")[:1500]
            raise RuntimeError(
                f"Falha na API em {metodo} {endpoint}: HTTP {erro.code} - {resposta_erro}"
            ) from erro
        return json.loads(conteudo) if conteudo else None

    def iniciar(self):
        resposta = self.requisicao("GET", "initSession", timeout=30)
        self.session_token = resposta["session_token"]
        self.headers["Session-Token"] = self.session_token
        perfil_id = self.config.get("profile_id", 0)
        if perfil_id:
            self.requisicao(
                "POST",
                "changeActiveProfile",
                {"profiles_id": int(perfil_id)},
            )

    def finalizar(self):
        if self.session_token:
            try:
                self.requisicao("GET", "killSession", timeout=15)
            except Exception:
                pass

    def criar(self, tipo, dados):
        retorno = self.requisicao("POST", tipo, {"input": dados})
        if isinstance(retorno, list):
            retorno = retorno[0]
        return int(retorno["id"])

    def atualizar(self, tipo, item_id, dados):
        self.requisicao(
            "PUT",
            f"{tipo}/{item_id}",
            {"input": {"id": item_id, "_disablenotifications": 1, "_disablenotif": 1, **dados}},
        )

    def listar_todos(self, tipo, limite=30000, tamanho_pagina=500):
        itens = []
        for inicio in range(0, limite, tamanho_pagina):
            fim = inicio + tamanho_pagina - 1
            retorno = self.requisicao("GET", f"{tipo}?range={inicio}-{fim}")
            pagina = retorno if isinstance(retorno, list) else []
            itens.extend(pagina)
            if len(pagina) < tamanho_pagina:
                break
        return itens


ALIAS_DIRETO = {
    "rh": ["DP"],
    "rh dp": ["DP"],
    "rh dp1": ["DP"],
    "dp rh": ["DP"],
    "dp": ["DP"],
    "departamento pessoal": ["DP"],
    "estagiaria rh": ["DP"],
    "estacao de limpeza": ["EL"],
    "e l estacao de limpeza": ["EL"],
    "essacao de limpeza": ["EL"],
    "armazem glp": ["Armazém"],
    "armazem sanca": ["Armazém"],
    "armazem adm": ["Armazém"],
    "adm armazem": ["Armazém"],
    "armazem logistica": ["Armazém"],
    "logistica sanca": ["Armazém"],
    "logistica": ["Armazém"],
    "filial sao jose dos pinhais": ["Filial - SJP"],
    "filial sjp": ["Filial - SJP"],
    "filialsjp": ["Filial - SJP"],
    "sjpinhais": ["Filial - SJP"],
    "adm filial sjp": ["Filial - SJP"],
    "filial bahia": ["Filial - BA"],
    "filial camacari": ["Filial - BA"],
    "seguranca do trabalho": ["SST"],
    "planejamento operacoes": ["Operacional"],
    "planejamento operacao": ["Operacional"],
    "planejamento operacos": ["Operacional"],
    "departamento de frota": ["Frota"],
    "programacao": ["Expedição"],
}


PADROES = [
    ("almoxarifado", "Almoxarifado"),
    ("cfm", "CFM"),
    ("diretoria", "Diretoria"),
    ("faturamento", "Faturamento"),
    ("manutencao", "Manutenção"),
    ("frota", "Frota"),
    ("financeiro", "Financeiro"),
    ("compras", "Compras"),
    ("rastreamento", "Rastreamento"),
    ("expedicao", "Expedição"),
    ("comercial", "Comercial"),
    ("programacao", "Expedição"),
    ("qualidade", "Qualidade"),
    ("operacional", "Operacional"),
    ("planejamento operac", "Operacional"),
    ("armazem", "Armazém"),
    ("tecnologia da informacao", "TI"),
    (" ti ", "TI"),
]


def nomes_candidatos_departamento(valor, mapa_config):
    """Transforma as variações do Excel em nomes de grupos do GLPI."""
    original = texto(valor)
    chave = normalizar(original)
    if not chave or chave in {".", "ignorar"}:
        return []

    for origem, destino in mapa_config.items():
        if normalizar(origem) == chave:
            if isinstance(destino, list):
                return [texto(item) for item in destino if texto(item)]
            return [texto(destino)] if texto(destino) else []

    if chave in ALIAS_DIRETO:
        return ALIAS_DIRETO[chave]

    encontrados = []
    texto_cercado = f" {chave} "
    for trecho, grupo in PADROES:
        corresponde = trecho in chave if trecho != " ti " else trecho in texto_cercado
        if corresponde and grupo not in encontrados:
            encontrados.append(grupo)

    if " estacao de limpeza" in texto_cercado or chave == "el":
        encontrados.append("EL")
    if any(marca in chave for marca in (" rh", "rh ", " dp", "dp ")):
        encontrados.append("DP")
    if "sst" in chave:
        encontrados.append("SST")
    if "sustentabilidade" in chave:
        encontrados.append("Sustentabilidade")
    if "borracharia" in chave:
        encontrados.append("Borracharia")
    if chave == "administrativo":
        encontrados.append("Administrativo")
    if "filial" in chave and "sjp" in chave:
        encontrados.append("Filial - SJP")
    if "filial" in chave and any(x in chave for x in ("bahia", "camacari")):
        encontrados.append("Filial - BA")

    if not encontrados:
        encontrados.append(original)
    return list(dict.fromkeys(encontrados))


def indice_grupos(grupos):
    indice = {}
    ambiguos = set()
    for grupo in grupos:
        if not grupo.get("id"):
            continue
        nomes = [texto(grupo.get("name")), texto(grupo.get("completename"))]
        completo = texto(grupo.get("completename"))
        if completo:
            nomes.append(re.split(r"\s*>\s*", completo)[-1])
        for nome in nomes:
            chave = normalizar(nome)
            if not chave:
                continue
            grupo_id = int(grupo["id"])
            if chave in indice and indice[chave] != grupo_id:
                ambiguos.add(chave)
            else:
                indice[chave] = grupo_id
    for chave in ambiguos:
        indice.pop(chave, None)
    return indice, ambiguos


def montar_plano(df, estado, mapa_config):
    plano = []
    sem_controle = 0
    sem_departamento = 0
    for indice, linha in df.iterrows():
        if not linha_valida(linha):
            continue
        referencia = f"EXCEL-{indice + 2}"
        controle = estado.get("importados", {}).get(referencia)
        if not controle or not controle.get("ticket_id"):
            sem_controle += 1
            continue
        ticket_ids = controle.get("ticket_ids") or [controle.get("ticket_id")]
        ticket_ids = list(
            dict.fromkeys(
                int(ticket_id)
                for ticket_id in ticket_ids
                if ticket_id
            )
        )
        if not ticket_ids:
            sem_controle += 1
            continue
        departamento = texto(linha.get("Departamento"))
        candidatos = nomes_candidatos_departamento(departamento, mapa_config)
        if not candidatos:
            sem_departamento += 1
            continue
        for ticket_id in ticket_ids:
            plano.append(
                {
                    "referencia": referencia,
                    "ticket_id": ticket_id,
                    "departamento": departamento,
                    "candidatos": candidatos,
                }
            )
    return plano, sem_controle, sem_departamento


def main():
    parser = argparse.ArgumentParser(
        description="Atribui setores como grupos requerentes dos chamados importados"
    )
    parser.add_argument("excel", help="Última planilha .xlsx")
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

    excel = Path(args.excel).resolve()
    config_path = Path(args.config).resolve()
    estado_path = Path(args.estado).resolve()
    if not excel.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel}")
    if not estado_path.exists():
        raise FileNotFoundError(f"Controle da importação não encontrado: {estado_path}")

    df = pd.read_excel(excel)
    df = normalizar_colunas(df)
    if "Departamento" not in df.columns:
        raise ValueError("A planilha não possui a coluna Departamento.")
    estado = carregar_json(estado_path)
    config = carregar_json(config_path) if config_path.exists() else {}
    if args.executar and not config:
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")
    mapa_config = config.get("department_group_map", {})
    plano, sem_controle, sem_departamento = montar_plano(df, estado, mapa_config)
    selecionados = plano[: args.limit] if args.limit > 0 else plano

    departamentos = sorted({item["departamento"] for item in plano}, key=normalizar)
    print(
        f"PLANO: {len(plano)} chamados com setor; "
        f"{sem_controle} linhas sem vínculo no arquivo de estado; "
        f"{sem_departamento} linhas sem departamento; "
        f"{len(selecionados)} chamados selecionados."
    )
    print(f"DEPARTAMENTOS: {len(departamentos)} variações encontradas.")

    if not args.executar:
        for item in selecionados[:50]:
            print(
                f"SIMULAÇÃO: {item['referencia']} -> GLPI #{item['ticket_id']} | "
                f"{item['departamento']} -> {', '.join(item['candidatos'])}"
            )
        if len(selecionados) > 50:
            print(f"... mais {len(selecionados) - 50} chamados na simulação.")
        print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
        return

    glpi = Glpi(config)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    pendencias_path = Path(f"atribuicao_setores_pendencias_{carimbo}.csv").resolve()
    sucessos = 0
    existentes = 0
    falhas = 0
    try:
        glpi.iniciar()
        grupos = glpi.listar_todos("Group")
        grupos_por_nome, ambiguos = indice_grupos(grupos)
        vinculos = glpi.listar_todos("Group_Ticket")
        existentes_glpi = {
            (
                int(v.get("tickets_id") or 0),
                int(v.get("groups_id") or 0),
                int(v.get("type") or 0),
            )
            for v in vinculos
        }

        for item in selecionados:
            resolvidos = []
            nao_encontrados = []
            for nome in item["candidatos"]:
                chave = normalizar(nome)
                grupo_id = grupos_por_nome.get(chave)
                if grupo_id:
                    resolvidos.append((grupo_id, nome))
                else:
                    motivo = "nome ambíguo no GLPI" if chave in ambiguos else "grupo não encontrado"
                    nao_encontrados.append(f"{nome} ({motivo})")

            if not resolvidos:
                detalhe = ", ".join(nao_encontrados) or "departamento não identificado"
                registrar_csv(
                    pendencias_path,
                    item["referencia"],
                    item["ticket_id"],
                    item["departamento"],
                    "PENDENTE",
                    detalhe,
                )
                print(
                    f"PENDENTE: {item['referencia']} / GLPI #{item['ticket_id']} | "
                    f"{item['departamento']} -> {detalhe}"
                )
                falhas += 1
                continue

            for grupo_id, nome_grupo in dict(resolvidos).items():
                chave_vinculo = (item["ticket_id"], grupo_id, 1)
                if chave_vinculo in existentes_glpi:
                    existentes += 1
                    print(
                        f"JÁ OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                        f"já possui o setor {nome_grupo}."
                    )
                    continue
                try:
                    glpi.criar(
                        "Group_Ticket",
                        {
                            "tickets_id": item["ticket_id"],
                            "groups_id": grupo_id,
                            "type": 1,
                            "_disablenotif": 1,
                        },
                    )
                    existentes_glpi.add(chave_vinculo)
                    sucessos += 1
                    print(
                        f"OK: {item['referencia']} / GLPI #{item['ticket_id']} "
                        f"recebeu o setor {nome_grupo}."
                    )
                except Exception as erro:
                    falhas += 1
                    registrar_csv(
                        pendencias_path,
                        item["referencia"],
                        item["ticket_id"],
                        item["departamento"],
                        "ERRO API",
                        f"{nome_grupo}: {erro}",
                    )
                    print(
                        f"ERRO: {item['referencia']} / GLPI #{item['ticket_id']} | "
                        f"{nome_grupo}: {erro}"
                    )
            if nao_encontrados:
                registrar_csv(
                    pendencias_path,
                    item["referencia"],
                    item["ticket_id"],
                    item["departamento"],
                    "PARCIAL",
                    ", ".join(nao_encontrados),
                )
            time.sleep(float(config.get("interval_seconds", 0.15)))
    finally:
        glpi.finalizar()

    estado.setdefault("sincronizacao_setores", {})["ultima_execucao"] = {
        "data": datetime.now().isoformat(timespec="seconds"),
        "planilha": excel.name,
        "chamados_selecionados": len(selecionados),
        "vinculos_criados": sucessos,
        "vinculos_ja_existentes": existentes,
        "falhas": falhas,
    }
    salvar_json(estado_path, estado)
    print(
        f"CONCLUÍDO: {sucessos} vínculos criados; {existentes} já existentes; "
        f"{falhas} falhas."
    )
    if pendencias_path.exists():
        print(f"PENDÊNCIAS: {pendencias_path.name}")


if __name__ == "__main__":
    main()
