#!/usr/bin/env python3
"""Move o vínculo perfil+entidade dos usuários de uma entidade para outra.

Uso típico deste projeto: todos os usuários hoje têm um vínculo apontando
para a entidade raiz (ex.: "TQUIM"). Este script troca a entidade desse
vínculo existente para uma entidade filha (ex.: "TQUIM > TI"), mantendo o
mesmo perfil e a mesma opção de recursividade de cada usuário - só o campo
entities_id do vínculo é alterado.

Não cria usuário, não cria perfil, não mexe em chamados. Só atualiza os
vínculos (Profile_User) que hoje apontam para a entidade de origem.

Só usa a biblioteca padrão do Python (sem pandas/openpyxl), não precisa de
ambiente virtual nem de instalação de dependências.
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


def texto(valor):
    if valor is None:
        return ""
    return str(valor).strip()


def normalizar(valor):
    valor = unicodedata.normalize("NFKD", texto(valor))
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    valor = valor.casefold().replace("–", "-").replace("—", "-")
    return re.sub(r"\s*>\s*", ">", valor).strip()


def carregar_json(caminho):
    with open(caminho, "r", encoding="utf-8-sig") as arquivo:
        return json.load(arquivo)


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
            self.requisicao("POST", "changeActiveProfile", {"profiles_id": int(perfil_id)})
        # A sessão abre com a "entidade ativa" padrão do usuário, que pode
        # estar restrita a uma entidade filha. Força a entidade raiz (ID 0)
        # com acesso recursivo para enxergar/alterar vínculos em todas as
        # entidades. Se a conta não tiver esse acesso, ignora e deixa a
        # resolução de entidades reportar o erro claro mais adiante.
        try:
            self.requisicao("POST", "changeActiveEntity", {"entities_id": 0, "is_recursive": True})
        except Exception:
            pass

    def finalizar(self):
        if self.session_token:
            try:
                self.requisicao("GET", "killSession", timeout=15)
            except Exception:
                pass

    def atualizar(self, tipo, item_id, dados):
        self.requisicao(
            "PUT",
            f"{tipo}/{item_id}",
            {"input": {"id": item_id, **dados}},
        )

    def listar_todos(self, tipo, limite=5000, tamanho_pagina=500):
        itens = []
        for inicio in range(0, limite, tamanho_pagina):
            fim = inicio + tamanho_pagina - 1
            retorno = self.requisicao("GET", f"{tipo}?range={inicio}-{fim}")
            pagina = retorno if isinstance(retorno, list) else []
            itens.extend(pagina)
            if len(pagina) < tamanho_pagina:
                break
        return itens


def indice_entidades(entidades):
    # Cuidado: a entidade raiz normalmente tem ID 0, que é "falsy" em Python -
    # "if not entidade.get('id')" descartaria ela por engano.
    indice = {}
    for entidade in entidades:
        if entidade.get("id") is None:
            continue
        completo = texto(entidade.get("completename")) or texto(entidade.get("name"))
        indice[normalizar(completo)] = int(entidade["id"])
    return indice


def resolver_entidade(indice, nome_completo, rotulo):
    chave = normalizar(nome_completo)
    if chave not in indice:
        raise RuntimeError(
            f"Entidade {rotulo} '{nome_completo}' não encontrada no GLPI. "
            "Confira o nome completo (ex.: 'TQUIM > TI') no config.json."
        )
    return indice[chave]


def montar_plano(vinculos, usuarios_por_id, entidade_origem_id, excluir_logins):
    excluidos = {normalizar(login) for login in excluir_logins}
    plano = []
    ignorados = []
    for vinculo in vinculos:
        if int(vinculo.get("entities_id") if vinculo.get("entities_id") is not None else -1) != entidade_origem_id:
            continue
        if not vinculo.get("id") or not vinculo.get("users_id"):
            continue
        users_id = int(vinculo["users_id"])
        usuario = usuarios_por_id.get(users_id, {})
        login = texto(usuario.get("name")) or f"ID {users_id}"
        if normalizar(login) in excluidos:
            ignorados.append(login)
            continue
        plano.append(
            {
                "vinculo_id": int(vinculo["id"]),
                "users_id": users_id,
                "login": login,
                "profiles_id": vinculo.get("profiles_id"),
                "is_dynamic": bool(vinculo.get("is_dynamic")),
            }
        )
    return plano, ignorados


def registrar_erro(caminho, login, vinculo_id, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(["Login", "Vínculo (Profile_User)", "Resultado", "Detalhe"])
        escritor.writerow([login, vinculo_id, "ERRO", detalhe])


def main():
    parser = argparse.ArgumentParser(
        description="Move o vínculo perfil+entidade dos usuários para outra entidade"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--limit",
        type=int,
        default=10,
        help="Máximo de usuários; use 0 para todos (padrão: 10)",
    )
    parser.add_argument(
        "--executar",
        action="store_true",
        help="Aplica no GLPI; sem esta opção, apenas simula",
    )
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    if not config_path.exists():
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")
    config = carregar_json(config_path)

    origem_nome = texto(config.get("origem_entity_completename")) or "TQUIM"
    destino_nome = texto(config.get("destino_entity_completename")) or "TQUIM > TI"
    excluir_logins = config.get("excluir_logins") or []

    glpi = Glpi(config)
    try:
        glpi.iniciar()
        entidades = glpi.listar_todos("Entity")
        indice = indice_entidades(entidades)
        origem_id = resolver_entidade(indice, origem_nome, "de origem")
        destino_id = resolver_entidade(indice, destino_nome, "de destino")
        print(f"ENTIDADES: origem '{origem_nome}' = ID {origem_id}; destino '{destino_nome}' = ID {destino_id}.")

        limite_varredura = int(config.get("scan_range", 5000))
        usuarios = glpi.listar_todos("User", limite=limite_varredura)
        usuarios_por_id = {int(u["id"]): u for u in usuarios if u.get("id")}
        vinculos = glpi.listar_todos("Profile_User", limite=limite_varredura)

        plano, ignorados = montar_plano(vinculos, usuarios_por_id, origem_id, excluir_logins)
        selecionados = plano[: args.limit] if args.limit > 0 else plano

        dinamicos = sum(1 for item in plano if item["is_dynamic"])
        print(
            f"PLANO: {len(plano)} vínculo(s) na entidade de origem; "
            f"{len(ignorados)} ignorado(s) por estar na lista excluir_logins; "
            f"{len(selecionados)} selecionado(s) neste comando."
        )
        if dinamicos:
            print(
                f"AVISO: {dinamicos} vínculo(s) selecionado(s) são marcados como dinâmicos "
                "(is_dynamic) - normalmente controlados por sincronização de diretório/regras "
                "de importação. Se houver uma sincronização automática rodando, ela pode "
                "recolocar esses usuários na entidade de origem depois."
            )

        if not args.executar:
            for item in selecionados[:50]:
                print(f"SIMULAÇÃO: {item['login']} -> passaria para '{destino_nome}'.")
            if len(selecionados) > 50:
                print(f"... mais {len(selecionados) - 50} usuário(s) na simulação.")
            print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
            return

        print("\nATENÇÃO: esta ação altera o acesso por entidade de TODOS os usuários listados acima.")
        confirmacao = input("Digite APLICAR para continuar: ").strip()
        if confirmacao != "APLICAR":
            print("Operação cancelada. Nenhuma alteração foi feita.")
            return

        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        erros_path = Path(f"atribuicao_entidade_usuarios_erros_{carimbo}.csv").resolve()
        sucessos = 0
        falhas = 0
        for item in selecionados:
            try:
                glpi.atualizar("Profile_User", item["vinculo_id"], {"entities_id": destino_id})
                sucessos += 1
                print(f"OK: {item['login']} movido para '{destino_nome}'.")
            except Exception as erro:
                falhas += 1
                registrar_erro(erros_path, item["login"], item["vinculo_id"], str(erro))
                print(f"ERRO: {item['login']} | {erro}")
            time.sleep(float(config.get("interval_seconds", 0.15)))

        print(f"\nCONCLUÍDO: {sucessos} usuário(s) movido(s); {falhas} falha(s).")
        if falhas:
            print(f"ERROS: {erros_path}")
    finally:
        glpi.finalizar()


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        print(f"Erro: {erro}")
        raise SystemExit(1)
