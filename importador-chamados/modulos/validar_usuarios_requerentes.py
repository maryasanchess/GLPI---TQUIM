#!/usr/bin/env python3
"""Validação interativa dos solicitantes que não casam automaticamente.

Esta etapa apenas consulta usuários do GLPI e grava as escolhas em
``requester_user_map`` dentro do ``config.json``. Nenhum chamado é alterado e
nenhuma notificação é enviada.
"""

import argparse
import difflib
import warnings
from pathlib import Path

import pandas as pd

from atribuir_setores_chamados_glpi import (
    Glpi,
    carregar_json,
    normalizar,
    salvar_json,
    texto,
)
from atribuir_usuarios_requerentes_glpi import (
    criar_indices_usuarios,
    extrair_email,
    montar_plano,
    resolver_usuario,
)


warnings.filterwarnings(
    "ignore",
    message=r"Cell .* is marked as a date.*",
    category=UserWarning,
)


def dados_usuario(usuario):
    primeiro_nome = texto(usuario.get("firstname"))
    sobrenome = texto(usuario.get("realname"))
    nome = " ".join(filter(None, [primeiro_nome, sobrenome]))
    nome = nome or texto(usuario.get("completename")) or texto(usuario.get("name"))
    login = texto(usuario.get("name"))
    email = extrair_email(usuario.get("email")) or extrair_email(login)
    return {
        "id": int(usuario.get("id") or 0),
        "nome": nome,
        "login": login,
        "email": email,
    }


def pontuacao_candidato(item, candidato):
    alvos = [normalizar(item["nome"])]
    if item["email"]:
        alvos.extend(
            [
                normalizar(item["email"]),
                normalizar(item["email"].split("@", 1)[0]),
            ]
        )

    opcoes = [
        normalizar(candidato["nome"]),
        normalizar(candidato["login"]),
        normalizar(candidato["email"]),
    ]
    if candidato["email"]:
        opcoes.append(normalizar(candidato["email"].split("@", 1)[0]))

    melhor = 0.0
    for alvo in filter(None, alvos):
        tokens_alvo = set(alvo.split())
        for opcao in filter(None, opcoes):
            similaridade = difflib.SequenceMatcher(None, alvo, opcao).ratio()
            tokens_opcao = set(opcao.split())
            uniao = tokens_alvo | tokens_opcao
            sobreposicao = len(tokens_alvo & tokens_opcao) / len(uniao) if uniao else 0
            melhor = max(melhor, similaridade, sobreposicao)
    return melhor


def sugerir_usuarios(item, usuarios, limite=5):
    candidatos = []
    for usuario in usuarios:
        candidato = dados_usuario(usuario)
        if not candidato["id"]:
            continue
        candidato["pontuacao"] = pontuacao_candidato(item, candidato)
        candidatos.append(candidato)
    candidatos.sort(
        key=lambda candidato: (
            -candidato["pontuacao"],
            normalizar(candidato["nome"]),
            candidato["id"],
        )
    )
    return candidatos[:limite]


def agrupar_pendentes(plano, indices, mapa_config):
    grupos = {}
    for item in plano:
        usuario_id, _ = resolver_usuario(item, indices, mapa_config)
        if usuario_id:
            continue
        chave = item["email"] or normalizar(item["nome"])
        grupo = grupos.setdefault(
            chave,
            {
                "chave": chave,
                "nome": item["nome"],
                "email": item["email"],
                "itens": [],
            },
        )
        grupo["itens"].append(item)
    return list(grupos.values())


def mostrar_candidato(numero, candidato):
    email = candidato["email"] or "sem e-mail"
    login = candidato["login"] or "sem login"
    percentual = round(candidato["pontuacao"] * 100)
    print(
        f"  {numero}) {candidato['nome']} | login: {login} | "
        f"e-mail: {email} | ID: {candidato['id']} | semelhança: {percentual}%"
    )


def main():
    parser = argparse.ArgumentParser(
        description="Valida interativamente os usuários requerentes pendentes"
    )
    parser.add_argument("excel", help="Última planilha .xlsx")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--estado", default="importacao_glpi_estado.json")
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    config_path = Path(args.config).resolve()
    estado_path = Path(args.estado).resolve()
    if not excel_path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")
    if not config_path.exists():
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")
    if not estado_path.exists():
        raise FileNotFoundError(f"Controle da importação não encontrado: {estado_path}")

    df = pd.read_excel(excel_path)
    estado = carregar_json(estado_path)
    config = carregar_json(config_path)
    plano, _, _ = montar_plano(df, estado)

    glpi = Glpi(config)
    try:
        print("Conectando ao GLPI apenas para consultar os usuários...")
        glpi.iniciar()
        usuarios = glpi.listar_todos("User")
    finally:
        glpi.finalizar()

    indices = criar_indices_usuarios(usuarios)
    mapa_config = config.get("requester_user_map")
    if not isinstance(mapa_config, dict):
        mapa_config = {}
        config["requester_user_map"] = mapa_config
    pendentes = agrupar_pendentes(plano, indices, mapa_config)
    usuarios_por_id = {
        int(usuario.get("id")): usuario
        for usuario in usuarios
        if usuario.get("id")
    }

    print(f"USUÁRIOS GLPI CONSULTADOS: {len(indices['ids'])}.")
    print(f"SOLICITANTES QUE PRECISAM DE ESCOLHA: {len(pendentes)}.")
    print("Esta validação não altera chamados e não envia notificações.")

    if not pendentes:
        print("Todos os solicitantes já foram localizados ou validados.")
        return

    salvos = 0
    ignorados = 0
    for posicao, item in enumerate(pendentes, start=1):
        print()
        print("=" * 72)
        print(f"VALIDAÇÃO {posicao} DE {len(pendentes)}")
        print(f"Planilha: {item['nome']}")
        print(f"E-mail: {item['email'] or 'não informado'}")
        exemplos = ", ".join(
            f"#{registro['ticket_id']}" for registro in item["itens"][:8]
        )
        complemento = "..." if len(item["itens"]) > 8 else ""
        print(
            f"Chamados relacionados: {len(item['itens'])} "
            f"({exemplos}{complemento})"
        )

        sugestoes = sugerir_usuarios(item, usuarios)
        print("Opções mais parecidas encontradas no GLPI:")
        for numero, candidato in enumerate(sugestoes, start=1):
            mostrar_candidato(numero, candidato)
        print("  0) Deixar pendente e passar para o próximo")
        print("  I) Informar manualmente o ID do usuário")
        print("  X) Salvar as escolhas e sair")

        while True:
            escolha = input("Escolha uma opção: ").strip().casefold()
            if escolha == "x":
                salvar_json(config_path, config)
                print(
                    f"Validação interrompida. {salvos} escolhas foram salvas "
                    "no config.json."
                )
                return
            if escolha == "0":
                ignorados += 1
                print("Mantido como pendente.")
                break
            if escolha == "i":
                valor_id = input("Digite o ID do usuário no GLPI: ").strip()
                try:
                    usuario_id = int(valor_id)
                except ValueError:
                    print("ID inválido. Digite apenas números.")
                    continue
                usuario = usuarios_por_id.get(usuario_id)
                if not usuario:
                    print("Esse ID não foi encontrado entre os usuários do GLPI.")
                    continue
                escolhido = dados_usuario(usuario)
            else:
                try:
                    indice_escolhido = int(escolha) - 1
                    escolhido = sugestoes[indice_escolhido]
                except (ValueError, IndexError):
                    print("Opção inválida.")
                    continue

            chave_mapa = item["email"] or item["nome"]
            mapa_config[chave_mapa] = escolhido["id"]
            salvar_json(config_path, config)
            salvos += 1
            print(
                f"SALVO: {item['nome']} -> {escolhido['nome']} "
                f"(GLPI ID {escolhido['id']})."
            )
            break

    salvar_json(config_path, config)
    print()
    print(
        f"VALIDAÇÃO CONCLUÍDA: {salvos} escolhas salvas; "
        f"{ignorados} solicitantes mantidos como pendentes."
    )
    print(
        "Agora execute novamente a simulação e o teste dos usuários requerentes."
    )


if __name__ == "__main__":
    main()
