#!/usr/bin/env python3
"""Adiciona os solicitantes da planilha como usuários requerentes no GLPI.

Os usuários precisam existir previamente no GLPI. O programa não cria usuários,
não remove atores existentes, evita duplicidades e desativa a notificação
gerada pela inclusão do vínculo.
"""

import argparse
import csv
import re
import time
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

from atribuir_setores_chamados_glpi import (
    Glpi,
    carregar_json,
    linha_valida,
    normalizar,
    normalizar_colunas,
    salvar_json,
    texto,
)


TIPO_REQUERENTE = 1
PADRAO_EMAIL = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

warnings.filterwarnings(
    "ignore",
    message=r"Cell .* is marked as a date.*",
    category=UserWarning,
)


def extrair_email(valor):
    encontrado = PADRAO_EMAIL.search(texto(valor))
    return encontrado.group(0).casefold() if encontrado else ""


def nome_solicitante(linha):
    return " ".join(
        filter(
            None,
            [
                texto(linha.get("Nome do Solicitante")),
                texto(linha.get("Sobrenome do Solicitante")),
            ],
        )
    )


def montar_plano(df, estado):
    plano = []
    sem_vinculo = 0
    sem_nome = 0
    for indice, linha in df.iterrows():
        if not linha_valida(linha):
            continue
        referencia = f"EXCEL-{indice + 2}"
        controle = estado.get("importados", {}).get(referencia)
        if not controle or not controle.get("ticket_id"):
            sem_vinculo += 1
            continue

        nome = nome_solicitante(linha)
        if not nome:
            sem_nome += 1
            continue
        email = extrair_email(linha.get("Endereço de e-mail"))
        ticket_ids = controle.get("ticket_ids") or [controle.get("ticket_id")]
        ticket_ids = list(
            dict.fromkeys(int(ticket_id) for ticket_id in ticket_ids if ticket_id)
        )
        for ticket_id in ticket_ids:
            plano.append(
                {
                    "referencia": referencia,
                    "ticket_id": ticket_id,
                    "nome": nome,
                    "email": email,
                }
            )
    return plano, sem_vinculo, sem_nome


def adicionar_indice(indice, chave, usuario_id):
    chave = normalizar(chave)
    if chave:
        indice.setdefault(chave, set()).add(int(usuario_id))


def adicionar_email(indice, valor, usuario_id):
    email = extrair_email(valor)
    if email:
        indice.setdefault(email, set()).add(int(usuario_id))


def criar_indices_usuarios(usuarios):
    ids_validos = set()
    por_nome = {}
    por_email = {}
    por_login = {}
    for usuario in usuarios:
        usuario_id = int(usuario.get("id") or 0)
        if not usuario_id:
            continue
        ids_validos.add(usuario_id)

        login = texto(usuario.get("name"))
        primeiro_nome = texto(usuario.get("firstname"))
        sobrenome = texto(usuario.get("realname"))
        nome_completo = texto(usuario.get("completename"))

        adicionar_indice(por_login, login, usuario_id)
        adicionar_email(por_email, login, usuario_id)
        adicionar_email(por_email, usuario.get("email"), usuario_id)
        emails = usuario.get("emails")
        if isinstance(emails, list):
            for email in emails:
                adicionar_email(por_email, email, usuario_id)

        if primeiro_nome and sobrenome:
            adicionar_indice(
                por_nome,
                f"{primeiro_nome} {sobrenome}",
                usuario_id,
            )
            adicionar_indice(
                por_nome,
                f"{sobrenome} {primeiro_nome}",
                usuario_id,
            )
        adicionar_indice(por_nome, nome_completo, usuario_id)
        adicionar_indice(por_nome, login, usuario_id)

    return {
        "ids": ids_validos,
        "nome": por_nome,
        "email": por_email,
        "login": por_login,
    }


def unico(indice, chave):
    candidatos = indice.get(chave, set())
    if len(candidatos) == 1:
        return next(iter(candidatos)), ""
    if len(candidatos) > 1:
        return None, f"{len(candidatos)} usuários correspondentes"
    return None, ""


def resolver_usuario(item, indices, mapa_config):
    nome_normalizado = normalizar(item["nome"])
    email = item["email"]

    destinos_configurados = []
    for origem, destino in mapa_config.items():
        origem_email = extrair_email(origem)
        if (
            (origem_email and origem_email == email)
            or normalizar(origem) == nome_normalizado
        ):
            destinos_configurados.append(destino)
    if len(destinos_configurados) > 1:
        return None, "mais de um mapeamento configurado"
    if destinos_configurados:
        destino = destinos_configurados[0]
        try:
            usuario_id = int(destino)
        except (TypeError, ValueError):
            destino_email = extrair_email(destino)
            if destino_email:
                usuario_id, motivo = unico(indices["email"], destino_email)
            else:
                chave_destino = normalizar(destino)
                usuario_id, motivo = unico(indices["login"], chave_destino)
                if not usuario_id and not motivo:
                    usuario_id, motivo = unico(indices["nome"], chave_destino)
            if not usuario_id:
                return None, f"mapeamento configurado não localizado: {motivo or destino}"
        if usuario_id not in indices["ids"]:
            return None, f"ID configurado não existe no GLPI: {usuario_id}"
        return usuario_id, "mapeamento configurado"

    if email:
        usuario_id, motivo = unico(indices["email"], email)
        if usuario_id:
            return usuario_id, "e-mail"
        if motivo:
            return None, f"e-mail ambíguo: {motivo}"

        local_email = email.split("@", 1)[0]
        usuario_id, motivo = unico(indices["login"], normalizar(local_email))
        if usuario_id:
            return usuario_id, "login do e-mail"
        if motivo:
            return None, f"login ambíguo: {motivo}"

    usuario_id, motivo = unico(indices["nome"], nome_normalizado)
    if usuario_id:
        return usuario_id, "nome completo"
    if motivo:
        return None, f"nome ambíguo: {motivo}"
    return None, "usuário não encontrado no GLPI"


def registrar_pendencia(caminho, item, resultado, detalhe):
    novo = not caminho.exists()
    with open(caminho, "a", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        if novo:
            escritor.writerow(
                [
                    "Referência",
                    "Chamado GLPI",
                    "Solicitante Excel",
                    "E-mail Excel",
                    "Resultado",
                    "Detalhe",
                ]
            )
        escritor.writerow(
            [
                item["referencia"],
                item["ticket_id"],
                item["nome"],
                item["email"],
                resultado,
                detalhe,
            ]
        )


def main():
    parser = argparse.ArgumentParser(
        description="Adiciona os solicitantes como usuários requerentes no GLPI"
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

    excel_path = Path(args.excel).resolve()
    config_path = Path(args.config).resolve()
    estado_path = Path(args.estado).resolve()
    if not excel_path.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")
    if not estado_path.exists():
        raise FileNotFoundError(f"Controle da importação não encontrado: {estado_path}")

    df = pd.read_excel(excel_path)
    df = normalizar_colunas(df)
    if "Nome do Solicitante" not in df.columns:
        raise ValueError("A planilha não possui a coluna Nome do Solicitante.")

    estado = carregar_json(estado_path)
    config = carregar_json(config_path) if config_path.exists() else {}
    if args.executar and not config:
        raise FileNotFoundError(f"Configuração não encontrada: {config_path}")

    plano, sem_vinculo, sem_nome = montar_plano(df, estado)
    selecionados = plano[: args.limit] if args.limit > 0 else plano
    print(
        f"PLANO REQUERENTES: {len(plano)} chamados com solicitante; "
        f"{sem_vinculo} linhas sem vínculo no estado; {sem_nome} linhas sem nome; "
        f"{len(selecionados)} chamados selecionados."
    )

    if not args.executar:
        for item in selecionados[:50]:
            complemento = f" <{item['email']}>" if item["email"] else ""
            print(
                f"SIMULAÇÃO: {item['referencia']} -> GLPI #{item['ticket_id']} | "
                f"{item['nome']}{complemento} será usuário requerente."
            )
        if len(selecionados) > 50:
            print(f"... mais {len(selecionados) - 50} chamados na simulação.")
        print("Nenhuma alteração foi realizada. Acrescente --executar para aplicar.")
        return

    glpi = Glpi(config)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    pendencias_path = Path(f"usuarios_requerentes_pendencias_{carimbo}.csv").resolve()
    sucessos = 0
    existentes = 0
    nao_encontrados = 0
    falhas = 0
    try:
        glpi.iniciar()
        usuarios = glpi.listar_todos("User")
        indices = criar_indices_usuarios(usuarios)
        mapa_config = config.get("requester_user_map", {})
        print(f"USUÁRIOS GLPI: {len(indices['ids'])} usuários consultados.")

        vinculos = glpi.listar_todos("Ticket_User")
        existentes_glpi = {
            (
                int(vinculo.get("tickets_id") or 0),
                int(vinculo.get("users_id") or 0),
                int(vinculo.get("type") or 0),
            )
            for vinculo in vinculos
        }

        for item in selecionados:
            usuario_id, criterio = resolver_usuario(item, indices, mapa_config)
            if not usuario_id:
                nao_encontrados += 1
                registrar_pendencia(
                    pendencias_path,
                    item,
                    "PENDENTE",
                    criterio,
                )
                print(
                    f"PENDENTE: {item['referencia']} / GLPI #{item['ticket_id']} | "
                    f"{item['nome']}: {criterio}."
                )
                continue

            chave_vinculo = (item["ticket_id"], usuario_id, TIPO_REQUERENTE)
            if chave_vinculo in existentes_glpi:
                existentes += 1
                print(
                    f"JÁ OK: {item['referencia']} / GLPI #{item['ticket_id']} | "
                    f"{item['nome']} já é usuário requerente."
                )
                continue

            try:
                try:
                    glpi.criar(
                        "Ticket_User",
                        {
                            "tickets_id": item["ticket_id"],
                            "users_id": usuario_id,
                            "type": TIPO_REQUERENTE,
                            "_disablenotif": 1,
                        },
                    )
                except Exception as erro_inicial:
                    # O GLPI recusa adicionar um novo ator a um chamado já
                    # fechado; reabre, tenta de novo e restaura o status.
                    if "ERROR_GLPI_ADD" not in str(erro_inicial):
                        raise
                    ticket_atual = glpi.requisicao("GET", f"Ticket/{item['ticket_id']}")
                    status_original = int(ticket_atual.get("status") or 0)
                    glpi.atualizar("Ticket", item["ticket_id"], {"status": 1})
                    try:
                        glpi.criar(
                            "Ticket_User",
                            {
                                "tickets_id": item["ticket_id"],
                                "users_id": usuario_id,
                                "type": TIPO_REQUERENTE,
                                "_disablenotif": 1,
                            },
                        )
                    finally:
                        if status_original:
                            glpi.atualizar("Ticket", item["ticket_id"], {"status": status_original})
                existentes_glpi.add(chave_vinculo)
                sucessos += 1
                print(
                    f"OK: {item['referencia']} / GLPI #{item['ticket_id']} | "
                    f"{item['nome']} foi adicionado como requerente ({criterio})."
                )
            except Exception as erro:
                falhas += 1
                registrar_pendencia(
                    pendencias_path,
                    item,
                    "ERRO API",
                    str(erro),
                )
                print(
                    f"ERRO: {item['referencia']} / GLPI #{item['ticket_id']} | {erro}"
                )
            time.sleep(float(config.get("interval_seconds", 0.15)))
    finally:
        glpi.finalizar()

    estado.setdefault("sincronizacao_usuarios_requerentes", {})[
        "ultima_execucao"
    ] = {
        "data": datetime.now().isoformat(timespec="seconds"),
        "planilha": excel_path.name,
        "chamados_selecionados": len(selecionados),
        "vinculos_criados": sucessos,
        "vinculos_ja_existentes": existentes,
        "usuarios_nao_encontrados": nao_encontrados,
        "falhas_api": falhas,
    }
    salvar_json(estado_path, estado)

    print(
        f"CONCLUÍDO REQUERENTES: {sucessos} vínculos criados; "
        f"{existentes} já existentes; {nao_encontrados} usuários pendentes; "
        f"{falhas} falhas de API."
    )
    if nao_encontrados or falhas:
        print(f"PENDÊNCIAS: {pendencias_path}")


if __name__ == "__main__":
    main()
