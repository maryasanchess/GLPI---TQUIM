#!/usr/bin/env python3
"""Configura o acesso à API e reconstrói o controle EXCEL-N -> chamado GLPI."""

import getpass
import json
import re
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path


API_PADRAO = "https://SEU-SERVIDOR-GLPI/api.php/v1"


def ler_json(caminho):
    with open(caminho, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def salvar_json(caminho, dados):
    temporario = caminho.with_suffix(".tmp")
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)
    temporario.replace(caminho)


def valor_valido(config, campo):
    valor = str(config.get(campo, "")).strip()
    return bool(valor and "COLE_AQUI" not in valor)


def pedir_config(config_path):
    config_existente = ler_json(config_path) if config_path.exists() else {}
    if all(valor_valido(config_existente, campo) for campo in ("api_url", "app_token", "user_token")):
        usar = input("Foi encontrado um config.json preenchido. Usar esse arquivo? [S/n]: ").strip().casefold()
        if usar in {"", "s", "sim"}:
            return config_existente

    print("\nCONFIGURAÇÃO DA API DO GLPI")
    print("Os tokens não aparecerão na tela enquanto você digita.\n")
    atual_url = str(config_existente.get("api_url") or API_PADRAO).strip()
    api_url = input(f"URL da API [{atual_url}]: ").strip() or atual_url
    app_token = getpass.getpass("App-Token: ").strip()
    user_token = getpass.getpass("User-Token: ").strip()
    if not app_token or not user_token:
        raise ValueError("App-Token e User-Token são obrigatórios.")

    return {
        "api_url": api_url.rstrip("/"),
        "app_token": app_token,
        "user_token": user_token,
        "profile_id": 0,
        "interval_seconds": 0.15,
        "scan_existing_tickets": False,
        "duplicate_scan_range": 30000,
        "duplicate_scan_page_size": 500,
        "local_attachment_dir": "anexos_forms",
        "auto_create_categories": False,
        "technician_map": {},
        "category_map": {},
        "department_group_map": {}
    }


class Glpi:
    def __init__(self, config):
        self.api_url = config["api_url"].rstrip("/")
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
            detalhe = erro.read().decode("utf-8", errors="replace")[:1500]
            raise RuntimeError(
                f"Falha na API em {metodo} {endpoint}: HTTP {erro.code} - {detalhe}"
            ) from erro
        return json.loads(conteudo) if conteudo else None

    def iniciar(self):
        resposta = self.requisicao("GET", "initSession", timeout=30)
        self.session_token = resposta["session_token"]
        self.headers["Session-Token"] = self.session_token

    def finalizar(self):
        if self.session_token:
            try:
                self.requisicao("GET", "killSession", timeout=15)
            except Exception:
                pass

    def listar_todos(self, tipo, limite=30000, pagina=500):
        itens = []
        for inicio in range(0, limite, pagina):
            fim = inicio + pagina - 1
            retorno = self.requisicao("GET", f"{tipo}?range={inicio}-{fim}")
            lote = retorno if isinstance(retorno, list) else []
            itens.extend(lote)
            print(f"Consultados {len(itens)} itens do GLPI...")
            if len(lote) < pagina:
                break
        return itens


def reconstruir_estado(glpi):
    tickets = glpi.listar_todos("Ticket")
    encontrados = {}
    for ticket in tickets:
        titulo = str(ticket.get("name") or "")
        achado = re.search(r"\[MIGRAÇÃO\s+(EXCEL-\d+)\]", titulo, flags=re.IGNORECASE)
        if not achado or not ticket.get("id"):
            continue
        referencia = achado.group(1).upper()
        ticket_id = int(ticket["id"])
        encontrados.setdefault(referencia, set()).add(ticket_id)

    if len(encontrados) < 100:
        raise RuntimeError(
            f"Apenas {len(encontrados)} chamados migrados foram encontrados. "
            "A reconstrução foi bloqueada para evitar um arquivo incompleto."
        )

    agora = datetime.now().isoformat(timespec="seconds")
    importados = {}
    duplicados = {}
    for referencia, ids_encontrados in sorted(
        encontrados.items(),
        key=lambda item: int(item[0].split("-")[1]),
    ):
        ticket_ids = sorted(ids_encontrados)
        controle = {
            "ticket_id": ticket_ids[0],
            "reconstruido_em": agora,
        }
        if len(ticket_ids) > 1:
            controle["ticket_ids"] = ticket_ids
            controle["referencia_duplicada_no_glpi"] = True
            duplicados[referencia] = ticket_ids
        importados[referencia] = controle

    return {
        "importados": importados,
        "reconstrucao": {
            "data": agora,
            "origem": "Títulos [MIGRAÇÃO EXCEL-N] consultados no GLPI",
            "quantidade": len(importados),
            "quantidade_tickets": sum(len(ids) for ids in encontrados.values()),
            "referencias_duplicadas": duplicados,
        },
    }


def main():
    config_path = Path("config.json").resolve()
    estado_path = Path("importacao_glpi_estado.json").resolve()
    config = pedir_config(config_path)
    salvar_json(config_path, config)
    print(f"\nConfiguração salva em {config_path.name}.")

    glpi = Glpi(config)
    try:
        print("Conectando ao GLPI...")
        glpi.iniciar()
        print("Conexão realizada. Reconstruindo o arquivo de estado...")
        estado = reconstruir_estado(glpi)
    finally:
        glpi.finalizar()

    if estado_path.exists():
        backup = estado_path.with_name(
            f"backup_importacao_glpi_estado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        backup.write_bytes(estado_path.read_bytes())
        print(f"Backup do estado anterior criado: {backup.name}")
    salvar_json(estado_path, estado)
    print(
        f"\nSUCESSO: {estado['reconstrucao']['quantidade']} vínculos "
        f"EXCEL-N -> GLPI foram gravados em {estado_path.name}."
    )
    duplicados = estado["reconstrucao"].get("referencias_duplicadas", {})
    if duplicados:
        print("\nATENÇÃO: referências duplicadas foram preservadas:")
        for referencia, ticket_ids in duplicados.items():
            print(f"- {referencia}: chamados GLPI {ticket_ids}")
        print("O programa de setores processará todos esses IDs.")


if __name__ == "__main__":
    main()
