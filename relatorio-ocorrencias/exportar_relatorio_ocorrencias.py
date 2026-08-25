# -*- coding: utf-8 -*-
"""Exporta um relatório em Excel de todos os chamados da entidade TQUIM > Ocorrências
do ano corrente, com os campos nativos do chamado e os 37 campos personalizados do
plugin Fields. O arquivo sai com uma aba "Anual" (todos os chamados do ano) e uma
aba por mês que tiver pelo menos um chamado.

Uso:
    python exportar_relatorio_ocorrencias.py [caminho_saida.xlsx] [ano]

Requer config.json na mesma pasta (mesmo formato usado nos outros scripts do projeto):
    {"api_url": "...", "app_token": "...", "user_token": "..."}
"""
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ENTITY_OCORRENCIAS = 6
CAMPO_DATA_ABERTURA = 15  # search-option usado tanto pra coluna quanto pra agrupar por mês

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

# (search-option id, cabeçalho, formatador opcional)
YESNO = {0: "Não", 1: "Sim", None: ""}
STATUS = {
    1: "Novo", 2: "Em atendimento (atribuído)", 3: "Em atendimento (planejado)",
    4: "Pendente", 5: "Solucionado", 6: "Fechado",
}
PRIORIDADE = {
    1: "Muito baixa", 2: "Baixa", 3: "Média", 4: "Alta", 5: "Muito alta", 6: "Crítica",
}

NATIVE_COLUMNS = [
    (2, "ID", None),
    (1, "Título", None),
    (7, "Categoria", None),
    (3, "Prioridade", lambda v: PRIORIDADE.get(_as_int(v), v)),
    (12, "Status", lambda v: STATUS.get(_as_int(v), v)),
    (CAMPO_DATA_ABERTURA, "Data de abertura", None),
    (119, "Fotos (Qtd. Anexos)", None),
]

PLUGIN_COLUMNS = [
    (76666, "Placa Tração", None),
    (76667, "Placa Semi-Reboque", None),
    (76668, "Colaborador Motorista", None),
    (76669, "OC / CT-e / NF", None),
    (76670, "Produto", None),
    (76671, "Origem", None),
    (76672, "Destino", None),
    (76673, "Cliente", None),
    (76674, "Início da Jornada", None),
    (76675, "Fim da Jornada", None),
    (76676, "Início do Carregamento", None),
    (76677, "Fim do Carregamento", None),
    (76678, "Impacto ao Cliente", lambda v: YESNO.get(_as_int(v), v)),
    (76679, "Inserir na avaliação motorista", lambda v: YESNO.get(_as_int(v), v)),
    (76680, "Situação da Carga", None),
    (76681, "Motivo", None),
    (76682, "Cod. Multa", None),
    (76683, "S.A.", None),
    (76684, "Plano de Ações", None),
    (76685, "Levantamento de Custos", None),
    (76686, "Quantidade que vazou", None),
    (76687, "Custo Total", None),
    (76688, "Acionamento SuatransPamcary", lambda v: YESNO.get(_as_int(v), v)),
    (76689, "Houve vazamento", lambda v: YESNO.get(_as_int(v), v)),
    (76690, "Tipo de NC", None),
    (76691, "Classificação", None),
    (76692, "Cargo/Função", None),
    (76693, "Depto/Setor", None),
    (76694, "Cód. Frota Tração", None),
    (76695, "Cód. Frota Semi Reboque", None),
    (76696, "Local da Ocorrência", None),
    (76697, "Descrição da Ocorrência", None),
    (76698, "Providências já Tomadas", None),
    (76699, "Responsável pela Análise / Ações Corretivas", None),
    (76700, "Outras Observações", None),
    (76701, "Custos da NC", None),
    (76702, "Responsável (Cliente/Motorista)", None),
]

ALL_COLUMNS = NATIVE_COLUMNS + PLUGIN_COLUMNS

# Cores do modelo de abertura de Ocorrência (mesmas do e-mail de notificação)
COR_CABECALHO = "FFFF99"
COR_TEXTO_CABECALHO = "333333"
COR_LINHA_PAR = "FFFACD"
COR_BORDA = "D4D9E2"

LARGURAS = {
    "ID": 8, "Título": 32, "Categoria": 22, "Prioridade": 12, "Status": 20,
    "Data de abertura": 16, "Fotos (Qtd. Anexos)": 18,
}


def _as_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


def api_request(base, app_token, method, path, headers=None, data=None):
    url = base + path
    h = {"App-Token": app_token, "Content-Type": "application/json"}
    if headers:
        h.update(headers)
    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, method=method, headers=h)
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Erro na API ({e.code}): {e.read().decode()}") from e


def buscar_chamados(cfg):
    base = cfg["api_url"]
    app_token = cfg["app_token"]
    user_token = cfg["user_token"]

    init = api_request(base, app_token, "GET", "/initSession", {"Authorization": "user_token " + user_token})
    session_token = init["session_token"]

    def sreq(method, path, data=None):
        return api_request(base, app_token, method, path, {"Session-Token": session_token}, data)

    try:
        params = [
            ("criteria[0][field]", "80"),
            ("criteria[0][searchtype]", "equals"),
            ("criteria[0][value]", str(ENTITY_OCORRENCIAS)),
            ("range", "0-9999"),
        ]
        for col_id, _label, _fmt in ALL_COLUMNS:
            params.append(("forcedisplay[]", str(col_id)))

        qs = urllib.parse.urlencode(params)
        res = sreq("GET", f"/search/Ticket?{qs}")
        return res.get("data", [])
    finally:
        sreq("GET", "/killSession")


def mes_da_linha(linha):
    valor = linha.get(str(CAMPO_DATA_ABERTURA))
    if not valor:
        return None
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d")
    except ValueError:
        return None


def filtrar_por_ano(linhas, ano):
    return [l for l in linhas if (d := mes_da_linha(l)) and d.year == ano]


def agrupar_por_mes(linhas):
    grupos = {}
    for linha in linhas:
        data = mes_da_linha(linha)
        if not data:
            continue
        grupos.setdefault(data.month, []).append(linha)
    return grupos


def escrever_aba(ws, linhas):
    fonte_padrao = Font(name="Arial", size=10)
    fonte_titulo = Font(name="Arial", size=14, bold=True, color=COR_TEXTO_CABECALHO)
    fonte_subtitulo = Font(name="Arial", size=9, italic=True, color="555555")
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color=COR_TEXTO_CABECALHO)
    borda = Border(*(Side(style="thin", color=COR_BORDA),) * 4)

    n_cols = len(ALL_COLUMNS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    titulo = ws.cell(row=1, column=1, value=f"TQUIM > OCORRÊNCIAS — {ws.title.upper()}")
    titulo.font = fonte_titulo
    titulo.fill = PatternFill("solid", fgColor=COR_CABECALHO)
    titulo.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    sub = ws.cell(row=2, column=1, value=f"Gerado em {datetime.now():%d/%m/%Y %H:%M} — {len(linhas)} chamado(s)")
    sub.font = fonte_subtitulo
    sub.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 16

    header_row = 4
    for col_idx, (_id, label, _fmt) in enumerate(ALL_COLUMNS, start=1):
        c = ws.cell(row=header_row, column=col_idx, value=label)
        c.font = fonte_cabecalho
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda
    ws.row_dimensions[header_row].height = 32
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    for i, linha in enumerate(linhas):
        row_idx = header_row + 1 + i
        fill = PatternFill("solid", fgColor=COR_LINHA_PAR) if i % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, (col_id, _label, fmt) in enumerate(ALL_COLUMNS, start=1):
            valor = linha.get(str(col_id))
            if fmt:
                valor = fmt(valor)
            if valor is None:
                valor = ""
            c = ws.cell(row=row_idx, column=col_idx, value=valor)
            c.font = fonte_padrao
            c.fill = fill
            c.border = borda
            c.alignment = Alignment(vertical="center", wrap_text=False)

    for col_idx, (_id, label, _fmt) in enumerate(ALL_COLUMNS, start=1):
        letra = get_column_letter(col_idx)
        ws.column_dimensions[letra].width = LARGURAS.get(label, 20)
    ws.sheet_view.showGridLines = False


def montar_planilha(linhas_ano, ano, caminho_saida):
    wb = Workbook()
    wb.remove(wb.active)

    ws_anual = wb.create_sheet(f"Anual {ano}")
    escrever_aba(ws_anual, linhas_ano)

    grupos = agrupar_por_mes(linhas_ano)
    for mes in sorted(grupos):
        ws_mes = wb.create_sheet(f"{MESES[mes]} - {ano}")
        escrever_aba(ws_mes, grupos[mes])

    wb.save(caminho_saida)


def main():
    pasta = Path(__file__).resolve().parent
    cfg = json.loads((pasta / "config.json").read_text(encoding="utf-8-sig"))

    ano = int(sys.argv[2]) if len(sys.argv) > 2 else datetime.now().year
    if len(sys.argv) > 1:
        saida = Path(sys.argv[1])
    else:
        saida = pasta / f"Relatorio_Ocorrencias_{ano}.xlsx"

    print("Buscando chamados da entidade Ocorrências...")
    linhas = buscar_chamados(cfg)
    linhas_ano = filtrar_por_ano(linhas, ano)
    print(f"{len(linhas_ano)} chamado(s) de {ano} encontrado(s). Montando planilha...")
    montar_planilha(linhas_ano, ano, saida)
    print(f"Pronto: {saida}")


if __name__ == "__main__":
    main()
