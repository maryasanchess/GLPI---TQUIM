# -*- coding: utf-8 -*-
"""Exporta um relatório em Excel de todos os chamados da entidade TQUIM > Ocorrências
do ano corrente, com os campos nativos do chamado e os campos personalizados do
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
from openpyxl.worksheet.datavalidation import DataValidation

ENTITY_OCORRENCIAS = 6
CAMPO_DATA_ABERTURA = 15  # search-option usado tanto pra coluna quanto pra agrupar por mês
CAMPO_ID = 2

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

YESNO = {0: "Não", 1: "Sim", None: ""}


def _as_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


def _yesno(v):
    return YESNO.get(_as_int(v), v)


def _jornada(linha):
    partes = []
    inicio_j = linha.get("76674")
    fim_j = linha.get("76675")
    inicio_c = linha.get("76676")
    fim_c = linha.get("76677")
    if inicio_j or fim_j:
        partes.append(f"Jornada: {inicio_j or '-'} até {fim_j or '-'}")
    if inicio_c or fim_c:
        partes.append(f"Carregamento: {inicio_c or '-'} até {fim_c or '-'}")
    return " | ".join(partes)


# (search-option id ou None se calculado, cabeçalho, formatador opcional)
# Ordem e nomes conforme definido pela TQUIM.
COLUNAS = [
    (CAMPO_ID, "ID", None),
    (7, "Tipo", None),
    (CAMPO_DATA_ABERTURA, "Data", None),
    (76666, "Placa Tração", None),
    (76694, "Código Placa Tração", None),
    (76667, "Placa Semi-Reboque", None),
    (76695, "Código Placa Semi-Reboque", None),
    (76668, "Colaborador", None),
    (76692, "Cargo/Função", None),
    (76693, "Depto/Setor", None),
    (76679, "Inserir na avaliação motorista?", _yesno),
    (76669, "OC/CT-e", None),
    (76670, "Produto", None),
    (76680, "Situação da Carga (Granel/Embalado/Carregado/Vazio)", None),
    (76671, "Origem", None),
    (76672, "Destino", None),
    (76673, "Cliente", None),
    (76678, "Impacto ao Cliente?", _yesno),
    (76696, "Local da ocorrência", None),
    (76699, "Responsável pela Análise", None),
    (76697, "Descrição da Ocorrência", None),
    (76698, "Ações Imediatas", None),
    (76700, "Observações e comentários", None),
    (None, "Descrição da Jornada", _jornada),
    (76681, "Motivo", None),
    (76702, "Responsável (Cliente/Motorista)", None),
    (76684, "Plano de ações", None),
    (76688, "Acionamento Suatrans/Pamcary?", _yesno),
    (76689, "Houve vazamento?", _yesno),
    (76686, "Quantidade que vazou", None),
    (None, "Fotos", None),  # preenchido à parte, via lista de anexos do chamado
    (76685, "Levantamento de Custos (descritivo)", None),
    (76687, "Custo Total", None),
    (76683, "S.A.", None),
    (76691, "Classificação", None),  # deixado em branco/editável para a Qualidade preencher na planilha
]

COL_FOTOS_IDX = next(i for i, (_id, label, _f) in enumerate(COLUNAS) if label == "Fotos")

# Cores do modelo de abertura de Ocorrência (mesmas do e-mail de notificação)
COR_CABECALHO = "FFFF99"
COR_TEXTO_CABECALHO = "333333"
COR_LINHA_PAR = "FFFACD"
COR_BORDA = "D4D9E2"

LARGURAS = {
    "ID": 8, "Tipo": 22, "Data": 16, "Descrição da Ocorrência": 40,
    "Ações Imediatas": 30, "Observações e comentários": 30,
    "Descrição da Jornada": 34, "Fotos": 30,
    "Levantamento de Custos (descritivo)": 30,
}


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


def buscar_chamados(app_ctx):
    base, app_token, sreq = app_ctx
    params = [
        ("criteria[0][field]", "80"),
        ("criteria[0][searchtype]", "equals"),
        ("criteria[0][value]", str(ENTITY_OCORRENCIAS)),
        ("range", "0-9999"),
    ]
    for col_id, _label, _fmt in COLUNAS:
        if col_id is not None:
            params.append(("forcedisplay[]", str(col_id)))
    qs = urllib.parse.urlencode(params)
    res = sreq("GET", f"/search/Ticket?{qs}")
    return res.get("data", [])


def buscar_anexos(app_ctx, ticket_id):
    base, app_token, sreq = app_ctx
    try:
        docs = sreq("GET", f"/Ticket/{ticket_id}/Document_Item")
    except RuntimeError:
        return ""
    nomes = []
    for d in docs:
        doc_id = d.get("documents_id")
        if not doc_id:
            continue
        try:
            doc = sreq("GET", f"/Document/{doc_id}")
            nomes.append(doc.get("filename") or doc.get("name") or f"documento-{doc_id}")
        except RuntimeError:
            continue
    return ", ".join(nomes)


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

    n_cols = len(COLUNAS)

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
    for col_idx, (_id, label, _fmt) in enumerate(COLUNAS, start=1):
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
        for col_idx, (col_id, _label, fmt) in enumerate(COLUNAS, start=1):
            if col_id is None:
                valor = fmt(linha) if fmt else linha.get("_fotos", "")
            else:
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

    for col_idx, (_id, label, _fmt) in enumerate(COLUNAS, start=1):
        letra = get_column_letter(col_idx)
        ws.column_dimensions[letra].width = LARGURAS.get(label, 20)
    ws.sheet_view.showGridLines = False

    # Coluna "Classificação" fica com o campo desativado no GLPI (preenchida
    # pela Qualidade direto aqui na planilha) — deixa pronta pra receber uma
    # lista suspensa assim que os valores forem definidos.
    classificacao_col = next(i for i, (_id, label, _f) in enumerate(COLUNAS, start=1) if label == "Classificação")
    letra = get_column_letter(classificacao_col)
    if len(linhas) > 0:
        dv = DataValidation(type="list", formula1='"Leve,Moderada,Grave,Crítica"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{letra}{header_row + 1}:{letra}{header_row + len(linhas)}")


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

    base = cfg["api_url"]
    app_token = cfg["app_token"]
    user_token = cfg["user_token"]
    init = api_request(base, app_token, "GET", "/initSession", {"Authorization": "user_token " + user_token})
    session_token = init["session_token"]

    def sreq(method, path, data=None):
        return api_request(base, app_token, method, path, {"Session-Token": session_token}, data)

    app_ctx = (base, app_token, sreq)

    try:
        print("Buscando chamados da entidade Ocorrências...")
        linhas = buscar_chamados(app_ctx)
        linhas_ano = filtrar_por_ano(linhas, ano)
        print(f"{len(linhas_ano)} chamado(s) de {ano} encontrado(s). Buscando anexos...")
        for linha in linhas_ano:
            tid = linha.get(str(CAMPO_ID))
            linha["_fotos"] = buscar_anexos(app_ctx, tid) if tid else ""
        print("Montando planilha...")
        montar_planilha(linhas_ano, ano, saida)
        print(f"Pronto: {saida}")
    finally:
        sreq("GET", "/killSession")


if __name__ == "__main__":
    main()
