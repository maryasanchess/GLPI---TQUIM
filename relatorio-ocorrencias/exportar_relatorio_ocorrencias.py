# -*- coding: utf-8 -*-
"""Exporta um relatório em Excel de todos os chamados da entidade TQUIM > Ocorrências
do ano corrente, com os campos nativos do chamado e os campos personalizados do
plugin Fields.

O arquivo sai com três partes:
  - "Anual <ano>": painel do ano — números do período, consulta interativa por
    mês e por tipo (listas suspensas), quebra por mês e por tipo com gráficos, e
    as instruções de como pedir o relatório por e-mail;
  - uma aba por mês que tiver pelo menos um chamado, com a listagem detalhada;
  - "Dados" (oculta): a listagem do ano inteiro, que é a base somada pelo painel.

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
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
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


def _num(v):
    """O GLPI devolve o custo como texto; vira número pra poder somar no indicador."""
    if v is None or str(v).strip() == "":
        return ""
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return v          # não era número: mantém o que veio


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
    (76687, "Custo Total", _num),
    (76683, "S.A.", None),
    (76691, "Classificação", None),  # deixado em branco/editável para a Qualidade preencher na planilha
]

COL_FOTOS_IDX = next(i for i, (_id, label, _f) in enumerate(COLUNAS) if label == "Fotos")

# Cores do modelo de abertura de Ocorrência (mesmas do e-mail de notificação)
COR_CABECALHO = "FFFF99"
COR_TEXTO_CABECALHO = "333333"
COR_LINHA_PAR = "FFFACD"
COR_BORDA = "D4D9E2"

# Cores do painel anual (aba do indicador)
COR_INDICADOR = "1F4E79"
COR_INTERATIVO = "1E7A5F"
COR_PEDIDO = "C0622A"

# Como pedir o relatório fora do dia agendado — mesmo gatilho aceito pelo
# verificar_solicitacao_relatorio.php, que lê os chamados abertos pelo coletor.
EMAIL_PEDIDO = "chamados.ti@tquim.com.br"
PALAVRA_CHAVE_PEDIDO = "RELATORIO OCORRENCIAS"

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
    # "Julho - 2026" -> "Ocorrências - Julho 2026"
    titulo = ws.cell(row=1, column=1, value=f"Ocorrências - {ws.title.replace(' - ', ' ')}")
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


def escrever_dados(ws, linhas):
    """Listagem crua do ano, numa aba oculta: é a base que o indicador soma.
    Cabeçalho na linha 1 e dados a partir da 2, pra manter as faixas previsíveis."""
    ws.append([label for _id, label, _fmt in COLUNAS])
    for c in ws[1]:
        c.font = Font(name="Arial", size=10, bold=True, color=COR_TEXTO_CABECALHO)
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)

    for linha in linhas:
        valores = []
        for col_id, _label, fmt in COLUNAS:
            if col_id is None:
                valor = fmt(linha) if fmt else linha.get("_fotos", "")
            else:
                valor = linha.get(str(col_id))
                if fmt:
                    valor = fmt(valor)
            valores.append("" if valor is None else valor)
        ws.append(valores)

    col_custo = next(i for i, (_id, label, _f) in enumerate(COLUNAS, start=1) if label == "Custo Total")
    for r in range(2, len(linhas) + 2):
        ws.cell(row=r, column=col_custo).number_format = "#,##0.00"

    ws.freeze_panes = "A2"
    if linhas:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS))}{len(linhas) + 1}"


def escrever_indicador(ws, linhas, ano):
    """Painel do ano: números do período, consulta interativa por mês/tipo,
    quebra por mês e por tipo, e como pedir o relatório por e-mail."""
    ws.sheet_view.showGridLines = False

    f_titulo = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    f_sub = Font(name="Arial", size=9, italic=True, color="555555")
    f_secao = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    f_rotulo = Font(name="Arial", size=9, bold=True, color="444444")
    f_txt = Font(name="Arial", size=10)
    f_bold = Font(name="Arial", size=10, bold=True)
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)

    for col, larg in zip("ABCDEFGH", [22, 14, 14, 4, 26, 14, 14, 14]):
        ws.column_dimensions[col].width = larg
    for col in ("T", "U", "V"):
        ws.column_dimensions[col].hidden = True

    def faixa(row, texto, cor):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        c = ws.cell(row=row, column=1, value=texto)
        c.font = f_secao
        c.fill = PatternFill("solid", fgColor=cor)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 22

    def bloco(row_rot, row_val, itens, cor_fundo, cor_valor, tamanho):
        for col, rotulo, formula, fmt in itens:
            col2 = get_column_letter(column_index_from_string(col) + 1)
            ws.merge_cells(f"{col}{row_rot}:{col2}{row_rot}")
            ws.merge_cells(f"{col}{row_val}:{col2}{row_val}")
            r = ws[f"{col}{row_rot}"]
            r.value, r.font = rotulo, f_rotulo
            r.alignment = Alignment(horizontal="center")
            r.fill = PatternFill("solid", fgColor=cor_fundo)
            v = ws[f"{col}{row_val}"]
            v.value = formula
            v.font = Font(name="Arial", size=tamanho, bold=True, color=cor_valor)
            v.number_format = fmt
            v.alignment = Alignment(horizontal="center")
            v.fill = PatternFill("solid", fgColor=cor_fundo)

    def coluna(label):
        return get_column_letter(next(i for i, (_id, lb, _f) in enumerate(COLUNAS, start=1) if lb == label))

    fim = len(linhas) + 1
    faixa_de = lambda label: f"Dados!${coluna(label)}$2:${coluna(label)}${fim}"
    R_TIPO, R_DATA = faixa_de("Tipo"), faixa_de("Data")
    R_IMP, R_AVAL = faixa_de("Impacto ao Cliente?"), faixa_de("Inserir na avaliação motorista?")
    R_CUSTO = faixa_de("Custo Total")

    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = f"Ocorrências - Anual {ano}"
    t.font = f_titulo
    t.fill = PatternFill("solid", fgColor=COR_INDICADOR)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    ws.merge_cells("A2:H2")
    ws["A2"] = f"Gerado em {datetime.now():%d/%m/%Y %H:%M} — {len(linhas)} chamado(s)"
    ws["A2"].font = f_sub
    ws["A2"].alignment = Alignment(horizontal="center")

    faixa(4, "PANORAMA DO ANO", COR_INDICADOR)
    bloco(5, 6, [
        ("A", "Total de ocorrências", f"=COUNTA({R_TIPO})", "0"),
        ("C", "Meses com registro", "=SUMPRODUCT((B17:B28>0)*1)", "0"),
        ("E", "Com impacto ao cliente", f'=COUNTIF({R_IMP},"Sim")', "0"),
        ("G", "Custo total lançado (R$)", f"=SUM({R_CUSTO})", "#,##0"),
    ], "F2F2F2", COR_INDICADOR, 20)
    ws.row_dimensions[6].height = 30

    ws.merge_cells("A7:H7")
    ws["A7"] = ('O custo considera apenas os registros com "Custo Total" preenchido. A listagem '
                "completa está na aba Dados (oculta: botão direito em uma aba > Reexibir).")
    ws["A7"].font = f_sub

    faixa(9, "CONSULTA INTERATIVA  —  escolha nas listas abaixo", COR_INTERATIVO)
    ws["A10"], ws["E10"] = "Mês:", "Tipo:"
    ws["A10"].font = ws["E10"].font = f_bold
    ws.merge_cells("F10:H10")
    ws["B10"], ws["F10"] = "(todos)", "(todos)"
    for cel in ("B10", "F10"):
        ws[cel].font = Font(name="Arial", size=10, bold=True, color="C00000")
        ws[cel].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[cel].border = borda
        ws[cel].alignment = Alignment(horizontal="center")

    meses_presentes = sorted({str(l.get(str(CAMPO_DATA_ABERTURA)))[:7] for l in linhas
                              if l.get(str(CAMPO_DATA_ABERTURA))})
    tipos = sorted({str(l.get("7")).strip() for l in linhas if l.get("7")})

    for i, m in enumerate(meses_presentes, start=2):
        ws[f"T{i}"] = m
    ws["U2"] = "(todos)"
    for i, m in enumerate(meses_presentes, start=3):
        ws[f"U{i}"] = MESES[int(m[5:7])]
    ws["V2"] = "(todos)"
    for i, tp in enumerate(tipos, start=3):
        ws[f"V{i}"] = tp

    dv_mes = DataValidation(type="list", formula1=f"=$U$2:$U${len(meses_presentes) + 2}", allow_blank=False)
    ws.add_data_validation(dv_mes)
    dv_mes.add("B10")
    dv_tipo = DataValidation(type="list", formula1=f"=$V$2:$V${len(tipos) + 2}", allow_blank=False)
    ws.add_data_validation(dv_tipo)
    dv_tipo.add("F10")

    # "(todos)" vira o curinga "*"; um mês vira o prefixo "AAAA-MM*" da data
    ws["T20"] = (f'=IF($B$10="(todos)","*",INDEX($T$2:$T${len(meses_presentes) + 1},'
                 f'MATCH($B$10,$U$3:$U${len(meses_presentes) + 2},0))&"*")')
    ws["T21"] = '=IF($F$10="(todos)","*",$F$10)'

    bloco(12, 13, [
        ("A", "Ocorrências no filtro", f"=COUNTIFS({R_DATA},$T$20,{R_TIPO},$T$21)", "0"),
        ("C", "Com impacto ao cliente", f'=COUNTIFS({R_DATA},$T$20,{R_TIPO},$T$21,{R_IMP},"Sim")', "0"),
        ("E", "Na avaliação do motorista", f'=COUNTIFS({R_DATA},$T$20,{R_TIPO},$T$21,{R_AVAL},"Sim")', "0"),
        ("G", "Custo no filtro (R$)", f"=SUMIFS({R_CUSTO},{R_DATA},$T$20,{R_TIPO},$T$21)", "#,##0"),
    ], "DCE6F1", COR_INTERATIVO, 18)
    ws.row_dimensions[13].height = 28

    faixa(15, "OCORRÊNCIAS POR MÊS", COR_INDICADOR)
    for cel, txt in (("A16", "Mês"), ("B16", "Qtde")):
        ws[cel] = txt
        ws[cel].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        ws[cel].fill = PatternFill("solid", fgColor="7F7F7F")
        ws[cel].alignment = Alignment(horizontal="center")
    for i, mes in enumerate(range(1, 13)):
        lin = 17 + i
        ws[f"A{lin}"] = MESES[mes]
        ws[f"B{lin}"] = f'=COUNTIF({R_DATA},"{ano}-{mes:02d}*")'
        for cel in (f"A{lin}", f"B{lin}"):
            ws[cel].font = f_txt
            ws[cel].border = borda
        ws[f"B{lin}"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add(
        "B17:B28", DataBarRule(start_type="num", start_value=0, end_type="max", color=COR_INDICADOR))

    g_mes = BarChart()
    g_mes.type, g_mes.legend = "col", None
    g_mes.title = "Ocorrências por mês"
    g_mes.height, g_mes.width = 7.5, 13
    g_mes.add_data(Reference(ws, min_col=2, min_row=16, max_row=28), titles_from_data=True)
    g_mes.set_categories(Reference(ws, min_col=1, min_row=17, max_row=28))
    ws.add_chart(g_mes, "E16")

    faixa(30, "OCORRÊNCIAS POR TIPO", COR_INDICADOR)
    for cel, txt in (("A31", "Tipo"), ("B31", "Qtde")):
        ws[cel] = txt
        ws[cel].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        ws[cel].fill = PatternFill("solid", fgColor="7F7F7F")
        ws[cel].alignment = Alignment(horizontal="center")

    contagem = Counter(str(l.get("7")).strip() for l in linhas if l.get("7"))
    top = [t for t, _ in contagem.most_common(10)]
    for i, tp in enumerate(top):
        lin = 32 + i
        ws[f"A{lin}"] = tp
        ws[f"A{lin}"].alignment = Alignment(wrap_text=True)
        ws[f"B{lin}"] = f"=COUNTIF({R_TIPO},A{lin})"
        ws[f"B{lin}"].alignment = Alignment(horizontal="center")
        for cel in (f"A{lin}", f"B{lin}"):
            ws[cel].font = f_txt
            ws[cel].border = borda
    fim_tipo = 31 + len(top)
    lin_demais = fim_tipo + 1
    ws[f"A{lin_demais}"] = "Demais tipos"
    ws[f"B{lin_demais}"] = f"=COUNTA({R_TIPO})-SUM(B32:B{fim_tipo})"
    for cel in (f"A{lin_demais}", f"B{lin_demais}"):
        ws[cel].font = Font(name="Arial", size=10, italic=True)
        ws[cel].border = borda
    ws[f"B{lin_demais}"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add(
        f"B32:B{lin_demais}", DataBarRule(start_type="num", start_value=0, end_type="max", color=COR_PEDIDO))

    if top:
        g_tipo = BarChart()
        g_tipo.type, g_tipo.legend = "bar", None
        g_tipo.title = "Top 10 tipos de ocorrência"
        g_tipo.height, g_tipo.width = 9, 13
        g_tipo.add_data(Reference(ws, min_col=2, min_row=31, max_row=fim_tipo), titles_from_data=True)
        g_tipo.set_categories(Reference(ws, min_col=1, min_row=32, max_row=fim_tipo))
        ws.add_chart(g_tipo, "E31")

    rodape = lin_demais + 2
    faixa(rodape, "COMO RECEBER ESTE RELATÓRIO POR E-MAIL", COR_PEDIDO)
    instrucoes = [
        (f"1. Envie um e-mail para {EMAIL_PEDIDO}", True),
        (f"2. Escreva no ASSUNTO: {PALAVRA_CHAVE_PEDIDO}", True),
        ("3. O relatório do ano corrente chega como anexo, em resposta ao próprio e-mail.", False),
        ("", False),
        ("O pedido precisa partir de um endereço autorizado. O corpo do e-mail pode ir vazio —", False),
        ("o que vale é a palavra-chave no assunto. Fora isso, o relatório também é enviado", False),
        ("automaticamente todo dia 1º de cada mês.", False),
    ]
    for i, (texto, forte) in enumerate(instrucoes):
        lin = rodape + 1 + i
        ws.merge_cells(start_row=lin, start_column=1, end_row=lin, end_column=8)
        c = ws.cell(row=lin, column=1, value=texto)
        c.font = Font(name="Arial", size=10, bold=forte, color=COR_INDICADOR if forte else "444444")
        c.alignment = Alignment(indent=1)


def montar_planilha(linhas_ano, ano, caminho_saida):
    wb = Workbook()
    wb.remove(wb.active)

    ws_anual = wb.create_sheet(f"Anual {ano}")

    grupos = agrupar_por_mes(linhas_ano)
    for mes in sorted(grupos):
        ws_mes = wb.create_sheet(f"{MESES[mes]} - {ano}")
        escrever_aba(ws_mes, grupos[mes])

    ws_dados = wb.create_sheet("Dados")
    escrever_dados(ws_dados, linhas_ano)
    ws_dados.sheet_state = "hidden"

    escrever_indicador(ws_anual, linhas_ano, ano)

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
