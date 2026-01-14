import dash
from dash import html, dcc, dash_table, Input, Output, State

import pandas as pd
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from pytz import timezone
import os


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/fiscais",
    name="Fiscais",
    title="Fiscais",
)


# --------------------------------------------------
# URL da planilha de Fiscais
# --------------------------------------------------
URL_FISCAIS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Fiscais"
)

# nomes originais no CSV (linha 4 é o cabeçalho)
COL_SETOR = "Setor"
COL_CONTRATO = "CONTRATO"
COL_OBJETO = "OBJETO"
COL_CONTRATADA = "CONTRATADA"
COL_FINAL_VIG = "Unnamed: 16"  # Final da Vigência
COL_LINK_COMPRASNET = "COMPRASNET Contratos"


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados_fiscais():
    df = pd.read_csv(URL_FISCAIS, skiprows=3, header=0)
    df.columns = [c.strip() for c in df.columns]

    # Colunas de servidores
    col_servidores_raw = [c for c in df.columns if c.startswith("SERVIDOR")]

    cols_keep = [
        COL_SETOR,
        COL_CONTRATO,
        COL_OBJETO,
        COL_CONTRATADA,
        COL_FINAL_VIG,
        COL_LINK_COMPRASNET,
    ] + col_servidores_raw

    df = df[cols_keep]

    df = df.rename(
        columns={
            COL_SETOR: "Setor",
            COL_CONTRATO: "Contrato",
            COL_OBJETO: "Objeto",
            COL_CONTRATADA: "Contratada",
            COL_FINAL_VIG: "Final da Vigência",
            COL_LINK_COMPRASNET: "Link Comprasnet",
        }
    )

    # Lista de servidores únicos (sem vazios) para o dropdown
    if col_servidores_raw:
        todos_serv = pd.Series(df[col_servidores_raw].values.ravel("K"), dtype="object")
        servidores_unicos = sorted(
            s.strip()
            for s in todos_serv.unique()
            if isinstance(s, str) and s.strip() != ""
        )
    else:
        servidores_unicos = []

    # Coluna agregada Servidores para exibição na tabela (sem nomes vazios)
    if col_servidores_raw:
        def junta_servidores(row):
            nomes = []
            for c in col_servidores_raw:
                v = row.get(c)
                if isinstance(v, str):
                    v = v.strip()
                else:
                    v = ""
                if v:
                    nomes.append(v)
            return "; ".join(nomes)

        df["Servidores"] = df.apply(junta_servidores, axis=1)
    else:
        df["Servidores"] = ""

    # Conversão e status pela Final da Vigência
    df["Final da Vigência"] = pd.to_datetime(
        df["Final da Vigência"], dayfirst=True, errors="coerce"
    )

    hoje = datetime.now().date()

    def calcular_status(data_final):
        if pd.isna(data_final):
            return ""
        dias = (data_final.date() - hoje).days
        if dias > 10:
            return "Vigente"
        if dias < 0:
            return "Vencido"
        return "Próximo do Vencimento"

    df["Status"] = df["Final da Vigência"].apply(calcular_status)

    # Formata datas para exibição
    df["Final da Vigência"] = df["Final da Vigência"].dt.strftime("%d/%m/%Y").fillna("")

    # guarda a lista de servidores únicos
    df._lista_servidores_unicos = servidores_unicos

    return df


df_fiscais_base = carregar_dados_fiscais()
SERVIDORES_UNICOS_FIS = getattr(df_fiscais_base, "_lista_servidores_unicos", [])


dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

# --------------------------------------------------
# Estilo dos botões (fundo azul, texto branco)
# --------------------------------------------------
botao_style = {
    "backgroundColor": "#0b2b57",
    "color": "white",
    "padding": "8px 16px",
    "border": "none",
    "borderRadius": "4px",
    "cursor": "pointer",
    "fontSize": "12px",
    "fontWeight": "bold",
    "marginRight": "6px",
}


# --------------------------------------------------
# Função auxiliar: filtros em cascata independentes
# --------------------------------------------------
def filtrar_fiscais(
    servidores_texto,
    servidores_drop,
    contrato_texto,
    contrato_drop,
    contratada_texto,
    contratada_drop,
    status,
):
    dff = df_fiscais_base.copy()

    # Servidores (digitação)
    if servidores_texto and str(servidores_texto).strip():
        termo = str(servidores_texto).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Servidores (dropdown)
    if servidores_drop:
        termo = str(servidores_drop).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Contrato (texto)
    if contrato_texto and str(contrato_texto).strip():
        termo = str(contrato_texto).strip().lower()
        dff = dff[
            dff["Contrato"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Contrato (dropdown)
    if contrato_drop:
        dff = dff[dff["Contrato"] == contrato_drop]

    # Contratada (texto)
    if contratada_texto and str(contratada_texto).strip():
        termo = str(contratada_texto).strip().lower()
        dff = dff[
            dff["Contratada"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Contratada (dropdown)
    if contratada_drop:
        dff = dff[dff["Contratada"] == contratada_drop]

    # Status
    if status:
        dff = dff[dff["Status"] == status]

    # remove linhas sem texto em Status
    dff = dff[dff["Status"].astype(str).str.strip() != ""]

    # Ordena por Final da Vigência (mais recente em cima)
    dff["_fim_vig_dt"] = pd.to_datetime(
        dff["Final da Vigência"], dayfirst=True, errors="coerce"
    )
    dff = dff.sort_values("_fim_vig_dt", ascending=False).drop(
        columns=["_fim_vig_dt"]
    )

    return dff


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_fiscais",
            className="filtros-sticky",
            children=[
                # Linha 1: Servidores + Contrato (texto e dropdown)
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        # Servidores (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores (digitação)"),
                                dcc.Input(
                                    id="filtro_servidores_texto_fis",
                                    type="text",
                                    placeholder="Digite parte do nome do servidor",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Servidores (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores"),
                                dcc.Dropdown(
                                    id="filtro_servidores_dropdown_fis",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in SERVIDORES_UNICOS_FIS
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Contrato (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato (digitação)"),
                                dcc.Input(
                                    id="filtro_contrato_texto_fis",
                                    type="text",
                                    placeholder="Digite parte do contrato",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Contrato (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato"),
                                dcc.Dropdown(
                                    id="filtro_contrato_dropdown_fis",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in sorted(
                                            df_fiscais_base["Contrato"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(c).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                    ],
                ),
                # Linha 2: Contratada (texto/drop), Status, botões
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-end",
                        "marginTop": "4px",
                    },
                    children=[
                        # Contratada (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contratada (digitação)"),
                                dcc.Input(
                                    id="filtro_contratada_texto_fis",
                                    type="text",
                                    placeholder="Digite parte da contratada",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Contratada (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contratada"),
                                dcc.Dropdown(
                                    id="filtro_contratada_dropdown_fis",
                                    options=[
                                        {"label": e, "value": e}
                                        for e in sorted(
                                            df_fiscais_base["Contratada"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(e).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Status
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="filtro_status_fis",
                                    options=[
                                        {"label": "Vigente", "value": "Vigente"},
                                        {
                                            "label": "Próximo do Vencimento",
                                            "value": "Próximo do Vencimento",
                                        },
                                        {"label": "Vencido", "value": "Vencido"},
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Botões
                        html.Div(
                            style={
                                "display": "flex",
                                "gap": "10px",
                                "flexShrink": 0,
                            },
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_fis",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_fis",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                dcc.Download(id="download_relatorio_fis"),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        dash_table.DataTable(
            id="tabela_fiscais",
            columns=[
                {
                    "name": "Contrato",
                    "id": "Contrato_markdown",
                    "presentation": "markdown",
                },
                {"name": "Setor", "id": "Setor"},
                {"name": "Objeto", "id": "Objeto"},
                {"name": "Contratada", "id": "Contratada"},
                {"name": "Final da Vigência", "id": "Final da Vigência"},
                {"name": "Servidores", "id": "Servidores"},
                {"name": "Status", "id": "Status"},
            ],
            data=[],
            row_selectable=False,
            cell_selectable=False,
            style_table={
                "overflowX": "auto",
                "overflowY": "auto",
                "height": "calc(100vh - 200px)",
                "minHeight": "300px",
                "position": "relative",
            },
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
                "minWidth": "80px",
                "maxWidth": "260px",
                "whiteSpace": "normal",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "textAlign": "center",
                "position": "sticky",
                "top": 0,
                "zIndex": 5,
            },
            style_cell_conditional=[
                {
                    "if": {"column_id": "Contrato_markdown"},
                    "textAlign": "center",
                },
            ],
            style_data_conditional=[
                # Zebra: linhas pares/ímpares
                {
                    "if": {"row_index": "odd"},
                    "backgroundColor": "#f0f0f0",
                },
                {
                    "if": {"row_index": "even"},
                    "backgroundColor": "white",
                },
                # Status = Vencido
                {
                    "if": {
                        "filter_query": '{Status} = "Vencido"',
                    },
                    "backgroundColor": "#ffcccc",
                    "color": "black",
                },
                # Status = Próximo do Vencimento
                {
                    "if": {
                        "filter_query": '{Status} = "Próximo do Vencimento"',
                    },
                    "backgroundColor": "#ffffcc",
                    "color": "black",
                },
            ],
            css=[
                {
                    "selector": "p",
                    "rule": "margin: 0; text-align: center;",
                },
            ],
        ),
        dcc.Store(id="store_dados_fis"),
    ]
)


# --------------------------------------------------
# Callback: filtros (tabela + store)
# --------------------------------------------------
@dash.callback(
    Output("tabela_fiscais", "data"),
    Output("store_dados_fis", "data"),
    Input("filtro_servidores_texto_fis", "value"),
    Input("filtro_servidores_dropdown_fis", "value"),
    Input("filtro_contrato_texto_fis", "value"),
    Input("filtro_contrato_dropdown_fis", "value"),
    Input("filtro_contratada_texto_fis", "value"),
    Input("filtro_contratada_dropdown_fis", "value"),
    Input("filtro_status_fis", "value"),
)
def atualizar_tabela_fiscais(
    servidores_texto,
    servidores_drop,
    contrato_texto,
    contrato_drop,
    contratada_texto,
    contratada_drop,
    status,
):
    dff = filtrar_fiscais(
        servidores_texto,
        servidores_drop,
        contrato_texto,
        contrato_drop,
        contratada_texto,
        contratada_drop,
        status,
    )

    dff = dff.copy()

    def mk_link(row):
        url = row.get("Link Comprasnet")
        contrato = row.get("Contrato")
        if isinstance(url, str) and url.strip() and isinstance(contrato, str):
            return f"[{contrato}]({url.strip()})"
        return ""

    dff["Contrato_markdown"] = dff.apply(mk_link, axis=1)
    dff = dff[dff["Contrato_markdown"].str.strip() != ""]

    cols = [
        "Contrato_markdown",
        "Setor",
        "Objeto",
        "Contratada",
        "Final da Vigência",
        "Servidores",
        "Status",
    ]
    cols = [c for c in cols if c in dff.columns]

    return dff[cols].to_dict("records"), dff.to_dict("records")


# --------------------------------------------------
# Callback: opções dos filtros (cascata)
# --------------------------------------------------
@dash.callback(
    Output("filtro_servidores_dropdown_fis", "options"),
    Output("filtro_contrato_dropdown_fis", "options"),
    Output("filtro_contratada_dropdown_fis", "options"),
    Input("filtro_servidores_texto_fis", "value"),
    Input("filtro_servidores_dropdown_fis", "value"),
    Input("filtro_contrato_texto_fis", "value"),
    Input("filtro_contrato_dropdown_fis", "value"),
    Input("filtro_contratada_texto_fis", "value"),
    Input("filtro_contratada_dropdown_fis", "value"),
    Input("filtro_status_fis", "value"),
)
def atualizar_opcoes_filtros_fis(
    servidores_texto,
    servidores_drop,
    contrato_texto,
    contrato_drop,
    contratada_texto,
    contratada_drop,
    status,
):
    dff = filtrar_fiscais(
        servidores_texto,
        servidores_drop,
        contrato_texto,
        contrato_drop,
        contratada_texto,
        contratada_drop,
        status,
    )

    # Servidores: extrai únicos da coluna agregada "Servidores"
    servidores_list = []
    for serv_str in dff["Servidores"].unique():
        if isinstance(serv_str, str) and serv_str.strip():
            for s in serv_str.split(";"):
                s = s.strip()
                if s and s not in servidores_list:
                    servidores_list.append(s)
    servidores_list.sort()

    op_servidores = [{"label": s, "value": s} for s in servidores_list]
    op_contrato = [
        {"label": c, "value": c}
        for c in sorted(dff["Contrato"].dropna().unique())
        if str(c).strip()
    ]
    op_contratada = [
        {"label": e, "value": e}
        for e in sorted(dff["Contratada"].dropna().unique())
        if str(e).strip()
    ]

    return op_servidores, op_contrato, op_contratada


# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_servidores_texto_fis", "value"),
    Output("filtro_servidores_dropdown_fis", "value"),
    Output("filtro_contrato_texto_fis", "value"),
    Output("filtro_contrato_dropdown_fis", "value"),
    Output("filtro_contratada_texto_fis", "value"),
    Output("filtro_contratada_dropdown_fis", "value"),
    Output("filtro_status_fis", "value"),
    Input("btn_limpar_filtros_fis", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_fis(n):
    return None, None, None, None, None, None, None


# --------------------------------------------------
# PDF – estilos
# --------------------------------------------------
wrap_style_data = ParagraphStyle(
    name="wrap_fiscais_data",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
    textColor=colors.black,
)

wrap_style_header = ParagraphStyle(
    name="wrap_fiscais_header",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
    textColor=colors.white,
)


def wrap_data(text):
    return Paragraph(str(text), wrap_style_data)


def wrap_header(text):
    return Paragraph(str(text), wrap_style_header)


# --------------------------------------------------
# Callback: gerar PDF de fiscais (padrão unificado)
# --------------------------------------------------
@dash.callback(
    Output("download_relatorio_fis", "data"),
    Input("btn_download_relatorio_fis", "n_clicks"),
    State("store_dados_fis", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_fiscais(n, dados_fis):
    if not n or not dados_fis:
        return None

    df = pd.DataFrame(dados_fis)

    buffer = BytesIO()
    pagesize = landscape(A4)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.3 * inch,
        leftMargin=0.3 * inch,
        topMargin=1.3 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    # Data / Hora (topo direito)
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")

    story.append(
        Table(
            [[
                Paragraph(
                    data_hora,
                    ParagraphStyle(
                        "data_topo_fiscais",
                        fontSize=9,
                        alignment=TA_RIGHT,
                        textColor="#333333",
                    ),
                )
            ]],
            colWidths=[pagesize[0] - 0.6 * inch],
        )
    )
    story.append(Spacer(1, 0.15 * inch))

    # Cabeçalho: Logo esq | Instituição | Logo dir
    logo_esq = (
        Image("assets/brasaobrasil.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/brasaobrasil.png")
        else ""
    )

    logo_dir = (
        Image("assets/simbolo_RGB.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/simbolo_RGB.png")
        else ""
    )

    texto_instituicao = (
        "<b><font color='#0b2b57' size=13>Ministério da Educação</font></b><br/>"
        "<b><font color='#0b2b57' size=13>Universidade Federal de Itajubá</font></b><br/>"
        "<font color='#0b2b57' size=11>Diretoria de Compras e Contratos</font>"
    )

    instituicao = Paragraph(
        texto_instituicao,
        ParagraphStyle(
            "instituicao_fiscais",
            alignment=TA_CENTER,
            leading=16,
        ),
    )

    cabecalho = Table(
        [[logo_esq, instituicao, logo_dir]],
        colWidths=[
            1.4 * inch,
            4.2 * inch,
            1.4 * inch,
        ],
    )

    cabecalho.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    story.append(cabecalho)
    story.append(Spacer(1, 0.25 * inch))

    # Título
    titulo = Paragraph(
        "RELATÓRIO DE FISCAIS DE CONTRATOS<br/>",
        ParagraphStyle(
            "titulo_fiscais",
            alignment=TA_CENTER,
            fontSize=10,
            leading=14,
            textColor=colors.black,
        ),
    )

    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    story.append(
        Paragraph(f"Total de registros: {len(df)}", styles["Normal"])
    )
    story.append(Spacer(1, 0.15 * inch))

    # Colunas do PDF
    cols = [
        "Setor",
        "Contrato",
        "Objeto",
        "Contratada",
        "Final da Vigência",
        "Servidores",
        "Status",
    ]

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    df_pdf = df.copy()

    header = [wrap_header(c) for c in cols]
    table_data = [header]

    status_values = df["Status"].fillna("").tolist()

    for _, row in df_pdf[cols].iterrows():
        linha = [wrap_data(row[c]) for c in cols]
        table_data.append(linha)

    # Larguras de coluna (ajustadas)
    page_width = pagesize[0] - 0.6 * inch
    col_widths = [
        0.75 * inch,  # Setor
        0.85 * inch,  # Contrato
        2.3 * inch,   # Objeto
        1.9 * inch,   # Contratada
        0.9 * inch,   # Final da Vigência
        1.9 * inch,   # Servidores
        1.0 * inch,   # Status
    ]

    tbl = Table(table_data, colWidths=col_widths[: len(cols)], repeatRows=1)

    table_styles = [
        # Header
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
    ]

    # Cores por status (linha inteira)
    for i, status in enumerate(status_values, 1):
        status_str = str(status).strip().lower()
        if "vencido" in status_str:
            table_styles.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#ffcccc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#cc0000"))
            )
        elif "próximo do vencimento" in status_str or "proximo do vencimento" in status_str:
            table_styles.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#ffffcc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#cc8800"))
            )

    tbl.setStyle(TableStyle(table_styles))
    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(buffer.getvalue(), "fiscais.pdf")
