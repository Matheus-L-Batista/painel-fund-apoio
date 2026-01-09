import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from datetime import datetime

from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
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
    path="/contratos",
    name="Contratos",
    title="Contratos",
)

# --------------------------------------------------
# URL da planilha de Contratos
# --------------------------------------------------
URL_CONTRATOS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Grupo%20da%20Cont."
)

# nomes exatos das colunas originais no CSV
COL_CONTRATO = "Contrato"
COL_SETOR = "Setor"
COL_MENU_GRUPO = "MENU Grupo"
COL_OBJETO_ORIG = (
    "UNIVERSIDADE FEDERAL DE ITAJUBÁ Diretoria de Compras e Contratos "
    "Campus Itajubá CONTRATOS ATIVOS - ALIMENTAÇÃO DO BI Objeto"
)
COL_EMPRESA = "Empresa Contratada"
COL_INICIO_VIG = "Início da Vigência"
COL_TERMINO_EXEC = "Término da Execução"
COL_TERMINO_VIG = "Termino da Vigência"  # igual na planilha
COL_LINK_COMPRASNET = "Comprasnet Contratos"

# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados_contratos():
    df = pd.read_csv(URL_CONTRATOS, header=0)
    df.columns = [c.strip() for c in df.columns]

    if COL_LINK_COMPRASNET not in df.columns:
        df[COL_LINK_COMPRASNET] = ""

    df = df.rename(
        columns={
            COL_CONTRATO: "Contrato",
            COL_SETOR: "Setor",
            COL_MENU_GRUPO: "Grupo",
            COL_OBJETO_ORIG: "Objeto",
            COL_EMPRESA: "Empresa Contratada",
            COL_INICIO_VIG: "Início da Vigência",
            COL_TERMINO_EXEC: "Término da Execução",
            COL_TERMINO_VIG: "Término da Vigência",
            COL_LINK_COMPRASNET: "Link Comprasnet",
        }
    )

    # Converte datas para datetime para cálculo do status
    for col in ["Início da Vigência", "Término da Execução", "Término da Vigência"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    hoje = datetime.now().date()

    def calcular_status(data_termino_exec):
        if pd.isna(data_termino_exec):
            return ""
        dias = (data_termino_exec.date() - hoje).days
        if dias > 10:
            return "Vigente"
        if dias < 0:
            return "Vencido"
        return "Próximo do Vencimento"

    df["Status da Vigência"] = df["Término da Execução"].apply(calcular_status)

    # Formata datas para string dd/mm/aaaa para exibição
    for col in ["Início da Vigência", "Término da Execução", "Término da Vigência"]:
        if col in df.columns:
            df[col] = df[col].dt.strftime("%d/%m/%Y").fillna("")

    return df

df_contratos_base = carregar_dados_contratos()

# --------------------------------------------------
# Função auxiliar: filtros em cascata independentes
# --------------------------------------------------
def filtrar_contratos(
    contrato_texto,
    contrato_drop,
    setor_texto,
    setor_drop,
    grupo,
    empresa_texto,
    empresa_drop,
    status_vig,
):
    dff = df_contratos_base.copy()

    # Contrato (texto)
    if contrato_texto and str(contrato_texto).strip():
        termo = str(contrato_texto).strip().lower()
        dff = dff[dff["Contrato"].astype(str).str.lower().str.contains(termo, na=False)]

    # Contrato (dropdown)
    if contrato_drop:
        dff = dff[dff["Contrato"] == contrato_drop]

    # Setor (texto)
    if setor_texto and str(setor_texto).strip():
        termo = str(setor_texto).strip().lower()
        dff = dff[dff["Setor"].astype(str).str.lower().str.contains(termo, na=False)]

    # Setor (dropdown)
    if setor_drop:
        dff = dff[dff["Setor"] == setor_drop]

    # Grupo
    if grupo:
        dff = dff[dff["Grupo"] == grupo]

    # Empresa (texto)
    if empresa_texto and str(empresa_texto).strip():
        termo = str(empresa_texto).strip().lower()
        dff = dff[
            dff["Empresa Contratada"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Empresa (dropdown)
    if empresa_drop:
        dff = dff[dff["Empresa Contratada"] == empresa_drop]

    # Status
    if status_vig:
        dff = dff[dff["Status da Vigência"] == status_vig]

    # remove linhas sem texto em Status da Vigência
    dff = dff[dff["Status da Vigência"].astype(str).str.strip() != ""]

    # Ordena por Término da Execução (mais recente em cima)
    dff["_termino_exec_dt"] = pd.to_datetime(
        dff["Término da Execução"], dayfirst=True, errors="coerce"
    )
    dff = dff.sort_values("_termino_exec_dt", ascending=False).drop(
        columns=["_termino_exec_dt"]
    )

    return dff

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_contratos",
            className="filtros-sticky",
            children=[
                # Linha 1: Contrato + Setor (texto e dropdown)
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato (digitação)"),
                                dcc.Input(
                                    id="filtro_contrato_texto",
                                    type="text",
                                    placeholder="Digite parte do contrato",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato"),
                                dcc.Dropdown(
                                    id="filtro_contrato_dropdown",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in sorted(
                                            df_contratos_base["Contrato"]
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
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Setor (digitação)"),
                                dcc.Input(
                                    id="filtro_setor_texto_ct",
                                    type="text",
                                    placeholder="Digite parte do setor",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Setor"),
                                dcc.Dropdown(
                                    id="filtro_setor_dropdown_ct",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in sorted(
                                            df_contratos_base["Setor"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(s).strip() != ""
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
                # Linha 2: Empresa (texto/drop), Grupo, Status, botões
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-end",
                        "marginTop": "4px",
                    },
                    children=[
                        # Empresa (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Empresa (digitação)"),
                                dcc.Input(
                                    id="filtro_empresa_texto",
                                    type="text",
                                    placeholder="Digite parte do nome da empresa",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Empresa (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Empresa Contratada"),
                                dcc.Dropdown(
                                    id="filtro_empresa",
                                    options=[
                                        {"label": e, "value": e}
                                        for e in sorted(
                                            df_contratos_base["Empresa Contratada"]
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
                        # Grupo
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Grupo"),
                                dcc.Dropdown(
                                    id="filtro_grupo",
                                    options=[
                                        {"label": g, "value": g}
                                        for g in sorted(
                                            df_contratos_base["Grupo"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(g).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Status da Vigência
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Status da Vigência"),
                                dcc.Dropdown(
                                    id="filtro_status_vig",
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
                                    id="btn_limpar_filtros_contratos",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_contratos",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                dcc.Download(id="download_relatorio_contratos"),
                            ],
                        ),
                    ],
                ),
            ],
        ),

        dash_table.DataTable(
            id="tabela_contratos",
            columns=[
                {
                    "name": "Contrato",
                    "id": "Contrato_markdown",
                    "presentation": "markdown",
                },
                {"name": "Setor", "id": "Setor"},
                {"name": "Grupo", "id": "Grupo"},
                {"name": "Objeto", "id": "Objeto"},
                {"name": "Empresa Contratada", "id": "Empresa Contratada"},
                {"name": "Início da Vigência", "id": "Início da Vigência"},
                {"name": "Término da Execução", "id": "Término da Execução"},
                {"name": "Término da Vigência", "id": "Término da Vigência"},
                {"name": "Status da Vigência", "id": "Status da Vigência"},
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
                {
                    "if": {
                        "filter_query": '{Status da Vigência} = "Vencido"'
                    },
                    "backgroundColor": "#ffcccc",
                    "color": "black",
                },
                {
                    "if": {
                        "filter_query": '{Status da Vigência} = "Próximo do Vencimento"'
                    },
                    "backgroundColor": "#ffffcc",
                    "color": "black",
                },
            ],
            css=[
                dict(
                    selector="p",
                    rule="margin: 0; text-align: center;",
                ),
            ],
        ),
        dcc.Store(id="store_dados_contratos"),
    ]
)

# --------------------------------------------------
# Callback: filtros (tabela + store)
# --------------------------------------------------
@dash.callback(
    Output("tabela_contratos", "data"),
    Output("store_dados_contratos", "data"),
    Input("filtro_contrato_texto", "value"),
    Input("filtro_contrato_dropdown", "value"),
    Input("filtro_setor_texto_ct", "value"),
    Input("filtro_setor_dropdown_ct", "value"),
    Input("filtro_grupo", "value"),
    Input("filtro_empresa_texto", "value"),
    Input("filtro_empresa", "value"),
    Input("filtro_status_vig", "value"),
)
def atualizar_tabela_contratos(
    contrato_texto,
    contrato_drop,
    setor_texto,
    setor_drop,
    grupo,
    empresa_texto,
    empresa_drop,
    status_vig,
):
    dff = filtrar_contratos(
        contrato_texto,
        contrato_drop,
        setor_texto,
        setor_drop,
        grupo,
        empresa_texto,
        empresa_drop,
        status_vig,
    )

    def mk_link(row):
        url = row.get("Link Comprasnet")
        contrato = row.get("Contrato")
        if isinstance(url, str) and url.strip() and isinstance(contrato, str):
            return f"[{contrato}]({url.strip()})"
        return ""

    dff = dff.copy()
    dff["Contrato_markdown"] = dff.apply(mk_link, axis=1)
    dff = dff[dff["Contrato_markdown"].str.strip() != ""]

    cols = [
        "Contrato_markdown",
        "Setor",
        "Grupo",
        "Objeto",
        "Empresa Contratada",
        "Início da Vigência",
        "Término da Execução",
        "Término da Vigência",
        "Status da Vigência",
    ]
    cols = [c for c in cols if c in dff.columns]

    return dff[cols].to_dict("records"), dff.to_dict("records")

# --------------------------------------------------
# Callback: opções dos filtros (cascata)
# --------------------------------------------------
@dash.callback(
    Output("filtro_contrato_dropdown", "options"),
    Output("filtro_setor_dropdown_ct", "options"),
    Output("filtro_empresa", "options"),
    Output("filtro_grupo", "options"),
    Input("filtro_contrato_texto", "value"),
    Input("filtro_contrato_dropdown", "value"),
    Input("filtro_setor_texto_ct", "value"),
    Input("filtro_setor_dropdown_ct", "value"),
    Input("filtro_grupo", "value"),
    Input("filtro_empresa_texto", "value"),
    Input("filtro_empresa", "value"),
    Input("filtro_status_vig", "value"),
)
def atualizar_opcoes_filtros(
    contrato_texto,
    contrato_drop,
    setor_texto,
    setor_drop,
    grupo,
    empresa_texto,
    empresa_drop,
    status_vig,
):
    dff = filtrar_contratos(
        contrato_texto,
        contrato_drop,
        setor_texto,
        setor_drop,
        grupo,
        empresa_texto,
        empresa_drop,
        status_vig,
    )

    op_contrato = [
        {"label": c, "value": c}
        for c in sorted(dff["Contrato"].dropna().unique())
        if str(c).strip()
    ]
    op_setor = [
        {"label": s, "value": s}
        for s in sorted(dff["Setor"].dropna().unique())
        if str(s).strip()
    ]
    op_empresa = [
        {"label": e, "value": e}
        for e in sorted(dff["Empresa Contratada"].dropna().unique())
        if str(e).strip()
    ]
    op_grupo = [
        {"label": g, "value": g}
        for g in sorted(dff["Grupo"].dropna().unique())
        if str(g).strip()
    ]

    return op_contrato, op_setor, op_empresa, op_grupo

# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_contrato_texto", "value"),
    Output("filtro_contrato_dropdown", "value"),
    Output("filtro_setor_texto_ct", "value"),
    Output("filtro_setor_dropdown_ct", "value"),
    Output("filtro_grupo", "value"),
    Output("filtro_empresa_texto", "value"),
    Output("filtro_empresa", "value"),
    Output("filtro_status_vig", "value"),
    Input("btn_limpar_filtros_contratos", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_contratos(n):
    return None, None, None, None, None, None, None, None

# --------------------------------------------------
# Callback: gerar PDF de contratos
# --------------------------------------------------
wrap_style = ParagraphStyle(
    name="wrap_contratos",
    fontSize=7,
    leading=8,
    spaceAfter=2,
    wordWrap="CJK",  # Quebra agressiva de palavras
)

simple_style = ParagraphStyle(
    name="simple_contratos",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

def simple(text):
    return Paragraph(str(text), simple_style)

@dash.callback(
    Output("download_relatorio_contratos", "data"),
    Input("btn_download_relatorio_contratos", "n_clicks"),
    State("store_dados_contratos", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_contratos(n, dados_contratos):
    if not n or not dados_contratos:
        return None

    df = pd.DataFrame(dados_contratos)

    buffer = BytesIO()
    pagesize = landscape(A4)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.15 * inch,  # Reduzido para ganhar espaço
        leftMargin=0.15 * inch,   # Reduzido para ganhar espaço
        topMargin=1.3 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    # Data e hora (topo direito)
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora_brasilia = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")
    data_top_table = Table(
        [
            [
                Paragraph(
                    data_hora_brasilia,
                    ParagraphStyle(
                        "data_topo_contratos",
                        fontSize=9,
                        alignment=TA_RIGHT,
                        textColor="#333333",
                    ),
                )
            ]
        ],
        colWidths=[pagesize[0] - 0.3 * inch],
    )
    data_top_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(data_top_table)
    story.append(Spacer(1, 0.1 * inch))

    # Logos
    logos_path = []
    if os.path.exists(os.path.join("assets", "brasaobrasil.png")):
        logos_path.append(os.path.join("assets", "brasaobrasil.png"))
    if os.path.exists(os.path.join("assets", "simbolo_RGB.png")):
        logos_path.append(os.path.join("assets", "simbolo_RGB.png"))

    if logos_path:
        logos = []
        for logo_file in logos_path:
            if os.path.exists(logo_file):
                logo = Image(logo_file, width=1.2 * inch, height=1.2 * inch)
                logos.append(logo)

        if logos:
            if len(logos) == 2:
                logo_table = Table(
                    [[logos[0], logos[1]]],
                    colWidths=[
                        pagesize[0] / 2 - 0.15 * inch,
                        pagesize[0] / 2 - 0.15 * inch,
                    ],
                )
            else:
                logo_table = Table(
                    [[logos[0]]],
                    colWidths=[pagesize[0] - 0.3 * inch],
                )

            logo_table.setStyle(
                TableStyle(
                    [
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ]
                )
            )
            story.append(logo_table)
            story.append(Spacer(1, 0.15 * inch))

    # Título
    titulo_texto = "RELATÓRIO DE CONTRATOS\n"
    titulo_paragraph = Paragraph(
        titulo_texto,
        ParagraphStyle(
            "titulo_contratos",
            fontSize=10,
            alignment=TA_CENTER,
            textColor="#0b2b57",
            spaceAfter=4,
            leading=14,
        ),
    )
    titulo_table = Table(
        [[titulo_paragraph]],
        colWidths=[pagesize[0] - 0.3 * inch],
    )
    titulo_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(titulo_table)
    story.append(Spacer(1, 0.15 * inch))

    # Quantidade de registros
    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.1 * inch))

    # Preparação da tabela de dados
    cols = [
        "Contrato",
        "Setor",
        "Grupo",
        "Objeto",
        "Empresa Contratada",
        "Início da Vigência",
        "Término da Execução",
        "Término da Vigência",
        "Status da Vigência",
    ]
    cols = [c for c in cols if c in df.columns]

    df_pdf = df.copy()

    header = cols
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        linha = []
        for c in cols:
            valor = str(row[c]).strip()

            if c in ["Objeto", "Empresa Contratada"]:
                linha.append(wrap(valor))
            elif c in [
                "Início da Vigência",
                "Término da Execução",
                "Término da Vigência",
                "Status da Vigência",
            ]:
                linha.append(simple(valor))
            else:
                linha.append(simple(valor))

        table_data.append(linha)

    # Larguras das colunas
    col_widths = [
        0.7 * inch,   # Contrato
        0.75 * inch,  # Setor
        0.85 * inch,  # Grupo
        2.3 * inch,   # Objeto
        1.9 * inch,   # Empresa Contratada
        0.9 * inch,   # Início da Vigência
        1.2 * inch,   # Término da Execução
        1.2 * inch,   # Término da Vigência
        1.2 * inch,   # Status da Vigência
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_list = [
        # Header
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
        # Grid
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        # Dados
        ("ALIGN", (0, 1), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        # Quebra de palavras
        ("WORDWRAP", (0, 0), (-1, -1), True),
        # Padding
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        # Zebra
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
    ]

    # Cores por status
    status_col_idx = cols.index("Status da Vigência") if "Status da Vigência" in cols else -1

    if status_col_idx != -1:
        for row_idx, row_data in enumerate(table_data[1:], start=1):
            status_value = str(row_data[status_col_idx]).lower()

            if "vencido" in status_value:
                style_list.append(
                    ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#ffcccc"))
                )
                style_list.append(
                    ("TEXTCOLOR", (0, row_idx), (-1, row_idx), colors.HexColor("#cc0000"))
                )
            elif "próximo do vencimento" in status_value:
                style_list.append(
                    ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#ffffcc"))
                )
                style_list.append(
                    ("TEXTCOLOR", (0, row_idx), (-1, row_idx), colors.HexColor("#cc8800"))
                )

    tbl.setStyle(TableStyle(style_list))
    story.append(tbl)
    doc.build(story)
    buffer.seek(0)

    from dash import dcc
    return dcc.send_bytes(buffer.getvalue(), "contratos_paisagem.pdf")
