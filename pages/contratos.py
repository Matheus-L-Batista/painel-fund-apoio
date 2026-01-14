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
        dff = dff[
            dff["Contrato"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Contrato (dropdown)
    if contrato_drop:
        dff = dff[dff["Contrato"] == contrato_drop]

    # Setor (texto)
    if setor_texto and str(setor_texto).strip():
        termo = str(setor_texto).strip().lower()
        dff = dff["Setor"].astype(str).str.lower().str.contains(termo, na=False)
        dff = df_contratos_base.loc[dff]

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
                                    style=botao_style,
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_contratos",
                                    n_clicks=0,
                                    style=botao_style,
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
                # zebra: linhas pares cinza claro, ímpares branco
                {
                    "if": {"row_index": "even"},
                    "backgroundColor": "#f5f5f5",
                    "color": "black",
                },
                {
                    "if": {"row_index": "odd"},
                    "backgroundColor": "white",
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
# Estilos para PDF
# --------------------------------------------------
wrap_style_data = ParagraphStyle(
    name="wrap_contratos_data",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
    textColor=colors.black,
)

wrap_style_header = ParagraphStyle(
    name="wrap_contratos_header",
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
# Callback: gerar PDF de contratos
# --------------------------------------------------
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
        rightMargin=0.3 * inch,
        leftMargin=0.3 * inch,
        topMargin=1.3 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    # Data / Hora
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")

    story.append(
        Table(
            [
                [
                    Paragraph(
                        data_hora,
                        ParagraphStyle(
                            "data_topo_contratos",
                            fontSize=9,
                            alignment=TA_RIGHT,
                            textColor="#333333",
                        ),
                    )
                ]
            ],
            colWidths=[pagesize[0] - 0.6 * inch],
        )
    )
    story.append(Spacer(1, 0.15 * inch))

    # Cabeçalho: Logo | Texto | Logo
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
        "<b><font color='#0b2b57' size=13>Universidade Federal de Itajubá</font></b><br/>"
        "<font color='#0b2b57' size=11>Diretoria de Compras e Contratos</font>"
    )

    instituicao = Paragraph(
        texto_instituicao,
        ParagraphStyle(
            "instituicao",
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
        "RELATÓRIO DE CONTRATOS ATIVOS - UASG: 153030 - Campus Itajubá<br/>",
        ParagraphStyle(
            "titulo_contratos",
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

    # Tabela
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

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    df_pdf = df.copy()

    header = [wrap_header(c) for c in cols]
    table_data = [header]

    status_values = df["Status da Vigência"].fillna("").tolist()

    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap_data(row[c]) for c in cols])

    page_width = pagesize[0] - 0.6 * inch
    col_widths = [
        0.8 * inch,   # Contrato
        0.9 * inch,   # Setor
        0.9 * inch,   # Grupo
        2.2 * inch,   # Objeto
        1.8 * inch,   # Empresa Contratada
        1.0 * inch,   # Início da Vigência
        1.1 * inch,   # Término da Execução
        1.1 * inch,   # Término da Vigência
        1.1 * inch,   # Status da Vigência
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    table_styles = [
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
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
    ]

    # Cores por status
    for i, status in enumerate(status_values, 1):
        status_str = str(status).strip().lower()
        if "vencido" in status_str:
            table_styles.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#ffcccc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#cc0000"))
            )
        elif "próximo do vencimento" in status_str:
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

    return dcc.send_bytes(buffer.getvalue(), "relatorio_contratos.pdf")
