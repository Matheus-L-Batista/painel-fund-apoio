import dash

from dash import html, dcc, dash_table, Input, Output, State, no_update

import pandas as pd

from datetime import date, datetime
from pytz import timezone

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

import os


# --------------------------------------------------
# Registro da página
# --------------------------------------------------

dash.register_page(
    __name__,
    path="/fracionamento_pdm",
    name="Fracionamento de Despesas PDM",
    title="Fracionamento de Despesas PDM",
)


# --------------------------------------------------
# URL da planilha
# --------------------------------------------------

URL_LIMITE_GASTO_ITA = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=Limite%20de%20Gasto%20-%20Itajub%C3%A1"
)

COL_PDM = "PDM"
COL_DESC_ORIG = "Descrição.1"
COL_VALOR_EMPENHADO_ORIG = "Unnamed: 7"

DATA_HOJE = date.today().strftime("%d/%m/%Y")

# Limite da dispensa 2026
VALOR_LIMITE_2026 = 65492.11


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------


def carregar_dados_limite_pdm():
    df = pd.read_csv(URL_LIMITE_GASTO_ITA)
    df.columns = [c.strip() for c in df.columns]

    if COL_PDM not in df.columns:
        df[COL_PDM] = ""

    if COL_DESC_ORIG not in df.columns:
        df[COL_DESC_ORIG] = ""

    df[COL_PDM] = (
        df[COL_PDM]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .str.zfill(5)
    )

    if COL_VALOR_EMPENHADO_ORIG in df.columns:
        df["Valor Empenhado"] = (
            df[COL_VALOR_EMPENHADO_ORIG]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df["Valor Empenhado"] = pd.to_numeric(
            df["Valor Empenhado"], errors="coerce"
        )
    else:
        df["Valor Empenhado"] = 0.0

    valor_limite = VALOR_LIMITE_2026
    df["Limite da Dispensa"] = valor_limite
    df["Saldo para contratação"] = df["Limite da Dispensa"] - df["Valor Empenhado"]

    df = df.rename(columns={COL_DESC_ORIG: "Descrição"})

    return df


df_limite_pdm_base = carregar_dados_limite_pdm()

PDMS_UNICOS = sorted(
    c
    for c in df_limite_pdm_base[COL_PDM].dropna().unique()
    if isinstance(c, str) and c.strip() != ""
)

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
    style={
        "display": "flex",
        "flexDirection": "row",
        "width": "100%",
        "gap": "10px",
        "background": "linear-gradient(to bottom, #f5f5f5 0, #f5f5f5 33%, white 33%, white 100%)",
    },
    children=[
        # Coluna esquerda
        html.Div(
            id="coluna_esquerda_pdm",
            style={
                "flex": "1 1 33%",
                "borderRight": "1px solid #ccc",
                "padding": "5px",
                "minWidth": "280px",
                "fontSize": "13px",
                "textAlign": "justify",
                "backgroundColor": "#f0f4fa",
            },
            children=[
                html.P("Prezado requisitante,"),
                html.Br(),
                html.P(
                    "Em atenção ao acórdão nº 324/2009 Plenário TCU, "
                    "“Planeje adequadamente as compras e a contratação de serviços durante o "
                    "exercício financeiro, de forma a evitar a prática de fracionamento de despesas”."
                ),
                html.Br(),
                html.P("Assim dispõe a IN SEGES/ME nº 67/2021:"),
                html.Br(),
                html.P(
                    "Art. 4º Os órgãos e entidades adotarão a dispensa de licitação, na forma "
                    "eletrônica, nas seguintes hipóteses:"
                ),
                html.P(
                    "[...] § 2º Considera-se ramo de atividade a linha de fornecimento registrada "
                    "pelo fornecedor quando do seu cadastramento no Sistema de Cadastramento "
                    "Unificado de Fornecedores (Sicaf), vinculada:"
                ),
                html.P(
                    "I - à classe de materiais, utilizando o Padrão Descritivo de Materiais (PDM) do "
                    "Sistema de Catalogação de Material do Governo federal; ou"
                ),
                html.P(
                    "II - à descrição dos serviços ou das obras, constante do Sistema de Catalogação "
                    "de Serviços ou de Obras do Governo federal. (NR)"
                ),
                html.Br(),
                html.P("Em resumo: Para materiais - PDM; para serviços - CATSER."),
                html.Br(),
                html.P(
                    [
                        "Para obtenção do PDM: no catálogo de compras disponível em ",
                        html.A(
                            "https://catalogo.compras.gov.br/cnbs-web/busca",
                            href="https://catalogo.compras.gov.br/cnbs-web/busca",
                            target="_blank",
                            style={
                                "color": "#1d4ed8",
                                "textDecoration": "underline",
                            },
                        ),
                        ", informar o número do CATMAT. Exemplo para o CATMAT 605322: a consulta "
                        "retornará PDM: 8320. Esse é o número que deverá ser considerado.",
                    ]
                ),
                html.Br(),
                html.P("Exemplo para a necessidade de contratação de três itens:"),
                html.P(
                    "1) o somatório do valor obtido na pesquisa de mercado para cada um dos itens "
                    "multiplicado por seu quantitativo não poderá exceder o limite da dispensa."
                ),
                html.P(
                    "2) O valor por item deverá obrigatoriamente ser igual ou inferior ao saldo para "
                    "contratação (PDM ou CATSER) desse item."
                ),
                html.Br(),
                html.P(
                    "Os valores informados na tabela são os já empenhados no exercício por PDM ou CATSER."
                ),
                html.Br(),
                html.P(
                    "O processo de compra deverá vir instruído já na modalidade DISPENSA DE LICITAÇÃO. "
                    "A tela de consulta (Relatório PDF) deverá estar apensado ao processo, que será "
                    "conferido pelo Setor de Compras e, somente a partir do resultado dessa conferência, "
                    "o processo prosseguirá.",
                    style={"color": "red"},
                ),
            ],
        ),
        # Coluna direita
        html.Div(
            id="coluna_direita_pdm",
            style={
                "flex": "2 1 67%",
                "padding": "5px",
                "minWidth": "400px",
            },
            children=[
                html.Div(
                    id="barra_filtros_limite_itajuba_pdm",
                    className="filtros-sticky",
                    children=[
                        # Primeira linha: PDM (digitação)
                        html.Div(
                            style={
                                "display": "flex",
                                "flexWrap": "wrap",
                                "gap": "10px",
                                "alignItems": "flex-start",
                            },
                            children=[
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 260px",
                                        "maxHeight": "60px",
                                    },
                                    children=[
                                        html.Label("PDM (digitação)"),
                                        dcc.Input(
                                            id="filtro_pdm_texto_itajuba",
                                            type="text",
                                            placeholder="Digite parte do PDM, selecione na lista e, após a seleção, apague o texto digitado.",
                                            style={
                                                "width": "100%",
                                                "marginBottom": "6px",
                                            },
                                        ),
                                    ],
                                ),
                            ],
                        ),
                        # Segunda linha: checklist PDM
                        html.Div(
                            style={
                                "marginTop": "4px",
                                "display": "flex",
                                "flexWrap": "wrap",
                                "gap": "10px",
                                "alignItems": "flex-start",
                            },
                            children=[
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 100%",
                                        "maxHeight": "130px",
                                        "overflowY": "auto",
                                        "border": "1px solid #d1d5db",
                                        "borderRadius": "4px",
                                        "padding": "4px",
                                        "fontSize": "11px",
                                    },
                                    children=[
                                        html.Label("PDM (lista)"),
                                        dcc.Checklist(
                                            id="filtro_pdm_lista_itajuba",
                                            options=[
                                                {"label": c, "value": c}
                                                for c in PDMS_UNICOS
                                            ],
                                            value=[],
                                            style={
                                                "display": "flex",
                                                "flexWrap": "wrap",
                                                "columnGap": "8px",
                                                "rowGap": "2px",
                                            },
                                            inputStyle={"marginRight": "4px"},
                                            labelStyle={
                                                "display": "inline-block",
                                                "width": "18%",
                                                "fontSize": "11px",
                                            },
                                        ),
                                    ],
                                ),
                            ],
                        ),
                        # Terceira linha: título + texto + card + botões
                        html.Div(
                            style={
                                "marginTop": "8px",
                                "display": "flex",
                                "alignItems": "center",
                                "gap": "16px",
                                "flexWrap": "wrap",
                                "justifyContent": "space-between",
                            },
                            children=[
                                html.Div(
                                    style={
                                        "display": "flex",
                                        "flexDirection": "column",
                                        "gap": "2px",
                                        "maxWidth": "520px",
                                    },
                                    children=[
                                        html.H4(
                                            "Limite de Gasto – Itajubá por PDM",
                                            style={"margin": "0px"},
                                        ),
                                        html.Div(
                                            [
                                                html.Span(
                                                    "O valor global do processo de compra não poderá exceder esse limite."
                                                ),
                                                html.Br(),
                                                html.Span(
                                                    "O valor de cada item não poderá exceder o Saldo para Contratação."
                                                ),
                                            ],
                                            style={
                                                "color": "red",
                                                "fontSize": "12px",
                                            },
                                        ),
                                    ],
                                ),
                                # CARD DO LIMITE DA DISPENSA 2026
                                html.Div(
                                    style={
                                        "border": "1px solid #d1d5db",
                                        "borderRadius": "6px",
                                        "padding": "6px 10px",
                                        "backgroundColor": "#ecfdf5",
                                        "minWidth": "220px",
                                        "boxShadow": "0 1px 2px rgba(0,0,0,0.05)",
                                        "fontSize": "12px",
                                    },
                                    children=[
                                        html.Div(
                                            "Limite da dispensa (2026)",
                                            style={
                                                "fontWeight": "bold",
                                                "color": "#166534",
                                                "marginBottom": "2px",
                                                "textAlign": "center",
                                            },
                                        ),
                                        html.Div(
                                            "R$ 65.492,11",
                                            style={
                                                "fontSize": "18px",
                                                "fontWeight": "bold",
                                                "color": "#14532d",
                                                "textAlign": "center",
                                            },
                                        ),
                                    ],
                                ),
                                html.Div(
                                    style={
                                        "display": "flex",
                                        "alignItems": "center",
                                        "gap": "10px",
                                        "flexWrap": "wrap",
                                    },
                                    children=[
                                        html.Button(
                                            "Limpar filtros",
                                            id="btn_limpar_filtros_limite_itajuba_pdm",
                                            n_clicks=0,
                                            className="filtros-button",
                                        ),
                                        html.Button(
                                            "Baixar Relatório PDF",
                                            id="btn_download_relatorio_limite_itajuba_pdm",
                                            n_clicks=0,
                                            className="filtros-button",
                                            style={"marginLeft": "4px"},
                                        ),
                                        html.Div(
                                            style={
                                                "padding": "6px 12px",
                                                "borderRadius": "4px",
                                                "backgroundColor": "#f3f4f6",
                                                "border": "1px solid #d1d5db",
                                                "fontSize": "12px",
                                            },
                                            children=[
                                                html.Span(
                                                    f"Data da consulta: {DATA_HOJE}"
                                                ),
                                            ],
                                        ),
                                        dcc.Download(
                                            id="download_relatorio_limite_itajuba_pdm"
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    ],
                ),
                dash_table.DataTable(
                    id="tabela_limite_itajuba_pdm",
                    columns=[
                        {"name": "PDM", "id": COL_PDM},
                        {"name": "Descrição", "id": "Descrição"},
                        {
                            "name": "Valor Empenhado (R$)",
                            "id": "Valor Empenhado_fmt",
                        },
                        {
                            "name": "Limite da Dispensa (R$)",
                            "id": "Limite da Dispensa_fmt",
                        },
                        {
                            "name": "Saldo para contratação (R$)",
                            "id": "Saldo para contratação_fmt",
                        },
                    ],
                    data=[],
                    row_selectable=False,
                    cell_selectable=False,
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "height": "calc(100vh - 350px)",
                        "minHeight": "300px",
                        "position": "relative",
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "4px",
                        "fontSize": "11px",
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
                    style_data_conditional=[
                        {
                            "if": {
                                "filter_query": "{Saldo para contratação} <= 0",
                            },
                            "backgroundColor": "#ffcccc",
                            "color": "#cc0000",
                        }
                    ],
                ),
                dcc.Store(id="store_dados_limite_itajuba_pdm"),
            ],
        ),
    ],
)


# --------------------------------------------------
# Callbacks de filtros e tabela
# --------------------------------------------------


@dash.callback(
    Output("filtro_pdm_lista_itajuba", "options"),
    Input("filtro_pdm_texto_itajuba", "value"),
    State("filtro_pdm_lista_itajuba", "value"),
)
def atualizar_opcoes_pdm(pdm_texto, valores_selecionados):
    base = PDMS_UNICOS

    if not pdm_texto or not str(pdm_texto).strip():
        opcoes = [{"label": c, "value": c} for c in base]
    else:
        termo = str(pdm_texto).strip().lower()
        filtradas = [c for c in base if termo in str(c).lower()]

        if valores_selecionados:
            for v in valores_selecionados:
                if v in base and v not in filtradas:
                    filtradas.append(v)

        opcoes = [{"label": c, "value": c} for c in sorted(filtradas)]

    return opcoes


@dash.callback(
    Output("tabela_limite_itajuba_pdm", "data"),
    Output("store_dados_limite_itajuba_pdm", "data"),
    Input("filtro_pdm_lista_itajuba", "value"),
    Input("filtro_pdm_texto_itajuba", "value"),
)
def atualizar_tabela_limite_itajuba_pdm(pdm_lista, pdm_texto):
    dff = df_limite_pdm_base.copy()

    # Filtro por PDM digitado
    if pdm_texto and str(pdm_texto).strip():
        termo_pdm = str(pdm_texto).strip().lower()
        dff = dff[
            dff[COL_PDM]
            .astype(str)
            .str.lower()
            .str.contains(termo_pdm, na=False)
        ]

    # Filtro pela checklist de PDM
    if pdm_lista:
        dff = dff[dff[COL_PDM].isin(pdm_lista)]

    cols_tabela = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado",
        "Limite da Dispensa",
        "Saldo para contratação",
    ]

    for c in cols_tabela:
        if c not in dff.columns:
            dff[c] = pd.NA

    dff_display = dff[cols_tabela].copy()

    def fmt_moeda(v):
        if pd.isna(v):
            return ""
        return "R$ " + (
            f"{v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    dff_display["Valor Empenhado_fmt"] = dff_display["Valor Empenhado"].apply(
        fmt_moeda
    )
    dff_display["Limite da Dispensa_fmt"] = dff_display[
        "Limite da Dispensa"
    ].apply(fmt_moeda)
    dff_display["Saldo para contratação_fmt"] = dff_display[
        "Saldo para contratação"
    ].apply(fmt_moeda)

    cols_tabela_display = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado_fmt",
        "Limite da Dispensa_fmt",
        "Saldo para contratação",
        "Saldo para contratação_fmt",
    ]

    return (
        dff_display[cols_tabela_display].to_dict("records"),
        dff.to_dict("records"),
    )


@dash.callback(
    Output("filtro_pdm_texto_itajuba", "value"),
    Output("filtro_pdm_lista_itajuba", "value"),
    Input("btn_limpar_filtros_limite_itajuba_pdm", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_limite_itajuba_pdm(n):
    return None, []


# --------------------------------------------------
# PDF
# --------------------------------------------------

# Estilo para as células de dados (texto preto)
wrap_style_data_pdm = ParagraphStyle(
    name="wrap_limite_itajuba_pdm_data",
    fontSize=8,
    leading=10,
    spaceAfter=4,
    alignment=TA_CENTER,
    textColor=colors.black,
)

# Estilo para o cabeçalho (texto branco)
wrap_style_header_pdm = ParagraphStyle(
    name="wrap_limite_itajuba_pdm_header",
    fontSize=8,
    leading=10,
    spaceAfter=4,
    alignment=TA_CENTER,
    textColor=colors.white,
)

def wrap_data_pdm(text):
    return Paragraph(str(text), wrap_style_data_pdm)

def wrap_header_pdm(text):
    return Paragraph(str(text), wrap_style_header_pdm)


@dash.callback(
    Output("download_relatorio_limite_itajuba_pdm", "data"),
    Input("btn_download_relatorio_limite_itajuba_pdm", "n_clicks"),
    State("store_dados_limite_itajuba_pdm", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_limite_itajuba_pdm(n, dados):

    if not n or not dados:
        return None

    df = pd.DataFrame(dados)

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

    # -------- Data / Hora --------
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora_brasilia = datetime.now(tz_brasilia).strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    data_top_table = Table(
        [[
            Paragraph(
                data_hora_brasilia,
                ParagraphStyle(
                    "data_topo",
                    fontSize=9,
                    alignment=TA_RIGHT,
                    textColor="#333333",
                ),
            )
        ]],
        colWidths=[pagesize[0] - 0.6 * inch],
    )

    data_top_table.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ])
    )

    story.append(data_top_table)
    story.append(Spacer(1, 0.15 * inch))

    # -------- Cabeçalho: Logo | Texto | Logo --------
    logo_esq = (
        Image("assets/brasaobrasil.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/brasaobrasil.png") else ""
    )

    logo_dir = (
        Image("assets/simbolo_RGB.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/simbolo_RGB.png") else ""
    )

    texto_instituicao = (
        "<b><font color='#0b2b57' size=13>Universidade Federal de Itajubá</font></b><br/>"
        "<font color='#0b2b57' size=11>Diretoria de Compras e Contratos</font>"
    )

    instituicao = Paragraph(
        texto_instituicao,
        ParagraphStyle(
            "instituicao_pdm",
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
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ])
    )

    story.append(cabecalho)
    story.append(Spacer(1, 0.25 * inch))

    # -------- Título --------
    titulo_texto = (
        "Consulta Fracionamento de Despesa 2026 - PDM (Material): UASG: 153030 - Campus Itajubá<br/>"
        
    )

    titulo_paragraph = Paragraph(
        titulo_texto,
        ParagraphStyle(
            "titulo_consulta_pdm",
            fontSize=10,
            alignment=TA_CENTER,
            textColor="#0b2b57",
            spaceAfter=4,
            leading=14,
        ),
    )

    story.append(titulo_paragraph)
    story.append(Spacer(1, 0.15 * inch))

    story.append(
        Paragraph(f"Total de registros: {len(df)}", styles["Normal"])
    )
    story.append(Spacer(1, 0.1 * inch))

    # -------- Tabela --------
    cols = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado",
        "Limite da Dispensa",
        "Saldo para contratação",
    ]

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    def fmt_moeda_pdf(v):
        if pd.isna(v):
            return ""
        return "R$ " + (
            f"{v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    df_pdf = df.copy()
    for col in cols[2:]:
        df_pdf[col] = df_pdf[col].apply(fmt_moeda_pdf)

    header = [wrap_header_pdm(col) for col in cols]
    table_data = [header]

    saldo_values = df["Saldo para contratação"].fillna(0).tolist()

    for _, row in df_pdf[cols].iterrows():
        table_data.append(
            [wrap_data_pdm(row[c]) for c in cols]
        )

    page_width = pagesize[0] - 0.6 * inch

    col_widths = [
        0.9 * inch,
        page_width - (0.9 + 1.1 + 1.1 + 1.1) * inch,
        1.1 * inch,
        1.1 * inch,
        1.1 * inch,
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("WORDWRAP", (0, 0), (-1, -1), True),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]

    for row_idx in range(1, len(table_data)):
        table_styles.append(
            ("ALIGN", (1, row_idx), (1, row_idx), "LEFT")
        )

    for row_idx, saldo in enumerate(saldo_values, 1):
        if saldo <= 0:
            table_styles.append(
                ("BACKGROUND", (0, row_idx), (-1, row_idx),
                 colors.HexColor("#ffcccc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, row_idx), (-1, row_idx),
                 colors.HexColor("#cc0000"))
            )

    tbl.setStyle(TableStyle(table_styles))
    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(
        buffer.getvalue(),
        "limite_gasto_itajuba_pdm.pdf",
    )
