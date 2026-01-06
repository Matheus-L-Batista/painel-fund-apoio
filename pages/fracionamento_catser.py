import dash
from dash import html, dcc, dash_table, Input, Output, State, no_update

import pandas as pd
from datetime import date

from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors

# --------------------------------------------------
# Registro da página
# --------------------------------------------------

dash.register_page(
    __name__,
    path="/fracionamento_catser",
    name="Fracionamento de Despesas CATSER",
    title="Fracionamento de Despesas CATSER",
)

# --------------------------------------------------
# URL da planilha
# --------------------------------------------------

URL_LIMITE_GASTO_ITA = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=Limite%20de%20Gasto%20-%20Itajub%C3%A1"
)

COL_CATSER = "CATSER"
COL_DESC_ORIG = "Descrição"
COL_VALOR_EMPENHADO_ORIG = "Unnamed: 3"

DATA_HOJE = date.today().strftime("%d/%m/%Y")

# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------

def carregar_dados_limite():
    df = pd.read_csv(URL_LIMITE_GASTO_ITA)
    df.columns = [c.strip() for c in df.columns]

    if COL_CATSER not in df.columns:
        df[COL_CATSER] = ""
    if COL_DESC_ORIG not in df.columns:
        df[COL_DESC_ORIG] = ""

    df[COL_CATSER] = (
        df[COL_CATSER]
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
        df["Valor Empenhado"] = pd.to_numeric(df["Valor Empenhado"], errors="coerce")
    else:
        df["Valor Empenhado"] = 0.0

    valor_limite = 65492.11
    df["Limite da Dispensa"] = valor_limite
    df["Saldo para contratação"] = df["Limite da Dispensa"] - df["Valor Empenhado"]

    df = df.rename(columns={COL_DESC_ORIG: "Descrição"})

    return df

df_limite_base = carregar_dados_limite()

CATSERS_UNICOS = sorted(
    [
        c
        for c in df_limite_base[COL_CATSER].dropna().unique()
        if isinstance(c, str) and c.strip() != ""
    ]
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
    },
    children=[
        # Coluna esquerda
        html.Div(
            id="coluna_esquerda_catser",
            style={
                "flex": "1 1 33%",
                "borderRight": "1px solid #ccc",
                "padding": "5px",
                "minWidth": "280px",
                "fontSize": "12px",
                "textAlign": "justify",
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
                    "de Serviços ou de Obras do Governo federal.\" (NR)"
                ),
                html.Br(),
                html.P("Em resumo: Para materiais - PDM; para serviços - CATSER."),
                html.Br(),
                html.P(
                    children=[
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
                    "A tela de consulta (print da tela) deverá estar apensado ao processo, que será "
                    "conferido pelo Setor de Compras e, somente a partir do resultado dessa conferência, "
                    "o processo prosseguirá.",
                    style={"color": "red"},
                ),
            ],
        ),

        # Coluna direita
        html.Div(
            id="coluna_direita_catser",
            style={
                "flex": "2 1 67%",
                "padding": "5px",
                "minWidth": "400px",
            },
            children=[
                # Barra de filtros
                html.Div(
                    id="barra_filtros_limite_itajuba",
                    className="filtros-sticky",
                    children=[
                        # Primeira linha: filtros
                        html.Div(
                            style={
                                "display": "flex",
                                "flexWrap": "wrap",
                                "gap": "10px",
                                "alignItems": "flex-start",
                            },
                            children=[
                                # CATSER texto
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 260px",
                                        "maxHeight": "60px",
                                    },
                                    children=[
                                        html.Label("CATSER (digitação)"),
                                        dcc.Input(
                                            id="filtro_catser_texto_itajuba",
                                            type="text",
                                            placeholder="Digite parte do CATSER",
                                            style={
                                                "width": "100%",
                                                "marginBottom": "6px",
                                            },
                                        ),
                                    ],
                                ),
                                # CATSER checklist em 5 colunas, altura reduzida
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 260px",
                                        "maxHeight": "130px",  # metade aprox.
                                        "overflowY": "auto",
                                        "border": "1px solid #d1d5db",
                                        "borderRadius": "4px",
                                        "padding": "4px",
                                        "fontSize": "11px",
                                    },
                                    children=[
                                        html.Label("CATSER (lista)"),
                                        dcc.Checklist(
                                            id="filtro_catser_dropdown_itajuba",
                                            options=[
                                                {"label": c, "value": c}
                                                for c in CATSERS_UNICOS
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
                        # Segunda linha: botões + data
                        html.Div(
                            style={
                                "marginTop": "4px",
                                "display": "flex",
                                "alignItems": "center",
                                "gap": "10px",
                                "flexWrap": "wrap",
                            },
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_limite_itajuba",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_limite_itajuba",
                                    n_clicks=0,
                                    className="filtros-button",
                                    style={"marginLeft": "10px"},
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
                                    id="download_relatorio_limite_itajuba"
                                ),
                            ],
                        ),
                    ],
                ),
                html.H4("Limite de Gasto – Itajubá por CATSER"),
                dash_table.DataTable(
                    id="tabela_limite_itajuba",
                    columns=[
                        {"name": "CATSER", "id": COL_CATSER},
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
                            "if": {"column_id": "Saldo para contratação_fmt"},
                            "backgroundColor": "#f9f9f9",
                        }
                    ],
                ),
                dcc.Store(id="store_dados_limite_itajuba"),
            ],
        ),
    ],
)

# --------------------------------------------------
# Callbacks
# --------------------------------------------------

@dash.callback(
    Output("filtro_catser_dropdown_itajuba", "options"),
    Input("filtro_catser_texto_itajuba", "value"),
    State("filtro_catser_dropdown_itajuba", "value"),
)
def atualizar_opcoes_catser(catser_texto, valores_selecionados):
    base = CATSERS_UNICOS

    if not catser_texto or not str(catser_texto).strip():
        opcoes = [{"label": c, "value": c} for c in base]
    else:
        termo = str(catser_texto).strip().lower()
        filtradas = [c for c in base if termo in str(c).lower()]
        if valores_selecionados:
            for v in valores_selecionados:
                if v in base and v not in filtradas:
                    filtradas.append(v)
        opcoes = [{"label": c, "value": c} for c in sorted(filtradas)]

    return opcoes

@dash.callback(
    Output("tabela_limite_itajuba", "data"),
    Output("store_dados_limite_itajuba", "data"),
    Input("filtro_catser_texto_itajuba", "value"),
    Input("filtro_catser_dropdown_itajuba", "value"),
)
def atualizar_tabela_limite_itajuba(catser_texto, catser_lista):
    dff = df_limite_base.copy()

    if catser_texto and str(catser_texto).strip():
        termo = str(catser_texto).strip().lower()
        dff = dff[
            dff[COL_CATSER]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    if catser_lista:
        dff = dff[dff[COL_CATSER].isin(catser_lista)]

    cols_tabela = [
        COL_CATSER,
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

    dff_display["Valor Empenhado_fmt"] = dff_display["Valor Empenhado"].apply(fmt_moeda)
    dff_display["Limite da Dispensa_fmt"] = dff_display["Limite da Dispensa"].apply(fmt_moeda)
    dff_display["Saldo para contratação_fmt"] = dff_display["Saldo para contratação"].apply(
        fmt_moeda
    )

    cols_tabela_display = [
        COL_CATSER,
        "Descrição",
        "Valor Empenhado_fmt",
        "Limite da Dispensa_fmt",
        "Saldo para contratação_fmt",
    ]

    return dff_display[cols_tabela_display].to_dict("records"), dff.to_dict("records")

@dash.callback(
    Output("filtro_catser_texto_itajuba", "value"),
    Output("filtro_catser_dropdown_itajuba", "value"),
    Input("btn_limpar_filtros_limite_itajuba", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_limite_itajuba(n):
    return None, []

wrap_style = ParagraphStyle(
    name="wrap_limite_itajuba",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

@dash.callback(
    Output("download_relatorio_limite_itajuba", "data"),
    Input("btn_download_relatorio_limite_itajuba", "n_clicks"),
    State("store_dados_limite_itajuba", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_limite_itajuba(n, dados):
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
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Relatório – Limite de Gasto – Itajubá por CATSER",
        ParagraphStyle(
            "titulo_limite_itajuba",
            fontSize=16,
            alignment=TA_CENTER,
            textColor="#0b2b57",
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        COL_CATSER,
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
    for col in ["Valor Empenhado", "Limite da Dispensa", "Saldo para contratação"]:
        if col in df_pdf.columns:
            df_pdf[col] = df_pdf[col].apply(fmt_moeda_pdf)

    header = cols
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        linha = [wrap(row[c]) for c in cols]
        table_data.append(linha)

    page_width = pagesize[0] - 0.6 * inch
    col_width = page_width / max(1, len(header))
    col_widths = [col_width] * len(header)

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )

    story.append(tbl)
    doc.build(story)
    buffer.seek(0)

    from dash import dcc

    return dcc.send_bytes(buffer.getvalue(), "limite_gasto_itajuba_catser.pdf")
