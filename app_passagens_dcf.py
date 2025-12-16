app = Dash(__name__)
server = app.server

# ----------------------------------------
# 1. IMPORTS
# ----------------------------------------

from dash import Dash, html, dcc, Input, Output, State, dash_table
import plotly.express as px
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
import plotly.io as pio

# PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Image, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors


# ----------------------------------------
# 2. CARREGAMENTO E PREPARAÇÃO DOS DADOS
# ----------------------------------------

URL = "https://docs.google.com/spreadsheets/d/1QJFSLpVO0bI-bsNdgiTWl8rOh1_h6_B7Q8F_SW66_yc/gviz/tq?tqx=out:csv&sheet=Passagens%20-%20DCF"


def carregar_dados():
    df = pd.read_csv(URL)

    # Padroniza nomes das colunas
    df.columns = [c.strip() for c in df.columns]

    # Converte datas
    df["Data Início da Viagem"] = pd.to_datetime(
        df["Data Início da Viagem"], format="%d/%m/%Y", errors="coerce"
    )

    # Converte valores monetários
    def conv_moeda(valor):
        if isinstance(valor, str):
            v = (
                valor.replace("R$", "")
                .replace(".", "")
                .replace(",", ".")
                .strip()
            )
            return float(v) if v not in ["", "-"] else 0.0
        return float(valor) if pd.notna(valor) else 0.0

    col_moeda = [
        "Valor das Diárias",
        "Valor da Viagem",
        "Valor da Passagem",
        "Valor Seguro Viagem",
        "Valor Restituição",
        "Custo com emissão de passagens dentro do prazo",
        "Custo com emissão de passagens em caráter de urgência",
    ]

    for col in col_moeda:
        df[col] = df[col].apply(conv_moeda)

    # Cria colunas de ano / mês
    df["Ano"] = df["Data Início da Viagem"].dt.year
    df["Mes"] = df["Data Início da Viagem"].dt.month

    return df


df = carregar_dados()


# ----------------------------------------
# 3. APLICATIVO DASH
# ----------------------------------------

app = Dash(__name__)

nomes_meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]


# ----------------------------------------
# 4. LAYOUT — COLUNA ESQUERDA FIXA
# ----------------------------------------

app.layout = html.Div([
    html.Div([

        # -------- COLUNA ESQUERDA — FIXA --------
        html.Div([
            html.Img(
                src="/assets/logo_unifei.png",
                style={"width": "80%", "marginBottom": "20px"}
            ),

            html.H3("Ano", style={"color": "white"}),
            dcc.Dropdown(
                id="filtro_ano",
                options=[{"label": int(a), "value": int(a)}
                         for a in sorted(df["Ano"].dropna().unique())],
                value=2025,
                clearable=False,
                style={"color": "black"}
            ),

            html.H3("Mês", style={"marginTop": "20px", "color": "white"}),
            dcc.Dropdown(
                id="filtro_mes",
                options=[{"label": m.capitalize(), "value": i}
                         for i, m in enumerate(nomes_meses, start=1)],
                value=None,
                placeholder="Todos",
                clearable=True,
                style={"color": "black"}
            ),

            html.H3("Unidade (Viagem)", style={"marginTop": "20px", "color": "white"}),
            dcc.Dropdown(
                id="filtro_unidade",
                options=[{"label": u, "value": u}
                         for u in sorted(df["Unidade (Viagem)"].unique())],
                value=None,
                placeholder="Todas",
                clearable=True,
                style={"color": "black"}
            ),

            html.Div([
                html.Button(
                    "Limpar filtros",
                    id="btn_limpar_filtros",
                    n_clicks=0,
                    style={"width": "100%", "marginTop": "20px"}
                ),
                html.Button(
                    "Baixar Relatório PDF",
                    id="btn_download_relatorio",
                    n_clicks=0,
                    style={"width": "100%", "marginTop": "10px"}
                ),
                dcc.Download(id="download_relatorio"),
            ])
        ],
            style={
                "width": "20%",
                "padding": "15px",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "height": "100vh",
                "position": "sticky",
                "top": 0,
                "overflow": "auto"
            }),

        # -------- COLUNA DIREITA (continuada na PARTE 2) --------
        # -------- COLUNA DIREITA — CONTEÚDO ROLÁVEL --------
        html.Div([

            html.H2("Gastos com Viagens", style={"textAlign": "center"}),

            # ----- CARDS -----
            html.Div(
                id="cards_container",
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "gap": "10px",
                    "marginBottom": "20px",
                    "flexWrap": "wrap"
                }
            ),

            # ----- GRÁFICOS LADO A LADO -----
            html.Div([
                dcc.Graph(id="grafico_pizza", style={"width": "50%"}),
                dcc.Graph(id="grafico_barras", style={"width": "50%"})
            ],
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "marginBottom": "25px"
                }
            ),

            # ----- RESUMO POR UNIDADE -----
            html.H4("Resumo por Unidade"),
            dash_table.DataTable(
                id="tabela_unidades",
                row_selectable=False,
                cell_selectable=False,
                active_cell=None,
                selected_cells=[],
                selected_rows=[],
                columns=[
                    {"name": "Unidade (Viagem)", "id": "Unidade (Viagem)"},
                    {"name": "Gasto com Diárias", "id": "Valor das Diárias"},
                    {"name": "Gasto com Passagem", "id": "Valor da Passagem"},
                    {"name": "Gasto com Restituição", "id": "Valor Restituição"},
                    {"name": "Gasto com Seguro Viagem", "id": "Valor Seguro Viagem"},
                ],
                data=[],
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "center", "padding": "4px"},
                style_header={
                    "fontWeight": "bold",
                    "backgroundColor": "#f0f0f0"
                }
            ),

            # ----- DETALHAMENTO PCDP -----
            html.H4("Detalhamento por Unidade e PCDP"),
            dash_table.DataTable(
                id="tabela_detalhe",
                row_selectable=False,
                cell_selectable=False,
                active_cell=None,
                selected_cells=[],
                selected_rows=[],
                columns=[
                    {"name": "Unidade (Viagem)", "id": "Unidade (Viagem)"},
                    {"name": "Número da PCDP", "id": "Número da PCDP"},
                    {"name": "Data Início da Viagem", "id": "Data Início da Viagem"},
                    {"name": "Custo passagens no prazo",
                     "id": "Custo com emissão de passagens dentro do prazo"},
                    {"name": "Custo passagens urgência",
                     "id": "Custo com emissão de passagens em caráter de urgência"},
                ],
                data=[],
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "center", "padding": "4px"},
                style_header={
                    "fontWeight": "bold",
                    "backgroundColor": "#f0f0f0"
                }
            ),

        ],
            style={
                "width": "80%",
                "padding": "15px",
                "overflowY": "auto"
            }
        )

    ], style={"display": "flex"}),

    # Store usado pelo PDF
    dcc.Store(id="store_graficos")
])
# ----------------------------------------
# 5. CALLBACK — Atualização geral (cards, gráficos, tabela resumo)
# ----------------------------------------

@app.callback(
    [
        Output("cards_container", "children"),
        Output("grafico_pizza", "figure"),
        Output("grafico_barras", "figure"),
        Output("tabela_unidades", "data"),
        Output("store_graficos", "data"),
    ],
    [
        Input("filtro_ano", "value"),
        Input("filtro_mes", "value"),
        Input("filtro_unidade", "value"),
    ],
)
def atualizar_pagina(ano, mes, unidade):
    dff = df.copy()

    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes:
        dff = dff[dff["Mes"] == mes]
    if unidade:
        dff = dff[dff["Unidade (Viagem)"] == unidade]

    total_viagem = dff["Valor da Viagem"].sum()
    total_prazo = dff["Custo com emissão de passagens dentro do prazo"].sum()
    total_urgencia = dff["Custo com emissão de passagens em caráter de urgência"].sum()
    total_diarias = dff["Valor das Diárias"].sum()
    total_seguro = dff["Valor Seguro Viagem"].sum()
    total_restit = dff["Valor Restituição"].sum()
    total_passagem = dff["Valor da Passagem"].sum()

    def f(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def card(titulo, valor):
        return html.Div([
            html.Div(titulo, style={"fontSize": "12px", "color": "#555", "textAlign": "center"}),
            html.Div(f(valor), style={"fontSize": "18px", "fontWeight": "bold", "color": "#b30000", "textAlign": "center"})
        ],
            style={
                "backgroundColor": "white",
                "padding": "10px 15px",
                "borderRadius": "4px",
                "boxShadow": "0 1px 3px rgba(0,0,0,0.2)",
                "flex": "1",
                "minWidth": "0"
            }
        )

    cards = [
        card("Total Viagens", total_viagem),
        card("Passagens no Prazo", total_prazo),
        card("Passagens Urgência", total_urgencia),
        card("Gasto em Diárias", total_diarias),
        card("Seguro Viagem", total_seguro),
        card("Restituições", total_restit),
    ]

    # --- Pizza ---
    pizza_df = pd.DataFrame({
        "Tipo": ["No prazo", "Urgência"],
        "Valor": [total_prazo, total_urgencia]
    })
    fig_pizza = px.pie(
        pizza_df,
        names="Tipo",
        values="Valor",
        hole=0.45,
        color="Tipo",
        color_discrete_map={"No prazo": "#003A70", "Urgência": "#DA291C"},
        title="Passagens — No prazo x Urgência"
    )
    fig_pizza.update_layout(title_x=0.5)

    # --- Barras ---
    barras_df = pd.DataFrame({
        "Categoria": ["Diárias", "Passagens"],
        "Valor": [total_diarias, total_passagem]
    })
    barras_df["Texto"] = barras_df["Valor"].apply(f)

    fig_barras = px.bar(
        barras_df,
        x="Categoria",
        y="Valor",
        text="Texto",
        color="Categoria",
        color_discrete_sequence=["#003A70", "#DA291C"],
        title="Comparativo: Diárias x Passagens"
    )
    fig_barras.update_layout(title_x=0.5, showlegend=False)

    # --- Resumo unidade ---
    resumo = (
        dff.groupby("Unidade (Viagem)", as_index=False)[
            ["Valor das Diárias", "Valor da Passagem", "Valor Restituição", "Valor Seguro Viagem"]
        ].sum()
    )

    for col in ["Valor das Diárias", "Valor da Passagem", "Valor Restituição", "Valor Seguro Viagem"]:
        resumo[col] = resumo[col].apply(f)

    dados_pdf = {
        "resumo": resumo.to_dict("records"),
        "filtros": {"ano": ano, "mes": mes, "unidade": unidade},
        "cards": {
            "total_viagem": total_viagem,
            "total_prazo": total_prazo,
            "total_urgencia": total_urgencia,
            "total_diarias": total_diarias,
            "total_seguro": total_seguro,
            "total_restit": total_restit,
        }
    }

    return cards, fig_pizza, fig_barras, resumo.to_dict("records"), dados_pdf


# ----------------------------------------
# 6. CALLBACK — Tabela de Detalhamento (somente filtros)
# ----------------------------------------

@app.callback(
    Output("tabela_detalhe", "data"),
    [
        Input("filtro_ano", "value"),
        Input("filtro_mes", "value"),
        Input("filtro_unidade", "value"),
    ],
)
def atualizar_detalhe(ano, mes, unidade):
    dff = df.copy()

    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes:
        dff = dff[dff["Mes"] == mes]
    if unidade:
        dff = dff[dff["Unidade (Viagem)"] == unidade]

    dff = dff[
        [
            "Unidade (Viagem)",
            "Número da PCDP",
            "Data Início da Viagem",
            "Custo com emissão de passagens dentro do prazo",
            "Custo com emissão de passagens em caráter de urgência",
        ]
    ].copy()

    dff["Data Início da Viagem"] = dff["Data Início da Viagem"].dt.strftime("%d/%m/%Y")

    def f(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    dff["Custo com emissão de passagens dentro do prazo"] = dff[
        "Custo com emissão de passagens dentro do prazo"
    ].apply(f)

    dff["Custo com emissão de passagens em caráter de urgência"] = dff[
        "Custo com emissão de passagens em caráter de urgência"
    ].apply(f)

    return dff.to_dict("records")


# ----------------------------------------
# 7. CALLBACK — Limpar filtros
# ----------------------------------------

@app.callback(
    [
        Output("filtro_ano", "value"),
        Output("filtro_mes", "value"),
        Output("filtro_unidade", "value"),
    ],
    Input("btn_limpar_filtros", "n_clicks"),
    prevent_initial_call=True,
)
def limpar(n):
    return 2025, None, None

# ----------------------------------------
# 8. CALLBACK — Geração do PDF COMPLETO (com DUAS tabelas)
# ----------------------------------------
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

# Estilo para quebra de linha automática
wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=8,
    leading=10,
    spaceAfter=4
)

def wrap(text):
    return Paragraph(str(text), wrap_style)


@app.callback(
    Output("download_relatorio", "data"),
    Input("btn_download_relatorio", "n_clicks"),
    [
        State("grafico_pizza", "figure"),
        State("grafico_barras", "figure"),
        State("tabela_unidades", "data"),
        State("tabela_detalhe", "data"),
        State("store_graficos", "data"),
    ],
    prevent_initial_call=True,
)
def gerar_pdf(n, fig_pizza, fig_barras, resumo, detalhe, dados_pdf):
    if not n:
        return None

    # Converter gráficos para PNG
    img_pizza = BytesIO(pio.to_image(fig_pizza, format="png", width=450, height=350))
    img_barras = BytesIO(pio.to_image(fig_barras, format="png", width=450, height=350))

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)

    styles = getSampleStyleSheet()
    story = []

    # Título
    titulo = Paragraph(
        "<b>Relatório de Gastos com Viagens</b>",
        ParagraphStyle("titulo", fontSize=22, alignment=TA_CENTER, textColor="#0b2b57")
    )
    story.append(titulo)
    story.append(Spacer(1, 0.3 * inch))

    # Filtros aplicados
    filtros = dados_pdf["filtros"]
    story.append(Paragraph(
        f"<b>Ano:</b> {filtros['ano']} — "
        f"<b>Mês:</b> {filtros['mes'] if filtros['mes'] else 'Todos'} — "
        f"<b>Unidade:</b> {filtros['unidade'] if filtros['unidade'] else 'Todas'}",
        styles["Normal"]
    ))
    story.append(Spacer(1, 0.3 * inch))

    # Gráficos
    story.append(Image(img_pizza, width=250, height=200))
    story.append(Image(img_barras, width=250, height=200))
    story.append(Spacer(1, 0.4 * inch))

    # Resumo por unidade
    story.append(Paragraph("<b>Resumo por Unidade</b>", styles["Heading2"]))

    table1 = [["Unidade", "Diárias", "Passagem", "Restituição", "Seguro"]]
    for r in resumo:
        table1.append([
            wrap(r["Unidade (Viagem)"]),
            wrap(r["Valor das Diárias"]),
            wrap(r["Valor da Passagem"]),
            wrap(r["Valor Restituição"]),
            wrap(r["Valor Seguro Viagem"]),
        ])

    # Larguras da tabela 1
    col_widths1 = [
        2.8 * inch,
        1.0 * inch,
        1.0 * inch,
        1.0 * inch,
        1.0 * inch
    ]

    tbl1 = Table(table1, colWidths=col_widths1)

    tbl1.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("WORDWRAP", (0, 0), (-1, -1), True),
    ]))

    story.append(tbl1)
    story.append(Spacer(1, 0.5 * inch))

    # Detalhamento
    story.append(Paragraph("<b>Detalhamento PCDP</b>", styles["Heading2"]))

    table2 = [["Unidade", "PCDP", "Data", "Prazo", "Urgência"]]
    for r in detalhe:
        table2.append([
            wrap(r["Unidade (Viagem)"]),
            wrap(r["Número da PCDP"]),
            wrap(r["Data Início da Viagem"]),
            wrap(r["Custo com emissão de passagens dentro do prazo"]),
            wrap(r["Custo com emissão de passagens em caráter de urgência"]),
        ])

    # Larguras da tabela 2
    col_widths2 = [
        2.8 * inch,
        1.0 * inch,
        1.0 * inch,
        1.2 * inch,
        1.2 * inch
    ]

    tbl2 = Table(table2, colWidths=col_widths2)

    tbl2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003A70")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("WORDWRAP", (0, 0), (-1, -1), True),
        ("FONTSIZE", (0, 0), (-1, -1), 8)
    ]))

    story.append(tbl2)

    # Finaliza PDF
    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(buffer.getvalue(), "relatorio_gastos_viagens.pdf")


# ----------------------------------------
# 9. Rodar App
# ----------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
