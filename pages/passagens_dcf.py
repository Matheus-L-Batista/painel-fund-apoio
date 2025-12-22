import dash
from dash import html, dcc, Input, Output, State, dash_table
import plotly.express as px
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Image, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/passagens-dcf",
    name="Passagens DCF",
    title="Gastos com Viagens",
)


# --------------------------------------------------
# URL da planilha
# --------------------------------------------------
URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1QJFSLpVO0bI-bsNdgiTWl8rOh1_h6_B7Q8F_SW66_yc/"
    "gviz/tq?tqx=out:csv&sheet=Passagens%20-%20DCF"
)


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados():
    df = pd.read_csv(URL)
    df.columns = [c.strip() for c in df.columns]
    df["Data Início da Viagem"] = pd.to_datetime(
        df["Data Início da Viagem"], format="%d/%m/%Y", errors="coerce"
    )

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

    df["Ano"] = df["Data Início da Viagem"].dt.year
    df["Mes"] = df["Data Início da Viagem"].dt.month
    return df


# 🔧 B) DF base inicial
df_base = carregar_dados()
ANO_PADRAO = int(sorted(df_base["Ano"].dropna().unique())[-1])

nomes_meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "10px",
    "whiteSpace": "normal",
}


# ----------------------------------------
# Layout (conteúdo da página)
# ----------------------------------------
layout = html.Div(
    children=[
        html.H2(
            "Gastos com Viagens",
            style={"textAlign": "center"},
        ),
        html.Div(
            style={"marginBottom": "20px"},
            children=[
                html.H3("Filtros", className="sidebar-title"),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        # Ano
                        html.Div(
                            style={"minWidth": "140px", "flex": "0 0 160px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_passagens",
                                    options=[
                                        {
                                            "label": int(a),
                                            "value": int(a),
                                        }
                                        for a in sorted(
                                            df_base["Ano"].dropna().unique()
                                        )
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style=dropdown_style,
                                    optionHeight=40,
                                ),
                            ],
                        ),
                        # Mês
                        html.Div(
                            style={"minWidth": "140px", "flex": "0 0 160px"},
                            children=[
                                html.Label("Mês"),
                                dcc.Dropdown(
                                    id="filtro_mes_passagens",
                                    options=[
                                        {
                                            "label": m.capitalize(),
                                            "value": i,
                                        }
                                        for i, m in enumerate(
                                            nomes_meses, start=1
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                    optionHeight=35,
                                ),
                            ],
                        ),
                        # Unidade (Viagem)
                        html.Div(
                            style={
                                "minWidth": "240px",
                                "flex": "1 1 280px",
                                "maxWidth": "480px",
                            },
                            children=[
                                html.Label("Unidade (Viagem)"),
                                dcc.Dropdown(
                                    id="filtro_unidade_passagens",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df_base[
                                                "Unidade (Viagem)"
                                            ].unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                    optionHeight=35,
                                ),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    style={"marginTop": "10px"},
                    children=[
                        html.Button(
                            "Limpar filtros",
                            id="btn_limpar_filtros_passagens",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_passagens",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_passagens"),
                    ],
                ),
            ],
        ),
        html.Div(
            id="cards_container_passagens",
            className="cards-container",
        ),
        html.Div(
            className="charts-row",
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "10px",
            },
            children=[
                dcc.Graph(
                    id="grafico_pizza_passagens",
                    style={"flex": "1 1 300px", "minWidth": "280px"},
                ),
                dcc.Graph(
                    id="grafico_barras_passagens",
                    style={"flex": "1 1 300px", "minWidth": "280px"},
                ),
            ],
        ),
        html.H4("Resumo por Unidade"),
        dash_table.DataTable(
            id="tabela_unidades_passagens",
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
                {
                    "name": "Gasto com Seguro Viagem",
                    "id": "Valor Seguro Viagem",
                },
            ],
            data=[],
            style_table={
                "overflowX": "auto",
                "overflowY": "auto",
                "maxHeight": "350px",
            },
            style_cell={"textAlign": "center", "padding": "4px"},
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#f0f0f0",
            },
        ),
        html.H4("Detalhamento por Unidade e PCDP"),
        dash_table.DataTable(
            id="tabela_detalhe_passagens",
            row_selectable=False,
            cell_selectable=False,
            active_cell=None,
            selected_cells=[],
            selected_rows=[],
            columns=[
                {"name": "Unidade (Viagem)", "id": "Unidade (Viagem)"},
                {"name": "Número da PCDP", "id": "Número da PCDP"},
                {
                    "name": "Data Início da Viagem",
                    "id": "Data Início da Viagem",
                },
                {
                    "name": "Custo passagens no prazo",
                    "id": "Custo com emissão de passagens dentro do prazo",
                },
                {
                    "name": "Custo passagens urgência",
                    "id": "Custo com emissão de passagens em caráter de urgência",
                },
            ],
            data=[],
            style_table={
                "overflowX": "auto",
                "overflowY": "auto",
                "maxHeight": "350px",
            },
            style_cell={"textAlign": "center", "padding": "4px"},
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#f0f0f0",
            },
        ),
        dcc.Store(id="store_graficos_passagens"),
    ],
)


# ----------------------------------------
# 5. CALLBACK — Atualização geral
# ----------------------------------------
@dash.callback(
    Output("cards_container_passagens", "children"),
    Output("grafico_pizza_passagens", "figure"),
    Output("grafico_barras_passagens", "figure"),
    Output("tabela_unidades_passagens", "data"),
    Output("store_graficos_passagens", "data"),
    Input("filtro_ano_passagens", "value"),
    Input("filtro_mes_passagens", "value"),
    Input("filtro_unidade_passagens", "value"),
    # 🔧 C) Interval como Input extra
    Input("interval-atualizacao", "n_intervals"),
)
def atualizar_pagina(ano, mes, unidade, n_intervals):
    # Atualiza o DF base apenas quando o intervalo dispara
    df = carregar_dados() if n_intervals is not None else df_base
    dff = df.copy()

    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes:
        dff = dff[dff["Mes"] == mes]
    if unidade:
        dff = dff[dff["Unidade (Viagem)"] == unidade]

    total_viagem = dff["Valor da Viagem"].sum()
    total_prazo = dff[
        "Custo com emissão de passagens dentro do prazo"
    ].sum()
    total_urgencia = dff[
        "Custo com emissão de passagens em caráter de urgência"
    ].sum()
    total_diarias = dff["Valor das Diárias"].sum()
    total_seguro = dff["Valor Seguro Viagem"].sum()
    total_restit = dff["Valor Restituição"].sum()
    total_passagem = dff["Valor da Passagem"].sum()

    def f(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    def card(titulo, valor):
        return html.Div(
            className="card",
            children=[
                html.Div(titulo, className="card-title"),
                html.Div(f(valor), className="card-value"),
            ],
        )

    cards = [
        card("Total Viagens", total_viagem),
        card("Passagens no Prazo", total_prazo),
        card("Passagens Urgência", total_urgencia),
        card("Gasto em Diárias", total_diarias),
        card("Seguro Viagem", total_seguro),
        card("Restituições", total_restit),
    ]

    pizza_df = pd.DataFrame(
        {"Tipo": ["No prazo", "Urgência"], "Valor": [total_prazo, total_urgencia]}
    )

    fig_pizza = px.pie(
        pizza_df,
        names="Tipo",
        values="Valor",
        hole=0.45,
        color="Tipo",
        color_discrete_map={"No prazo": "#003A70", "Urgência": "#DA291C"},
        title="Passagens — No prazo x Urgência",
    )
    fig_pizza.update_layout(title_x=0.5)
    fig_pizza.update_traces(
        texttemplate="%{label}<br>R$ %{value:,.2f} (%{percent})",
        hovertemplate="%{label}<br>R$ %{value:,.2f}",
        textposition="inside",
    )

    barras_df = pd.DataFrame(
        {"Categoria": ["Diárias", "Passagens"], "Valor": [total_diarias, total_passagem]}
    )

    fig_barras = px.bar(
        barras_df,
        x="Categoria",
        y="Valor",
        color="Categoria",
        color_discrete_sequence=["#003A70", "#DA291C"],
        title="Comparativo: Diárias x Passagens",
        text="Valor",
    )
    fig_barras.update_traces(
        texttemplate="R$ %{y:,.2f}",
        textposition="inside",
        hovertemplate="%{x}<br>R$ %{y:,.2f}",
    )
    fig_barras.update_layout(
        title_x=0.5,
        showlegend=False,
        yaxis_tickprefix="R$ ",
        yaxis_tickformat=",.2f",
    )

    resumo = dff.groupby(
        "Unidade (Viagem)", as_index=False
    )[
        [
            "Valor das Diárias",
            "Valor da Passagem",
            "Valor Restituição",
            "Valor Seguro Viagem",
        ]
    ].sum()

    for col in [
        "Valor das Diárias",
        "Valor da Passagem",
        "Valor Restituição",
        "Valor Seguro Viagem",
    ]:
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
        },
    }

    return cards, fig_pizza, fig_barras, resumo.to_dict("records"), dados_pdf


# ----------------------------------------
# 6. CALLBACK — Tabela de Detalhamento
# ----------------------------------------
@dash.callback(
    Output("tabela_detalhe_passagens", "data"),
    Input("filtro_ano_passagens", "value"),
    Input("filtro_mes_passagens", "value"),
    Input("filtro_unidade_passagens", "value"),
    # também atualiza com o intervalo
    Input("interval-atualizacao", "n_intervals"),
)
def atualizar_detalhe(ano, mes, unidade, n_intervals):
    df = carregar_dados() if n_intervals is not None else df_base
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
    dff["Data Início da Viagem"] = dff[
        "Data Início da Viagem"
    ].dt.strftime("%d/%m/%Y")

    def f(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    dff["Custo com emissão de passagens dentro do prazo"] = dff[
        "Custo com emissão de passagens dentro do prazo"
    ].apply(f)
    dff[
        "Custo com emissão de passagens em caráter de urgência"
    ] = dff[
        "Custo com emissão de passagens em caráter de urgência"
    ].apply(f)

    return dff.to_dict("records")


# ----------------------------------------
# 7. CALLBACK — Limpar filtros
# ----------------------------------------
@dash.callback(
    Output("filtro_ano_passagens", "value"),
    Output("filtro_mes_passagens", "value"),
    Output("filtro_unidade_passagens", "value"),
    Input("btn_limpar_filtros_passagens", "n_clicks"),
    prevent_initial_call=True,
)
def limpar(n):
    return ANO_PADRAO, None, None


# ----------------------------------------
# 8. CALLBACK — Geração do PDF
# ----------------------------------------
wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)


def wrap(text):
    return Paragraph(str(text), wrap_style)


@dash.callback(
    Output("download_relatorio_passagens", "data"),
    Input("btn_download_relatorio_passagens", "n_clicks"),
    State("grafico_pizza_passagens", "figure"),
    State("grafico_barras_passagens", "figure"),
    State("tabela_unidades_passagens", "data"),
    State("tabela_detalhe_passagens", "data"),
    State("store_graficos_passagens", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, fig_pizza, fig_barras, resumo, detalhe, dados_pdf):
    if not n:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Relatório de Gastos com Viagens",
        ParagraphStyle(
            "titulo", fontSize=22, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.3 * inch))

    filtros = dados_pdf["filtros"]
    story.append(
        Paragraph(
            f"Ano: {filtros['ano']} — "
            f"Mês: {filtros['mes'] if filtros['mes'] else 'Todos'} — "
            f"Unidade: {filtros['unidade'] if filtros['unidade'] else 'Todas'}",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 0.3 * inch))

    cards_vals = dados_pdf["cards"]

    def fmt(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    cards_data = [
        ["Total Viagens", fmt(cards_vals["total_viagem"])],
        ["Passagens no Prazo", fmt(cards_vals["total_prazo"])],
        ["Passagens Urgência", fmt(cards_vals["total_urgencia"])],
        ["Gasto em Diárias", fmt(cards_vals["total_diarias"])],
        ["Seguro Viagem", fmt(cards_vals["total_seguro"])],
        ["Restituições", fmt(cards_vals["total_restit"])],
    ]

    tbl_cards = Table(cards_data, colWidths=[3.0 * inch, 3.0 * inch])
    tbl_cards.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#0b2b57")),
            ]
        )
    )
    story.append(tbl_cards)
    story.append(Spacer(1, 0.4 * inch))

    story.append(Paragraph("Resumo por Unidade", styles["Heading2"]))
    table1 = [["Unidade", "Diárias", "Passagem", "Restituição", "Seguro"]]
    for r in resumo:
        table1.append(
            [
                wrap(r["Unidade (Viagem)"]),
                wrap(r["Valor das Diárias"]),
                wrap(r["Valor da Passagem"]),
                wrap(r["Valor Restituição"]),
                wrap(r["Valor Seguro Viagem"]),
            ]
        )

    col_widths1 = [2.8 * inch, 1.0 * inch, 1.0 * inch, 1.0 * inch, 1.0 * inch]
    tbl1 = Table(table1, colWidths=col_widths1)
    tbl1.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
            ]
        )
    )
    story.append(tbl1)
    story.append(Spacer(1, 0.5 * inch))

    story.append(Paragraph("Detalhamento PCDP", styles["Heading2"]))
    table2 = [["Unidade", "PCDP", "Data", "Prazo", "Urgência"]]
    for r in detalhe:
        table2.append(
            [
                wrap(r["Unidade (Viagem)"]),
                wrap(r["Número da PCDP"]),
                wrap(r["Data Início da Viagem"]),
                wrap(
                    r[
                        "Custo com emissão de passagens dentro do prazo"
                    ]
                ),
                wrap(
                    r[
                        "Custo com emissão de passagens em caráter de urgência"
                    ]
                ),
            ]
        )

    col_widths2 = [2.8 * inch, 1.0 * inch, 1.0 * inch, 1.2 * inch, 1.2 * inch]
    tbl2 = Table(table2, colWidths=col_widths2)
    tbl2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003A70")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(tbl2)

    doc.build(story)
    buffer.seek(0)

    from dash import dcc
    return dcc.send_bytes(buffer.getvalue(), "relatorio_gastos_viagens.pdf")
