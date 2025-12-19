# pages/execucao_orcamento_unifei.py
# Painel: Execução do Orçamento - UNIFEI

import dash
from dash import html, dcc, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px

from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors

# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/execucao-orcamento-unifei",
    name="Execução Orçamento UNIFEI",
    title="Execução do Orçamento - UNIFEI",
)

# --------------------------------------------------
# URL da planilha (UNIFEI)
# --------------------------------------------------
URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1MkiWDH-MBnLeSUlqV91qjzCVRTlTAVh9xYooENJ151o/"
    "gviz/tq?tqx=out:csv&sheet=Execucao%20do%20Orcamento%20Unifei"
)

# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados():
    df = pd.read_csv(URL)
    df.columns = [c.strip() for c in df.columns]

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

    col_valores = [
        "DESPESAS INSCRITAS EM RP NAO PROCESSADOS",
        "DESPESAS EMPENHADAS (CONTROLE EMPENHO)",
        "DESPESAS LIQUIDADAS (CONTROLE EMPENHO)",
        "DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)",
        "DESPESAS PAGAS (CONTROLE EMPENHO)",
    ]
    for c in col_valores:
        df[c + "_VAL"] = df[c].apply(conv_moeda)

    return df


df = carregar_dados()
ANO_PADRAO = int(sorted(df["Ano"].dropna().unique())[-1])

dropdown_style = {
    "color": "black",
    "marginBottom": "10px",
    "whiteSpace": "normal",
}

# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.H2(
            "Execução do Orçamento - UNIFEI",
            style={"textAlign": "center"},
        ),

        html.Div(
            style={"marginBottom": "20px"},
            children=[
                html.H3("Filtros", className="sidebar-title"),
                html.Div(
                    style={"display": "flex", "flexWrap": "wrap", "gap": "10px"},
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("UG Executora"),
                                dcc.Dropdown(
                                    id="filtro_ug_exec_unifei",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df["UG Executora"].dropna().unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "150px", "flex": "0 0 180px"},
                            children=[
                                html.Label("Mês"),
                                dcc.Dropdown(
                                    id="filtro_mes_unifei",
                                    options=[
                                        {"label": m, "value": m}
                                        for m in sorted(df["Mês"].dropna().unique())
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "120px", "flex": "0 0 150px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_unifei",
                                    options=[
                                        {"label": int(a), "value": int(a)}
                                        for a in sorted(df["Ano"].dropna().unique())
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Fonte Recursos Detalhada"),
                                dcc.Dropdown(
                                    id="filtro_fonte_unifei",
                                    options=[
                                        {"label": f, "value": f}
                                        for f in sorted(
                                            df["Fonte Recursos Detalhada"]
                                                .dropna()
                                                .unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Grupo Despesa"),
                                dcc.Dropdown(
                                    id="filtro_grupo_unifei",
                                    options=[
                                        {"label": g, "value": g}
                                        for g in sorted(df["GRUPO DESP"].dropna().unique())
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Natureza Despesa"),
                                dcc.Dropdown(
                                    id="filtro_nat_unifei",
                                    options=[
                                        {"label": n, "value": n}
                                        for n in sorted(df["NAT DESP"].dropna().unique())
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
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
                            id="btn_limpar_filtros_unifei",
                            n_clicks=0,
                            className="sidebar-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_unifei",
                            n_clicks=0,
                            className="sidebar-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_unifei"),
                    ],
                ),
            ],
        ),

        html.Div(
            id="cards_container_unifei",
            className="cards-container",
        ),

        html.Div(
            className="charts-row",
            children=[
                dcc.Graph(id="grafico_barras_grupo_unifei", style={"width": "50%"}),
                dcc.Graph(id="grafico_pizza_status_unifei", style={"width": "50%"}),
            ],
        ),

        html.H4("Detalhamento"),
        dash_table.DataTable(
            id="tabela_execucao_unifei",
            columns=[
                {"name": "UG Executora", "id": "UG Executora"},
                {
                    "name": "Fonte Recursos Detalhada",
                    "id": "Fonte Recursos Detalhada",
                },
                {"name": "Grupo Despesa", "id": "GRUPO DESP"},
                {"name": "Natureza Despesa", "id": "Natureza Despesa"},
                {
                    "name": "RP Não Processados",
                    "id": "DESPESAS INSCRITAS EM RP NAO PROCESSADOS",
                },
                {
                    "name": "Empenhadas",
                    "id": "DESPESAS EMPENHADAS (CONTROLE EMPENHO)",
                },
                {
                    "name": "Liquidadas",
                    "id": "DESPESAS LIQUIDADAS (CONTROLE EMPENHO)",
                },
                {
                    "name": "Liquidadas a Pagar",
                    "id": "DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)",
                },
                {
                    "name": "Pagas",
                    "id": "DESPESAS PAGAS (CONTROLE EMPENHO)",
                },
            ],
            data=[],
            style_table={"overflowX": "auto"},
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
                "whiteSpace": "normal",
                "height": "auto",
                "maxWidth": "220px",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
            },
        ),

        dcc.Store(id="store_pdf_unifei"),
    ],
)

# --------------------------------------------------
# Callback principal
# --------------------------------------------------
@dash.callback(
    Output("tabela_execucao_unifei", "data"),
    Output("cards_container_unifei", "children"),
    Output("grafico_barras_grupo_unifei", "figure"),
    Output("grafico_pizza_status_unifei", "figure"),
    Output("store_pdf_unifei", "data"),
    Input("filtro_ug_exec_unifei", "value"),
    Input("filtro_mes_unifei", "value"),
    Input("filtro_ano_unifei", "value"),
    Input("filtro_fonte_unifei", "value"),
    Input("filtro_grupo_unifei", "value"),
    Input("filtro_nat_unifei", "value"),
)
def atualizar_painel(ug_exec, mes, ano, fonte, grupo, nat):
    dff = df.copy()

    if ug_exec:
        dff = dff[dff["UG Executora"] == ug_exec]
    if mes:
        dff = dff[dff["Mês"] == mes]
    if ano:
        dff = dff[dff["Ano"] == ano]
    if fonte:
        dff = dff[dff["Fonte Recursos Detalhada"] == fonte]
    if grupo:
        dff = dff[dff["GRUPO DESP"] == grupo]
    if nat:
        dff = dff[dff["NAT DESP"] == nat]

    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    total_rp = dff["DESPESAS INSCRITAS EM RP NAO PROCESSADOS_VAL"].sum()
    total_emp = dff["DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"].sum()
    total_liq = dff["DESPESAS LIQUIDADAS (CONTROLE EMPENHO)_VAL"].sum()
    total_liq_pagar = dff["DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)_VAL"].sum()
    total_pagas = dff["DESPESAS PAGAS (CONTROLE EMPENHO)_VAL"].sum()

    def card(titulo, valor):
        return html.Div(
            className="card",
            children=[
                html.Div(titulo, className="card-title"),
                html.Div(fmt(valor), className="card-value"),
            ],
        )

    cards = [
        card("RP Não Processados", total_rp),
        card("Empenhadas", total_emp),
        card("Liquidadas", total_liq),
        card("Liquidadas a Pagar", total_liq_pagar),
        card("Pagas", total_pagas),
    ]

    dff_display = dff.copy()
    monetarias = [
        "DESPESAS INSCRITAS EM RP NAO PROCESSADOS",
        "DESPESAS EMPENHADAS (CONTROLE EMPENHO)",
        "DESPESAS LIQUIDADAS (CONTROLE EMPENHO)",
        "DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)",
        "DESPESAS PAGAS (CONTROLE EMPENHO)",
    ]
    for c in monetarias:
        dff_display[c] = dff_display[c + "_VAL"].apply(fmt)

    colunas_tabela = [
        "UG Executora",
        "Fonte Recursos Detalhada",
        "GRUPO DESP",
        "Natureza Despesa",
    ] + monetarias
    dff_display = dff_display[colunas_tabela]

    if not dff.empty:
        grp_grupo = (
            dff.groupby("GRUPO DESP", as_index=False)[
                "DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"
            ]
            .sum()
            .sort_values(
                "DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL", ascending=False
            )
        )
        fig_barras = px.bar(
            grp_grupo,
            x="GRUPO DESP",
            y="DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL",
            title="Despesas Empenhadas por Grupo de Despesa",
        )
        fig_barras.update_traces(
            marker_color="#003A70",
            text=[
                fmt(v)
                for v in grp_grupo[
                    "DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"
                ]
            ],
            textposition="inside",
            insidetextanchor="middle",
            hovertemplate="Grupo=%{x}<br>Empenhadas=R$ %{y:,.2f}<extra></extra>",
        )
        fig_barras.update_layout(
            xaxis_title="Grupo de Despesa",
            yaxis_title="Empenhadas (R$)",
            yaxis_tickprefix="R$ ",
            yaxis_tickformat=",.2f",
            title_y=0.95,
            uniformtext_minsize=10,
            uniformtext_mode="hide",
        )
    else:
        fig_barras = px.bar(title="Sem dados para os filtros selecionados")

    if total_emp + total_liq + total_pagas > 0:
        df_pizza = pd.DataFrame(
            {
                "Status": ["Empenhadas", "Liquidadas", "Pagas"],
                "Valor": [total_emp, total_liq, total_pagas],
            }
        )
        fig_pizza = px.pie(
            df_pizza,
            names="Status",
            values="Valor",
            title="Distribuição: Empenhadas x Liquidadas x Pagas",
            color="Status",
            color_discrete_map={
                "Empenhadas": "#003A70",
                "Liquidadas": "#DA291C",
                "Pagas": "#A2AAAD",
            },
        )
        fig_pizza.update_traces(
            texttemplate="%{label}<br>R$ %{value:,.2f}",
            hovertemplate="%{label}<br>R$ %{value:,.2f}<extra></extra>",
        )
        fig_pizza.update_layout(
            legend_title="Status",
            legend_orientation="h",
            legend_y=-0.1,
            title_y=0.95,
        )
    else:
        fig_pizza = px.pie(
            title="Sem valores para Empenhadas, Liquidadas e Pagas"
        )

    dados_pdf = {
        "tabela": dff_display.to_dict("records"),
        "totais": {
            "rp": float(total_rp),
            "emp": float(total_emp),
            "liq": float(total_liq),
            "liq_pagar": float(total_liq_pagar),
            "pagas": float(total_pagas),
        },
        "filtros": {
            "ug_exec": ug_exec,
            "mes": mes,
            "ano": ano,
            "fonte": fonte,
            "grupo": grupo,
            "nat": nat,
        },
    }

    return dff_display.to_dict("records"), cards, fig_barras, fig_pizza, dados_pdf

# --------------------------------------------------
# Limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_ug_exec_unifei", "value"),
    Output("filtro_mes_unifei", "value"),
    Output("filtro_ano_unifei", "value"),
    Output("filtro_fonte_unifei", "value"),
    Output("filtro_grupo_unifei", "value"),
    Output("filtro_nat_unifei", "value"),
    Input("btn_limpar_filtros_unifei", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros(n):
    return None, None, ANO_PADRAO, None, None, None

# --------------------------------------------------
# PDF
# --------------------------------------------------
wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=7,
    leading=8,
    spaceAfter=2,
    alignment=TA_LEFT,
)

def wrap(text):
    return Paragraph(str(text)[:200], wrap_style)

@dash.callback(
    Output("download_relatorio_unifei", "data"),
    Input("btn_download_relatorio_unifei", "n_clicks"),
    State("store_pdf_unifei", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, dados_pdf):
    if not n or not dados_pdf:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Relatório de Execução do Orçamento - UNIFEI",
        ParagraphStyle(
            "titulo", fontSize=18, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.15 * inch))

    f = dados_pdf["filtros"]
    story.append(
        Paragraph(
            f"<b>UG Executora:</b> {f['ug_exec'] if f['ug_exec'] else 'Todas'} | "
            f"<b>Mês:</b> {f['mes'] if f['mes'] else 'Todos'} | "
            f"<b>Ano:</b> {f['ano'] if f['ano'] else 'Todos'}",
            ParagraphStyle("filtros", fontSize=7, alignment=TA_LEFT),
        )
    )
    story.append(
        Paragraph(
            f"<b>Fonte:</b> {f['fonte'] if f['fonte'] else 'Todas'} | "
            f"<b>Grupo:</b> {f['grupo'] if f['grupo'] else 'Todos'} | "
            f"<b>Natureza:</b> {f['nat'] if f['nat'] else 'Todas'}",
            ParagraphStyle("filtros", fontSize=7, alignment=TA_LEFT),
        )
    )
    story.append(Spacer(1, 0.15 * inch))

    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    tot = dados_pdf["totais"]
    cards_data = [
        ["RP Não Processados", fmt(tot["rp"])],
        ["Empenhadas", fmt(tot["emp"])],
        ["Liquidadas", fmt(tot["liq"])],
        ["Liquidadas a Pagar", fmt(tot["liq_pagar"])],
        ["Pagas", fmt(tot["pagas"])],
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
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(tbl_cards)
    story.append(Spacer(1, 0.2 * inch))

    table_data = [
        [
            "UG Executora",
            "Fonte Recursos",
            "Grupo Despesa",
            "Natureza Despesa",
            "RP Não Proc.",
            "Empenhadas",
            "Liquidadas",
            "Liq. a Pagar",
            "Pagas",
        ]
    ]
    for r in dados_pdf["tabela"]:
        table_data.append(
            [
                wrap(r["UG Executora"]),
                wrap(r["Fonte Recursos Detalhada"]),
                wrap(r["GRUPO DESP"]),
                wrap(r["Natureza Despesa"]),
                wrap(r["DESPESAS INSCRITAS EM RP NAO PROCESSADOS"]),
                wrap(r["DESPESAS EMPENHADAS (CONTROLE EMPENHO)"]),
                wrap(r["DESPESAS LIQUIDADAS (CONTROLE EMPENHO)"]),
                wrap(r["DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)"]),
                wrap(r["DESPESAS PAGAS (CONTROLE EMPENHO)"]),
            ]
        )

    col_widths = [
        1.6 * inch,
        1.8 * inch,
        1.5 * inch,
        1.8 * inch,
        0.9 * inch,
        0.9 * inch,
        0.9 * inch,
        0.9 * inch,
        0.9 * inch,
    ]
    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, 0), 6),
                ("FONTSIZE", (0, 1), (-1, -1), 6),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                 [colors.white, colors.HexColor("#f5f5f5")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(tbl)

    doc.build(story)
    buffer.seek(0)
    from dash import dcc

    return dcc.send_bytes(buffer.getvalue(), "execucao_orcamento_unifei.pdf")
