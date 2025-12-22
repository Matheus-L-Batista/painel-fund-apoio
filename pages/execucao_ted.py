# pages/execucao_ted.py

# Painel: Execução do Orçamento - TED

import dash
from dash import html, dcc, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors
import datetime as dt


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/execucao-ted",
    name="Execução TED",
    title="Execução do Orçamento - TED",
)


# --------------------------------------------------
# URL da planilha (TED)
# --------------------------------------------------
URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1MkiWDH-MBnLeSUlqV91qjzCVRTlTAVh9xYooENJ151o/"
    "gviz/tq?tqx=out:csv&sheet=Execucao%20do%20Orcamento%20TED"
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


# DF base global (apenas cache, recarregado pelo Interval)
df_base = carregar_dados()
ANO_PADRAO = int(sorted(df_base["Ano"].dropna().unique())[-1])

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
            "Execução do Orçamento - TED",
            style={"textAlign": "center"},
        ),
        html.Div(
            style={"marginBottom": "20px"},
            children=[
                html.H3("Filtros", className="sidebar-title"),

                # 1ª linha: UO, UG, Ano, Mês
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "marginBottom": "8px",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Unidade Orçamentária"),
                                dcc.Dropdown(
                                    id="filtro_uo_ted",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df_base["Unidade Orçamentária"]
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
                                html.Label("UG Executora"),
                                dcc.Dropdown(
                                    id="filtro_ug_exec_ted",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df_base["UG EXEC"]
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
                            style={"minWidth": "120px", "flex": "0 0 150px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_ted",
                                    options=[
                                        {"label": int(a), "value": int(a)}
                                        for a in sorted(
                                            df_base["Ano"].dropna().unique()
                                        )
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "150px", "flex": "0 0 180px"},
                            children=[
                                html.Label("Mês"),
                                dcc.Dropdown(
                                    id="filtro_mes_ted",
                                    options=[
                                        {"label": m, "value": m}
                                        for m in sorted(
                                            df_base["Mês"].dropna().unique()
                                        )
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

                # 2ª linha: Fonte, Grupo, Natureza + botões
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-end",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Fonte Recursos Detalhada"),
                                dcc.Dropdown(
                                    id="filtro_fonte_ted",
                                    options=[
                                        {"label": f, "value": f}
                                        for f in sorted(
                                            df_base["FRD"].dropna().unique()
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
                                html.Label("Grupo da Despesa"),
                                dcc.Dropdown(
                                    id="filtro_grupo_ted",
                                    options=[
                                        {"label": g, "value": g}
                                        for g in sorted(
                                            df_base["GRUPO DESP"]
                                            .dropna()
                                            .unique()
                                        )
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
                                    id="filtro_nat_ted",
                                    options=[
                                        {"label": n, "value": n}
                                        for n in sorted(
                                            df_base["NAT DESP"]
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
                            style={
                                "display": "flex",
                                "gap": "10px",
                                "marginTop": "24px",
                            },
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_ted",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_ted",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                dcc.Download(id="download_relatorio_ted"),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        html.Div(
            id="cards_container_ted",
            className="cards-container",
        ),
        html.Div(
            className="charts-row",
            children=[
                dcc.Graph(
                    id="grafico_barras_grupo_ted", style={"width": "50%"}
                ),
                dcc.Graph(
                    id="grafico_pizza_status_ted", style={"width": "50%"}
                ),
            ],
        ),
        html.H4("Detalhamento"),
        dash_table.DataTable(
            id="tabela_execucao_ted",
            columns=[
                {
                    "name": "Unidade Orçamentária",
                    "id": "Unidade Orçamentária",
                },
                {
                    "name": "Fonte Recursos Detalhada",
                    "id": "Fonte Recursos Detalhada",
                },
                {"name": "Grupo da Despesa", "id": "GRUPO DESP"},
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
        dcc.Store(id="store_pdf_ted"),
    ],
)


# --------------------------------------------------
# Callback principal
# --------------------------------------------------
@dash.callback(
    Output("tabela_execucao_ted", "data"),
    Output("cards_container_ted", "children"),
    Output("grafico_barras_grupo_ted", "figure"),
    Output("grafico_pizza_status_ted", "figure"),
    Output("store_pdf_ted", "data"),
    Input("filtro_uo_ted", "value"),
    Input("filtro_ug_exec_ted", "value"),
    Input("filtro_ano_ted", "value"),
    Input("filtro_mes_ted", "value"),
    Input("filtro_fonte_ted", "value"),
    Input("filtro_grupo_ted", "value"),
    Input("filtro_nat_ted", "value"),
    Input("interval-atualizacao", "n_intervals"),  # novo Input
)
def atualizar_painel(uo, ugexec, ano, mes, fonte, grupo, nat, n_intervals):
    global df_base

    # Atualiza o df_base somente em um horário permitido (exemplo: 08h–20h)
    agora = dt.datetime.now().time()
    if dt.time(8, 0) <= agora <= dt.time(20, 0):
        if n_intervals is not None:
            df_base = carregar_dados()

    dff = df_base.copy()

    if uo:
        dff = dff[dff["Unidade Orçamentária"] == uo]
    if ugexec:
        dff = dff[dff["UG EXEC"] == ugexec]
    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes:
        dff = dff[dff["Mês"] == mes]
    if fonte:
        dff = dff[dff["FRD"] == fonte]
    if grupo:
        dff = dff[dff["GRUPO DESP"] == grupo]
    if nat:
        dff = dff[dff["NAT DESP"] == nat]

    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    total_rp = dff[
        "DESPESAS INSCRITAS EM RP NAO PROCESSADOS_VAL"
    ].sum()
    total_emp = dff["DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"].sum()
    total_liq = dff["DESPESAS LIQUIDADAS (CONTROLE EMPENHO)_VAL"].sum()
    total_liq_pagar = dff[
        "DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)_VAL"
    ].sum()
    total_pagas = dff["DESPESAS PAGAS (CONTROLE EMPENHO)_VAL"].sum()

    cards = [
        html.Div(
            className="card",
            children=[
                html.Div("RP Não Processados", className="card-title"),
                html.Div(fmt(total_rp), className="card-value"),
            ],
        ),
        html.Div(
            className="card",
            children=[
                html.Div("Empenhadas", className="card-title"),
                html.Div(fmt(total_emp), className="card-value"),
            ],
        ),
        html.Div(
            className="card",
            children=[
                html.Div("Liquidadas", className="card-title"),
                html.Div(fmt(total_liq), className="card-value"),
            ],
        ),
        html.Div(
            className="card",
            children=[
                html.Div("Liquidadas a Pagar", className="card-title"),
                html.Div(fmt(total_liq_pagar), className="card-value"),
            ],
        ),
        html.Div(
            className="card",
            children=[
                html.Div("Pagas", className="card-title"),
                html.Div(fmt(total_pagas), className="card-value"),
            ],
        ),
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
        "Unidade Orçamentária",
        "Fonte Recursos Detalhada",
        "GRUPO DESP",
        "Natureza Despesa",
    ] + monetarias
    dff_display = dff_display[colunas_tabela]

    # -----------------------------
    # GRÁFICO DE BARRAS POR GRUPO
    # -----------------------------
    if not dff.empty:
        grp_grupo = (
            dff.groupby("GRUPO DESP", as_index=False)[
                "DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"
            ]
            .sum()
            .sort_values(
                "DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL",
                ascending=False,
            )
        )

        valores = grp_grupo["DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL"].values
        limiar = 0.2 * valores.max() if valores.size > 0 else 0
        textpositions = [
            "inside" if v >= limiar else "outside"
            for v in valores
        ]

        fig_barras = px.bar(
            grp_grupo,
            x="GRUPO DESP",
            y="DESPESAS EMPENHADAS (CONTROLE EMPENHO)_VAL",
            title="Despesas Empenhadas por Grupo de Despesa",
        )
        fig_barras.update_traces(
            marker_color="#003A70",
            text=[fmt(v) for v in valores],
            textposition=textpositions,
            insidetextanchor="middle",
            hovertemplate="Grupo=%{x}<br>Empenhadas=R$ %{y:,.2f}",
            cliponaxis=False,
        )
        fig_barras.update_layout(
            xaxis_title="Grupo de Despesa",
            yaxis_title="Empenhadas (R$)",
            yaxis_tickprefix="R$ ",
            yaxis_tickformat=",.2f",
            title_x=0.5,
            title_y=0.9,
            uniformtext_minsize=10,
            uniformtext_mode="hide",
        )
    else:
        fig_barras = px.bar(
            title="Sem dados para os filtros selecionados"
        )
        fig_barras.update_layout(title_x=0.5, title_y=0.9)

    # -----------------------------
    # GRÁFICO DE PIZZA STATUS
    # -----------------------------
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
            hovertemplate="%{label}<br>R$ %{value:,.2f}",
        )
        fig_pizza.update_layout(
            legend_title="Status",
            legend_orientation="h",
            legend_y=-0.1,
            legend_x=0.5,
            legend_xanchor="center",
            title_x=0.5,
            title_y=0.9,
        )
    else:
        fig_pizza = px.pie(
            title="Sem valores para Empenhadas, Liquidadas e Pagas"
        )
        fig_pizza.update_layout(title_x=0.5, title_y=0.9)

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
            "uo": uo,
            "ugexec": ugexec,
            "ano": ano,
            "mes": mes,
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
    Output("filtro_uo_ted", "value"),
    Output("filtro_ug_exec_ted", "value"),
    Output("filtro_ano_ted", "value"),
    Output("filtro_mes_ted", "value"),
    Output("filtro_fonte_ted", "value"),
    Output("filtro_grupo_ted", "value"),
    Output("filtro_nat_ted", "value"),
    Input("btn_limpar_filtros_ted", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros(n):
    return None, None, ANO_PADRAO, None, None, None, None


# --------------------------------------------------
# PDF
# --------------------------------------------------
wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=5,
    leading=6,
    spaceAfter=0,
    alignment=TA_LEFT,
)


def wrap(text):
    return Paragraph(str(text)[:150], wrap_style)


@dash.callback(
    Output("download_relatorio_ted", "data"),
    Input("btn_download_relatorio_ted", "n_clicks"),
    State("store_pdf_ted", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, dados_pdf):
    if not n or not dados_pdf:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        topMargin=0.3 * inch,
        bottomMargin=0.3 * inch,
        leftMargin=0.3 * inch,
        rightMargin=0.3 * inch,
    )
    styles = getSampleStyleSheet()
    story = []

    # Título
    titulo = Paragraph(
        "Relatório de Execução do Orçamento - TED",
        ParagraphStyle(
            "titulo", fontSize=14, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.08 * inch))

    # Filtros
    f = dados_pdf["filtros"]
    story.append(
        Paragraph(
            f"UO: {f['uo'] if f['uo'] else 'Todas'} | "
            f"UG Exec: {f['ugexec'] if f['ugexec'] else 'Todas'} | "
            f"Ano: {f['ano'] if f['ano'] else 'Todos'} | "
            f"Mês: {f['mes'] if f['mes'] else 'Todos'}",
            ParagraphStyle("filtros", fontSize=6, alignment=TA_LEFT),
        )
    )
    story.append(
        Paragraph(
            f"Fonte: {f['fonte'] if f['fonte'] else 'Todas'} | "
            f"Grupo: {f['grupo'] if f['grupo'] else 'Todos'} | "
            f"Natureza: {f['nat'] if f['nat'] else 'Todas'}",
            ParagraphStyle("filtros", fontSize=6, alignment=TA_LEFT),
        )
    )
    story.append(Spacer(1, 0.08 * inch))

    # Cards/Totais
    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    tot = dados_pdf["totais"]
    cards_data = [
        ["RP Não Proc.", fmt(tot["rp"])],
        ["Empenhadas", fmt(tot["emp"])],
        ["Liquidadas", fmt(tot["liq"])],
        ["Liq. a Pagar", fmt(tot["liq_pagar"])],
        ["Pagas", fmt(tot["pagas"])],
    ]

    tbl_cards = Table(cards_data, colWidths=[1.5 * inch, 1.5 * inch])
    tbl_cards.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#0b2b57")),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(tbl_cards)
    story.append(Spacer(1, 0.1 * inch))

    # Tabela detalhada
    table_data = [
        [
            "UO",
            "Fonte<br>Recursos",
            "Grupo<br>Despesa",
            "Natureza<br>Despesa",
            "RP N.P.",
            "Empenha.",
            "Liquida.",
            "Liq. Pagar",
            "Pagas",
        ]
    ]

    for r in dados_pdf["tabela"]:
        table_data.append(
            [
                wrap(r["Unidade Orçamentária"]),
                wrap(r["Fonte Recursos Detalhada"]),
                wrap(r["GRUPO DESP"]),
                wrap(r["Natureza Despesa"]),
                wrap(r["DESPESAS INSCRITAS EM RP NAO PROCESSADOS"]),
                wrap(r["DESPESAS EMPENHADAS (CONTROLE EMPENHO)"]),
                wrap(r["DESPESAS LIQUIDADAS (CONTROLE EMPENHO)"]),
                wrap(
                    r["DESPESAS LIQUIDADAS A PAGAR(CONTROLE EMPENHO)"]
                ),
                wrap(r["DESPESAS PAGAS (CONTROLE EMPENHO)"]),
            ]
        )

    # Larguras otimizadas para landscape
    col_widths = [
        1.2 * inch,  # UO
        1.4 * inch,  # Fonte Recursos
        1.1 * inch,  # Grupo Desp
        1.3 * inch,  # Natureza Desp
        0.75 * inch, # RP N.P.
        0.75 * inch, # Empenha.
        0.75 * inch, # Liquida.
        0.8 * inch,  # Liq. Pagar
        0.75 * inch, # Pagas
    ]

    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, 0), 5),
                ("FONTSIZE", (0, 1), (-1, -1), 5),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, colors.HexColor("#f9f9f9")],
                ),
                ("LEFTPADDING", (0, 0), (-1, -1), 1),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )

    story.append(tbl)
    doc.build(story)
    buffer.seek(0)

    from dash import dcc
    return dcc.send_bytes(buffer.getvalue(), "execucao_orcamento_ted.pdf")
