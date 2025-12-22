# pages/dotacao.py


# Painel: Dotação Atualizada e Destaques Recebidos


import dash
from dash import html, dcc, Input, Output, State, dash_table
import pandas as pd
import plotly.express as px
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors  # [file:4]
import datetime as dt


# --------------------------------------------------
# Registro da página no Dash Pages
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/dotacao",
    name="Dotação Atualizada",
    title="Dotação Atualizada e Destaques",
)


# --------------------------------------------------
# 1. Dados
# --------------------------------------------------
URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1MkiWDH-MBnLeSUlqV91qjzCVRTlTAVh9xYooENJ151o/"
    "gviz/tq?tqx=out:csv&sheet=Dotacao%20Atualizada%20e%20Destaques%20Recebidos"
)


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

    df["DOTACAO ATUALIZADA_VAL"] = df["DOTACAO ATUALIZADA"].apply(conv_moeda)
    df["DESTAQUE RECEBIDO_VAL"] = df["DESTAQUE RECEBIDO"].apply(conv_moeda)
    return df  # [file:4]


# >>> DF base em vez de df global
df_base = carregar_dados()
ANO_PADRAO = int(sorted(df_base["ANO"].dropna().unique())[-1])


# --------------------------------------------------
# 2. Layout da página (só conteúdo principal)
# --------------------------------------------------
layout = html.Div(
    children=[
        html.H2(
            "Dotação Atualizada e Destaques Recebidos",
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
                                html.Label("Grupo da Despesa"),
                                dcc.Dropdown(
                                    id="filtro_grupo_dotacao",
                                    options=[
                                        {"label": g, "value": g}
                                        for g in sorted(
                                            df_base["GRUPO DA DESPESA"]
                                            .dropna()
                                            .unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style={
                                        "color": "black",
                                        "marginBottom": "10px",
                                        "whiteSpace": "normal",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "120px", "flex": "0 0 150px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_dotacao",
                                    options=[
                                        {"label": int(a), "value": int(a)}
                                        for a in sorted(
                                            df_base["ANO"].dropna().unique()
                                        )
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style={
                                        "color": "black",
                                        "marginBottom": "10px",
                                        "whiteSpace": "normal",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Unidade Orçamentária"),
                                dcc.Dropdown(
                                    id="filtro_unidade_dotacao",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df_base["UNIDADE ORÇAMENTÁRIA"]
                                            .dropna()
                                            .unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style={
                                        "color": "black",
                                        "marginBottom": "10px",
                                        "whiteSpace": "normal",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Fonte Recursos Detalhada"),
                                dcc.Dropdown(
                                    id="filtro_fonte_dotacao",
                                    options=[
                                        {"label": f, "value": f}
                                        for f in sorted(
                                            df_base["Fonte Recursos Detalhada"]
                                            .dropna()
                                            .unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style={
                                        "color": "black",
                                        "marginBottom": "10px",
                                        "whiteSpace": "normal",
                                    },
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
                            id="btn_limpar_filtros_dotacao",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_dotacao",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_dotacao"),
                    ],
                ),
            ],
        ),
        html.Div(
            id="cards_container_dotacao",
            className="cards-container",
        ),
        html.Div(
            className="charts-row",
            children=[
                dcc.Graph(id="grafico_pizza_dotacao", style={"width": "50%"}),
                dcc.Graph(id="grafico_pizza_destaque", style={"width": "50%"}),
            ],
        ),
        html.Div(
            className="charts-row",
            children=[
                dcc.Graph(
                    id="grafico_barra_dotacao_fonte", style={"width": "50%"}
                ),
                dcc.Graph(
                    id="grafico_barra_destaque_fonte", style={"width": "50%"}
                ),
            ],
        ),
        html.H4("Detalhamento"),
        dash_table.DataTable(
            id="tabela_dotacao",
            columns=[
                {"name": "GRUPO DA DESPESA", "id": "GRUPO DA DESPESA"},
                {"name": "ANO", "id": "ANO"},
                {
                    "name": "UNIDADE ORÇAMENTÁRIA",
                    "id": "UNIDADE ORÇAMENTÁRIA",
                },
                {
                    "name": "Fonte Recursos Detalhada",
                    "id": "Fonte Recursos Detalhada",
                },
                {"name": "DOTACAO ATUALIZADA", "id": "DOTACAO ATUALIZADA"},
                {"name": "DESTAQUE RECEBIDO", "id": "DESTAQUE RECEBIDO"},
            ],
            data=[],
            style_table={"overflowX": "auto"},
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
            },
        ),
        dcc.Store(id="store_pdf_dotacao"),
    ],
)


# --------------------------------------------------
# 3. Callback principal
# --------------------------------------------------
@dash.callback(
    Output("tabela_dotacao", "data"),
    Output("cards_container_dotacao", "children"),
    Output("grafico_pizza_dotacao", "figure"),
    Output("grafico_pizza_destaque", "figure"),
    Output("grafico_barra_dotacao_fonte", "figure"),
    Output("grafico_barra_destaque_fonte", "figure"),
    Output("store_pdf_dotacao", "data"),
    Input("filtro_grupo_dotacao", "value"),
    Input("filtro_ano_dotacao", "value"),
    Input("filtro_unidade_dotacao", "value"),
    Input("filtro_fonte_dotacao", "value"),
    Input("interval-atualizacao", "n_intervals"),  # novo Input
)
def atualizar_painel(grupo, ano, unidade, fonte, n_intervals):
    global df_base

    # Atualiza df_base somente em horário permitido (exemplo: 08h–20h)
    agora = dt.datetime.now().time()
    if dt.time(8, 0) <= agora <= dt.time(20, 0):
        if n_intervals is not None:
            df_base = carregar_dados()

    dff = df_base.copy()

    if ano:
        dff = dff[dff["ANO"] == ano]
    if grupo:
        dff = dff[dff["GRUPO DA DESPESA"] == grupo]
    if unidade:
        dff = dff[dff["UNIDADE ORÇAMENTÁRIA"] == unidade]
    if fonte:
        dff = dff[dff["Fonte Recursos Detalhada"] == fonte]

    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    total_dotacao = dff["DOTACAO ATUALIZADA_VAL"].sum()
    total_destaque = dff["DESTAQUE RECEBIDO_VAL"].sum()

    cards = [
        html.Div(
            className="card",
            children=[
                html.Div("Dotação Atualizada", className="card-title"),
                html.Div(fmt(total_dotacao), className="card-value"),
            ],
        ),
        html.Div(
            className="card",
            children=[
                html.Div("Destaques Recebidos", className="card-title"),
                html.Div(fmt(total_destaque), className="card-value"),
            ],
        ),
    ]

    dff_display = dff.copy()
    dff_display["DOTACAO ATUALIZADA"] = dff_display[
        "DOTACAO ATUALIZADA_VAL"
    ].apply(fmt)
    dff_display["DESTAQUE RECEBIDO"] = dff_display[
        "DESTAQUE RECEBIDO_VAL"
    ].apply(fmt)

    colunas = [
        "GRUPO DA DESPESA",
        "ANO",
        "UNIDADE ORÇAMENTÁRIA",
        "Fonte Recursos Detalhada",
        "DOTACAO ATUALIZADA",
        "DESTAQUE RECEBIDO",
    ]
    dff_display = dff_display[colunas]

    if not dff.empty:
        grp_dot_grupo = dff.groupby(
            "GRUPO DA DESPESA", as_index=False
        )["DOTACAO ATUALIZADA_VAL"].sum()
        fig_pizza_dot = px.pie(
            grp_dot_grupo,
            names="GRUPO DA DESPESA",
            values="DOTACAO ATUALIZADA_VAL",
            title="Dotação Atualizada por Grupo de Despesa",
        )
        fig_pizza_dot.update_traces(
            texttemplate="%{label}<br>R$ %{value:,.2f}",
            hovertemplate="%{label}<br>R$ %{value:,.2f}",
            marker=dict(colors=["#003A70", "#DA291C", "#A2AAAD"]),
        )
        fig_pizza_dot.update_layout(
            legend_title="Grupo da Despesa",
            legend_orientation="h",
            legend_y=-0.1,
            title_y=0.95,
        )
    else:
        fig_pizza_dot = px.pie(
            title="Sem dados para os filtros selecionados"
        )

    if not dff.empty:
        grp_des_grupo = dff.groupby(
            "GRUPO DA DESPESA", as_index=False
        )["DESTAQUE RECEBIDO_VAL"].sum()
        fig_pizza_des = px.pie(
            grp_des_grupo,
            names="GRUPO DA DESPESA",
            values="DESTAQUE RECEBIDO_VAL",
            title="Destaques Recebidos por Grupo de Despesa",
        )
        fig_pizza_des.update_traces(
            texttemplate="%{label}<br>R$ %{value:,.2f}",
            hovertemplate="%{label}<br>R$ %{value:,.2f}",
            marker=dict(colors=["#003A70", "#DA291C", "#A2AAAD"]),
        )
        fig_pizza_des.update_layout(
            legend_title="Grupo da Despesa",
            legend_orientation="h",
            legend_y=-0.1,
            title_y=0.95,
        )
    else:
        fig_pizza_des = px.pie(
            title="Sem dados para os filtros selecionados"
        )

    def texto_posicoes(valores):
        max_v = max(valores) if len(valores) else 0
        posicoes = []
        for v in valores:
            if max_v > 0 and v >= 0.3 * max_v:
                posicoes.append("inside")
            else:
                posicoes.append("outside")
        return posicoes

    if not dff.empty:
        grp_dot_fonte = dff.groupby(
            "Fonte Recursos Detalhada", as_index=False
        )["DOTACAO ATUALIZADA_VAL"].sum()
        fig_bar_dot = px.bar(
            grp_dot_fonte,
            x="DOTACAO ATUALIZADA_VAL",
            y="Fonte Recursos Detalhada",
            orientation="h",
            title="Dotação Atualizada por Fonte de Recursos Detalhada",
        )
        valores = grp_dot_fonte["DOTACAO ATUALIZADA_VAL"].tolist()
        posicoes = texto_posicoes(valores)
        fig_bar_dot.update_traces(
            marker_color="#003A70",
            hovertemplate="Fonte=%{y}<br>Dotação=R$ %{x:,.2f}",
            text=[fmt(v) for v in valores],
            textposition=posicoes,
            textfont_color="white",
        )
        fig_bar_dot.update_layout(
            xaxis_title="Dotação Atualizada (R$)",
            yaxis_title="Fonte Recursos Detalhada",
            xaxis_tickprefix="R$ ",
            xaxis_tickformat=",.2f",
            title_y=0.95,
            margin=dict(l=200, r=80, t=80, b=60),
            plot_bgcolor="#748092",
            paper_bgcolor="white",
            font_color="black",
        )
    else:
        fig_bar_dot = px.bar(
            title="Sem dados para os filtros selecionados"
        )

    if not dff.empty:
        grp_des_fonte = dff.groupby(
            "Fonte Recursos Detalhada", as_index=False
        )["DESTAQUE RECEBIDO_VAL"].sum()
        fig_bar_des = px.bar(
            grp_des_fonte,
            x="DESTAQUE RECEBIDO_VAL",
            y="Fonte Recursos Detalhada",
            orientation="h",
            title="Destaques Recebidos por Fonte de Recursos Detalhada",
        )
        valores_des = grp_des_fonte["DESTAQUE RECEBIDO_VAL"].tolist()
        posicoes_des = texto_posicoes(valores_des)
        fig_bar_des.update_traces(
            marker_color="#DA291C",
            hovertemplate="Fonte=%{y}<br>Destaque=R$ %{x:,.2f}",
            text=[fmt(v) for v in valores_des],
            textposition=posicoes_des,
            textfont_color="white",
        )
        fig_bar_des.update_layout(
            xaxis_title="Destaques Recebidos (R$)",
            yaxis_title="Fonte Recursos Detalhada",
            xaxis_tickprefix="R$ ",
            xaxis_tickformat=",.2f",
            title_y=0.95,
            margin=dict(l=200, r=80, t=80, b=60),
            plot_bgcolor="#748092",
            paper_bgcolor="white",
            font_color="black",
        )
    else:
        fig_bar_des = px.bar(
            title="Sem dados para os filtros selecionados"
        )

    dados_pdf = {
        "tabela": dff_display.to_dict("records"),
        "total_dotacao": float(total_dotacao),
        "total_destaque": float(total_destaque),
        "filtros": {
            "grupo": grupo,
            "ano": ano,
            "unidade": unidade,
            "fonte": fonte,
        },
    }

    return (
        dff_display.to_dict("records"),
        cards,
        fig_pizza_dot,
        fig_pizza_des,
        fig_bar_dot,
        fig_bar_des,
        dados_pdf,
    )


# --------------------------------------------------
# 4. Limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_ano_dotacao", "value"),
    Output("filtro_grupo_dotacao", "value"),
    Output("filtro_unidade_dotacao", "value"),
    Output("filtro_fonte_dotacao", "value"),
    Input("btn_limpar_filtros_dotacao", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros(n):
    return ANO_PADRAO, None, None, None


# --------------------------------------------------
# 5. PDF (cards + tabela)
# --------------------------------------------------
wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)


def wrap(text):
    return Paragraph(str(text), wrap_style)


@dash.callback(
    Output("download_relatorio_dotacao", "data"),
    Input("btn_download_relatorio_dotacao", "n_clicks"),
    State("store_pdf_dotacao", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, dados_pdf):
    if not n or not dados_pdf:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Relatório de Dotação e Destaques",
        ParagraphStyle(
            "titulo", fontSize=20, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.25 * inch))

    f = dados_pdf["filtros"]
    story.append(
        Paragraph(
            f"Ano: {f['ano'] if f['ano'] else 'Todos'} — "
            f"Grupo: {f['grupo'] if f['grupo'] else 'Todos'} — "
            f"Unidade: {f['unidade'] if f['unidade'] else 'Todas'} — "
            f"Fonte: {f['fonte'] if f['fonte'] else 'Todas'}",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 0.25 * inch))

    def fmt(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    cards_data = [
        ["Dotação Atualizada", fmt(dados_pdf["total_dotacao"])],
        ["Destaques Recebidos", fmt(dados_pdf["total_destaque"])],
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
                ("FONTSIZE", (0, 0), (-1, -1), 10),
            ]
        )
    )
    story.append(tbl_cards)
    story.append(Spacer(1, 0.35 * inch))

    table_data = [["GRUPO", "ANO", "UNIDADE", "FONTE", "DOTACAO", "DESTAQUE"]]
    for r in dados_pdf["tabela"]:
        table_data.append(
            [
                wrap(r["GRUPO DA DESPESA"]),
                wrap(r["ANO"]),
                wrap(r["UNIDADE ORÇAMENTÁRIA"]),
                wrap(r["Fonte Recursos Detalhada"]),
                wrap(r["DOTACAO ATUALIZADA"]),
                wrap(r["DESTAQUE RECEBIDO"]),
            ]
        )

    col_widths = [
        1.2 * inch,
        0.6 * inch,
        2.0 * inch,
        2.5 * inch,
        1.0 * inch,
        1.0 * inch,
    ]
    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )

    story.append(tbl)
    doc.build(story)
    buffer.seek(0)

    from dash import dcc
    return dcc.send_bytes(buffer.getvalue(), "dotacao_destaques.pdf")
