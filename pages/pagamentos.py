# Painel: Pagamentos Efetivados

import dash
from dash import html, dcc, Input, Output, State, dash_table
import pandas as pd
from datetime import datetime
from io import BytesIO
import plotly.express as px
from reportlab.lib.pagesizes import letter
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
    path="/pagamentos",
    name="Pagamentos Efetivados",
    title="Pagamentos Efetivados",
)


URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1KEEohPamH36URHpPjFjpVmSNOoK3429erayoPv6fcDo/"
    "gviz/tq?tqx=out:csv&sheet=Pagamentos%20Efetivados"
)


# ----------------------------------------
# 2. CARGA E TRATAMENTO DOS DADOS
# ----------------------------------------
def carregar_dados():
    df = pd.read_csv(URL)
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(
        columns={
            "Unnamed: 2": "DT ATESTE",
            "Unnamed: 3": "DT PGTO",
        }
    )
    df["DT ATESTE"] = pd.to_datetime(
        df["DT ATESTE"], format="%d/%m/%Y", errors="coerce"
    )
    df["DT PGTO"] = pd.to_datetime(
        df["DT PGTO"], format="%d/%m/%Y", errors="coerce"
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

    df["Valor"] = df["Valor"].apply(conv_moeda)

    mapa_meses = {
        "JANEIRO": 1,
        "FEVEREIRO": 2,
        "MARÇO": 3,
        "MARCO": 3,
        "ABRIL": 4,
        "MAIO": 5,
        "JUNHO": 6,
        "JULHO": 7,
        "AGOSTO": 8,
        "SETEMBRO": 9,
        "OUTUBRO": 10,
        "NOVEMBRO": 11,
        "DEZEMBRO": 12,
    }

    df["Ano"] = df["ANO"].astype(int)
    df["Mes"] = df["MÊS"].astype(str).str.upper().map(mapa_meses)
    return df


# 🔧 B) DF base inicial
df_base = carregar_dados()
ANO_PADRAO = int(sorted(df_base["Ano"].dropna().unique())[-1])


# ----------------------------------------
# 3. LISTA DE MESES (para o dropdown)
# ----------------------------------------
nomes_meses = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
]

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "10px",
    "whiteSpace": "normal",
}


# ----------------------------------------
# 4. LAYOUT DA PÁGINA (somente conteúdo)
# ----------------------------------------
layout = html.Div(
    children=[
        html.H2(
            "Pagamentos Efetivados",
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
                            style={"minWidth": "140px", "flex": "0 0 160px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_pagamentos",
                                    options=[
                                        {"label": int(a), "value": int(a)}
                                        for a in sorted(
                                            df_base["Ano"].dropna().unique()
                                        )
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style=dropdown_style,
                                    optionHeight=40,
                                    maxHeight=400,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "140px", "flex": "0 0 160px"},
                            children=[
                                html.Label("Mês"),
                                dcc.Dropdown(
                                    id="filtro_mes_pagamentos",
                                    options=[
                                        {"label": m.capitalize(), "value": i}
                                        for i, m in enumerate(
                                            nomes_meses, start=1
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                    optionHeight=40,
                                    maxHeight=400,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1"},
                            children=[
                                html.Label("Lista"),
                                dcc.Dropdown(
                                    id="filtro_lista_pagamentos",
                                    options=[
                                        {"label": u, "value": u}
                                        for u in sorted(
                                            df_base["LISTAS"].dropna().unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                    optionHeight=50,
                                    maxHeight=400,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "160px", "flex": "1"},
                            children=[
                                html.Label("Fonte"),
                                dcc.Dropdown(
                                    id="filtro_fonte_pagamentos",
                                    options=[
                                        {"label": str(u), "value": str(u)}
                                        for u in sorted(
                                            df_base["FONTE"].dropna().unique()
                                        )
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                    optionHeight=50,
                                    maxHeight=400,
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
                            id="btn_limpar_filtros_pagamentos",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_pagamentos",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_pagamentos"),
                    ],
                ),
            ],
        ),
        html.Div(
            className="charts-row",
            children=[
                dcc.Graph(id="grafico_lista_pagamentos", style={"width": "50%"}),
                dcc.Graph(id="grafico_fonte_pagamentos", style={"width": "50%"}),
            ],
        ),
        html.H4("Detalhamento de Pagamentos"),
        dash_table.DataTable(
            id="tabela_pagamentos",
            row_selectable=False,
            cell_selectable=False,
            active_cell=None,
            selected_cells=[],
            selected_rows=[],
            columns=[
                {"name": "DT ATESTE", "id": "DT ATESTE"},
                {"name": "DT PGTO", "id": "DT PGTO"},
                {"name": "Valor", "id": "Valor"},
                {"name": "FONTE", "id": "FONTE"},
                {"name": "LISTAS", "id": "LISTAS"},
                {"name": "RAZÃO SOCIAL", "id": "RAZÃO SOCIAL"},
            ],
            data=[],
            style_table={"overflowX": "auto"},
            style_cell={
                "textAlign": "center",
                "padding": "8px",
                "fontSize": "13px",
            },
            style_data_conditional=[
                {
                    "if": {"column_id": "Valor"},
                    "color": "#0b2b57",
                    "fontWeight": "bold",
                }
            ],
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "textAlign": "center",
            },
        ),
        dcc.Store(id="store_dados_pagamentos"),
    ],
)


# ----------------------------------------
# 5. CALLBACK — Atualização tabela + gráficos
# ----------------------------------------
@dash.callback(
    Output("tabela_pagamentos", "data"),
    Output("store_dados_pagamentos", "data"),
    Output("grafico_lista_pagamentos", "figure"),
    Output("grafico_fonte_pagamentos", "figure"),
    Input("filtro_ano_pagamentos", "value"),
    Input("filtro_mes_pagamentos", "value"),
    Input("filtro_lista_pagamentos", "value"),
    Input("filtro_fonte_pagamentos", "value"),
    # 🔧 C) Interval como Input extra
    Input("interval-atualizacao", "n_intervals"),
)
def atualizar_tabela(ano, mes, lista, fonte, n_intervals):
    # Atualiza df apenas quando o intervalo dispara
    df = carregar_dados() if n_intervals is not None else df_base
    dff = df.copy()

    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes:
        dff = dff[dff["Mes"] == mes]
    if lista:
        dff = dff[dff["LISTAS"] == lista]
    if fonte:
        dff = dff[dff["FONTE"].astype(str) == str(fonte)]

    dff_display = dff.copy()
    dff_display["DT ATESTE"] = dff_display["DT ATESTE"].dt.strftime("%d/%m/%Y")
    dff_display["DT PGTO"] = dff_display["DT PGTO"].dt.strftime("%d/%m/%Y")

    def formatar_moeda(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    dff_display["Valor"] = dff_display["Valor"].apply(formatar_moeda)

    colunas_exibir = [
        "DT ATESTE",
        "DT PGTO",
        "Valor",
        "FONTE",
        "LISTAS",
        "RAZÃO SOCIAL",
    ]
    dff_display = dff_display[colunas_exibir]

    dados_pdf = {
        "tabela": dff_display.to_dict("records"),
        "filtros": {"ano": ano, "mes": mes, "lista": lista, "fonte": fonte},
        "total_geral": dff["Valor"].sum() if not dff.empty else 0.0,
    }

    if not dff.empty:
        grp_lista = dff.groupby("LISTAS", as_index=False)["Valor"].sum()
    else:
        grp_lista = pd.DataFrame({"LISTAS": [], "Valor": []})

    fig_lista = px.line(
        grp_lista,
        x="LISTAS",
        y="Valor",
        markers=True,
        title="Total Pago por Lista",
    )
    fig_lista.update_traces(
        line_color="#003A70",
        hovertemplate="Lista=%{x}<br>Valor=R$ %{y:,.2f}",
    )
    fig_lista.update_layout(
        xaxis_title="Lista",
        yaxis_title="Valor (R$)",
        yaxis_tickprefix="R$ ",
        yaxis_tickformat=",.2f",
    )

    if not dff.empty:
        grp_fonte = dff.groupby("FONTE", as_index=False)["Valor"].sum()
    else:
        grp_fonte = pd.DataFrame({"FONTE": [], "Valor": []})

    fig_fonte = px.line(
        grp_fonte,
        x="FONTE",
        y="Valor",
        markers=True,
        title="Total por Fonte de Recurso",
    )
    fig_fonte.update_traces(
        line_color="#DA291C",
        hovertemplate="FONTE=%{x}<br>Valor=R$ %{y:,.2f}",
    )
    fig_fonte.update_layout(
        xaxis_title="Fonte",
        yaxis_title="Valor (R$)",
        yaxis_tickprefix="R$ ",
        yaxis_tickformat=",.2f",
    )

    return dff_display.to_dict("records"), dados_pdf, fig_lista, fig_fonte


# ----------------------------------------
# 6. CALLBACK — Limpar filtros
# ----------------------------------------
@dash.callback(
    Output("filtro_ano_pagamentos", "value"),
    Output("filtro_mes_pagamentos", "value"),
    Output("filtro_lista_pagamentos", "value"),
    Output("filtro_fonte_pagamentos", "value"),
    Input("btn_limpar_filtros_pagamentos", "n_clicks"),
    prevent_initial_call=True,
)
def limpar(n):
    ano_padrao = ANO_PADRAO
    return ano_padrao, None, None, None


# ----------------------------------------
# 7. CALLBACK — Geração do PDF
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
    Output("download_relatorio_pagamentos", "data"),
    Input("btn_download_relatorio_pagamentos", "n_clicks"),
    State("tabela_pagamentos", "data"),
    State("store_dados_pagamentos", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, tabela, dados_pdf):
    if not n or not dados_pdf:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Relatório de Pagamentos Efetivados",
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
            f"Lista: {filtros['lista'] if filtros['lista'] else 'Todas'} — "
            f"Fonte: {filtros['fonte'] if filtros['fonte'] else 'Todas'}",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 0.3 * inch))

    def fmt(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    story.append(
        Paragraph(
            f"Total Geral: {fmt(dados_pdf['total_geral'])}",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 0.2 * inch))

    table_data = [
        ["DT ATESTE", "DT PGTO", "Valor", "FONTE", "LISTAS", "RAZÃO SOCIAL"]
    ]
    for r in dados_pdf["tabela"]:
        table_data.append(
            [
                wrap(r["DT ATESTE"]),
                wrap(r["DT PGTO"]),
                wrap(r["Valor"]),
                wrap(r["FONTE"]),
                wrap(r["LISTAS"]),
                wrap(r["RAZÃO SOCIAL"][:30]),
            ]
        )

    col_widths = [
        1.0 * inch,
        1.0 * inch,
        1.0 * inch,
        1.0 * inch,
        1.0 * inch,
        1.3 * inch,
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
    return dcc.send_bytes(buffer.getvalue(), "pagamentos_efetivados.pdf")
