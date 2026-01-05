import dash
from dash import html, dcc, Input, Output, State, dash_table

import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
import plotly.express as px
from datetime import datetime


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/processos-de-compras",
    name="Processos de Compras",
    title="Processos de Compras",
)


URL_PROCESSOS = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=BI%20-%20Itajub%C3%A1"
)


# ----------------------------------------
# Carga e tratamento dos dados
# ----------------------------------------
def carregar_dados_processos():
    df = pd.read_csv(URL_PROCESSOS)
    df.columns = [c.strip() for c in df.columns]

    col_solicitante = "Solicitante"
    col_num_proc = "Numero do Processo"
    col_preco_estimado = "PREÇO ESTIMADO"
    col_valor_contratado = "Valor Contratado"
    col_objeto = "Objeto"
    col_modalidade = "Modalidade"
    col_ano = "Ano"
    col_status = "Status"
    col_classif_nc = "Classificação dos processos não concluídos"
    col_numero = "Número"
    col_data_entrada = "Data de Entrada"
    col_data_finalizacao = "Data finalização"
    col_contr_reinstr_com = (
        "CONTRATAÇÃO REINSTRUÍDA PELO PROCESSO Nº (com pontos e traços)"
    )

    for c in [
        col_solicitante,
        col_num_proc,
        col_preco_estimado,
        col_valor_contratado,
        col_objeto,
        col_modalidade,
        col_ano,
        col_status,
        col_classif_nc,
        col_numero,
        col_data_entrada,
        col_data_finalizacao,
        col_contr_reinstr_com,
    ]:
        if c not in df.columns:
            df[c] = ""

    def conv_moeda(v):
        if isinstance(v, str):
            v = (
                v.replace("R$", "")
                .replace(".", "")
                .replace(",", ".")
                .strip()
            )
            return float(v) if v not in ["", "-"] else 0.0
        return float(v) if pd.notna(v) else 0.0

    df[col_preco_estimado] = df[col_preco_estimado].apply(conv_moeda)
    df[col_valor_contratado] = df[col_valor_contratado].apply(conv_moeda)

    df["Data finalização"] = pd.to_datetime(
        df["Data finalização"], format="%d/%m/%Y", errors="coerce"
    )

    meses_map = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro",
    }

    df["Mes_finalizacao"] = df["Data finalização"].dt.month.map(meses_map)

    return df


df_proc_base = carregar_dados_processos()
ANO_ATUAL = datetime.now().year

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}


def formatar_moeda(v):
    return (
        f"R$ {v:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# ----------------------------------------
# Layout
# ----------------------------------------
layout = html.Div(
    children=[
        # Barra de filtros (sticky dentro da main-content)
        html.Div(
            id="barra_filtros_proc",
            className="filtros-sticky",
            children=[
                # Linha 1
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "200px", "flex": "1 1 240px"},
                            children=[
                                html.Label("Número do Processo"),
                                dcc.Input(
                                    id="filtro_num_proc",
                                    type="text",
                                    placeholder="Digite o número completo ou parte",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "120px", "flex": "0 0 140px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_proc",
                                    options=[
                                        {"label": str(a), "value": a}
                                        for a in sorted(
                                            df_proc_base["Ano"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a) != ""
                                    ],
                                    value=ANO_ATUAL,
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "150px", "flex": "0 0 170px"},
                            children=[
                                html.Label("Mês de Finalização"),
                                dcc.Dropdown(
                                    id="filtro_mes_finalizacao",
                                    options=[
                                        {"label": m.capitalize(), "value": m}
                                        for m in sorted(
                                            df_proc_base["Mes_finalizacao"]
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
                            style={"minWidth": "200px", "flex": "1 1 240px"},
                            children=[
                                html.Label("Solicitante"),
                                dcc.Dropdown(
                                    id="filtro_solicitante_proc",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in sorted(
                                            df_proc_base["Solicitante"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(s) != ""
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "260px", "flex": "2 1 320px"},
                            children=[
                                html.Label("Objeto"),
                                dcc.Dropdown(
                                    id="filtro_objeto_proc",
                                    options=[
                                        {"label": o, "value": o}
                                        for o in sorted(
                                            df_proc_base["Objeto"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(o) != ""
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
                # Linha 2
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                        "marginTop": "4px",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Modalidade"),
                                dcc.Dropdown(
                                    id="filtro_modalidade_proc",
                                    options=[
                                        {"label": m, "value": m}
                                        for m in sorted(
                                            df_proc_base["Modalidade"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(m) != ""
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
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="filtro_status_proc",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in sorted(
                                            df_proc_base["Status"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(s) != ""
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "260px", "flex": "2 1 320px"},
                            children=[
                                html.Label("Classificação (Não Concluídos)"),
                                dcc.Dropdown(
                                    id="filtro_classif_nc_proc",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in sorted(
                                            df_proc_base[
                                                "Classificação dos processos não concluídos"
                                            ]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(c) != ""
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
                # Botões
                html.Div(
                    style={"marginTop": "4px"},
                    children=[
                        html.Button(
                            "Limpar Filtros",
                            id="btn_limpar_filtros_proc",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_proc",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_proc"),
                    ],
                ),
            ],
        ),
        # Conteúdo
        html.Div(
            children=[
                html.Div(
                    id="cards_resumo_proc",
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "marginBottom": "15px",
                        "marginTop": "10px",
                    },
                ),
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "marginBottom": "15px",
                    },
                    children=[
                        dcc.Graph(
                            id="grafico_status_proc",
                            style={"flex": "1 1 320px", "minWidth": "300px"},
                        ),
                        dcc.Graph(
                            id="grafico_valor_mes_proc",
                            style={"flex": "2 1 420px", "minWidth": "340px"},
                        ),
                    ],
                ),
                html.H4("Tabela de Processos de Compras"),
                dash_table.DataTable(
                    id="tabela_proc",
                    columns=[
                        {"name": "Solicitante", "id": "Solicitante"},
                        {"name": "Número Do Processo", "id": "Numero do Processo"},
                        {"name": "Objeto", "id": "Objeto"},
                        {"name": "Modalidade", "id": "Modalidade"},
                        {"name": "Preço Estimado", "id": "PREÇO ESTIMADO_FMT"},
                        {"name": "Valor Contratado", "id": "Valor Contratado_FMT"},
                        {"name": "Status", "id": "Status"},
                        {"name": "Data De Entrada", "id": "Data de Entrada"},
                        {"name": "Data Finalização", "id": "Data finalização_FMT"},
                        {
                            "name": "Classificação (Não Concluídos)",
                            "id": "Classificação dos processos não concluídos",
                        },
                        {
                            "name": "Contratação Reinstruída Pelo Processo Nº",
                            "id": "CONTRATAÇÃO REINSTRUÍDA PELO PROCESSO Nº (com pontos e traços)",
                        },
                    ],
                    data=[],
                    row_selectable=False,
                    cell_selectable=False,
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "500px",
                        "position": "relative",  # para sticky header
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "minWidth": "80px",
                        "maxWidth": "220px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#0b2b57",
                        "color": "white",
                        "textAlign": "center",
                        "position": "sticky",
                        "top": 0,
                        "zIndex": 10,
                    },
                ),
                dcc.Store(id="store_dados_proc"),
            ],
        ),
    ],
)


# ----------------------------------------
# Callback: atualizar tabela + cards + gráficos
# ----------------------------------------
@dash.callback(
    Output("tabela_proc", "data"),
    Output("store_dados_proc", "data"),
    Output("cards_resumo_proc", "children"),
    Output("grafico_status_proc", "figure"),
    Output("grafico_valor_mes_proc", "figure"),
    Input("filtro_num_proc", "value"),
    Input("filtro_ano_proc", "value"),
    Input("filtro_mes_finalizacao", "value"),
    Input("filtro_solicitante_proc", "value"),
    Input("filtro_objeto_proc", "value"),
    Input("filtro_modalidade_proc", "value"),
    Input("filtro_status_proc", "value"),
    Input("filtro_classif_nc_proc", "value"),
)
def atualizar_tabela_proc(
    num_proc, ano, mes_finalizacao, solicitante, objeto, modalidade, status, classif_nc
):
    dff = df_proc_base.copy()

    if num_proc and str(num_proc).strip():
        termo = str(num_proc).strip()
        dff = dff[
            dff["Numero do Processo"]
            .astype(str)
            .str.contains(termo, case=False, na=False)
        ]

    if ano:
        dff = dff[dff["Ano"] == ano]
    if mes_finalizacao:
        dff = dff[dff["Mes_finalizacao"] == mes_finalizacao]
    if solicitante:
        dff = dff[dff["Solicitante"] == solicitante]
    if objeto:
        dff = dff[dff["Objeto"] == objeto]
    if modalidade:
        dff = dff[dff["Modalidade"] == modalidade]
    if status:
        dff = dff[dff["Status"] == status]
    if classif_nc:
        dff = dff[dff["Classificação dos processos não concluídos"] == classif_nc]

    dff_display = dff.copy()
    dff_display["PREÇO ESTIMADO_FMT"] = dff_display["PREÇO ESTIMADO"].apply(
        formatar_moeda
    )
    dff_display["Valor Contratado_FMT"] = dff_display["Valor Contratado"].apply(
        formatar_moeda
    )
    dff_display["Data de Entrada"] = pd.to_datetime(
        dff_display["Data de Entrada"], format="%d/%m/%Y", errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    dff_display["Data finalização_FMT"] = dff_display["Data finalização"].dt.strftime(
        "%d/%m/%Y"
    )

    # Ordenar pela Data de Entrada (mais recente para mais antiga)
    dff_display["Data_Entrada_dt"] = pd.to_datetime(
        dff_display["Data de Entrada"], format="%d/%m/%Y", errors="coerce"
    )
    dff_display = dff_display.sort_values(
        "Data_Entrada_dt", ascending=False
    ).reset_index(drop=True)

    total_valor_contratado = dff["Valor Contratado"].sum()
    qtd_processos = len(dff)
    media_por_processo = (
        total_valor_contratado / qtd_processos if qtd_processos > 0 else 0.0
    )

    concluidos = (dff["Status"] == "Concluído").sum()
    em_andamento = (dff["Status"] == "Em Andamento").sum()
    nao_concluidos = (dff["Status"] == "Não Concluído").sum()

    card_style = {
        "flex": "1 1 220px",
        "backgroundColor": "#ffffff",
        "padding": "14px",
        "textAlign": "center",
        "minHeight": "20px",
    }

    cards = [
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    formatar_moeda(total_valor_contratado),
                    style={"color": "#c00000", "margin": "0", "fontSize": "20px"},
                ),
                html.Div("Valor Contratado", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    formatar_moeda(media_por_processo),
                    style={"color": "#003A70", "margin": "0", "fontSize": "20px"},
                ),
                html.Div("Média por Processo Concluído", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(qtd_processos, style={"margin": "0", "fontSize": "20px"}),
                html.Div("Número de Processos", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(concluidos, style={"margin": "0", "fontSize": "20px"}),
                html.Div("Processos Concluídos", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(em_andamento, style={"margin": "0", "fontSize": "20px"}),
                html.Div("Processos Em Andamento", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(nao_concluidos, style={"margin": "0", "fontSize": "20px"}),
                html.Div("Processos Não Concluídos", style={"fontSize": "15px"}),
            ],
        ),
    ]

    if dff.empty:
        fig_status = px.pie(title="Porcentagem de Status")
        fig_valor_mes = px.bar(title="Valor dos Processos Concluídos por Mês")
    else:
        grp_status = (
            dff.groupby("Status", as_index=False)["Numero do Processo"]
            .count()
            .rename(columns={"Numero do Processo": "Qtd"})
        )

        fig_status = px.pie(
            grp_status,
            names="Status",
            values="Qtd",
            hole=0.6,
            title="Porcentagem de Status",
        )

        fig_status.update_traces(
            marker=dict(
                colors=["#003A70", "#DA291C", "#A2AAAD"],
                line=dict(color="#ECEDEF", width=2),
            ),
            textposition="outside",
            texttemplate="%{label} %{value} (%{percent:.2%})",
        )

        fig_status.update_layout(
            title_x=0.5,
            plot_bgcolor="#FFFFFF",
            paper_bgcolor="#FFFFFF",
            showlegend=True,
        )

        dff_conc = dff[dff["Status"] == "Concluído"].copy()

        if dff_conc.empty:
            fig_valor_mes = px.bar(title="Valor dos Processos Concluídos por Mês")
        else:
            grp_mes = (
                dff_conc.groupby("Mes_finalizacao", as_index=False)["Valor Contratado"]
                .sum()
            )

            meses_ordem = [
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

            grp_mes["Mes_finalizacao"] = pd.Categorical(
                grp_mes["Mes_finalizacao"],
                categories=meses_ordem,
                ordered=True,
            )

            grp_mes = grp_mes.sort_values("Mes_finalizacao")
            grp_mes["Valor_fmt"] = grp_mes["Valor Contratado"].apply(formatar_moeda)

            fig_valor_mes = px.bar(
                grp_mes,
                x="Valor Contratado",
                y="Mes_finalizacao",
                orientation="h",
                text="Valor_fmt",
                title="Valor dos Processos Concluídos por Mês",
            )

            fig_valor_mes.update_traces(
                marker_color="#003A70",
                textposition="outside",
            )

            fig_valor_mes.update_layout(
                title_x=0.5,
                xaxis_title="Valor Contratado (R$)",
                yaxis_title="Mês de Finalização",
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#FFFFFF",
            )

    cols_tabela = [
        "Solicitante",
        "Numero do Processo",
        "Objeto",
        "Modalidade",
        "PREÇO ESTIMADO_FMT",
        "Valor Contratado_FMT",
        "Status",
        "Data de Entrada",
        "Data finalização_FMT",
        "Classificação dos processos não concluídos",
        "CONTRATAÇÃO REINSTRUÍDA PELO PROCESSO Nº (com pontos e traços)",
    ]

    return (
        dff_display[cols_tabela].to_dict("records"),
        dff.to_dict("records"),
        cards,
        fig_status,
        fig_valor_mes,
    )


# ----------------------------------------
# Callback: limpar filtros
# ----------------------------------------
@dash.callback(
    Output("filtro_num_proc", "value"),
    Output("filtro_ano_proc", "value"),
    Output("filtro_mes_finalizacao", "value"),
    Output("filtro_solicitante_proc", "value"),
    Output("filtro_objeto_proc", "value"),
    Output("filtro_modalidade_proc", "value"),
    Output("filtro_status_proc", "value"),
    Output("filtro_classif_nc_proc", "value"),
    Input("btn_limpar_filtros_proc", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_proc(n):
    return None, ANO_ATUAL, None, None, None, None, None, None


# ----------------------------------------
# Callback: gerar PDF (paisagem, header repetido)
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
    Output("download_relatorio_proc", "data"),
    Input("btn_download_relatorio_proc", "n_clicks"),
    State("store_dados_proc", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_proc(n, dados_proc):
    if not n or not dados_proc:
        return None
    df = pd.DataFrame(dados_proc)

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
        "Relatório de Processos de Compras",
        ParagraphStyle(
            "titulo", fontSize=16, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )

    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    story.append(
        Paragraph(
            f"Total de registros: {len(df)}",
            styles["Normal"],
        )
    )

    story.append(Spacer(1, 0.15 * inch))

    cols = [
        "Solicitante",
        "Numero do Processo",
        "Objeto",
        "Modalidade",
        "PREÇO ESTIMADO",
        "Valor Contratado",
        "Status",
        "Data de Entrada",
        "Data finalização",
        "Classificação dos processos não concluídos",
        "CONTRATAÇÃO REINSTRUÍDA PELO PROCESSO Nº (com pontos e traços)",
    ]

    df_pdf = df.copy()
    df_pdf["PREÇO ESTIMADO"] = df_pdf["PREÇO ESTIMADO"].apply(formatar_moeda)
    df_pdf["Valor Contratado"] = df_pdf["Valor Contratado"].apply(formatar_moeda)
    df_pdf["Data de Entrada"] = pd.to_datetime(
        df_pdf["Data de Entrada"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    df_pdf["Data finalização"] = pd.to_datetime(
        df_pdf["Data finalização"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    header = cols
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap(row[c]) for c in cols])

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
    return dcc.send_bytes(buffer.getvalue(), "processos_de_compras_paisagem.pdf")
