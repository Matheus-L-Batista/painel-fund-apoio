import dash
from dash import html, dcc, Input, Output, State, dash_table
from dash.exceptions import PreventUpdate
import pandas as pd
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
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.lib import colors
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
from pytz import timezone
import os

# --------------------------------------------------
# Função para verificar se estamos na página de processos de compras
# --------------------------------------------------
def verificar_pagina_processos_compras():
    """Verifica se o callback está sendo executado na página de processos de compras"""
    try:
        if not dash.ctx.triggered:
            # Permite execução inicial
            return True
        
        # Componentes específicos da página de processos de compras
        componentes_processos = {
            'filtro_num_proc', 'filtro_ano_proc', 'filtro_mes_finalizacao',
            'filtro_solicitante_proc', 'filtro_objeto_proc', 'filtro_modalidade_proc',
            'filtro_status_proc', 'filtro_classif_nc_proc',
            'btn_limpar_filtros_proc', 'btn_download_relatorio_proc'
        }
        
        # Obtém o ID do componente que disparou o callback
        triggered = dash.ctx.triggered[0]
        triggered_id = triggered['prop_id'].split('.')[0]
        
        # Verifica se é um componente da página de processos de compras
        return triggered_id in componentes_processos
    except Exception:
        # Em caso de erro, permite a execução (segurança para inicialização)
        return True

# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/processos-de-compras",
    name="Processos de Compras",
    title="Processos de Compras",
)

# URL da planilha de Processos de Compras (BI Itajubá)
URL_PROCESSOS = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=BI%20-%20Itajub%C3%A1"
)

# --------------------------------------------------
# Carga de dados e utilitários
# --------------------------------------------------
def carregar_dados_processos():
    """
    Lê a planilha de processos de compras e faz:
    - garantia de existência de colunas esperadas
    - conversão de campos monetários para float
    - conversão de datas e criação da coluna de mês de finalização (texto)
    """
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

    # Garante todas as colunas presentes, mesmo se a planilha mudar
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
        """
        Converte string no formato brasileiro de moeda
        (R$, pontos de milhar, vírgula decimal) em float.
        """
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

    # Converte a data de finalização para datetime
    df["Data finalização"] = pd.to_datetime(
        df["Data finalização"], format="%d/%m/%Y", errors="coerce"
    )

    # Mapeamento numérico -> mês por extenso (minúsculo)
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

# Força ano padrão 2026
ANO_PADRAO = 2026

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

# --------------------------------------------------
# Estilo unificado dos botões
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


def formatar_moeda(v):
    """
    Formata float em moeda brasileira com prefixo R$.
    """
    return (
        f"R$ {v:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# Lista fixa de meses em ordem cronológica
MESES_ORDENADOS = [
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

# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        # Barra de filtros
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
                        # Filtro: Número do Processo (dropdown)
                        html.Div(
                            style={
                                "minWidth": "200px",
                                "flex": "1 1 240px",
                            },
                            children=[
                                html.Label("Número do Processo"),
                                dcc.Dropdown(
                                    id="filtro_num_proc",
                                    options=[],
                                    value=None,
                                    placeholder="Selecione um número de processo...",
                                    clearable=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Filtro: Ano (sempre obrigatório, default = 2026)
                        html.Div(
                            style={
                                "minWidth": "120px",
                                "flex": "0 0 140px",
                            },
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_proc",
                                    options=[
                                        {
                                            "label": str(a),
                                            "value": a,
                                        }
                                        for a in sorted(
                                            df_proc_base["Ano"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a) != ""
                                    ],
                                    value=ANO_PADRAO,
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Filtro: Mês de Finalização
                        html.Div(
                            style={
                                "minWidth": "150px",
                                "flex": "0 0 170px",
                            },
                            children=[
                                html.Label("Mês de Finalização"),
                                dcc.Dropdown(
                                    id="filtro_mes_finalizacao",
                                    options=[],
                                    value=None,
                                    placeholder="Selecione um mês...",
                                    clearable=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Filtro: Solicitante
                        html.Div(
                            style={
                                "minWidth": "200px",
                                "flex": "1 1 240px",
                            },
                            children=[
                                html.Label("Solicitante"),
                                dcc.Dropdown(
                                    id="filtro_solicitante_proc",
                                    options=[],
                                    value=None,
                                    placeholder="Selecione um solicitante...",
                                    clearable=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Filtro: Objeto
                        html.Div(
                            style={
                                "minWidth": "260px",
                                "flex": "2 1 320px",
                            },
                            children=[
                                html.Label("Objeto"),
                                dcc.Dropdown(
                                    id="filtro_objeto_proc",
                                    options=[],
                                    value=None,
                                    placeholder="Selecione um objeto...",
                                    clearable=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                    ],
                ),
                # Linha 2 + botões à direita
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                        "marginTop": "4px",
                        "justifyContent": "space-between",
                    },
                    children=[
                        # Coluna esquerda: filtros
                        html.Div(
                            style={
                                "display": "flex",
                                "flexWrap": "wrap",
                                "gap": "10px",
                                "alignItems": "flex-start",
                                "flex": "1 1 auto",
                            },
                            children=[
                                # Filtro: Modalidade
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 260px",
                                    },
                                    children=[
                                        html.Label("Modalidade"),
                                        dcc.Dropdown(
                                            id="filtro_modalidade_proc",
                                            options=[],
                                            value=None,
                                            placeholder="Selecione uma modalidade...",
                                            clearable=True,
                                            searchable=True,
                                            style=dropdown_style,
                                        ),
                                    ],
                                ),
                                # Filtro: Status
                                html.Div(
                                    style={
                                        "minWidth": "220px",
                                        "flex": "1 1 260px",
                                    },
                                    children=[
                                        html.Label("Status"),
                                        dcc.Dropdown(
                                            id="filtro_status_proc",
                                            options=[],
                                            value=None,
                                            placeholder="Selecione um status...",
                                            clearable=True,
                                            searchable=True,
                                            style=dropdown_style,
                                        ),
                                    ],
                                ),
                                # Filtro: Classificação (Não Concluídos)
                                html.Div(
                                    style={
                                        "minWidth": "260px",
                                        "flex": "2 1 320px",
                                    },
                                    children=[
                                        html.Label(
                                            "Classificação (Não Concluídos)"
                                        ),
                                        dcc.Dropdown(
                                            id="filtro_classif_nc_proc",
                                            options=[],
                                            value=None,
                                            placeholder="Selecione uma classificação...",
                                            clearable=True,
                                            searchable=True,
                                            style=dropdown_style,
                                        ),
                                    ],
                                ),
                            ],
                        ),
                        # Coluna direita: botões
                        html.Div(
                            style={
                                "display": "flex",
                                "flexWrap": "wrap",
                                "gap": "6px",
                                "justifyContent": "flex-end",
                                "alignItems": "center",
                            },
                            children=[
                                html.Button(
                                    "Limpar Filtros",
                                    id="btn_limpar_filtros_proc",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_proc",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                dcc.Download(id="download_relatorio_proc"),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        # Conteúdo principal
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
                            style={
                                "flex": "1 1 320px",
                                "minWidth": "300px",
                                "height": "320px",
                            },
                        ),
                        dcc.Graph(
                            id="grafico_valor_mes_proc",
                            style={
                                "flex": "2 1 420px",
                                "minWidth": "340px",
                                "height": "320px",
                            },
                        ),
                    ],
                ),
                html.H4("Tabela de Processos de Compras"),
                dash_table.DataTable(
                    id="tabela_proc",
                    columns=[
                        {"name": "Solicitante", "id": "Solicitante"},
                        {
                            "name": "Número Do Processo",
                            "id": "Numero do Processo",
                        },
                        {"name": "Objeto", "id": "Objeto"},
                        {"name": "Modalidade", "id": "Modalidade"},
                        {
                            "name": "Preço Estimado",
                            "id": "PREÇO ESTIMADO_FMT",
                        },
                        {
                            "name": "Valor Contratado",
                            "id": "Valor Contratado_FMT",
                        },
                        {"name": "Status", "id": "Status"},
                        {
                            "name": "Data De Entrada",
                            "id": "Data de Entrada",
                        },
                        {
                            "name": "Data Finalização",
                            "id": "Data finalização_FMT",
                        },
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
                        "position": "relative",
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
                    # Linhas alternando branco e cinza
                    style_data_conditional=[
                        {
                            "if": {"row_index": "odd"},
                            "backgroundColor": "#ffffff",
                        },
                        {
                            "if": {"row_index": "even"},
                            "backgroundColor": "#f0f0f0",
                        },
                    ],
                ),
                # Store com os dados filtrados (base para PDF)
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
    num_proc,
    ano,
    mes_finalizacao,
    solicitante,
    objeto,
    modalidade,
    status,
    classif_nc,
):
    # VERIFICAÇÃO: Só executa se estiver na página de processos de compras
    if not verificar_pagina_processos_compras():
        raise PreventUpdate
    
    # -------------------------
    # Filtro principal
    # -------------------------
    dff = df_proc_base.copy()
    mask = pd.Series(True, index=dff.index)

    # Filtro por número de processo (texto parcial)
    if num_proc and str(num_proc).strip():
        termo = str(num_proc).strip()
        mask &= (
            dff["Numero do Processo"]
            .astype(str)
            .str.contains(termo, case=False, na=False)
        )

    # Ano (sempre aplicado se não nulo)
    if ano:
        mask &= dff["Ano"] == ano
    if mes_finalizacao:
        mask &= dff["Mes_finalizacao"] == mes_finalizacao
    if solicitante:
        mask &= dff["Solicitante"] == solicitante
    if objeto:
        mask &= dff["Objeto"] == objeto
    if modalidade:
        mask &= dff["Modalidade"] == modalidade
    if status:
        mask &= dff["Status"] == status
    if classif_nc:
        mask &= (
            dff["Classificação dos processos não concluídos"] == classif_nc
        )

    dff = dff[mask]

    # -------------------------
    # Formatação da tabela
    # -------------------------
    dff_display = dff.copy()
    dff_display["PREÇO ESTIMADO_FMT"] = dff_display["PREÇO ESTIMADO"].apply(
        formatar_moeda
    )
    dff_display["Valor Contratado_FMT"] = dff_display[
        "Valor Contratado"
    ].apply(formatar_moeda)

    # Data de Entrada
    dff_display["Data de Entrada"] = pd.to_datetime(
        dff_display["Data de Entrada"],
        format="%d/%m/%Y",
        errors="coerce",
    ).dt.strftime("%d/%m/%Y")

    # Data finalização
    dff_display["Data finalização_FMT"] = dff_display["Data finalização"].dt.strftime(
        "%d/%m/%Y"
    )

    # Campo auxiliar para ordenação
    dff_display["Data_Entrada_dt"] = pd.to_datetime(
        dff_display["Data de Entrada"], format="%d/%m/%Y", errors="coerce"
    )

    dff_display = (
        dff_display.sort_values("Data_Entrada_dt", ascending=False)
        .reset_index(drop=True)
    )

    # -------------------------
    # Cards resumo
    # -------------------------
    total_valor_contratado = dff["Valor Contratado"].sum()
    qtd_processos = len(dff)
    concluidos = (dff["Status"] == "Concluído").sum()
    media_por_processo = (
        total_valor_contratado / concluidos if concluidos > 0 else 0.0
    )

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
                    style={
                        "color": "#c00000",
                        "margin": "0",
                        "fontSize": "20px",
                    },
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
                    style={
                        "color": "#003A70",
                        "margin": "0",
                        "fontSize": "20px",
                    },
                ),
                html.Div(
                    "Média por Processo Concluído",
                    style={"fontSize": "15px"},
                ),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    qtd_processos,
                    style={"margin": "0", "fontSize": "20px"},
                ),
                html.Div("Número de Processos", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    concluidos,
                    style={"margin": "0", "fontSize": "20px"},
                ),
                html.Div("Processos Concluídos", style={"fontSize": "15px"}),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    em_andamento,
                    style={"margin": "0", "fontSize": "20px"},
                ),
                html.Div(
                    "Processos Em Andamento", style={"fontSize": "15px"}
                ),
            ],
        ),
        html.Div(
            className="card-resumo",
            style=card_style,
            children=[
                html.H4(
                    nao_concluidos,
                    style={"margin": "0", "fontSize": "20px"},
                ),
                html.Div(
                    "Processos Não Concluídos", style={"fontSize": "15px"}
                ),
            ],
        ),
    ]

    # -------------------------
    # Gráficos
    # -------------------------
    if dff.empty:
        fig_status = px.pie(title="Porcentagem de Status")
        fig_valor_mes = px.bar(title="Processos Concluídos por Ano")
    else:
        # Gráfico de status (pizza)
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

        # --- gráfico anual usa filtros exceto ano ---
        dff_global = df_proc_base.copy()
        mask_global = pd.Series(True, index=dff_global.index)

        if num_proc and str(num_proc).strip():
            termo = str(num_proc).strip()
            mask_global &= (
                dff_global["Numero do Processo"]
                .astype(str)
                .str.contains(termo, case=False, na=False)
            )
        if mes_finalizacao:
            mask_global &= (
                dff_global["Mes_finalizacao"] == mes_finalizacao
            )
        if solicitante:
            mask_global &= dff_global["Solicitante"] == solicitante
        if objeto:
            mask_global &= dff_global["Objeto"] == objeto
        if modalidade:
            mask_global &= dff_global["Modalidade"] == modalidade
        if status:
            mask_global &= dff_global["Status"] == status
        if classif_nc:
            mask_global &= (
                dff_global["Classificação dos processos não concluídos"]
                == classif_nc
            )

        dff_global = dff_global[mask_global]
        dff_conc_global = dff_global[
            dff_global["Status"] == "Concluído"
        ].copy()

        if dff_conc_global.empty:
            fig_valor_mes = px.bar(
                title="Processos Concluídos por Ano"
            )
        else:
            grp_ano = (
                dff_conc_global.groupby("Ano", as_index=False)
                .agg(
                    Valor_Contratado_Total=("Valor Contratado", "sum"),
                    Qtd_Processos=("Numero do Processo", "count"),
                )
            )

            grp_ano["Media_Por_Processo"] = (
                grp_ano["Valor_Contratado_Total"]
                / grp_ano["Qtd_Processos"]
            )

            grp_ano["Valor_Contratado_Total_FMT"] = grp_ano[
                "Valor_Contratado_Total"
            ].apply(formatar_moeda)
            grp_ano["Media_Por_Processo_FMT"] = grp_ano[
                "Media_Por_Processo"
            ].apply(formatar_moeda)

            hovertemplate_ano = (
                "Ano: %{x}<br>"
                + "Valor Contratado Total: %{customdata[0]}<br>"
                + "Média por Processo: %{customdata[1]}<br>"
                + "Número de Processos Concluídos: %{customdata[2]}<extra></extra>"
            )

            customdata = grp_ano[
                [
                    "Valor_Contratado_Total_FMT",
                    "Media_Por_Processo_FMT",
                    "Qtd_Processos",
                ]
            ].values

            fig_valor_mes = make_subplots(
                specs=[[{"secondary_y": True}]]
            )

            fig_valor_mes.add_trace(
                go.Bar(
                    x=grp_ano["Ano"],
                    y=grp_ano["Valor_Contratado_Total"],
                    name="Valor Contratado Total",
                    marker_color="red",
                    width=0.4,
                    customdata=customdata,
                    hovertemplate=hovertemplate_ano,
                ),
                secondary_y=False,
            )

            fig_valor_mes.add_trace(
                go.Bar(
                    x=grp_ano["Ano"],
                    y=grp_ano["Media_Por_Processo"],
                    name="Média por Processo",
                    marker_color="blue",
                    width=0.4,
                    offset=-0.2,
                    customdata=customdata,
                    hovertemplate=hovertemplate_ano,
                ),
                secondary_y=False,
            )

            fig_valor_mes.add_trace(
                go.Scatter(
                    x=grp_ano["Ano"],
                    y=grp_ano["Qtd_Processos"],
                    name="Número de Processos Concluídos",
                    mode="lines+markers",
                    line=dict(color="green", width=3),
                    customdata=customdata,
                    hovertemplate=hovertemplate_ano,
                ),
                secondary_y=True,
            )

            fig_valor_mes.update_layout(
                barmode="overlay",
                title="Processos Concluídos por Ano",
                title_x=0.5,
                xaxis_title="Ano",
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#FFFFFF",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.2,
                    x=0.5,
                    xanchor="center",
                    title_text="",  # remove texto da legenda
                ),
            )

            fig_valor_mes.update_yaxes(
                title_text="Valores (R$)", secondary_y=False
            )
            fig_valor_mes.update_yaxes(
                title_text="Número de Processos Concluídos",
                secondary_y=True,
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
# Callback: filtros em cascata
# ----------------------------------------
@dash.callback(
    Output("filtro_num_proc", "options"),
    Output("filtro_mes_finalizacao", "options"),
    Output("filtro_solicitante_proc", "options"),
    Output("filtro_objeto_proc", "options"),
    Output("filtro_modalidade_proc", "options"),
    Output("filtro_status_proc", "options"),
    Output("filtro_classif_nc_proc", "options"),
    Input("filtro_ano_proc", "value"),
    Input("filtro_mes_finalizacao", "value"),
    Input("filtro_solicitante_proc", "value"),
    Input("filtro_objeto_proc", "value"),
    Input("filtro_modalidade_proc", "value"),
    Input("filtro_status_proc", "value"),
    Input("filtro_classif_nc_proc", "value"),
    Input("filtro_num_proc", "value"),
)
def atualizar_opcoes_filtros(
    ano,
    mes_finalizacao,
    solicitante,
    objeto,
    modalidade,
    status,
    classif_nc,
    num_proc,
):
    """
    Gera opções de dropdown em cascata a partir de um único filtro global.
    A ordem de seleção dos filtros não importa.
    """
    # VERIFICAÇÃO: Só executa se estiver na página de processos de compras
    if not verificar_pagina_processos_compras():
        raise PreventUpdate
    
    dff = df_proc_base.copy()
    mask = pd.Series(True, index=dff.index)

    # Aplica todos os filtros
    if ano:
        mask &= dff["Ano"] == ano
    if mes_finalizacao:
        mask &= dff["Mes_finalizacao"] == mes_finalizacao
    if solicitante:
        mask &= dff["Solicitante"] == solicitante
    if objeto:
        mask &= dff["Objeto"] == objeto
    if modalidade:
        mask &= dff["Modalidade"] == modalidade
    if status:
        mask &= dff["Status"] == status
    if classif_nc:
        mask &= (
            dff["Classificação dos processos não concluídos"] == classif_nc
        )
    if num_proc:
        mask &= dff["Numero do Processo"] == num_proc

    dff = dff[mask]

    # Opções para Número do Processo
    op_num_proc = [
        {"label": str(p), "value": str(p)}
        for p in sorted(dff["Numero do Processo"].dropna().unique())
        if str(p) != ""
    ]

    # Opções para Mês de Finalização (respeitando a ordem cronológica)
    meses_disponiveis = dff["Mes_finalizacao"].dropna().unique().tolist()
    op_mes_finalizacao = [
        {"label": m.capitalize(), "value": m}
        for m in MESES_ORDENADOS
        if m in meses_disponiveis
    ]

    # Opções para Solicitante
    op_solicitante = [
        {"label": str(s), "value": str(s)}
        for s in sorted(dff["Solicitante"].dropna().unique())
        if str(s) != ""
    ]

    # Opções para Objeto
    op_objeto = [
        {"label": str(o), "value": str(o)}
        for o in sorted(dff["Objeto"].dropna().unique())
        if str(o) != ""
    ]

    # Opções para Modalidade
    op_modalidade = [
        {"label": str(m), "value": str(m)}
        for m in sorted(dff["Modalidade"].dropna().unique())
        if str(m) != ""
    ]

    # Opções para Status
    op_status = [
        {"label": str(s), "value": str(s)}
        for s in sorted(dff["Status"].dropna().unique())
        if str(s) != ""
    ]

    # Opções para Classificação
    op_classif = [
        {"label": str(c), "value": str(c)}
        for c in sorted(
            dff["Classificação dos processos não concluídos"]
            .dropna()
            .unique()
        )
        if str(c) != ""
    ]

    return (
        op_num_proc,
        op_mes_finalizacao,
        op_solicitante,
        op_objeto,
        op_modalidade,
        op_status,
        op_classif,
    )

# ----------------------------------------
# Callback: limpar filtros (volta sempre para ano 2026)
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
    """
    Limpa todos os filtros e retorna o ano para ANO_PADRAO (2026).
    """
    # VERIFICAÇÃO: Só executa se estiver na página de processos de compras
    if not verificar_pagina_processos_compras():
        raise PreventUpdate
    
    return None, ANO_PADRAO, None, None, None, None, None, None

# ====================================================
# FUNÇÕES AUXILIARES PARA PDF – COMPRAS
# ====================================================
def formatar_moeda(valor):
    """
    Formata um valor numérico como moeda brasileira (R$ X.XXX,XX).
    """
    try:
        valor_float = float(valor) if isinstance(valor, str) else valor
        return (
            f"R$ {valor_float:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except (ValueError, TypeError):
        return str(valor)


def criar_card_elemento(titulo, valor, cor):
    """
    Cria um elemento de card para PDF.
    """
    card_content = [
        [
            Paragraph(
                f"{valor}",
                ParagraphStyle(
                    "card_valor",
                    alignment=TA_CENTER,
                    spaceAfter=4,
                ),
            )
        ],
        [
            Paragraph(
                f"{titulo}",
                ParagraphStyle(
                    "card_titulo",
                    alignment=TA_CENTER,
                    textColor="#666666",
                    spaceAfter=0,
                ),
            )
        ],
    ]

    card_table = Table(card_content, colWidths=[1.5 * inch])
    card_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#FFFFFF"),
                ),
                (
                    "BORDER",
                    (0, 0),
                    (-1, -1),
                    1,
                    colors.HexColor("#DDDDDD"),
                ),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return card_table


def criar_cards_resumo_pdf(story, df, pagesize):
    """
    Cria cards de resumo no PDF com os mesmos dados dos cards HTML.
    """
    df_num = df.copy()
    df_num["Valor Contratado"] = pd.to_numeric(
        df_num["Valor Contratado"], errors="coerce"
    ).fillna(0)

    total_valor_contratado = df_num["Valor Contratado"].sum()
    qtd_processos = len(df_num)
    media_por_processo = (
        total_valor_contratado / qtd_processos if qtd_processos > 0 else 0.0
    )

    concluidos = (df_num["Status"] == "Concluído").sum()
    em_andamento = (df_num["Status"] == "Em Andamento").sum()
    nao_concluidos = (df_num["Status"] == "Não Concluído").sum()

    story.append(Spacer(1, 0.08 * inch))

    card_data = [
        [
            criar_card_elemento(
                "Valor Contratado",
                formatar_moeda(total_valor_contratado),
                "#c00000",
            ),
            criar_card_elemento(
                "Média por Processo",
                formatar_moeda(media_por_processo),
                "#003A70",
            ),
            criar_card_elemento(
                "Número de Processos",
                str(qtd_processos),
                "#333333",
            ),
            criar_card_elemento(
                "Processos Concluídos",
                str(concluidos),
                "#003A70",
            ),
            criar_card_elemento(
                "Processos Em Andamento",
                str(em_andamento),
                "#F4A000",
            ),
            criar_card_elemento(
                "Processos Não Concluídos",
                str(nao_concluidos),
                "#DA291C",
            ),
        ]
    ]

    card_width = (pagesize[0] - 0.3 * inch) / 6 - 0.05 * inch
    cards_table = Table(
        card_data,
        colWidths=[
            card_width,
            card_width,
            card_width,
            card_width,
            card_width,
            card_width,
        ],
    )

    cards_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("GRID", (0, 0), (-1, -1), 0, colors.transparent),
            ]
        )
    )

    story.append(cards_table)
    story.append(Spacer(1, 0.15 * inch))

# Estilos PDF
wrap_style_compras = ParagraphStyle(
    name="wrap_compras",
    fontSize=7,
    leading=9,
    spaceAfter=2,
    wordWrap="CJK",
)

simple_style_compras = ParagraphStyle(
    name="simple_compras",
    fontSize=7,
    alignment=TA_CENTER,
)

header_cell_style_compras = ParagraphStyle(
    name="header_cell_compras",
    fontSize=7,
    alignment=TA_CENTER,
    fontName="Helvetica-Bold",
    textColor=colors.white,
)


def wrap_pdf_compras(text):
    return Paragraph(str(text), wrap_style_compras)


def simple_pdf_compras(text):
    return Paragraph(str(text), simple_style_compras)


def header_pdf_compras(text):
    return Paragraph(str(text), header_cell_style_compras)

# Cabeçalho PDF – Compras
def adicionar_cabecalho_compras(story, df, styles):
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
        "<b><font color='#0b2b57' size=13>Ministério da Educação</font></b><br/>"
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
        colWidths=[1.4 * inch, 4.2 * inch, 1.4 * inch],
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

    titulo = Paragraph(
        "RELATÓRIO DE PROCESSOS DE COMPRAS",
        ParagraphStyle(
            "titulo_compras",
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

# Tabela de dados no PDF
def criar_tabela_dados_compras(story, df, pagesize):
    if df.empty:
        return

    story.append(Spacer(1, 0.08 * inch))

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
        "CONTRATAÇÃO REINSTRUÍDA PELO PROCESSO Nº",
    ]

    cols = [c for c in cols if c in df.columns]
    df_pdf = df.copy()

    header = [header_pdf_compras(c) for c in cols]
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        linha = []
        for c in cols:
            valor = "" if pd.isna(row[c]) else str(row[c]).strip()
            if c in ["Objeto"]:
                linha.append(wrap_pdf_compras(valor))
            else:
                linha.append(simple_pdf_compras(valor))
        table_data.append(linha)

    col_widths = [
        0.7 * inch,
        1.2 * inch,
        1.2 * inch,
        1.2 * inch,
        1.1 * inch,
        1.1 * inch,
        0.9 * inch,
        0.9 * inch,
        0.9 * inch,
        1.2 * inch,
        1.0 * inch,
    ]
    col_widths = col_widths[: len(cols)]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_list = [
        (
            "BACKGROUND",
            (0, 0),
            (-1, 0),
            colors.HexColor("#0b2b57"),
        ),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        (
            "LINEBELOW",
            (0, 0),
            (-1, 0),
            1.5,
            colors.HexColor("#0b2b57"),
        ),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 1), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        # Linhas alternadas em branco e cinza claro
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -1),
            [colors.white, colors.HexColor("#f0f0f0")],
        ),
        ("WORDWRAP", (0, 0), (-1, -1), True),
    ]

    tbl.setStyle(TableStyle(style_list))
    story.append(tbl)

# --------------------------------------------------
# CALLBACK: GERAR PDF DE PROCESSOS DE COMPRAS
# --------------------------------------------------
@dash.callback(
    Output("download_relatorio_proc", "data"),
    Input("btn_download_relatorio_proc", "n_clicks"),
    State("store_dados_proc", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_proc(n, dados_proc):
    """
    Gera o relatório em PDF com base nos dados filtrados atualmente na tabela.
    """
    # VERIFICAÇÃO: Só executa se estiver na página de processos de compras
    if not verificar_pagina_processos_compras():
        raise PreventUpdate
    
    if not n or not dados_proc:
        return None

    df = pd.DataFrame(dados_proc)

    # Dataframe numérico para os cards
    df_cards = df.copy()
    df_cards["Valor Contratado"] = pd.to_numeric(
        df_cards["Valor Contratado"], errors="coerce"
    ).fillna(0)

    # Dataframe para tabela do PDF
    df_pdf = df.copy()
    df_pdf["PREÇO ESTIMADO"] = df_pdf["PREÇO ESTIMADO"].apply(formatar_moeda)
    df_pdf["Valor Contratado"] = df_pdf["Valor Contratado"].apply(
        formatar_moeda
    )
    df_pdf["Data de Entrada"] = pd.to_datetime(
        df_pdf["Data de Entrada"], format="%d/%m/%Y", errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    df_pdf["Data finalização"] = pd.to_datetime(
        df_pdf["Data finalização"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    buffer = BytesIO()
    pagesize = landscape(A4)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.15 * inch,
        leftMargin=0.15 * inch,
        topMargin=0.2 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    adicionar_cabecalho_compras(story, df_pdf, styles)
    criar_cards_resumo_pdf(story, df_cards, pagesize)
    criar_tabela_dados_compras(story, df_pdf, pagesize)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(
        buffer.getvalue(),
        f"processos_compras_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf",
    )