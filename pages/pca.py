import dash
from dash import html, dcc, dash_table, Input, Output, State
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
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.lib import colors

from datetime import datetime
from pytz import timezone
import os

dash.register_page(
    __name__,
    path="/pca",
    name="PCA",
    title="PCA",
)

URL_PCA = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=PCA%20-%20BI"
)

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
    "position": "relative",
    "zIndex": 1000,
}

# Estilo comum para botões (fundo azul, texto branco)
button_style = {
    "backgroundColor": "#0b2b57",
    "color": "white",
    "padding": "8px 16px",
    "border": "none",
    "borderRadius": "4px",
    "cursor": "pointer",
    "fontSize": "14px",
    "fontWeight": "bold",
}


def conv_moeda_br(v):
    if isinstance(v, str):
        v = v.strip()
        if v == "":
            return None
        v = v.replace(".", "").replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def formatar_moeda(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# --------------------------------------------------
# Carga e tratamento de dados
# --------------------------------------------------
def carregar_dados_pca():
    df = pd.read_csv(URL_PCA, header=0)
    df.columns = [c.strip() for c in df.columns]

    base_cols = [
        "Ano",
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Valor Total",
        "Saldo",
        "Item",
        "Código Classe / Grupo",
        "Nome Classe/Grupo",
        "Código PDM material",
        "Nome do PDM material",
        "Processo",
        "Observações",
        "Objeto",
        "SRP ou Outro Valor",
        "Valor",
    ]

    for c in base_cols:
        if c not in df.columns:
            df[c] = None

    df["Valor Total"] = df["Valor Total"].apply(conv_moeda_br)
    df["Saldo"] = df["Saldo"].apply(conv_moeda_br)

    for c in df.columns:
        if c.startswith("Valor"):
            df[c] = df[c].apply(conv_moeda_br)
        if c.startswith("SRP ou Outro Valor"):
            df[c] = df[c].apply(conv_moeda_br)

    df["Ano"] = df["Ano"].astype("string")

    # Colunas de texto
    for c in [
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Item",
        "Código Classe / Grupo",
        "Nome Classe/Grupo",
        "Nome do PDM material",
        "Processo",
        "Observações",
        "Objeto",
    ]:
        if c in df.columns:
            df[c] = df[c].astype("string")

    # Conversão específica de Código PDM material para inteiro (nullable)
    if "Código PDM material" in df.columns:
        df["Código PDM material"] = (
            df["Código PDM material"]
            .apply(lambda x: str(x).strip() if pd.notna(x) else "")
        )
        df["Código PDM material"] = pd.to_numeric(
            df["Código PDM material"].replace({"": None}), errors="coerce"
        ).astype("Int64")

    return df


df_pca_base = carregar_dados_pca()

# ------------------ Tabela 1: Planejamento ------------------
df_planejamento = df_pca_base.copy()
df_planejamento["Planejado"] = df_planejamento["Valor Total"]
df_planejamento["Executado"] = df_planejamento["Planejado"] - df_planejamento["Saldo"]

# ------------------ Tabela 2: Processos (explodindo colunas) ------------------
cols_grupo0 = [
    "Ano",
    "Área requisitante",
    "Material ou Serviço",
    "DFD",
    "Item",
    "Valor Total",
    "Saldo",
    "Processo",
    "Observações",
    "Objeto",
    "SRP ou Outro Valor",
    "Valor",
]

for c in cols_grupo0:
    if c not in df_pca_base.columns:
        df_pca_base[c] = None

grupo0 = df_pca_base[cols_grupo0].copy()


def gerar_grupo(indice: int) -> pd.DataFrame:
    suf = f".{indice}"
    col_processo = f"Processo{suf}"
    col_observ = f"Observações{suf}"
    col_objeto = f"Objeto{suf}"
    col_srp = f"SRP ou Outro Valor{suf}"
    col_valor = f"Valor{suf}"

    colunas_originais = [
        "Ano",
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Item",
        "Valor Total",
        "Saldo",
        col_processo,
        col_observ,
        col_objeto,
        col_srp,
        col_valor,
    ]

    for c in colunas_originais:
        if c not in df_pca_base.columns:
            df_pca_base[c] = None

    tabela_sel = df_pca_base[colunas_originais].copy()
    tabela_ren = tabela_sel.rename(
        columns={
            col_processo: "Processo",
            col_observ: "Observações",
            col_objeto: "Objeto",
            col_srp: "SRP ou Outro Valor",
            col_valor: "Valor",
        }
    )
    return tabela_ren


grupos_dinamicos = [gerar_grupo(i) for i in range(1, 32)]
tabela_processos_unida = pd.concat([grupo0] + grupos_dinamicos, ignore_index=True)

for c in [
    "Área requisitante",
    "Material ou Serviço",
    "DFD",
    "Item",
    "Processo",
    "Observações",
    "Objeto",
]:
    tabela_processos_unida[c] = tabela_processos_unida[c].astype("string")

tabela_processos_unida["Valor"] = tabela_processos_unida["Valor"].apply(conv_moeda_br)

# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_pca",
            className="filtros-sticky",
            children=[
                # Linha 1 de filtros
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "120px", "flex": "0 0 140px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_pca",
                                    options=[
                                        {"label": str(a), "value": str(a)}
                                        for a in sorted(
                                            df_pca_base["Ano"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a).strip() != ""
                                    ],
                                    value="2026",
                                    placeholder=None,
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Material ou Serviço"),
                                dcc.Dropdown(
                                    id="filtro_tipo_pca",
                                    options=[
                                        {"label": t, "value": t}
                                        for t in sorted(
                                            df_pca_base["Material ou Serviço"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(t).strip() != ""
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
                                html.Label("Área requisitante"),
                                dcc.Dropdown(
                                    id="filtro_area_pca",
                                    options=[
                                        {"label": a, "value": a}
                                        for a in sorted(
                                            df_pca_base["Área requisitante"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Nome Classe/Grupo (digitação)"),
                                dcc.Input(
                                    id="filtro_classe_texto_pca",
                                    type="text",
                                    placeholder=(
                                        "Digite parte do nome da classe/grupo"
                                    ),
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("DFD (digitação)"),
                                dcc.Input(
                                    id="filtro_dfd_texto_pca",
                                    type="text",
                                    placeholder="Digite parte do DFD",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
                # Linha 2: botões + cartões
                html.Div(
                    style={
                        "display": "flex",
                        "alignItems": "center",
                        "justifyContent": "space-between",
                        "marginTop": "6px",
                        "flexWrap": "wrap",
                        "gap": "10px",
                    },
                    children=[
                        html.Div(
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_pca",
                                    n_clicks=0,
                                    style={**button_style, "marginRight": "10px"},
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_pca",
                                    n_clicks=0,
                                    style=button_style,
                                ),
                                dcc.Download(id="download_relatorio_pca"),
                            ],
                        ),
                        html.Div(
                            style={
                                "display": "flex",
                                "justifyContent": "center",
                                "flexGrow": 1,
                                "gap": "10px",
                                "flexWrap": "wrap",
                            },
                            children=[
                                html.Div(
                                    id="card_planejado_pca",
                                    style={
                                        "minWidth": "180px",
                                        "padding": "10px 15px",
                                        "backgroundColor": "#f5f5f5",
                                        "textAlign": "center",
                                    },
                                ),
                                html.Div(
                                    id="card_executado_pca",
                                    style={
                                        "minWidth": "180px",
                                        "padding": "10px 15px",
                                        "backgroundColor": "#f5f5f5",
                                        "textAlign": "center",
                                    },
                                ),
                                html.Div(
                                    id="card_saldo_pca",
                                    style={
                                        "minWidth": "180px",
                                        "padding": "10px 15px",
                                        "backgroundColor": "#f5f5f5",
                                        "textAlign": "center",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        html.Div(
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "10px",
                "marginTop": "10px",
            },
            children=[
                html.Div(
                    style={"flex": "1 1 50%", "minWidth": "300px"},
                    children=[
                        html.H4("Planejamento (PCA)"),
                        dash_table.DataTable(
                            id="tabela_pca_planejamento",
                            columns=[
                                {"name": "DFD", "id": "DFD"},
                                {
                                    "name": "Área requisitante",
                                    "id": "Área requisitante",
                                },
                                {
                                    "name": "Material ou Serviço",
                                    "id": "Material ou Serviço",
                                },
                                {"name": "Item", "id": "Item"},
                                {
                                    "name": "Nome Classe/Grupo",
                                    "id": "Nome Classe/Grupo",
                                },
                                {
                                    "name": "Código PDM material",
                                    "id": "Código PDM material",
                                },
                                {
                                    "name": "Nome do PDM material",
                                    "id": "Nome do PDM material",
                                },
                                {
                                    "name": "Planejado",
                                    "id": "Planejado_fmt",
                                },
                                {
                                    "name": "Executado",
                                    "id": "Executado_fmt",
                                },
                                {"name": "Saldo", "id": "Saldo_fmt"},
                            ],
                            data=[],
                            fixed_rows={"headers": True},
                            style_table={
                                "overflowX": "auto",
                                "overflowY": "auto",
                                "maxHeight": "420px",
                                "width": "100%",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "6px",
                                "fontSize": "12px",
                                "whiteSpace": "normal",
                            },
                            style_cell_conditional=[
                                {
                                    "if": {"column_id": "DFD"},
                                    "width": "6%",
                                },
                                {
                                    "if": {"column_id": "Área requisitante"},
                                    "width": "12%",
                                },
                                {
                                    "if": {"column_id": "Material ou Serviço"},
                                    "width": "10%",
                                },
                                {
                                    "if": {"column_id": "Item"},
                                    "width": "5%",
                                },
                                {
                                    "if": {"column_id": "Nome Classe/Grupo"},
                                    "width": "20%",
                                },
                                {
                                    "if": {"column_id": "Código PDM material"},
                                    "width": "10%",
                                },
                                {
                                    "if": {"column_id": "Nome do PDM material"},
                                    "width": "16%",
                                },
                                {
                                    "if": {"column_id": "Planejado_fmt"},
                                    "width": "7%",
                                },
                                {
                                    "if": {"column_id": "Executado_fmt"},
                                    "width": "7%",
                                },
                                {
                                    "if": {"column_id": "Saldo_fmt"},
                                    "width": "8%",
                                },
                            ],
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                            },
                            style_data_conditional=[
                                {
                                    "if": {"filter_query": "{Saldo_num} <= 0"},
                                    "backgroundColor": "#ffcccc",
                                },
                            ],
                        ),
                    ],
                ),
                html.Div(
                    style={"flex": "1 1 50%", "minWidth": "300px"},
                    children=[
                        html.H4("Processos vinculados ao PCA"),
                        dash_table.DataTable(
                            id="tabela_pca_processos",
                            columns=[
                                {"name": "DFD", "id": "DFD"},
                                {
                                    "name": "Área requisitante",
                                    "id": "Área requisitante",
                                },
                                {
                                    "name": "Material ou Serviço",
                                    "id": "Material ou Serviço",
                                },
                                {"name": "Item", "id": "Item"},
                                {"name": "Processo", "id": "Processo"},
                                {"name": "Objeto", "id": "Objeto"},
                                {
                                    "name": "Observações",
                                    "id": "Observações",
                                },
                                {"name": "Valor", "id": "Valor_fmt"},
                            ],
                            data=[],
                            fixed_rows={"headers": True},
                            style_table={
                                "overflowX": "auto",
                                "overflowY": "auto",
                                "maxHeight": "420px",
                                "width": "100%",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "6px",
                                "fontSize": "12px",
                                "whiteSpace": "normal",
                            },
                            style_cell_conditional=[
                                {
                                    "if": {"column_id": "DFD"},
                                    "width": "10%",
                                },
                                {
                                    "if": {"column_id": "Área requisitante"},
                                    "width": "12%",
                                },
                                {
                                    "if": {"column_id": "Material ou Serviço"},
                                    "width": "10%",
                                },
                                {
                                    "if": {"column_id": "Item"},
                                    "width": "6%",
                                },
                                {
                                    "if": {"column_id": "Processo"},
                                    "width": "12%",
                                },
                                {
                                    "if": {"column_id": "Objeto"},
                                    "width": "25%",
                                    "textAlign": "left",
                                },
                                {
                                    "if": {"column_id": "Observações"},
                                    "width": "15%",
                                    "textAlign": "left",
                                },
                                {
                                    "if": {"column_id": "Valor_fmt"},
                                    "width": "10%",
                                },
                            ],
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                            },
                        ),
                    ],
                ),
            ],
        ),
        dcc.Store(id="store_dados_pca_processos"),
    ],
)

# --------------------------------------------------
# Callbacks de dados + cartões
# --------------------------------------------------
@dash.callback(
    Output("tabela_pca_planejamento", "data"),
    Output("tabela_pca_processos", "data"),
    Output("store_dados_pca_processos", "data"),
    Output("card_planejado_pca", "children"),
    Output("card_executado_pca", "children"),
    Output("card_saldo_pca", "children"),
    Input("filtro_ano_pca", "value"),
    Input("filtro_classe_texto_pca", "value"),
    Input("filtro_dfd_texto_pca", "value"),
    Input("filtro_area_pca", "value"),
    Input("filtro_tipo_pca", "value"),
)
def atualizar_tabelas_pca(
    ano,
    classe_texto,
    dfd_texto,
    area,
    tipo,
):
    dff_plan = df_planejamento.copy()
    dff_proc = tabela_processos_unida.copy()

    if ano:
        dff_plan = dff_plan[dff_plan["Ano"] == str(ano)]
        dff_proc = dff_proc[dff_proc["Ano"] == str(ano)]

    if classe_texto and str(classe_texto).strip():
        termo = str(classe_texto).strip().lower()
        if "Nome Classe/Grupo" in dff_plan.columns:
            dff_plan = dff_plan[
                dff_plan["Nome Classe/Grupo"]
                .astype(str)
                .str.lower()
                .str.contains(termo, na=False)
            ]
        if "Nome Classe/Grupo" in dff_proc.columns:
            dff_proc = dff_proc[
                dff_proc["Nome Classe/Grupo"]
                .astype(str)
                .str.lower()
                .str.contains(termo, na=False)
            ]

    if dfd_texto and str(dfd_texto).strip():
        termo = str(dfd_texto).strip().lower()
        dff_plan = dff_plan[
            dff_plan["DFD"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]
        dff_proc = dff_proc[
            dff_proc["DFD"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    if area:
        dff_plan = dff_plan[dff_plan["Área requisitante"] == area]
        dff_proc = dff_proc[dff_proc["Área requisitante"] == area]

    if tipo:
        dff_plan = dff_plan[dff_plan["Material ou Serviço"] == tipo]
        dff_proc = dff_proc[dff_proc["Material ou Serviço"] == tipo]

    # remove linhas sem processo
    dff_proc = dff_proc[
        (dff_proc["Processo"].astype(str).str.strip() != "")
        & (
            dff_proc["Processo"]
            .astype(str)
            .str.strip()
            .str.lower()
            != "nan"
        )
        & (dff_proc["Processo"].notna())
    ]

    # Item inteiro
    dff_plan["Item"] = (
        dff_plan["Item"]
        .fillna("")
        .astype(str)
        .apply(
            lambda x: str(int(float(x))) if x not in ["", "nan"] else ""
        )
    )

    dff_proc["Item"] = (
        dff_proc["Item"]
        .fillna("")
        .astype(str)
        .apply(
            lambda x: str(int(float(x))) if x not in ["", "nan"] else ""
        )
    )

    dff_plan["Saldo_num"] = dff_plan["Saldo"]

    # Exibir Código PDM material como inteiro (sem NaN)
    if "Código PDM material" in dff_plan.columns:
        dff_plan["Código PDM material"] = dff_plan["Código PDM material"].apply(
            lambda x: "" if pd.isna(x) else int(x)
        )

    def marca_executado(v):
        if v is None or pd.isna(v):
            return ""
        try:
            v = float(v)
        except (TypeError, ValueError):
            return ""
        marcador = " ✔" if v > 0 else ""
        return formatar_moeda(v) + marcador

    dff_plan["Planejado_fmt"] = dff_plan["Planejado"].apply(formatar_moeda)
    dff_plan["Executado_fmt"] = dff_plan["Executado"].apply(marca_executado)
    dff_plan["Saldo_fmt"] = dff_plan["Saldo"].apply(formatar_moeda)

    dff_proc["Valor_fmt"] = dff_proc["Valor"].apply(formatar_moeda)

    cols_planejamento = [
        "DFD",
        "Área requisitante",
        "Material ou Serviço",
        "Item",
        "Nome Classe/Grupo",
        "Código PDM material",
        "Nome do PDM material",
        "Planejado_fmt",
        "Executado_fmt",
        "Saldo_fmt",
        "Saldo_num",
        "Planejado",
        "Executado",
        "Saldo",
    ]

    dados_planejamento = (
        dff_plan[cols_planejamento].fillna("").to_dict("records")
    )

    cols_processos = [
        "DFD",
        "Área requisitante",
        "Material ou Serviço",
        "Item",
        "Processo",
        "Objeto",
        "Observações",
        "Valor_fmt",
    ]

    dados_processos_df = dff_proc[cols_processos].fillna("")
    dados_processos = dados_processos_df.to_dict("records")

    total_planejado = dff_plan["Planejado"].sum()
    total_executado = dff_plan["Executado"].sum()
    total_saldo = dff_plan["Saldo"].sum()

    card_planejado = html.Div(
        [
            html.Div(
                formatar_moeda(total_planejado),
                style={
                    "color": "#c0392b",
                    "fontSize": "20px",
                    "fontWeight": "bold",
                },
            ),
            html.Div("Planejado"),
        ]
    )

    card_executado = html.Div(
        [
            html.Div(
                formatar_moeda(total_executado),
                style={
                    "color": "#0b2b57",
                    "fontSize": "20px",
                    "fontWeight": "bold",
                },
            ),
            html.Div("Executado"),
        ]
    )

    card_saldo = html.Div(
        [
            html.Div(
                formatar_moeda(total_saldo),
                style={
                    "color": "#2c3e50",
                    "fontSize": "20px",
                    "fontWeight": "bold",
                },
            ),
            html.Div("Saldo"),
        ]
    )

    return (
        dados_planejamento,
        dados_processos,
        dados_processos_df.to_dict("records"),
        card_planejado,
        card_executado,
        card_saldo,
    )


@dash.callback(
    Output("filtro_ano_pca", "value"),
    Output("filtro_classe_texto_pca", "value"),
    Output("filtro_dfd_texto_pca", "value"),
    Output("filtro_area_pca", "value"),
    Output("filtro_tipo_pca", "value"),
    Input("btn_limpar_filtros_pca", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_pca(n):
    return "2026", None, None, None, None


# --------------------------------------------------
# PDF - estilos para PCA
# --------------------------------------------------
wrap_style_pca = ParagraphStyle(
    name="wrap_pca_pdf",
    fontSize=7,
    leading=8,
    spaceAfter=2,
    wordWrap="CJK",
)

simple_style_pca = ParagraphStyle(
    name="simple_pca_pdf",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
)


def wrap_pdf(text):
    return Paragraph(str(text), wrap_style_pca)


def simple_pdf(text):
    return Paragraph(str(text), simple_style_pca)


# --------------------------------------------------
# Callback: gerar PDF do PCA
# --------------------------------------------------
@dash.callback(
    Output("download_relatorio_pca", "data"),
    Input("btn_download_relatorio_pca", "n_clicks"),
    State("store_dados_pca_processos", "data"),
    State("tabela_pca_planejamento", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_pca(n, dados_processos, dados_planejamento):
    from dash import dcc

    if not n or (not dados_processos and not dados_planejamento):
        return None

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

    # Data e hora
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora_brasilia = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")
    data_top_table = Table(
        [
            [
                Paragraph(
                    data_hora_brasilia,
                    ParagraphStyle(
                        "data_topo_pca",
                        fontSize=9,
                        alignment=TA_RIGHT,
                        textColor="#333333",
                    ),
                )
            ]
        ],
        colWidths=[pagesize[0] - 0.3 * inch],
    )

    data_top_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    story.append(data_top_table)
    story.append(Spacer(1, 0.1 * inch))

    # Cabeçalho: Logo esq | Instituição | Logo dir
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
            "instituicao_fiscais",
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

    # Título principal
    titulo = Paragraph(
        "RELATÓRIO DE PLANEJAMENTO DE CONTRATAÇÃO ANUAL (PCA)<br/>",
        ParagraphStyle(
            "titulo_fiscais",
            alignment=TA_CENTER,
            fontSize=10,
            leading=14,
            textColor=colors.black,
        ),
    )

    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    # ========================
    # TABELA 1: PLANEJAMENTO
    # ========================
    if dados_planejamento:
        df_plan = pd.DataFrame(dados_planejamento)

        story.append(
            Paragraph(
                "PLANEJAMENTO (PCA)",
                ParagraphStyle(
                    "subtitulo_plan",
                    fontSize=9,
                    alignment=TA_LEFT,
                    textColor="#0b2b57",
                    fontName="Helvetica-Bold",
                    spaceAfter=6,
                ),
            )
        )

        story.append(
            Paragraph(
                f"Total de registros: {len(df_plan)}",
                styles["Normal"],
            )
        )

        story.append(Spacer(1, 0.08 * inch))

        cols_plan = [
            "DFD",
            "Área requisitante",
            "Material ou Serviço",
            "Item",
            "Nome Classe/Grupo",
            "Planejado_fmt",
            "Executado_fmt",
            "Saldo_fmt",
        ]

        cols_plan = [c for c in cols_plan if c in df_plan.columns]
        df_plan_filtered = df_plan[cols_plan].copy()

        header_plan = [
            "DFD",
            "Área requisitante",
            "Material ou Serviço",
            "Item",
            "Nome Classe/Grupo",
            "Planejado",
            "Executado",
            "Saldo",
        ]
        table_data_plan = [header_plan]

        for _, row in df_plan_filtered.iterrows():
            linha = []
            for c in cols_plan:
                valor = str(row[c]).strip()
                if c in ["Nome Classe/Grupo"]:
                    linha.append(wrap_pdf(valor))
                else:
                    linha.append(simple_pdf(valor))
            table_data_plan.append(linha)

        col_widths_plan = [
            0.7 * inch,  # DFD
            1.1 * inch,  # Área requisitante
            1.0 * inch,  # Material ou Serviço
            0.5 * inch,  # Item
            2.0 * inch,  # Nome Classe/Grupo
            0.95 * inch,  # Planejado
            0.95 * inch,  # Executado
            0.95 * inch,  # Saldo
        ]

        tbl_plan = Table(
            table_data_plan, colWidths=col_widths_plan, repeatRows=1
        )

        style_list_plan = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, 0), 7),
            ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (0, 1), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 1), (-1, -1), 7),
            ("WORDWRAP", (0, 0), (-1, -1), True),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#f0f0f0")],
            ),
        ]

        # Destacar linha em vermelho quando saldo <= 0
        saldo_col_index = cols_plan.index("Saldo_fmt") if "Saldo_fmt" in cols_plan else None

        if saldo_col_index is not None:
            for row_idx in range(1, len(table_data_plan)):
                try:
                    saldo_paragraph = table_data_plan[row_idx][saldo_col_index]
                    saldo_str = getattr(saldo_paragraph, "text", str(saldo_paragraph))
                    saldo_str = (
                        saldo_str.replace("R$", "")
                        .replace(".", "")
                        .replace(",", ".")
                        .strip()
                    )
                    saldo_valor = float(saldo_str) if saldo_str not in ("", "-") else 0.0

                    if saldo_valor <= 0:
                        style_list_plan.append(
                            ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#ffcccc"))
                        )
                        style_list_plan.append(
                            ("TEXTCOLOR", (0, row_idx), (-1, row_idx), colors.HexColor("#cc0000"))
                        )
                except (ValueError, IndexError):
                    pass

        tbl_plan.setStyle(TableStyle(style_list_plan))
        story.append(tbl_plan)
        story.append(Spacer(1, 0.2 * inch))

    # ==============================
    # TABELA 2: PROCESSOS VINCULADOS
    # ==============================
    if dados_processos:
        df_proc = pd.DataFrame(dados_processos)

        story.append(
            Paragraph(
                "PROCESSOS VINCULADOS AO PCA",
                ParagraphStyle(
                    "subtitulo_proc",
                    fontSize=9,
                    alignment=TA_LEFT,
                    textColor="#0b2b57",
                    fontName="Helvetica-Bold",
                    spaceAfter=6,
                ),
            )
        )

        story.append(
            Paragraph(
                f"Total de registros: {len(df_proc)}",
                styles["Normal"],
            )
        )

        story.append(Spacer(1, 0.08 * inch))

        cols_proc = [
            "DFD",
            "Área requisitante",
            "Material ou Serviço",
            "Item",
            "Processo",
            "Objeto",
            "Observações",
            "Valor_fmt",
        ]
        cols_proc = [c for c in cols_proc if c in df_proc.columns]
        df_proc_filtered = df_proc[cols_proc].copy()

        header_proc = [
            "DFD",
            "Área requisitante",
            "Material ou Serviço",
            "Item",
            "Processo",
            "Objeto",
            "Observações",
            "Valor",
        ]
        table_data_proc = [header_proc]

        for _, row in df_proc_filtered.iterrows():
            linha = []
            for c in cols_proc:
                valor = str(row[c]).strip()
                if c in ["Objeto", "Observações"]:
                    linha.append(wrap_pdf(valor))
                else:
                    linha.append(simple_pdf(valor))
            table_data_proc.append(linha)

        col_widths_proc = [
            0.7 * inch,  # DFD
            1.1 * inch,  # Área requisitante
            1.0 * inch,  # Material ou Serviço
            0.5 * inch,  # Item
            1.1 * inch,  # Processo
            2.0 * inch,  # Objeto
            2.0 * inch,  # Observações
            0.95 * inch,  # Valor
        ]

        tbl_proc = Table(
            table_data_proc, colWidths=col_widths_proc, repeatRows=1
        )

        style_list_proc = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, 0), 7),
            ("FONTWEIGHT", (0, 0), (-1, 0), "bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (0, 1), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 1), (-1, -1), 7),
            ("WORDWRAP", (0, 0), (-1, -1), True),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#f0f0f0")],
            ),
        ]

        # Opcional: destacar linha quando Valor <= 0
        valor_col_index = cols_proc.index("Valor_fmt") if "Valor_fmt" in cols_proc else None

        if valor_col_index is not None:
            for row_idx in range(1, len(table_data_proc)):
                try:
                    valor_paragraph = table_data_proc[row_idx][valor_col_index]
                    valor_str = getattr(valor_paragraph, "text", str(valor_paragraph))
                    valor_str = (
                        valor_str.replace("R$", "")
                        .replace(".", "")
                        .replace(",", ".")
                        .strip()
                    )
                    valor_numerico = float(valor_str) if valor_str not in ("", "-") else 0.0

                    if valor_numerico <= 0:
                        style_list_proc.append(
                            ("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#ffcccc"))
                        )
                        style_list_proc.append(
                            ("TEXTCOLOR", (0, row_idx), (-1, row_idx), colors.HexColor("#cc0000"))
                        )
                except (ValueError, IndexError):
                    pass

        tbl_proc.setStyle(TableStyle(style_list_proc))
        story.append(tbl_proc)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(
        f"relatorio_pca_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
    )
