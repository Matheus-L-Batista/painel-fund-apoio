import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from datetime import datetime

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

    for c in [
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Item",
        "Código Classe / Grupo",
        "Nome Classe/Grupo",
        "Código PDM material",
        "Nome do PDM material",
        "Processo",
        "Observações",
        "Objeto",
    ]:
        if c in df.columns:
            df[c] = df[c].astype("string")

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
                                    # instrução da imagem: ano já selecionado,
                                    # sem placeholder "Todos" e sem clearable
                                    value="2025",
                                    placeholder=None,
                                    clearable=False,
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
                                    placeholder="Digite parte do nome da classe/grupo",
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
                                html.Label("Nome Classe/Grupo"),
                                dcc.Dropdown(
                                    id="filtro_classe_pca",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in sorted(
                                            df_pca_base["Nome Classe/Grupo"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(c).strip() != ""
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
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("DFD"),
                                dcc.Dropdown(
                                    id="filtro_dfd_pca",
                                    options=[
                                        {"label": d, "value": d}
                                        for d in sorted(
                                            df_pca_base["DFD"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(d).strip() != ""
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
                html.Div(
                    style={
                        "display": "flex",
                        "marginTop": "4px",
                        "alignItems": "center",
                        "gap": "10px",
                    },
                    children=[
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
                            style={"marginTop": "18px"},
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_pca",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_pca",
                                    n_clicks=0,
                                    className="filtros-button",
                                    style={"marginLeft": "10px"},
                                ),
                                dcc.Download(id="download_relatorio_pca"),
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
                                {"name": "Planejado", "id": "Planejado_fmt"},
                                {
                                    "name": "Executado",
                                    "id": "Executado_fmt",
                                },
                                {"name": "Saldo", "id": "Saldo_fmt"},
                            ],
                            data=[],
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
                                {"if": {"column_id": "DFD"}, "width": "6%"},
                                {
                                    "if": {"column_id": "Área requisitante"},
                                    "width": "12%",
                                },
                                {
                                    "if": {"column_id": "Material ou Serviço"},
                                    "width": "10%",
                                },
                                {"if": {"column_id": "Item"}, "width": "5%"},
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
                                {"if": {"column_id": "Saldo_fmt"}, "width": "8%"},
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
                                {"if": {"column_id": "DFD"}, "width": "10%"},
                                {
                                    "if": {"column_id": "Área requisitante"},
                                    "width": "12%",
                                },
                                {
                                    "if": {"column_id": "Material ou Serviço"},
                                    "width": "10%",
                                },
                                {"if": {"column_id": "Item"}, "width": "6%"},
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
# Callbacks
# --------------------------------------------------

@dash.callback(
    Output("tabela_pca_planejamento", "data"),
    Output("tabela_pca_processos", "data"),
    Output("store_dados_pca_processos", "data"),
    Input("filtro_ano_pca", "value"),
    Input("filtro_classe_texto_pca", "value"),
    Input("filtro_classe_pca", "value"),
    Input("filtro_dfd_texto_pca", "value"),
    Input("filtro_dfd_pca", "value"),
    Input("filtro_area_pca", "value"),
    Input("filtro_tipo_pca", "value"),
)
def atualizar_tabelas_pca(
    ano,
    classe_texto,
    classe_select,
    dfd_texto,
    dfd_select,
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

    if classe_select:
        if "Nome Classe/Grupo" in dff_plan.columns:
            dff_plan = dff_plan[dff_plan["Nome Classe/Grupo"] == classe_select]
        if "Nome Classe/Grupo" in dff_proc.columns:
            dff_proc = dff_proc[dff_proc["Nome Classe/Grupo"] == classe_select]

    if dfd_texto and str(dfd_texto).strip():
        termo = str(dfd_texto).strip().lower()
        dff_plan = dff_plan[
            dff_plan["DFD"].astype(str).str.lower().str.contains(termo, na=False)
        ]
        dff_proc = dff_proc[
            dff_proc["DFD"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    if dfd_select:
        dff_plan = dff_plan[dff_plan["DFD"] == dfd_select]
        dff_proc = dff_proc[dff_proc["DFD"] == dfd_select]

    if area:
        dff_plan = dff_plan[dff_plan["Área requisitante"] == area]
        dff_proc = dff_proc[dff_proc["Área requisitante"] == area]

    if tipo:
        dff_plan = dff_plan[dff_plan["Material ou Serviço"] == tipo]
        dff_proc = dff_proc[dff_proc["Material ou Serviço"] == tipo]

    # remove linhas sem processo
    dff_proc = dff_proc[dff_proc["Processo"].astype(str).str.strip() != ""]

    # Item inteiro
    dff_plan["Item"] = (
        dff_plan["Item"]
        .fillna("")
        .astype(str)
        .apply(lambda x: str(int(float(x))) if x not in ["", "nan"] else "")
    )
    dff_proc["Item"] = (
        dff_proc["Item"]
        .fillna("")
        .astype(str)
        .apply(lambda x: str(int(float(x))) if x not in ["", "nan"] else "")
    )

    dff_plan["Saldo_num"] = dff_plan["Saldo"]

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
    ]

    dados_planejamento = dff_plan[cols_planejamento].fillna("").to_dict("records")

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

    return dados_planejamento, dados_processos, dados_processos_df.to_dict("records")

@dash.callback(
    Output("filtro_ano_pca", "value"),
    Output("filtro_classe_texto_pca", "value"),
    Output("filtro_classe_pca", "value"),
    Output("filtro_dfd_texto_pca", "value"),
    Output("filtro_dfd_pca", "value"),
    Output("filtro_area_pca", "value"),
    Output("filtro_tipo_pca", "value"),
    Input("btn_limpar_filtros_pca", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_pca(n):
    # instrução da imagem: voltar sempre para 2025
    return "2025", None, None, None, None, None, None

wrap_style_pca = ParagraphStyle(
    name="wrap_pca",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)

def wrap_text_pca(text):
    return Paragraph(str(text), wrap_style_pca)

@dash.callback(
    Output("download_relatorio_pca", "data"),
    Input("btn_download_relatorio_pca", "n_clicks"),
    State("store_dados_pca_processos", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_pca(n, dados_pca):
    if not n or not dados_pca:
        return None
    df = pd.DataFrame(dados_pca)

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
        "Relatório PCA - Processos vinculados",
        ParagraphStyle(
            "titulo_pca",
            fontSize=16,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0b2b57"),
        ),
    )

    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        "DFD",
        "Área requisitante",
        "Material ou Serviço",
        "Item",
        "Processo",
        "Objeto",
        "Observações",
        "Valor",
    ]

    cols = [c for c in cols if c in df.columns]
    df_pdf = df.copy()
    if "Valor" in df_pdf.columns:
        df_pdf["Valor"] = df_pdf["Valor"].apply(formatar_moeda)

    header = cols
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap_text_pca(row[c]) for c in cols])

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

    return dcc.send_bytes(buffer.getvalue(), "pca_processos_paisagem.pdf")
