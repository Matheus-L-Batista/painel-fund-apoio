import dash
from dash import html, dcc, dash_table, Input, Output, State, callback
import pandas as pd
from datetime import datetime
from dash.exceptions import PreventUpdate

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
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from pytz import timezone
import os


# --------------------------------------------------
# Função para verificar se estamos na página de contratos
# --------------------------------------------------
def verificar_pagina_contratos():
    """Verifica se o callback está sendo executado na página de contratos"""
    try:
        if not dash.ctx.triggered:
            # Permite execução inicial
            return True

        # Componentes específicos da página de contratos
        componentes_contratos = {
            "filtro_contrato",
            "filtro_objeto",
            "filtro_setor",
            "filtro_grupo",
            "filtro_empresa",
            "filtro_status_vig",
            "btn_limpar_filtros_contratos",
            "btn_download_relatorio_contratos",
        }

        # Obtém o ID do componente que disparou o callback
        triggered = dash.ctx.triggered[0]
        triggered_id = triggered["prop_id"].split(".")[0]

        # Verifica se é um componente da página de contratos
        return triggered_id in componentes_contratos
    except Exception:
        # Em caso de erro, permite a execução (segurança para inicialização)
        return True


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/contratos",
    name="Contratos",
    title="Contratos",
)


# --------------------------------------------------
# URL da planilha de Contratos
# --------------------------------------------------
URL_CONTRATOS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Grupo%20da%20Cont."
)

# Grupo fixo a exibir
GRUPO_FIXO = "FUNDAÇÃO DE APOIO"


# nomes exatos das colunas originais no CSV
COL_CONTRATO = "Contrato"
COL_SETOR = "Setor"
COL_MENU_GRUPO = "MENU Grupo"
COL_OBJETO_ORIG = (
    "UNIVERSIDADE FEDERAL DE ITAJUBÁ Diretoria de Compras e Contratos "
    "Campus Itajubá CONTRATOS ATIVOS - ALIMENTAÇÃO DO BI Objeto"
)
COL_EMPRESA = "Empresa Contratada"
COL_INICIO_VIG = "Início da Vigência"
COL_TERMINO_EXEC = "Término da Execução"
COL_TERMINO_VIG = "Termino da Vigência"  # igual na planilha
COL_LINK_COMPRASNET = "Comprasnet Contratos"


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados_contratos():
    df = pd.read_csv(URL_CONTRATOS, header=0)
    df.columns = [c.strip() for c in df.columns]

    if COL_LINK_COMPRASNET not in df.columns:
        df[COL_LINK_COMPRASNET] = ""

    df = df.rename(
        columns={
            COL_CONTRATO: "Contrato",
            COL_SETOR: "Setor",
            COL_MENU_GRUPO: "Grupo",
            COL_OBJETO_ORIG: "Objeto",
            COL_EMPRESA: "Empresa Contratada",
            COL_INICIO_VIG: "Início da Vigência",
            COL_TERMINO_EXEC: "Término da Execução",
            COL_TERMINO_VIG: "Término da Vigência",
            COL_LINK_COMPRASNET: "Link Comprasnet",
        }
    )

    # ✅ FILTRO FIXO: somente FUNDAÇÃO DE APOIO
    df = df[df["Grupo"].astype(str).str.strip().str.upper() == GRUPO_FIXO.upper()]

    # Converte datas para datetime para cálculo do status
    for col in ["Início da Vigência", "Término da Execução", "Término da Vigência"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

    hoje = datetime.now().date()

    def calcular_status(data_termino_exec):
        if pd.isna(data_termino_exec):
            return ""
        dias = (data_termino_exec.date() - hoje).days
        if dias > 10:
            return "Vigente"
        if dias < 0:
            return "Vencido"
        return "Próximo do Vencimento"

    df["Status da Vigência"] = df["Término da Execução"].apply(calcular_status)

    # Formata datas para string dd/mm/aaaa para exibição
    for col in ["Início da Vigência", "Término da Execução", "Término da Vigência"]:
        if col in df.columns:
            df[col] = df[col].dt.strftime("%d/%m/%Y").fillna("")

    return df


df_contratos_base = carregar_dados_contratos()


# --------------------------------------------------
# Função auxiliar: filtros em cascata independentes
# --------------------------------------------------
def filtrar_contratos(
    contrato_texto,
    objeto_texto,
    setor,
    grupo,
    empresa,
    status_vig,
):
    dff = df_contratos_base.copy()

    # Contrato (texto) - busca parcial conforme digitação
    if contrato_texto and str(contrato_texto).strip():
        termo = str(contrato_texto).strip().lower()
        dff = dff[
            dff["Contrato"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Objeto (texto) - busca parcial conforme digitação
    if objeto_texto and str(objeto_texto).strip():
        termo = str(objeto_texto).strip().lower()
        dff = dff[
            dff["Objeto"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Setor - aceita lista de valores
    if setor:
        if isinstance(setor, str):
            setor = [setor]
        dff = dff[dff["Setor"].isin(setor)]

    # Grupo - aceita lista de valores (mesmo que agora só exista 1)
    if grupo:
        if isinstance(grupo, str):
            grupo = [grupo]
        dff = dff[dff["Grupo"].isin(grupo)]

    # Empresa - aceita lista de valores
    if empresa:
        if isinstance(empresa, str):
            empresa = [empresa]
        dff = dff[dff["Empresa Contratada"].isin(empresa)]

    # Status - aceita lista de valores
    if status_vig:
        if isinstance(status_vig, str):
            status_vig = [status_vig]
        dff = dff[dff["Status da Vigência"].isin(status_vig)]

    # remove linhas sem texto em Status da Vigência
    dff = dff[dff["Status da Vigência"].astype(str).str.strip() != ""]

    # Ordena por Término da Execução (mais recente em cima)
    dff["_termino_exec_dt"] = pd.to_datetime(
        dff["Término da Execução"], dayfirst=True, errors="coerce"
    )
    dff = dff.sort_values("_termino_exec_dt", ascending=False).drop(
        columns=["_termino_exec_dt"]
    )

    return dff


dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

input_style = {
    "width": "100%",
    "padding": "8px",
    "border": "1px solid #ccc",
    "borderRadius": "4px",
    "fontSize": "12px",
    "marginBottom": "6px",
}

# --------------------------------------------------
# Estilo dos botões (fundo azul, texto branco)
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


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_contratos",
            className="filtros-sticky",
            children=[
                # Linha 1: Contrato, Objeto, Setor
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato"),
                                dcc.Input(
                                    id="filtro_contrato",
                                    type="text",
                                    placeholder="Digite parte do número do contrato...",
                                    value="",
                                    style=input_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Objeto"),
                                dcc.Input(
                                    id="filtro_objeto",
                                    type="text",
                                    placeholder="Digite parte do objeto do contrato...",
                                    value="",
                                    style=input_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Setor"),
                                dcc.Dropdown(
                                    id="filtro_setor",
                                    options=[
                                        {"label": str(setor), "value": str(setor)}
                                        for setor in sorted(
                                            df_contratos_base["Setor"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(setor).strip() != ""
                                    ],
                                    value=[],
                                    placeholder="Selecione um ou mais setores...",
                                    clearable=True,
                                    multi=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                    ],
                ),
                # Linha 2: Empresa, Grupo, Status, botões
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-end",
                        "marginTop": "4px",
                    },
                    children=[
                        # Empresa
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Empresa Contratada"),
                                dcc.Dropdown(
                                    id="filtro_empresa",
                                    options=[
                                        {
                                            "label": str(empresa)[:80] + "..."
                                            if len(str(empresa)) > 80
                                            else str(empresa),
                                            "value": str(empresa),
                                        }
                                        for empresa in sorted(
                                            df_contratos_base["Empresa Contratada"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(empresa).strip() != ""
                                    ],
                                    value=[],
                                    placeholder="Selecione uma ou mais empresas...",
                                    clearable=True,
                                    multi=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Grupo
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Grupo"),
                                dcc.Dropdown(
                                    id="filtro_grupo",
                                    options=[
                                        {"label": str(grupo), "value": str(grupo)}
                                        for grupo in sorted(
                                            df_contratos_base["Grupo"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(grupo).strip() != ""
                                    ],
                                    value=[],
                                    placeholder="Selecione um ou mais grupos...",
                                    clearable=True,
                                    multi=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Status da Vigência
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Status da Vigência"),
                                dcc.Dropdown(
                                    id="filtro_status_vig",
                                    options=[
                                        {"label": "Vigente", "value": "Vigente"},
                                        {
                                            "label": "Próximo do Vencimento",
                                            "value": "Próximo do Vencimento",
                                        },
                                        {"label": "Vencido", "value": "Vencido"},
                                    ],
                                    value=[],
                                    placeholder="Selecione um ou mais status...",
                                    clearable=True,
                                    multi=True,
                                    searchable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Botões
                        html.Div(
                            style={
                                "display": "flex",
                                "gap": "10px",
                                "flexShrink": 0,
                            },
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_contratos",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_contratos",
                                    n_clicks=0,
                                    style=botao_style,
                                ),
                                dcc.Download(id="download_relatorio_contratos"),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        dash_table.DataTable(
            id="tabela_contratos",
            columns=[
                {
                    "name": "Contrato",
                    "id": "Contrato_Link",
                    "type": "text",
                    "presentation": "markdown",
                },
                {"name": "Setor", "id": "Setor"},
                {"name": "Grupo", "id": "Grupo"},
                {"name": "Objeto", "id": "Objeto"},
                {"name": "Empresa Contratada", "id": "Empresa Contratada"},
                {"name": "Início da Vigência", "id": "Início da Vigência"},
                {"name": "Término da Execução", "id": "Término da Execução"},
                {"name": "Término da Vigência", "id": "Término da Vigência"},
                {"name": "Status da Vigência", "id": "Status da Vigência"},
            ],
            data=[],
            markdown_options={"html": True},
            row_selectable=False,
            cell_selectable=False,
            style_table={
                "overflowX": "auto",
                "overflowY": "auto",
                "height": "calc(100vh - 200px)",
                "minHeight": "300px",
                "position": "relative",
            },
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
                "minWidth": "80px",
                "maxWidth": "260px",
                "whiteSpace": "normal",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "textAlign": "center",
                "position": "sticky",
                "top": 0,
                "zIndex": 5,
            },
            style_cell_conditional=[
                {"if": {"column_id": "Contrato_Link"}, "textAlign": "center"},
            ],
            style_data_conditional=[
                # Zebra: linhas pares/ímpares
                {"if": {"row_index": "odd"}, "backgroundColor": "#f0f0f0"},
                {"if": {"row_index": "even"}, "backgroundColor": "white"},
                # Status = Vencido
                {
                    "if": {"filter_query": '{Status da Vigência} = "Vencido"'},
                    "backgroundColor": "#ffcccc",
                    "color": "black",
                },
                # Status = Próximo do Vencimento
                {
                    "if": {
                        "filter_query": '{Status da Vigência} = "Próximo do Vencimento"'
                    },
                    "backgroundColor": "#ffffcc",
                    "color": "black",
                },
            ],
            css=[
                dict(selector="p", rule="margin: 0; text-align: center;"),
            ],
        ),
        dcc.Store(id="store_dados_contratos"),
    ]
)


# --------------------------------------------------
# Callback: filtros (tabela + store) - ATUALIZAÇÃO EM TEMPO REAL
# --------------------------------------------------
@callback(
    Output("tabela_contratos", "data"),
    Output("store_dados_contratos", "data"),
    Input("filtro_contrato", "value"),
    Input("filtro_objeto", "value"),
    Input("filtro_setor", "value"),
    Input("filtro_grupo", "value"),
    Input("filtro_empresa", "value"),
    Input("filtro_status_vig", "value"),
    prevent_initial_call=False,
)
def atualizar_tabela_contratos(
    contrato_texto,
    objeto_texto,
    setor,
    grupo,
    empresa,
    status_vig,
):
    if not verificar_pagina_contratos():
        raise PreventUpdate

    dff = filtrar_contratos(
        contrato_texto,
        objeto_texto,
        setor,
        grupo,
        empresa,
        status_vig,
    )

    dff = dff.copy()

    # Criar coluna com hyperlink HTML para a coluna Contrato
    if "Link Comprasnet" in dff.columns:
        dff["Contrato_Link"] = dff.apply(
            lambda row: (
                f'<a href="{row["Link Comprasnet"]}" target="_blank" '
                f'style="color: #0b2b57; text-decoration: none; font-weight: bold;">'
                f'{row["Contrato"]}</a>'
            )
            if pd.notna(row["Link Comprasnet"])
            and str(row["Link Comprasnet"]).strip()
            and str(row["Link Comprasnet"]).startswith(("http://", "https://"))
            else row["Contrato"],
            axis=1,
        )
    else:
        dff["Contrato_Link"] = dff["Contrato"]

    cols = [
        "Contrato_Link",
        "Setor",
        "Grupo",
        "Objeto",
        "Empresa Contratada",
        "Início da Vigência",
        "Término da Execução",
        "Término da Vigência",
        "Status da Vigência",
    ]
    cols = [c for c in cols if c in dff.columns]

    return dff[cols].to_dict("records"), dff.to_dict("records")


# --------------------------------------------------
# Callback: opções dos filtros (cascata) - ATUALIZAÇÃO EM TEMPO REAL
# --------------------------------------------------
@callback(
    Output("filtro_setor", "options"),
    Output("filtro_grupo", "options"),
    Output("filtro_empresa", "options"),
    Input("filtro_contrato", "value"),
    Input("filtro_objeto", "value"),
    Input("filtro_setor", "value"),
    Input("filtro_grupo", "value"),
    Input("filtro_empresa", "value"),
    Input("filtro_status_vig", "value"),
    prevent_initial_call=False,
)
def atualizar_opcoes_filtros(
    contrato_texto,
    objeto_texto,
    setor,
    grupo,
    empresa,
    status_vig,
):
    if not verificar_pagina_contratos():
        raise PreventUpdate

    dff = filtrar_contratos(
        contrato_texto,
        objeto_texto,
        setor,
        grupo,
        empresa,
        status_vig,
    )

    op_setor = [
        {"label": str(s), "value": str(s)}
        for s in sorted(dff["Setor"].dropna().unique())
        if str(s).strip()
    ]

    op_grupo = [
        {"label": str(g), "value": str(g)}
        for g in sorted(dff["Grupo"].dropna().unique())
        if str(g).strip()
    ]

    op_empresa = []
    for emp in sorted(dff["Empresa Contratada"].dropna().unique()):
        emp_str = str(emp)
        label = emp_str[:80] + "..." if len(emp_str) > 80 else emp_str
        op_empresa.append({"label": label, "value": emp_str})

    return op_setor, op_grupo, op_empresa


# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@callback(
    Output("filtro_contrato", "value", allow_duplicate=True),
    Output("filtro_objeto", "value", allow_duplicate=True),
    Output("filtro_setor", "value", allow_duplicate=True),
    Output("filtro_grupo", "value", allow_duplicate=True),
    Output("filtro_empresa", "value", allow_duplicate=True),
    Output("filtro_status_vig", "value", allow_duplicate=True),
    Input("btn_limpar_filtros_contratos", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_contratos(n):
    if not verificar_pagina_contratos():
        raise PreventUpdate

    return "", "", [], [], [], []


# --------------------------------------------------
# Estilos para PDF
# --------------------------------------------------
wrap_style_data = ParagraphStyle(
    name="wrap_contratos_data",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
    textColor=colors.black,
)

wrap_style_header = ParagraphStyle(
    name="wrap_contratos_header",
    fontSize=7,
    leading=8,
    alignment=TA_CENTER,
    textColor=colors.white,
)


def wrap_data(text):
    return Paragraph(str(text), wrap_style_data)


def wrap_header(text):
    return Paragraph(str(text), wrap_style_header)


# --------------------------------------------------
# Callback: gerar PDF de contratos
# --------------------------------------------------
@callback(
    Output("download_relatorio_contratos", "data"),
    Input("btn_download_relatorio_contratos", "n_clicks"),
    State("store_dados_contratos", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_contratos(n, dados_contratos):
    if not verificar_pagina_contratos():
        raise PreventUpdate

    if not n or not dados_contratos:
        return None

    df = pd.DataFrame(dados_contratos)

    buffer = BytesIO()
    pagesize = landscape(A4)

    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.3 * inch,
        leftMargin=0.3 * inch,
        topMargin=0.2 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")

    story.append(
        Table(
            [
                [
                    Paragraph(
                        data_hora,
                        ParagraphStyle(
                            "data_topo_contratos",
                            fontSize=9,
                            alignment=TA_RIGHT,
                            textColor="#333333",
                        ),
                    )
                ]
            ],
            colWidths=[pagesize[0] - 0.6 * inch],
        )
    )
    story.append(Spacer(1, 0.15 * inch))

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
        "RELATÓRIO DE CONTRATOS ATIVOS - UASG: 153030 - Campus Itajubá<br/>",
        ParagraphStyle(
            "titulo_contratos",
            alignment=TA_CENTER,
            fontSize=10,
            leading=14,
            textColor=colors.black,
        ),
    )

    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        "Contrato",
        "Setor",
        "Grupo",
        "Objeto",
        "Empresa Contratada",
        "Início da Vigência",
        "Término da Execução",
        "Término da Vigência",
        "Status da Vigência",
    ]

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    df_pdf = df.copy()

    header = [wrap_header(c) for c in cols]
    table_data = [header]

    status_values = df["Status da Vigência"].fillna("").tolist()

    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap_data(row[c]) for c in cols])

    col_widths = [
        0.8 * inch,  # Contrato
        0.9 * inch,  # Setor
        0.9 * inch,  # Grupo
        2.2 * inch,  # Objeto
        1.8 * inch,  # Empresa Contratada
        1.0 * inch,  # Início da Vigência
        1.1 * inch,  # Término da Execução
        1.1 * inch,  # Término da Vigência
        1.1 * inch,  # Status da Vigência
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
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

    for i, status in enumerate(status_values, 1):
        status_str = str(status).strip().lower()
        if "vencido" in status_str:
            table_styles.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#ffcccc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#cc0000"))
            )
        elif "próximo do vencimento" in status_str or "proximo do vencimento" in status_str:
            table_styles.append(
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#ffffcc"))
            )
            table_styles.append(
                ("TEXTCOLOR", (0, i), (-1, i), colors.HexColor("#cc8800"))
            )

    tbl.setStyle(TableStyle(table_styles))
    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(
        buffer.getvalue(),
        f"relatorio_contratos_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf",
    )
