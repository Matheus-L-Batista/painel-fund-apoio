import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd

from reportlab.lib.enums import TA_CENTER, TA_RIGHT
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
from reportlab.lib import colors

from datetime import datetime
from pytz import timezone
import os


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/portarias_planejamento",
    name="Portarias – Planejamento",
    title="Portarias – Planejamento",
)


# --------------------------------------------------
# URL da planilha de Portarias
# --------------------------------------------------
URL_PORTARIAS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Check%20List"
)

# nome EXATO da coluna de link no CSV
NOME_COL_LINK_ORIGINAL = "Link do documento\nEquipe de Planejamento"


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados_portarias():
    df = pd.read_csv(URL_PORTARIAS, header=1)
    df.columns = [c.strip() for c in df.columns]

    df = df.rename(
        columns={
            "Unnamed: 5": "Data",
            "N° / ANO": "N°/ANO da Portaria",
            "ORIGEM": "Setor de Origem",
        }
    )

    # Colunas de servidores (1..15) se existirem
    cols_serv = [str(i) for i in range(1, 16) if str(i) in df.columns]

    # Concatena servidores em uma única coluna
    if cols_serv:
        df["Servidores"] = (
            df[cols_serv]
            .astype(str)
            .replace({"nan": ""})
            .agg("; ".join, axis=1)
            .str.replace(r"(; )+$", "", regex=True)
        )
    else:
        df["Servidores"] = ""

    if "TIPO" not in df.columns:
        df["TIPO"] = ""

    # Tipos específicos desta página
    tipos_validos = [
        "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO",
        "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO - TI",
    ]
    df = df[df["TIPO"].isin(tipos_validos)]

    if NOME_COL_LINK_ORIGINAL not in df.columns:
        df[NOME_COL_LINK_ORIGINAL] = ""

    # mantém apenas linhas com link válido
    df = df[
        df[NOME_COL_LINK_ORIGINAL]
        .astype(str)
        .str.strip()
        .str.startswith("http")
    ]

    # Ordena pela Data (mais recente em cima)
    df["Data_dt"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df = df.sort_values("Data_dt", ascending=False).drop(columns=["Data_dt"])

    # lista de servidores únicos após o filtro
    if cols_serv:
        todos_serv = pd.Series(df[cols_serv].values.ravel("K"), dtype="object")
        servidores_unicos = sorted(
            [s for s in todos_serv.unique() if isinstance(s, str) and s.strip() != ""]
        )
    else:
        servidores_unicos = []

    # armazena lista de servidores únicos como atributo
    df._lista_servidores_unicos = servidores_unicos

    return df


df_portarias_base = carregar_dados_portarias()
SERVIDORES_UNICOS = getattr(df_portarias_base, "_lista_servidores_unicos", [])

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

# --------------------------------------------------
# Estilo unificado dos botões (fundo azul, texto branco)
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
            id="barra_filtros_port_planej",
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
                        # N°/ANO da Portaria (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("N°/Ano da Portaria"),
                                dcc.Input(
                                    id="filtro_numero_ano_planej",
                                    type="text",
                                    placeholder="Digite parte do número/ano",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Setor de Origem
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Setor de Origem"),
                                dcc.Dropdown(
                                    id="filtro_setor_dropdown_planej",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in sorted(
                                            df_portarias_base["Setor de Origem"]
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
                        # Servidor (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores (digitação)"),
                                dcc.Input(
                                    id="filtro_servidor_texto_planej",
                                    type="text",
                                    placeholder="Digite parte do nome",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Servidor (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores"),
                                dcc.Dropdown(
                                    id="filtro_servidor_dropdown_planej",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in SERVIDORES_UNICOS
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Tipo
                        html.Div(
                            style={"minWidth": "220px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Tipo"),
                                dcc.Dropdown(
                                    id="filtro_tipo_planej",
                                    options=[
                                        {"label": "Todos", "value": "TODOS"},
                                        {
                                            "label": "PLANEJAMENTO DA CONTRATAÇÃO",
                                            "value": "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO",
                                        },
                                        {
                                            "label": "PLANEJAMENTO DA CONTRATAÇÃO - TI",
                                            "value": "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO - TI",
                                        },
                                    ],
                                    value="TODOS",
                                    clearable=False,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    style={"marginTop": "4px"},
                    children=[
                        html.Button(
                            "Limpar filtros",
                            id="btn_limpar_filtros_port_planej",
                            n_clicks=0,
                            style=botao_style,
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_port_planej",
                            n_clicks=0,
                            style=botao_style,
                        ),
                        dcc.Download(id="download_relatorio_port_planej"),
                    ],
                ),
            ],
        ),
        # Texto de orientação
        html.Div(
            style={
                "marginTop": "15px",
                "marginBottom": "15px",
                "textAlign": "center",
                "color": "#b30000",
                "fontSize": "14px",
                "whiteSpace": "normal",
            },
            children=[
                html.Span(
                    "Portarias válidas para composição das Equipes de Planejamento da Contratação (inclusive TI)",
                    style={"fontWeight": "bold"},
                ),
            ],
        ),
        # Tabela
        dash_table.DataTable(
            id="tabela_portarias_planej",
            columns=[
                {"name": "Data", "id": "Data"},
                {"name": "N°/ANO da Portaria", "id": "N°/ANO da Portaria"},
                {"name": "Setor de Origem", "id": "Setor de Origem"},
                {"name": "Servidores", "id": "Servidores"},
                {"name": "TIPO", "id": "TIPO"},
                {
                    "name": "Link",
                    "id": "Link_markdown",
                    "presentation": "markdown",
                },
            ],
            data=[],
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
            style_data={
                "color": "black",
                "backgroundColor": "white",
            },
            style_data_conditional=[
                {
                    "if": {"row_index": "odd"},
                    "backgroundColor": "rgb(240, 240, 240)",
                },
            ],
            style_cell_conditional=[
                {
                    "if": {"column_id": "Link_markdown"},
                    "textAlign": "center",
                },
            ],
            css=[
                dict(
                    selector="p",
                    rule="margin: 0; text-align: center;",
                ),
            ],
        ),
        dcc.Store(id="store_dados_port_planej"),
    ]
)


# --------------------------------------------------
# Callback: aplicar filtros + link clicável (máscara única)
# --------------------------------------------------
@dash.callback(
    Output("tabela_portarias_planej", "data"),
    Output("store_dados_port_planej", "data"),
    Input("filtro_numero_ano_planej", "value"),
    Input("filtro_setor_dropdown_planej", "value"),
    Input("filtro_servidor_texto_planej", "value"),
    Input("filtro_servidor_dropdown_planej", "value"),
    Input("filtro_tipo_planej", "value"),
)
def atualizar_tabela_portarias_planej(
    numero_ano_texto,
    setor_drop,
    servidor_texto,
    servidor_drop,
    tipo_sel,
):
    """
    Aplica todos os filtros em um único dataframe base (df_portarias_base),
    usando uma máscara booleana combinada. A ordem dos filtros não importa.
    """
    dff = df_portarias_base.copy()

    mask = pd.Series(True, index=dff.index)

    # Tipo
    if tipo_sel and tipo_sel != "TODOS":
        mask &= dff["TIPO"] == tipo_sel

    # Nº/ANO da Portaria (contains, case-insensitive)
    if numero_ano_texto and str(numero_ano_texto).strip():
        termo = str(numero_ano_texto).strip().lower()
        mask &= (
            dff["N°/ANO da Portaria"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    # Setor de Origem (igualdade)
    if setor_drop:
        mask &= dff["Setor de Origem"] == setor_drop

    # Servidores por texto (contains em string concatenada)
    if servidor_texto and str(servidor_texto).strip():
        termo = str(servidor_texto).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    # Servidor pelo dropdown (contains)
    if servidor_drop:
        termo = str(servidor_drop).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    dff = dff[mask].copy()

    # Garante apenas linhas com link válido
    dff = dff[
        dff[NOME_COL_LINK_ORIGINAL]
        .astype(str)
        .str.strip()
        .str.startswith("http")
    ]

    dff_display = dff.copy()

    def formatar_link(url):
        if isinstance(url, str) and url.strip():
            return f"[Link]({url.strip()})"
        return ""

    dff_display["Link_markdown"] = dff_display[NOME_COL_LINK_ORIGINAL].apply(
        formatar_link
    )

    cols_tabela = [
        "Data",
        "N°/ANO da Portaria",
        "Setor de Origem",
        "Servidores",
        "TIPO",
        "Link_markdown",
    ]

    return dff_display[cols_tabela].to_dict("records"), dff.to_dict("records")


# --------------------------------------------------
# Callback: filtros em cascata (ordem-invariante)
# --------------------------------------------------
@dash.callback(
    Output("filtro_setor_dropdown_planej", "options"),
    Output("filtro_servidor_dropdown_planej", "options"),
    Output("filtro_tipo_planej", "options"),
    Input("filtro_numero_ano_planej", "value"),
    Input("filtro_setor_dropdown_planej", "value"),
    Input("filtro_servidor_texto_planej", "value"),
    Input("filtro_servidor_dropdown_planej", "value"),
    Input("filtro_tipo_planej", "value"),
)
def atualizar_opcoes_filtros_portarias(
    numero_ano_texto,
    setor_drop,
    servidor_texto,
    servidor_drop,
    tipo_sel,
):
    """
    Atualiza as opções de Setor, Servidores (dropdown) e Tipo
    em cascata, usando um único filtro global. A ordem dos filtros
    não importa.
    """
    dff = df_portarias_base.copy()

    mask = pd.Series(True, index=dff.index)

    # Tipo
    if tipo_sel and tipo_sel != "TODOS":
        mask &= dff["TIPO"] == tipo_sel

    # Nº/ANO da Portaria
    if numero_ano_texto and str(numero_ano_texto).strip():
        termo = str(numero_ano_texto).strip().lower()
        mask &= (
            dff["N°/ANO da Portaria"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    # Setor de Origem
    if setor_drop:
        mask &= dff["Setor de Origem"] == setor_drop

    # Servidor por texto
    if servidor_texto and str(servidor_texto).strip():
        termo = str(servidor_texto).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    # Servidor dropdown
    if servidor_drop:
        termo = str(servidor_drop).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    dff = dff[mask].copy()

    # Opções de Setor de Origem
    op_setor = [
        {"label": s, "value": s}
        for s in sorted(dff["Setor de Origem"].dropna().unique())
        if str(s).strip() != ""
    ]

    # Opções de Servidores (lista única a partir do subconjunto filtrado)
    cols_serv = [str(i) for i in range(1, 16) if str(i) in dff.columns]
    if cols_serv:
        todos_serv = pd.Series(dff[cols_serv].values.ravel("K"), dtype="object")
        servidores_unicos_filtrados = sorted(
            [
                s
                for s in todos_serv.unique()
                if isinstance(s, str) and s.strip() != ""
            ]
        )
    else:
        servidores_unicos_filtrados = []

    op_servidor = [
        {"label": s, "value": s}
        for s in servidores_unicos_filtrados
    ]

    # Opções de Tipo, mantendo a opção "TODOS" sempre
    tipos_presentes = sorted(
        [t for t in dff["TIPO"].dropna().unique() if str(t).strip() != ""]
    )
    op_tipo = [{"label": "Todos", "value": "TODOS"}]

    mapa_rotulos_tipo = {
        "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO": "PLANEJAMENTO DA CONTRATAÇÃO",
        "PORTARIA DE PLANEJAMENTO DA CONTRATAÇÃO - TI": "PLANEJAMENTO DA CONTRATAÇÃO - TI",
    }

    for t in tipos_presentes:
        op_tipo.append(
            {
                "label": mapa_rotulos_tipo.get(t, t),
                "value": t,
            }
        )

    return op_setor, op_servidor, op_tipo


# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_numero_ano_planej", "value"),
    Output("filtro_setor_dropdown_planej", "value"),
    Output("filtro_servidor_texto_planej", "value"),
    Output("filtro_servidor_dropdown_planej", "value"),
    Output("filtro_tipo_planej", "value"),
    Input("btn_limpar_filtros_port_planej", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_port_planej(n):
    return None, None, None, None, "TODOS"


# --------------------------------------------------
# Estilos para o PDF (planejamento)
# --------------------------------------------------
wrap_style_data = ParagraphStyle(
    name="wrap_planej_data",
    fontSize=7,
    leading=8,
    spaceAfter=2,
    wordWrap="CJK",
    alignment=TA_CENTER,
)

wrap_style_header = ParagraphStyle(
    name="wrap_planej_header",
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
# Callback: gerar PDF (layout igual ao de contratos)
# --------------------------------------------------
@dash.callback(
    Output("download_relatorio_port_planej", "data"),
    Input("btn_download_relatorio_port_planej", "n_clicks"),
    State("store_dados_port_planej", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_port_planej(n, dados_port):
    if not n or not dados_port:
        return None

    df = pd.DataFrame(dados_port)

    buffer = BytesIO()
    pagesize = landscape(A4)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.3 * inch,
        leftMargin=0.3 * inch,
        topMargin=1.3 * inch,
        bottomMargin=0.4 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    # --------------------------------------------------
    # Data / Hora (topo direito)
    # --------------------------------------------------
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")

    story.append(
        Table(
            [[Paragraph(
                data_hora,
                ParagraphStyle(
                    "data_topo",
                    fontSize=9,
                    alignment=TA_RIGHT,
                    textColor="#333333",
                ),
            )]],
            colWidths=[pagesize[0] - 0.6 * inch],
        )
    )
    story.append(Spacer(1, 0.15 * inch))

    # --------------------------------------------------
    # Cabeçalho: Logo esq | Instituição | Logo dir
    # --------------------------------------------------
    logo_esq = (
        Image("assets/brasaobrasil.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/brasaobrasil.png") else ""
    )

    logo_dir = (
        Image("assets/simbolo_RGB.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/simbolo_RGB.png") else ""
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
        colWidths=[
            1.4 * inch,
            4.2 * inch,
            1.4 * inch,
        ],
    )

    cabecalho.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ])
    )

    story.append(cabecalho)
    story.append(Spacer(1, 0.25 * inch))

    # --------------------------------------------------
    # Título
    # --------------------------------------------------
    titulo = Paragraph(
        "Portarias vigentes – Equipes de Planejamento da Contratação (inclusive TI)<br/>",
        ParagraphStyle(
            "titulo",
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

    # --------------------------------------------------
    # Preparação da tabela de dados
    # --------------------------------------------------
    cols = [
        "Data",
        "N°/ANO da Portaria",
        "Setor de Origem",
        "Servidores",
        "TIPO",
    ]

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    df_pdf = df.copy()

    header = [wrap_header(c) for c in cols]
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap_data(row[c]) for c in cols])

    # --------------------------------------------------
    # Larguras das colunas (ajustadas para landscape)
    # --------------------------------------------------
    page_width = pagesize[0] - 0.6 * inch
    col_widths = [
        0.9 * inch,               # Data
        1.2 * inch,               # N°/ANO da Portaria
        1.6 * inch,               # Setor de Origem
        3.0 * inch,               # Servidores
        page_width - (0.9 + 1.2 + 1.6 + 3.0) * inch,  # TIPO (restante)
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    table_styles = [
        # Header
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        # Padding
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        # Zebra
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
    ]

    tbl.setStyle(TableStyle(table_styles))
    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    return dcc.send_bytes(
        buffer.getvalue(),
        "relatorio_portarias_planejamento.pdf",
    )
