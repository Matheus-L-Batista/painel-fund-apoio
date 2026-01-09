import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd

from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/portarias_agentedecompras",
    name="portarias_agentedecompras",
    title="portarias_agentedecompras",
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
NOME_COL_LINK_ORIGINAL = (
    "Link do documento\nAgentes de Compras e\nContratos tipo empenho"
)

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

    cols_serv = [str(i) for i in range(1, 16) if str(i) in df.columns]

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

    tipos_validos = ["AGENTES DE COMPRAS", "CONTRATOS TIPO EMPENHO"]
    df = df[df["TIPO"].isin(tipos_validos)]

    if NOME_COL_LINK_ORIGINAL not in df.columns:
        df[NOME_COL_LINK_ORIGINAL] = ""

    df = df[
        df[NOME_COL_LINK_ORIGINAL]
        .astype(str)
        .str.strip()
        .str.startswith("http")
    ]

    # Ordena pela Data (mais recente em cima)
    df["Data_dt"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df = df.sort_values("Data_dt", ascending=False).drop(columns=["Data_dt"])

    if cols_serv:
        todos_serv = pd.Series(df[cols_serv].values.ravel("K"), dtype="object")
        servidores_unicos = sorted(
            [s for s in todos_serv.unique() if isinstance(s, str) and s.strip() != ""]
        )
    else:
        servidores_unicos = []

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
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_port",
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
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("N°/Ano da Portaria"),
                                dcc.Input(
                                    id="filtro_numero_ano",
                                    type="text",
                                    placeholder="Digite parte do número/ano",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Setor de Origem"),
                                dcc.Dropdown(
                                    id="filtro_setor_dropdown",
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
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores (digitação)"),
                                dcc.Input(
                                    id="filtro_servidor_texto",
                                    type="text",
                                    placeholder="Digite parte do nome",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores"),
                                dcc.Dropdown(
                                    id="filtro_servidor_dropdown",
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
                        html.Div(
                            style={"minWidth": "220px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Tipo"),
                                dcc.Dropdown(
                                    id="filtro_tipo",
                                    options=[
                                        {"label": "Todos", "value": "TODOS"},
                                        {
                                            "label": "AGENTES DE COMPRAS",
                                            "value": "AGENTES DE COMPRAS",
                                        },
                                        {
                                            "label": "CONTRATOS TIPO EMPENHO",
                                            "value": "CONTRATOS TIPO EMPENHO",
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
                            id="btn_limpar_filtros_port",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_port",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_port"),
                    ],
                ),
            ],
        ),
        # Texto
        html.Div(
            style={
                "marginTop": "15px",
                "marginBottom": "15px",
                "display": "flex",
                "justifyContent": "center",
                "alignItems": "baseline",
                "gap": "5px",
                "color": "#b30000",
                "fontSize": "14px",
                "whiteSpace": "nowrap",
            },
            children=[
                html.Span(
                    "Portarias válidas para vinculação dos servidores às notas de empenho",
                    style={"fontWeight": "bold"},
                ),
                html.Span(
                    "(fase que antecede o lançamento dos ",
                    style={"fontSize": "15px"},
                ),
                html.Span(
                    "Instrumentos de Cobrança",
                    style={
                        "fontSize": "15px",
                        "textDecoration": "underline",
                    },
                ),
                html.Span(
                    " no sistema contratos.gov.br)",
                    style={"fontSize": "15px"},
                ),
            ],
        ),
        dash_table.DataTable(
            id="tabela_portarias",
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
        dcc.Store(id="store_dados_port"),
    ]
)

# --------------------------------------------------
# Callback: aplicar filtros + link clicável (máscara única)
# --------------------------------------------------
@dash.callback(
    Output("tabela_portarias", "data"),
    Output("store_dados_port", "data"),
    Input("filtro_numero_ano", "value"),
    Input("filtro_setor_dropdown", "value"),
    Input("filtro_servidor_texto", "value"),
    Input("filtro_servidor_dropdown", "value"),
    Input("filtro_tipo", "value"),
)
def atualizar_tabela_portarias(
    numero_ano_texto,
    setor_drop,
    servidor_texto,
    servidor_drop,
    tipo_sel,
):
    """
    Aplica todos os filtros em um único dataframe base (df_portarias_base),
    usando máscara booleana combinada. A ordem dos filtros não importa.
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

    # Setor
    if setor_drop:
        mask &= dff["Setor de Origem"] == setor_drop

    # Servidor texto
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
    Output("filtro_setor_dropdown", "options"),
    Output("filtro_servidor_dropdown", "options"),
    Output("filtro_tipo", "options"),
    Input("filtro_numero_ano", "value"),
    Input("filtro_setor_dropdown", "value"),
    Input("filtro_servidor_texto", "value"),
    Input("filtro_servidor_dropdown", "value"),
    Input("filtro_tipo", "value"),
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
    em cascata, usando um único filtro global.
    """
    dff = df_portarias_base.copy()

    mask = pd.Series(True, index=dff.index)

    if tipo_sel and tipo_sel != "TODOS":
        mask &= dff["TIPO"] == tipo_sel

    if numero_ano_texto and str(numero_ano_texto).strip():
        termo = str(numero_ano_texto).strip().lower()
        mask &= (
            dff["N°/ANO da Portaria"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    if setor_drop:
        mask &= dff["Setor de Origem"] == setor_drop

    if servidor_texto and str(servidor_texto).strip():
        termo = str(servidor_texto).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    if servidor_drop:
        termo = str(servidor_drop).strip().lower()
        mask &= (
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        )

    dff = dff[mask].copy()

    # Opções de Setor
    op_setor = [
        {"label": s, "value": s}
        for s in sorted(dff["Setor de Origem"].dropna().unique())
        if str(s).strip() != ""
    ]

    # Opções de servidores (explodindo colunas 1..15 se existirem)
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

    # Opções de Tipo, mantendo "TODOS"
    tipos_presentes = sorted(
        [t for t in dff["TIPO"].dropna().unique() if str(t).strip() != ""]
    )
    op_tipo = [{"label": "Todos", "value": "TODOS"}]

    for t in tipos_presentes:
        op_tipo.append({"label": t, "value": t})

    return op_setor, op_servidor, op_tipo

# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_numero_ano", "value"),
    Output("filtro_setor_dropdown", "value"),
    Output("filtro_servidor_texto", "value"),
    Output("filtro_servidor_dropdown", "value"),
    Output("filtro_tipo", "value"),
    Input("btn_limpar_filtros_port", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_port(n):
    return None, None, None, None, "TODOS"

# --------------------------------------------------
# Callback: gerar PDF
# --------------------------------------------------
from datetime import datetime
from pytz import timezone
from reportlab.platypus import Image
import os

wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=8,
    leading=10,
    spaceAfter=4,
    alignment=TA_CENTER,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

@dash.callback(
    Output("download_relatorio_port", "data"),
    Input("btn_download_relatorio_port", "n_clicks"),
    State("store_dados_port", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_port(n, dados_port):
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

    # Data e hora
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora_brasilia = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")

    data_top_table = Table(
        [
            [
                Paragraph(
                    data_hora_brasilia,
                    ParagraphStyle(
                        "data_topo",
                        fontSize=9,
                        alignment=TA_RIGHT,
                        textColor="#333333",
                    ),
                )
            ]
        ],
        colWidths=[pagesize[0] - 0.6 * inch],
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

    # Logos
    logos_path = []
    if os.path.exists(os.path.join("assets", "brasaobrasil.png")):
        logos_path.append(os.path.join("assets", "brasaobrasil.png"))
    if os.path.exists(os.path.join("assets", "simbolo_RGB.png")):
        logos_path.append(os.path.join("assets", "simbolo_RGB.png"))

    if logos_path:
        logos = []
        for logo_file in logos_path:
            if os.path.exists(logo_file):
                logo = Image(logo_file, width=1.2 * inch, height=1.2 * inch)
                logos.append(logo)

        if logos:
            if len(logos) == 2:
                logo_table = Table(
                    [[logos[0], logos[1]]],
                    colWidths=[
                        pagesize[0] / 2 - 0.3 * inch,
                        pagesize[0] / 2 - 0.3 * inch,
                    ],
                )
            else:
                logo_table = Table(
                    [[logos[0]]],
                    colWidths=[pagesize[0] - 0.6 * inch],
                )

            logo_table.setStyle(
                TableStyle(
                    [
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ]
                )
            )

            story.append(logo_table)
            story.append(Spacer(1, 0.15 * inch))

    # Título
    titulo_texto = (
        "RELATÓRIO DE PORTARIAS<br/>"
        "AGENTES DE COMPRAS E CONTRATOS TIPO EMPENHO<br/>"
        "Campus Itajubá"
    )

    titulo_paragraph = Paragraph(
        titulo_texto,
        ParagraphStyle(
            "titulo_portarias",
            fontSize=10,
            alignment=TA_CENTER,
            textColor="#0b2b57",
            spaceAfter=4,
            leading=14,
        ),
    )

    titulo_table = Table(
        [[titulo_paragraph]],
        colWidths=[pagesize[0] - 0.6 * inch],
    )

    titulo_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )

    story.append(titulo_table)
    story.append(Spacer(1, 0.15 * inch))

    # Quantidade de registros
    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.1 * inch))

    # Colunas
    cols = [
        "Data",
        "N°/ANO da Portaria",
        "Setor de Origem",
        "TIPO",
        "Servidores",
    ]

    df_pdf = df.copy()

    header = cols
    table_data = [header]
    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap(row[c]) for c in cols])

    page_width = pagesize[0] - 0.6 * inch

    col_widths = [
        0.8 * inch,        # Data
        0.9 * inch,        # N°/ANO da Portaria
        1.0 * inch,        # Setor de Origem
        1.2 * inch,        # TIPO
        page_width - (0.8 + 0.9 + 1.0 + 1.2) * inch,  # Servidores
    ]

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)

    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("WORDWRAP", (0, 0), (-1, -1), True),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]

    for row_idx in range(len(table_data)):
        table_styles.append(
            ("ALIGN", (4, row_idx), (4, row_idx), "LEFT")
        )

    tbl.setStyle(TableStyle(table_styles))

    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    from dash import dcc

    return dcc.send_bytes(buffer.getvalue(), "portarias_.pdf")
