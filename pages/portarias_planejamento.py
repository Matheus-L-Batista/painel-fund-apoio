import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd

from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from io import BytesIO
from reportlab.lib.pagesizes import landscape, A4
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
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_port_planej",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_port_planej"),
                    ],
                ),
            ],
        ),

        # Texto de orientação (pode ser ajustado conforme desejar)
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
# Callback: aplicar filtros + link clicável
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
    numero_ano_texto, setor_drop, servidor_texto, servidor_drop, tipo_sel
):
    dff = df_portarias_base.copy()

    if tipo_sel and tipo_sel != "TODOS":
        dff = dff[dff["TIPO"] == tipo_sel]

    if numero_ano_texto and str(numero_ano_texto).strip():
        termo = str(numero_ano_texto).strip().lower()
        dff = dff[
            dff["N°/ANO da Portaria"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    if setor_drop:
        dff = dff[dff["Setor de Origem"] == setor_drop]

    if servidor_texto and str(servidor_texto).strip():
        termo = str(servidor_texto).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    if servidor_drop:
        termo = str(servidor_drop).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

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
# Callback: gerar PDF
# --------------------------------------------------
from datetime import datetime
from pytz import timezone
from reportlab.platypus import Image
import os

wrap_style = ParagraphStyle(
    name="wrap_planej",
    fontSize=8,
    leading=10,
    spaceAfter=4,
    alignment=TA_CENTER,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

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

    # Data e hora (topo direito)
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
        "EQUIPE DE PLANEJAMENTO DA CONTRATAÇÃO<br/>"
        "Campus Itajubá"
    )

    titulo_paragraph = Paragraph(
        titulo_texto,
        ParagraphStyle(
            "titulo_portarias_planej",
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

    # Colunas da tabela (sem Link)
    cols = [
        "Data",
        "N°/ANO da Portaria",
        "Setor de Origem",
        "TIPO",
        "Servidores",
    ]

    df_pdf = df.copy()

    # Monta tabela
    header = cols
    table_data = [header]
    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap(row[c]) for c in cols])

    page_width = pagesize[0] - 0.6 * inch
    
    # Larguras proporcionais das colunas
    col_widths = [
        0.8 * inch,        # Data
        0.9 * inch,        # N°/ANO da Portaria
        1.0 * inch,        # Setor de Origem
        1.2 * inch,        # TIPO
        page_width - (0.8 + 0.9 + 1.0 + 1.2) * inch,  # Servidores (resto)
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

    # Servidores alinhada à esquerda (coluna 4)
    for row_idx in range(len(table_data)):
        table_styles.append(
            ("ALIGN", (4, row_idx), (4, row_idx), "LEFT")
        )

    tbl.setStyle(TableStyle(table_styles))

    story.append(tbl)

    doc.build(story)
    buffer.seek(0)

    from dash import dcc

    return dcc.send_bytes(buffer.getvalue(), "portarias_planejamento.pdf")
