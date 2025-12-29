import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from datetime import datetime

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
    path="/fiscais",
    name="Fiscais",
    title="Fiscais",
)

# --------------------------------------------------
# URL da planilha de Fiscais
# --------------------------------------------------
URL_FISCAIS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Fiscais"
)

# nomes originais no CSV (linha 4 é o cabeçalho)
COL_SETOR = "Setor"
COL_CONTRATO = "CONTRATO"
COL_OBJETO = "OBJETO"
COL_CONTRATADA = "CONTRATADA"
COL_FINAL_VIG = "Unnamed: 16"  # Final da Vigência
COL_LINK_COMPRASNET = "COMPRASNET Contratos"

# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------
def carregar_dados_fiscais():
    df = pd.read_csv(URL_FISCAIS, skiprows=3, header=0)
    df.columns = [c.strip() for c in df.columns]

    # Colunas de servidores
    col_servidores_raw = [c for c in df.columns if c.startswith("SERVIDOR")]

    cols_keep = [
        COL_SETOR,
        COL_CONTRATO,
        COL_OBJETO,
        COL_CONTRATADA,
        COL_FINAL_VIG,
        COL_LINK_COMPRASNET,
    ] + col_servidores_raw

    df = df[cols_keep]

    df = df.rename(
        columns={
            COL_SETOR: "Setor",
            COL_CONTRATO: "Contrato",
            COL_OBJETO: "Objeto",
            COL_CONTRATADA: "Contratada",
            COL_FINAL_VIG: "Final da Vigência",
            COL_LINK_COMPRASNET: "Link Comprasnet",
        }
    )

    # Lista de servidores únicos (sem vazios) para o dropdown
    if col_servidores_raw:
        todos_serv = pd.Series(df[col_servidores_raw].values.ravel("K"), dtype="object")
        servidores_unicos = sorted(
            s.strip()
            for s in todos_serv.unique()
            if isinstance(s, str) and s.strip() != ""
        )
    else:
        servidores_unicos = []


    # Coluna agregada Servidores para exibição na tabela (sem nomes vazios)
    if col_servidores_raw:
        def junta_servidores(row):
            nomes = []
            for c in col_servidores_raw:
                v = row.get(c)
                if isinstance(v, str):
                    v = v.strip()
                else:
                    v = ""
                if v:  # só entra se não for vazio
                    nomes.append(v)
            return "; ".join(nomes)

        df["Servidores"] = df.apply(junta_servidores, axis=1)
    else:
        df["Servidores"] = ""

    # Conversão e status pela Final da Vigência
    df["Final da Vigência"] = pd.to_datetime(
        df["Final da Vigência"], dayfirst=True, errors="coerce"
    )

    hoje = datetime.now().date()

    def calcular_status(data_final):
        if pd.isna(data_final):
            return ""
        dias = (data_final.date() - hoje).days
        if dias > 10:
            return "Vigente"
        if dias < 0:
            return "Vencido"
        return "Próximo do Vencimento"

    df["Status"] = df["Final da Vigência"].apply(calcular_status)

    df["Final da Vigência"] = df["Final da Vigência"].dt.strftime("%d/%m/%Y").fillna("")

    # guarda a lista de servidores únicos
    df._lista_servidores_unicos = servidores_unicos

    return df


df_fiscais_base = carregar_dados_fiscais()
SERVIDORES_UNICOS_FIS = getattr(df_fiscais_base, "_lista_servidores_unicos", [])

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
            id="barra_filtros_fiscais",
            className="filtros-sticky",
            children=[
                # Linha 1: Servidores e Contrato (texto + dropdown)
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        # Servidores (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores (digitação)"),
                                dcc.Input(
                                    id="filtro_servidores_texto_fis",
                                    type="text",
                                    placeholder="Digite parte do nome do servidor",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Servidores (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Servidores"),
                                dcc.Dropdown(
                                    id="filtro_servidores_dropdown_fis",
                                    options=[
                                        {"label": s, "value": s}
                                        for s in SERVIDORES_UNICOS_FIS
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Contrato (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato (digitação)"),
                                dcc.Input(
                                    id="filtro_contrato_texto_fis",
                                    type="text",
                                    placeholder="Digite parte do contrato",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Contrato (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato"),
                                dcc.Dropdown(
                                    id="filtro_contrato_dropdown_fis",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in sorted(
                                            df_fiscais_base["Contrato"]
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
                    ],
                ),
                # Linha 2: Contratada (texto+dropdown), Status + botões
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "10px",
                        "alignItems": "flex-end",
                        "marginTop": "4px",
                    },
                    children=[
                        # Contratada (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contratada (digitação)"),
                                dcc.Input(
                                    id="filtro_contratada_texto_fis",
                                    type="text",
                                    placeholder="Digite parte da contratada",
                                    style={"width": "100%", "marginBottom": "6px"},
                                ),
                            ],
                        ),
                        # Contratada (dropdown)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contratada"),
                                dcc.Dropdown(
                                    id="filtro_contratada_dropdown_fis",
                                    options=[
                                        {"label": e, "value": e}
                                        for e in sorted(
                                            df_fiscais_base["Contratada"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(e).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        # Status
                        html.Div(
                            style={"minWidth": "200px", "flex": "0 0 220px"},
                            children=[
                                html.Label("Status"),
                                dcc.Dropdown(
                                    id="filtro_status_fis",
                                    options=[
                                        {"label": "Vigente", "value": "Vigente"},
                                        {
                                            "label": "Próximo do Vencimento",
                                            "value": "Próximo do Vencimento",
                                        },
                                        {"label": "Vencido", "value": "Vencido"},
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
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
                                    id="btn_limpar_filtros_fis",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_fis",
                                    n_clicks=0,
                                    className="filtros-button",
                                ),
                                dcc.Download(id="download_relatorio_fis"),
                            ],
                        ),
                    ],
                ),
            ],
        ),

        html.H4("Fiscais de Contratos"),
        dash_table.DataTable(
            id="tabela_fiscais",
            columns=[
                {
                    "name": "Contrato",
                    "id": "Contrato_markdown",
                    "presentation": "markdown",
                },
                {"name": "Setor", "id": "Setor"},
                {"name": "Objeto", "id": "Objeto"},
                {"name": "Contratada", "id": "Contratada"},
                {"name": "Final da Vigência", "id": "Final da Vigência"},
                {"name": "Status", "id": "Status"},
                {"name": "Servidores", "id": "Servidores"},
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
        ),
        dcc.Store(id="store_dados_fis"),
    ]
)

# --------------------------------------------------
# Callback: filtros
# --------------------------------------------------
@dash.callback(
    Output("tabela_fiscais", "data"),
    Output("store_dados_fis", "data"),
    Input("filtro_servidores_texto_fis", "value"),
    Input("filtro_servidores_dropdown_fis", "value"),
    Input("filtro_contrato_texto_fis", "value"),
    Input("filtro_contrato_dropdown_fis", "value"),
    Input("filtro_contratada_texto_fis", "value"),
    Input("filtro_contratada_dropdown_fis", "value"),
    Input("filtro_status_fis", "value"),
)
def atualizar_tabela_fiscais(
    servidores_texto,
    servidores_drop,
    contrato_texto,
    contrato_drop,
    contratada_texto,
    contratada_drop,
    status,
):
    dff = df_fiscais_base.copy()

    # Servidores (digitação)
    if servidores_texto and str(servidores_texto).strip():
        termo = str(servidores_texto).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Servidores (dropdown) - contém o nome na string agregada
    if servidores_drop:
        termo = str(servidores_drop).strip().lower()
        dff = dff[
            dff["Servidores"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Contrato texto
    if contrato_texto and str(contrato_texto).strip():
        termo = str(contrato_texto).strip().lower()
        dff = dff[
            dff["Contrato"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Contrato dropdown
    if contrato_drop:
        dff = dff[dff["Contrato"] == contrato_drop]

    # Contratada texto
    if contratada_texto and str(contratada_texto).strip():
        termo = str(contratada_texto).strip().lower()
        dff = dff[
            dff["Contratada"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    # Contratada dropdown
    if contratada_drop:
        dff = dff[dff["Contratada"] == contratada_drop]

    # Status
    if status:
        dff = dff[dff["Status"] == status]

    dff = dff.copy()

    def mk_link(row):
        url = row.get("Link Comprasnet")
        contrato = row.get("Contrato")
        if isinstance(url, str) and url.strip() and isinstance(contrato, str):
            return f"[{contrato}]({url.strip()})"
        return str(contrato) if contrato is not None else ""

    dff["Contrato_markdown"] = dff.apply(mk_link, axis=1)

    cols = [
        "Contrato_markdown",
        "Setor",
        "Objeto",
        "Contratada",
        "Final da Vigência",
        "Servidores",
        "Status",
    ]

    return dff[cols].to_dict("records"), dff.to_dict("records")

# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_servidores_texto_fis", "value"),
    Output("filtro_servidores_dropdown_fis", "value"),
    Output("filtro_contrato_texto_fis", "value"),
    Output("filtro_contrato_dropdown_fis", "value"),
    Output("filtro_contratada_texto_fis", "value"),
    Output("filtro_contratada_dropdown_fis", "value"),
    Output("filtro_status_fis", "value"),
    Input("btn_limpar_filtros_fis", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_fis(n):
    return None, None, None, None, None, None, None

# --------------------------------------------------
# Callback: gerar PDF de fiscais
# --------------------------------------------------
wrap_style = ParagraphStyle(
    name="wrap_fiscais",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

@dash.callback(
    Output("download_relatorio_fis", "data"),
    Input("btn_download_relatorio_fis", "n_clicks"),
    State("store_dados_fis", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_fiscais(n, dados_fis):
    if not n or not dados_fis:
        return None

    df = pd.DataFrame(dados_fis)

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
        "Relatório de Fiscais de Contratos",
        ParagraphStyle(
            "titulo_fiscais",
            fontSize=16,
            alignment=TA_CENTER,
            textColor="#0b2b57",
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))
    story.append(
        Paragraph(f"Total de registros: {len(df)}", styles["Normal"])
    )
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        "Setor",
        "Contrato",
        "Objeto",
        "Contratada",
        "Final da Vigência",
        "Servidores",
        "Status",
    ]
    cols = [c for c in cols if c in df.columns]

    df_pdf = df.copy()

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

    return dcc.send_bytes(buffer.getvalue(), "fiscais_paisagem.pdf")
