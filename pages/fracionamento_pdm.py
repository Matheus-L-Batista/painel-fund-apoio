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


# --------------------------------------------------
# Registro da página
# --------------------------------------------------


dash.register_page(
    __name__,
    path="/fracionamento_pdm",
    name="Fracionamento de Despesas PDM",
    title="Fracionamento de Despesas PDM",
)


# --------------------------------------------------
# URL da planilha
# --------------------------------------------------


URL_LIMITE_GASTO_ITA = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=Limite%20de%20Gasto%20-%20Itajub%C3%A1"
)

COL_PDM = "PDM"
COL_DESC_ORIG = "Descrição.1"  # será renomeada para "Descrição"


# --------------------------------------------------
# Carga e tratamento dos dados
# --------------------------------------------------


def carregar_dados_limite_pdm():
    df = pd.read_csv(URL_LIMITE_GASTO_ITA)
    df.columns = [c.strip() for c in df.columns]

    # Renomeia apenas a descrição
    renomeios = {
        COL_DESC_ORIG: "Descrição",
    }
    df = df.rename(columns=renomeios)

    # Garante colunas básicas
    if COL_PDM not in df.columns:
        df[COL_PDM] = ""
    if "Descrição" not in df.columns:
        df["Descrição"] = ""

    # PDM como texto com 5 dígitos (zeros à esquerda)
    df[COL_PDM] = (
        df[COL_PDM]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .str.zfill(5)
    )

    # Trata Valor Empenhado (se existir Unnamed: 7)
    if "Unnamed: 7" in df.columns:
        df["Valor Empenhado"] = (
            df["Unnamed: 7"]
            .astype(str)
            .str.replace(".", "", regex=False)   # milhar
            .str.replace(",", ".", regex=False)  # decimal
        )
        df["Valor Empenhado"] = pd.to_numeric(df["Valor Empenhado"], errors="coerce")
    else:
        df["Valor Empenhado"] = 0.0

    # Cria Limite da Dispensa fixo: R$ 65.492,11 -> 65492.11
    valor_limite = 65492.11
    df["Limite da Dispensa"] = valor_limite

    # Saldo = limite - empenhado
    df["Saldo para contratação"] = df["Limite da Dispensa"] - df["Valor Empenhado"]

    # Lista de PDM para dropdown
    pdms_unicos = sorted(
        [
            c
            for c in df[COL_PDM].dropna().unique()
            if isinstance(c, str) and c.strip() != ""
        ]
    )
    df._lista_pdms_unicos = pdms_unicos

    return df


df_limite_pdm_base = carregar_dados_limite_pdm()
PDMS_UNICOS = getattr(df_limite_pdm_base, "_lista_pdms_unicos", [])

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
            id="barra_filtros_limite_itajuba_pdm",
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
                        # Filtro PDM (digitação)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("PDM (digitação)"),
                                dcc.Input(
                                    id="filtro_pdm_texto_itajuba",
                                    type="text",
                                    placeholder="Digite parte do PDM",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                        # Filtro PDM (dropdown simples)
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("PDM"),
                                dcc.Dropdown(
                                    id="filtro_pdm_dropdown_itajuba",
                                    options=[
                                        {"label": c, "value": c}
                                        for c in PDMS_UNICOS
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
                    style={"marginTop": "4px"},
                    children=[
                        html.Button(
                            "Limpar filtros",
                            id="btn_limpar_filtros_limite_itajuba_pdm",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_limite_itajuba_pdm",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_limite_itajuba_pdm"),
                    ],
                ),
            ],
        ),
        html.H4("Limite de Gasto – Itajubá por PDM"),
        dash_table.DataTable(
            id="tabela_limite_itajuba_pdm",
            columns=[
                {"name": "PDM", "id": COL_PDM},
                {"name": "Descrição", "id": "Descrição"},
                {"name": "Valor Empenhado (R$)", "id": "Valor Empenhado_fmt"},
                {"name": "Limite da Dispensa (R$)", "id": "Limite da Dispensa_fmt"},
                {
                    "name": "Saldo para contratação (R$)",
                    "id": "Saldo para contratação_fmt",
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
            style_data_conditional=[
                {
                    "if": {"column_id": "Saldo para contratação_fmt"},
                    "backgroundColor": "#f9f9f9",
                }
            ],
        ),
        dcc.Store(id="store_dados_limite_itajuba_pdm"),
    ]
)


# --------------------------------------------------
# Callback: aplicar filtros
# --------------------------------------------------


@dash.callback(
    Output("tabela_limite_itajuba_pdm", "data"),
    Output("store_dados_limite_itajuba_pdm", "data"),
    Input("filtro_pdm_texto_itajuba", "value"),
    Input("filtro_pdm_dropdown_itajuba", "value"),
)
def atualizar_tabela_limite_itajuba_pdm(pdm_texto, pdm_drop):
    dff = df_limite_pdm_base.copy()

    # Filtro por PDM (texto)
    if pdm_texto and str(pdm_texto).strip():
        termo = str(pdm_texto).strip().lower()
        dff = dff[
            dff[COL_PDM]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    # Filtro por PDM (dropdown)
    if pdm_drop:
        dff = dff[dff[COL_PDM] == pdm_drop]

    cols_tabela = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado",
        "Limite da Dispensa",
        "Saldo para contratação",
    ]

    for c in cols_tabela:
        if c not in dff.columns:
            dff[c] = pd.NA

    dff_display = dff[cols_tabela].copy()

    # Formatação moeda R$
    def fmt_moeda(v):
        if pd.isna(v):
            return ""
        return "R$ " + (
            f"{v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    dff_display["Valor Empenhado_fmt"] = dff_display["Valor Empenhado"].apply(fmt_moeda)
    dff_display["Limite da Dispensa_fmt"] = dff_display["Limite da Dispensa"].apply(fmt_moeda)
    dff_display["Saldo para contratação_fmt"] = dff_display["Saldo para contratação"].apply(
        fmt_moeda
    )

    cols_tabela_display = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado_fmt",
        "Limite da Dispensa_fmt",
        "Saldo para contratação_fmt",
    ]

    return dff_display[cols_tabela_display].to_dict("records"), dff.to_dict("records")


# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------


@dash.callback(
    Output("filtro_pdm_texto_itajuba", "value"),
    Output("filtro_pdm_dropdown_itajuba", "value"),
    Input("btn_limpar_filtros_limite_itajuba_pdm", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_limite_itajuba_pdm(n):
    return None, None


# --------------------------------------------------
# Callback: gerar PDF
# --------------------------------------------------


wrap_style = ParagraphStyle(
    name="wrap_limite_itajuba_pdm",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)


def wrap(text):
    return Paragraph(str(text), wrap_style)


@dash.callback(
    Output("download_relatorio_limite_itajuba_pdm", "data"),
    Input("btn_download_relatorio_limite_itajuba_pdm", "n_clicks"),
    State("store_dados_limite_itajuba_pdm", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_limite_itajuba_pdm(n, dados):
    if not n or not dados:
        return None

    df = pd.DataFrame(dados)

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
        "Relatório – Limite de Gasto – Itajubá por PDM",
        ParagraphStyle(
            "titulo_limite_itajuba_pdm",
            fontSize=16,
            alignment=TA_CENTER,
            textColor="#0b2b57",
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        COL_PDM,
        "Descrição",
        "Valor Empenhado",
        "Limite da Dispensa",
        "Saldo para contratação",
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = ""

    def fmt_moeda_pdf(v):
        if pd.isna(v):
            return ""
        return "R$ " + (
            f"{v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    df_pdf = df.copy()
    for col in ["Valor Empenhado", "Limite da Dispensa", "Saldo para contratação"]:
        if col in df_pdf.columns:
            df_pdf[col] = df_pdf[col].apply(fmt_moeda_pdf)

    header = cols
    table_data = [header]

    for _, row in df_pdf[cols].iterrows():
        linha = [wrap(row[c]) for c in cols]
        table_data.append(linha)

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

    return dcc.send_bytes(buffer.getvalue(), "limite_gasto_itajuba_pdm.pdf")
