# pages/natureza_despesa_2024.py

import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors

# Painel: Naturezas de Despesa utilizadas em 2024 sem filtros

dash.register_page(
    __name__,
    path="/natureza-despesa-2024",
    name="Naturezas 2024",
    title="Naturezas de Despesa 2024",
)

URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1ofT3KdBLI26nDp2SsYePjAgaDIObHT3WDZRwb34g2EU/"
    "gviz/tq?tqx=out:csv&sheet=TODOS201"
)

def carregar_dados():
    df = pd.read_csv(URL)
    df.columns = [c.strip() for c in df.columns]
    df = df[["ND SOF", "TITULO"]]
    return df  # [file:2]

df = carregar_dados()

layout = html.Div(
    children=[
        html.H2(
            "Naturezas de Despesa utilizadas em 2024",
            style={"textAlign": "center"},
        ),
        html.Div(
            style={"marginBottom": "10px", "textAlign": "right"},
            children=[
                html.Button(
                    "Baixar Relatório PDF",
                    id="btn_download_relatorio_natureza_2024",
                    n_clicks=0,
                    className="filtros-button",
                ),
                dcc.Download(id="download_relatorio_natureza_2024"),
            ],
        ),
        html.Div(
            style={"maxWidth": "800px", "margin": "0 auto"},
            children=[
                dash_table.DataTable(
                    id="tabela_natureza_2024",
                    data=df.to_dict("records"),
                    columns=[{"name": c, "id": c} for c in df.columns],
                    style_table={
                        "overflowX": "auto",
                        "maxHeight": "80vh",
                        "overflowY": "auto",
                    },
                    style_cell={
                        "textAlign": "left",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                        "height": "auto",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#0b2b57",
                        "color": "white",
                    },
                    page_size=50,
                )
            ],
        ),
    ]
)

# ---------------- PDF callback ----------------

wrap_style = ParagraphStyle(
    name="wrap",
    fontSize=8,
    leading=10,
    spaceAfter=2,
    alignment=TA_LEFT,
)

def wrap(text):
    return Paragraph(str(text), wrap_style)

@dash.callback(
    Output("download_relatorio_natureza_2024", "data"),
    Input("btn_download_relatorio_natureza_2024", "n_clicks"),
    State("tabela_natureza_2024", "data"),
    prevent_initial_call=True,
)
def gerar_pdf(n, tabela):
    if not n or not tabela:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    story = []

    titulo = Paragraph(
        "Naturezas de Despesa utilizadas em 2024",
        ParagraphStyle(
            "titulo", fontSize=18, alignment=TA_CENTER, textColor="#0b2b57"
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    # Cabeçalho e linhas
    colunas = list(tabela[0].keys())
    header = [wrap(c) for c in colunas]
    table_data = [header]

    for r in tabela:
        row = [wrap(r.get(c, "")) for c in colunas]
        table_data.append(row)

    col_widths = [3.0 * inch, 7.0 * inch]  # ND SOF, TITULO
    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, colors.HexColor("#f5f5f5")],
                ),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )

    story.append(tbl)
    doc.build(story)
    buffer.seek(0)

    from dash import dcc
    return dcc.send_bytes(buffer.getvalue(), "naturezas_despesa_2024.pdf")
