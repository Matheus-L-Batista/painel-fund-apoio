# pages/natureza_despesa_2024.py
# Painel: Naturezas de Despesa utilizadas em 2024 (sem filtros)

import dash
from dash import html, dcc, dash_table
import pandas as pd

# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/natureza-despesa-2024",
    name="Naturezas 2024",
    title="Naturezas de Despesa 2024",
)

# --------------------------------------------------
# URL da planilha (aba TODOS 1)
# --------------------------------------------------
URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1ofT3KdBLI26nDp2SsYePjAgaDIObHT3WDZRwb34g2EU/"
    "gviz/tq?tqx=out:csv&sheet=TODOS%201"
)

# --------------------------------------------------
# Carga dos dados
# --------------------------------------------------
def carregar_dados():
    df = pd.read_csv(URL)
    df.columns = [c.strip() for c in df.columns]
    return df

df = carregar_dados()

# Mantém apenas as colunas desejadas
df = df[["ND SOF", "TITULO"]]

# --------------------------------------------------
# Layout (tabela centralizada e mais estreita)
# --------------------------------------------------
layout = html.Div(
    children=[
        html.H2(
            "Naturezas de Despesa utilizadas em 2024",
            style={"textAlign": "center"},
        ),
        html.Div(
            style={
                "maxWidth": "800px",   # largura máxima da área da tabela
                "margin": "0 auto",    # centraliza horizontalmente
            },
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
                ),
            ],
        ),
    ]
)
