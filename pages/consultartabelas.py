import dash
from dash import html, dash_table
import pandas as pd

# Registra a página no Dash Pages
dash.register_page(
    __name__,
    path="/consultartabelas",
    name="consultartabelas",
    title="Consultar Tabelas",
)

# URL da planilha (aba Fiscais)
URL_PORTARIAS = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=Fiscais"
)

# Função para carregar os dados a partir da 3ª linha
def carregar_dados_portarias():
    # header=2 -> usa a 3ª linha como cabeçalho
    df = pd.read_csv(URL_PORTARIAS, header=3)
    df.columns = [c.strip() for c in df.columns]
    return df

# Carrega os dados
df_portarias_base = carregar_dados_portarias()

# Layout da página
layout = html.Div(
    children=[
        html.H4("Colunas da planilha de Portarias"),
        html.Ul(
            [html.Li(col) for col in df_portarias_base.columns]
        ),

        html.H4("Tabela de Portarias"),
        dash_table.DataTable(
            id="tabela_portarias",
            columns=[{"name": c, "id": c} for c in df_portarias_base.columns],
            data=df_portarias_base.to_dict("records"),
            row_selectable=False,
            cell_selectable=False,
            style_table={
                "overflowX": "auto",
                "overflowY": "auto",
                "maxHeight": "500px",
                "position": "relative",
            },
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
                "minWidth": "80px",
                "maxWidth": "220px",
                "whiteSpace": "normal",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "textAlign": "center",
                "position": "sticky",
                "top": 0,
                "zIndex": 10,
            },
        ),
    ]
)
