import dash
from dash import html, dash_table
import pandas as pd

# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/consultartabelas",
    name="consultartabelas",
    title="Consultar Tabelas",
)

# --------------------------------------------------
# URL da planilha (ABA SEM ACENTO NA URL)
# --------------------------------------------------
URL_PORTARIAS = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=Limite%20de%20Gasto%20-%20Itajub%C3%A1"
)

# --------------------------------------------------
# Função de carga dos dados
# --------------------------------------------------
def carregar_dados_portarias():
    df = pd.read_csv(URL_PORTARIAS)
    df.columns = [c.strip() for c in df.columns]
    return df

# --------------------------------------------------
# Carrega os dados
# --------------------------------------------------
df_portarias_base = carregar_dados_portarias()

# --------------------------------------------------
# Layout
# --------------------------------------------------
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
            },
            style_cell={
                "textAlign": "center",
                "padding": "6px",
                "fontSize": "12px",
                "whiteSpace": "normal",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
                "position": "sticky",
                "top": 0,
                "zIndex": 10,
            },
        ),
    ]
)
