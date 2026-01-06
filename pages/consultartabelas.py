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

# DataFrame auxiliar com índice e nome da coluna
df_cols = pd.DataFrame(
    {
        "Índice": range(len(df_portarias_base.columns)),
        "Nome da coluna": list(df_portarias_base.columns),
    }
)


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.H4("Colunas da planilha de Portarias (índice e nome)"),
        dash_table.DataTable(
            id="tabela_colunas_portarias",
            columns=[
                {"name": "Índice", "id": "Índice"},
                {"name": "Nome da coluna", "id": "Nome da coluna"},
            ],
            data=df_cols.to_dict("records"),
            style_table={"maxHeight": "300px", "overflowY": "auto"},
            style_cell={
                "textAlign": "left",
                "padding": "4px",
                "fontSize": "12px",
                "whiteSpace": "normal",
            },
            style_header={
                "fontWeight": "bold",
                "backgroundColor": "#0b2b57",
                "color": "white",
            },
        ),

        html.H4("Tabela de Portarias (amostra)"),
        dash_table.DataTable(
            id="tabela_portarias",
            columns=[{"name": c, "id": c} for c in df_portarias_base.columns],
            data=df_portarias_base.head(20).to_dict("records"),
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
