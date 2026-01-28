import dash
from dash import html, dcc, dash_table, callback
from dash.dependencies import Input, Output
import pandas as pd
from datetime import datetime


# --------------------------------------------------
# Registro da página
# --------------------------------------------------
dash.register_page(
    __name__,
    path="/atas",
    name="Atas",
    title="Atas",
)

# --------------------------------------------------
# Planilha (aba única por GID) - MAIS CONFIÁVEL
# --------------------------------------------------
SHEET_ID = "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog"
GID_CONTROLE = "1976446622"

URL_CONTROLE_ATAS = (
    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export"
    f"?format=csv&gid={GID_CONTROLE}"
)

# --------------------------------------------------
# Carga e tratamento
# --------------------------------------------------
def carregar_base_controle() -> pd.DataFrame:
    # Cabeçalho começa na 2ª linha
    df = pd.read_csv(URL_CONTROLE_ATAS, header=1)

    # Limpa nomes das colunas (sem remover colunas antes do fatiamento)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def carregar_atas_vigentes() -> pd.DataFrame:
    df = carregar_base_controle()

    # A:E => índices 0..4 (5 colunas)
    if df.shape[1] < 5:
        return pd.DataFrame(columns=["Número", "Ata Vigente", "Data Inicial", "Data de Término", "Link_markdown"])

    df = df.iloc[:, 0:5].copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Padroniza nome
    df = df.rename(columns={"ATAS VIGENTES": "Ata Vigente"})

    # Remove colunas Unnamed (caso existam dentro do recorte)
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

    # Filtra somente vigentes
    if "Data de Término" in df.columns:
        df["Data de Término_dt"] = pd.to_datetime(
            df["Data de Término"], dayfirst=True, errors="coerce"
        )
        hoje = datetime.now().date()
        df = df[df["Data de Término_dt"].notna()]
        df = df[df["Data de Término_dt"].dt.date >= hoje]
        df["Data de Término"] = df["Data de Término_dt"].dt.strftime("%d/%m/%Y")

    # Link em markdown (valida URL)
    def formatar_link(url) -> str:
        url = str(url).strip()
        return f"[link]({url})" if url.startswith("http") else ""

    if "Link" in df.columns:
        df["Link_markdown"] = df["Link"].apply(formatar_link)
    else:
        df["Link_markdown"] = ""

    cols = ["Número", "Ata Vigente", "Data Inicial", "Data de Término", "Link_markdown"]
    return df[[c for c in cols if c in df.columns]]


def carregar_atas_andamento() -> pd.DataFrame:
    df = carregar_base_controle()

    # G:I => índices 6..8 (9ª coluna é índice 8)
    # Se a planilha vier com menos colunas, não quebra: devolve vazio
    if df.shape[1] < 9:
        return pd.DataFrame(columns=["Atas em Andamento", "Situação", "Previsão para estar disponível"])

    df = df.iloc[:, 6:9].copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Remove Unnamed do recorte
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

    # Padroniza nomes
    df = df.rename(
        columns={
            "ATAS EM ANDAMENTO": "Atas em Andamento",
            "Situação ": "Situação",
            "Previsão para estar disponível": "Previsão para estar disponível",
        }
    )

    cols = ["Atas em Andamento", "Situação", "Previsão para estar disponível"]
    return df[[c for c in cols if c in df.columns]]


# --------------------------------------------------
# Estilos
# --------------------------------------------------
header_style = {
    "fontWeight": "bold",
    "backgroundColor": "#0b2b57",
    "color": "white",
    "position": "sticky",
    "top": 0,
    "zIndex": 1,
}

cell_style = {
    "textAlign": "center",
    "padding": "6px",
    "fontSize": "12px",
    "whiteSpace": "normal",
    "height": "auto",
}

zebra_style = [{"if": {"row_index": "odd"}, "backgroundColor": "#f5f5f5"}]
datatable_links_css = [{"selector": "p", "rule": "margin: 0; text-align: center;"}]


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    style={"padding": "10px"},
    children=[
        # Atualiza ao abrir e a cada 10 minutos
        dcc.Interval(id="atas_refresh", interval=600000, n_intervals=0),

        # Mostra erro de carregamento na tela (sem esconder)
        html.Div(id="atas_erro", style={"color": "crimson", "textAlign": "center", "marginBottom": "8px"}),

        html.H3("Atas Vigentes", style={"textAlign": "center"}),

        dash_table.DataTable(
            id="tabela_atas_vigentes",
            columns=[
                {"name": "Número", "id": "Número"},
                {"name": "Ata Vigente", "id": "Ata Vigente"},
                {"name": "Data Inicial", "id": "Data Inicial"},
                {"name": "Data de Término", "id": "Data de Término"},
                {"name": "Link", "id": "Link_markdown", "presentation": "markdown"},
            ],
            data=[],
            style_table={"maxHeight": "450px", "overflowY": "auto", "overflowX": "auto"},
            style_cell=cell_style,
            style_header=header_style,
            style_data_conditional=zebra_style,
            css=datatable_links_css,
        ),

        html.H3("Atas em Andamento", style={"marginTop": "20px", "textAlign": "center"}),

        dash_table.DataTable(
            id="tabela_atas_andamento",
            columns=[
                {"name": "Atas em Andamento", "id": "Atas em Andamento"},
                {"name": "Situação", "id": "Situação"},
                {"name": "Previsão para estar disponível", "id": "Previsão para estar disponível"},
            ],
            data=[],
            style_table={"maxHeight": "220px", "overflowY": "auto", "overflowX": "auto"},
            style_cell=cell_style,
            style_header=header_style,
            style_data_conditional=zebra_style,
        ),
    ],
)


# --------------------------------------------------
# Callbacks
# --------------------------------------------------
@callback(
    Output("tabela_atas_vigentes", "data"),
    Output("tabela_atas_andamento", "data"),
    Output("atas_erro", "children"),
    Input("atas_refresh", "n_intervals"),
)
def atualizar_tabelas(_n):
    try:
        df_vig = carregar_atas_vigentes()
        df_and = carregar_atas_andamento()
        return df_vig.to_dict("records"), df_and.to_dict("records"), ""
    except Exception as e:
        # Mostra o erro na tela (e também imprime no log)
        msg = f"Erro ao carregar dados da planilha: {e}"
        print(f"[ATAS] {msg}")
        return [], [], msg
