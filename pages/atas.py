import dash
from dash import html, dcc, dash_table
import pandas as pd
from datetime import datetime

dash.register_page(
    __name__,
    path="/atas",
    name="Atas",
    title="Atas",
)

URL_ATAS_AND = (
    "https://docs.google.com/spreadsheets/d/"
    "1fEWJL85yZg3y-ea-qY29LjQpCDto1vuRxF6OYODqMNE/"
    "gviz/tq?tqx=out:csv&sheet=ATAS%20EM%20ANDAMENTO"
)
URL_ATAS_VIG = (
    "https://docs.google.com/spreadsheets/d/"
    "1fEWJL85yZg3y-ea-qY29LjQpCDto1vuRxF6OYODqMNE/"
    "gviz/tq?tqx=out:csv&sheet=ATAS%20VIGENTES"
)


def carregar_atas_andamento():
    df = pd.read_csv(URL_ATAS_AND)
    df = df[[c for c in df.columns if not c.startswith("Unnamed")]]
    df.columns = [c.strip() for c in df.columns]

    df = df.rename(
        columns={
            "ATAS EM ANDAMENTO": "Atas em Andamento",
            "Situação": "Situação",
            "Situação ": "Situação",
            "Previsão para estar disponível": "Previsão para estar disponível",
        }
    )

    cols = [
        "Atas em Andamento",
        "Situação",
        "Previsão para estar disponível",
    ]
    df = df[[c for c in cols if c in df.columns]]
    return df


def carregar_atas_vigentes():
    df = pd.read_csv(URL_ATAS_VIG)
    df = df[[c for c in df.columns if not c.startswith("Unnamed")]]
    df.columns = [c.strip() for c in df.columns]

    df = df.rename(
        columns={
            "ATAS VIGENTES": "Ata Vigente",
        }
    )

    # converte Data de Término e filtra atas já vencidas
    if "Data de Término" in df.columns:
        df["Data de Término_dt"] = pd.to_datetime(
            df["Data de Término"], dayfirst=True, errors="coerce"
        )
        hoje = datetime.now().date()
        df = df[df["Data de Término_dt"].dt.date >= hoje]
    else:
        df["Data de Término_dt"] = pd.NaT

    # link clicável
    if "Link" in df.columns:
        df["Link_markdown"] = df["Link"].apply(
            lambda url: "[link](" + str(url).strip() + ")" if str(url).strip() else ""
        )
    else:
        df["Link_markdown"] = ""

    return df


df_and = carregar_atas_andamento()
df_vig = carregar_atas_vigentes()

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
}

table_style = {
    "overflowX": "auto",
    "maxHeight": "400px",
}

layout = html.Div(
    children=[
        html.H3("Atas Vigentes"),
        dash_table.DataTable(
            id="tabela_atas_vigentes",
            columns=[
                {"name": "Número", "id": "Número"},
                {"name": "Ata Vigente", "id": "Ata Vigente"},
                {"name": "Data Inicial", "id": "Data Inicial"},
                {"name": "Data de Término", "id": "Data de Término"},
                {
                    "name": "Link",
                    "id": "Link_markdown",
                    "presentation": "markdown",
                },
            ],
            data=df_vig[
                [
                    c
                    for c in [
                        "Número",
                        "Ata Vigente",
                        "Data Inicial",
                        "Data de Término",
                        "Link_markdown",
                    ]
                    if c in df_vig.columns
                ]
            ].to_dict("records"),
            style_table=table_style,
            style_cell=cell_style,
            style_header=header_style,
        ),
        html.H3(style={"marginTop": "20px"}, children="Atas em Andamento"),
        dash_table.DataTable(
            id="tabela_atas_andamento",
            columns=[
                {
                    "name": "Atas em Andamento",
                    "id": "Atas em Andamento",
                },
                {"name": "Situação", "id": "Situação"},
                {
                    "name": "Previsão para estar disponível",
                    "id": "Previsão para estar disponível",
                },
            ],
            data=df_and.to_dict("records"),
            style_table=table_style,
            style_cell=cell_style,
            style_header=header_style,
        ),
    ]
)
