import dash
from dash import html, dcc, dash_table, Input, Output
import pandas as pd

dash.register_page(
    __name__,
    path="/pca",
    name="PCA",
    title="PCA",
)

URL_PCA = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=PCA%20-%20BI"
)

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}


def conv_moeda_br(v):
    if isinstance(v, str):
        v = v.strip()
        if v == "":
            return None
        v = v.replace(".", "").replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def formatar_moeda(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def carregar_dados_pca():
    df = pd.read_csv(URL_PCA, header=0)
    df.columns = [c.strip() for c in df.columns]

    base_cols = [
        "Ano",
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Valor Total",
        "Saldo",
        "Item",
        "Código Classe / Grupo",
        "Nome Classe/Grupo",
        "Código PDM material",
        "Nome do PDM material",
        "Processo",
        "Observações",
        "Objeto",
        "SRP ou Outro Valor",
        "Valor",
    ]
    for c in base_cols:
        if c not in df.columns:
            df[c] = None

    # converte campos monetários principais
    df["Valor Total"] = df["Valor Total"].apply(conv_moeda_br)
    df["Saldo"] = df["Saldo"].apply(conv_moeda_br)
    df["Valor"] = df["Valor"].apply(conv_moeda_br)

    df["Ano"] = df["Ano"].astype("string")
    for c in [
        "Área requisitante",
        "Material ou Serviço",
        "DFD",
        "Item",
        "Código Classe / Grupo",
        "Nome Classe/Grupo",
        "Código PDM material",
        "Nome do PDM material",
        "Processo",
        "Observações",
        "Objeto",
    ]:
        df[c] = df[c].astype("string")

    return df


df_pca_base = carregar_dados_pca()

# Planejamento
df_planejamento = df_pca_base.copy()
df_planejamento["Planejado"] = df_planejamento["Valor Total"]
df_planejamento["Executado"] = df_planejamento["Planejado"] - df_planejamento["Saldo"]

# Processos: usar apenas Valor, agregado por DFD + Item
df_processos = df_pca_base.copy()
df_processos["Valor_Proc"] = df_processos["Valor"]

df_processos_group = (
    df_processos.groupby(
        ["DFD", "Item", "Área requisitante", "Material ou Serviço"],
        as_index=False,
    ).agg(
        {
            "Valor_Proc": "sum",
            "Processo": "first",
            "Objeto": "first",
            "Observações": "first",
        }
    )
)
df_processos_group.rename(
    columns={"Valor_Proc": "Valor_Total_Processos"},
    inplace=True,
)


layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_pca",
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
                        html.Div(
                            style={"minWidth": "120px", "flex": "0 0 140px"},
                            children=[
                                html.Label("Ano"),
                                dcc.Dropdown(
                                    id="filtro_ano_pca",
                                    options=[
                                        {"label": str(a), "value": str(a)}
                                        for a in sorted(
                                            df_pca_base["Ano"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a).strip() != ""
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
                                html.Label("Nome Classe/Grupo (digitação)"),
                                dcc.Input(
                                    id="filtro_classe_texto_pca",
                                    type="text",
                                    placeholder="Digite parte do nome da classe/grupo",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("DFD (digitação)"),
                                dcc.Input(
                                    id="filtro_dfd_texto_pca",
                                    type="text",
                                    placeholder="Digite parte do DFD",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                    },
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Área requisitante"),
                                dcc.Dropdown(
                                    id="filtro_area_pca",
                                    options=[
                                        {"label": a, "value": a}
                                        for a in sorted(
                                            df_pca_base["Área requisitante"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(a).strip() != ""
                                    ],
                                    value=None,
                                    placeholder="Todas",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Material ou Serviço"),
                                dcc.Dropdown(
                                    id="filtro_tipo_pca",
                                    options=[
                                        {"label": t, "value": t}
                                        for t in sorted(
                                            df_pca_base["Material ou Serviço"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(t).strip() != ""
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
            ],
        ),
        html.Div(
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "10px",
                "marginTop": "10px",
            },
            children=[
                html.Div(
                    style={"flex": "1 1 50%", "minWidth": "300px"},
                    children=[
                        html.H4("Planejamento (PCA)"),
                        dash_table.DataTable(
                            id="tabela_pca_planejamento",
                            columns=[
                                {"name": "DFD", "id": "DFD"},
                                {
                                    "name": "Área requisitante",
                                    "id": "Área requisitante",
                                },
                                {
                                    "name": "Material ou Serviço",
                                    "id": "Material ou Serviço",
                                },
                                {"name": "Item", "id": "Item"},
                                {
                                    "name": "Nome Classe/Grupo",
                                    "id": "Nome Classe/Grupo",
                                },
                                {"name": "Planejado", "id": "Planejado_fmt"},
                                {"name": "Executado", "id": "Executado_fmt"},
                                {"name": "Saldo", "id": "Saldo_fmt"},
                            ],
                            data=[],
                            style_table={
                                "overflowX": "auto",
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
                            },
                        ),
                    ],
                ),
                html.Div(
                    style={"flex": "1 1 50%", "minWidth": "300px"},
                    children=[
                        html.H4(
                            "Processos vinculados (Valor da coluna Valor por DFD + Item)"
                        ),
                        dash_table.DataTable(
                            id="tabela_pca_processos",
                            columns=[
                                {"name": "DFD", "id": "DFD"},
                                {
                                    "name": "Área requisitante",
                                    "id": "Área requisitante",
                                },
                                {
                                    "name": "Material ou Serviço",
                                    "id": "Material ou Serviço",
                                },
                                {"name": "Item", "id": "Item"},
                                {"name": "Processo", "id": "Processo"},
                                {"name": "Objeto", "id": "Objeto"},
                                {
                                    "name": "Observações",
                                    "id": "Observações",
                                },
                                {
                                    "name": "Valor Total Processos",
                                    "id": "Valor_Total_Processos_fmt",
                                },
                            ],
                            data=[],
                            style_table={
                                "overflowX": "auto",
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
                            },
                        ),
                    ],
                ),
            ],
        ),
    ]
)


@dash.callback(
    Output("tabela_pca_planejamento", "data"),
    Output("tabela_pca_processos", "data"),
    Input("filtro_ano_pca", "value"),
    Input("filtro_classe_texto_pca", "value"),
    Input("filtro_dfd_texto_pca", "value"),
    Input("filtro_area_pca", "value"),
    Input("filtro_tipo_pca", "value"),
)
def atualizar_tabelas_pca(
    ano,
    classe_texto,
    dfd_texto,
    area,
    tipo,
):
    dff_plan = df_planejamento.copy()
    dff_proc = df_processos_group.copy()

    if ano:
        dff_plan = dff_plan[dff_plan["Ano"] == str(ano)]
        dff_proc = dff_proc[dff_proc["DFD"].str.contains(str(ano), na=False)]

    if classe_texto and str(classe_texto).strip():
        termo = str(classe_texto).strip().lower()
        dff_plan = dff_plan[
            dff_plan["Nome Classe/Grupo"]
            .astype(str)
            .str.lower()
            .str.contains(termo, na=False)
        ]

    if dfd_texto and str(dfd_texto).strip():
        termo = str(dfd_texto).strip().lower()
        dff_plan = dff_plan[
            dff_plan["DFD"].astype(str).str.lower().str.contains(termo, na=False)
        ]
        dff_proc = dff_proc[
            dff_proc["DFD"].astype(str).str.lower().str.contains(termo, na=False)
        ]

    if area:
        dff_plan = dff_plan[dff_plan["Área requisitante"] == area]
        dff_proc = dff_proc[dff_proc["Área requisitante"] == area]

    if tipo:
        dff_plan = dff_plan[dff_plan["Material ou Serviço"] == tipo]
        dff_proc = dff_proc[dff_proc["Material ou Serviço"] == tipo]

    dff_plan["Planejado_fmt"] = dff_plan["Planejado"].apply(formatar_moeda)
    dff_plan["Executado_fmt"] = dff_plan["Executado"].apply(formatar_moeda)
    dff_plan["Saldo_fmt"] = dff_plan["Saldo"].apply(formatar_moeda)

    dff_proc["Valor_Total_Processos_fmt"] = dff_proc[
        "Valor_Total_Processos"
    ].apply(formatar_moeda)

    cols_planejamento = [
        "DFD",
        "Área requisitante",
        "Material ou Serviço",
        "Item",
        "Nome Classe/Grupo",
        "Planejado_fmt",
        "Executado_fmt",
        "Saldo_fmt",
    ]
    dados_planejamento = dff_plan[cols_planejamento].fillna("").to_dict("records")

    cols_processos = [
        "DFD",
        "Área requisitante",
        "Material ou Serviço",
        "Item",
        "Processo",
        "Objeto",
        "Observações",
        "Valor_Total_Processos_fmt",
    ]
    dados_processos = dff_proc[cols_processos].fillna("").to_dict("records")

    return dados_planejamento, dados_processos
