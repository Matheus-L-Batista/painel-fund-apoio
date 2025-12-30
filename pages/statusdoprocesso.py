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
    path="/statusdoprocesso",
    name="Status do Processo",
    title="Status do Processo",
)


# --------------------------------------------------
# Fonte de dados (Consulta BI)
# --------------------------------------------------
URL_CONSULTA_BI = (
    "https://docs.google.com/spreadsheets/d/"
    "1YNg6WRww19Gf79ISjQtb8tkzjX2lscHirnR_F3wGjog/"
    "gviz/tq?tqx=out:csv&sheet=Consulta%20BI"
)


# --------------------------------------------------
# Carga e tratamento: espelha o CSV e empilha Data Mov, Data Mov.1, ...
# --------------------------------------------------
def carregar_dados_status():
    # lê a aba Consulta BI
    df = pd.read_csv(URL_CONSULTA_BI, header=0)
    df.columns = [c.strip() for c in df.columns]

    # colunas fixas de processo (base do Grupo0 no M)
    col_fixas = [
        "Linha",
        "Finalizado",
        "Processo",
        "Requisitante",
        "Objeto",
        "Modalidade",
        "Número",
        "Valor inicial",
        "Não concluído",
        "Entrada na DCC",
    ]
    for c in col_fixas:
        if c not in df.columns:
            df[c] = None

    # garante que a primeira Data Mov exista
    if "Data Mov" not in df.columns:
        df["Data Mov"] = None

    # garante colunas base de movimentação
    for c in ["E/S", "Deptº", "Ação"]:
        if c not in df.columns:
            df[c] = None

    # captura todas as colunas que começam com "Data Mov"
    # (padrão pandas: Data Mov, Data Mov.1, Data Mov.2, ...)
    data_cols = [c for c in df.columns if c.startswith("Data Mov")]

    grupos = []

    # 1) grupo base: usa a coluna "Data Mov" sem sufixo
    grupo0 = df[col_fixas + ["Data Mov", "E/S", "Deptº", "Ação"]].copy()
    grupos.append(grupo0)

    # 2) grupos adicionais: Data Mov.1, Data Mov.2, ...
    for col in data_cols:
        if col == "Data Mov":
            continue  # já tratado no grupo0

        # col = "Data Mov.1" => suf = ".1"
        suf = col[len("Data Mov"):]  # ".1", ".2", ...
        col_data = f"Data Mov{suf}"
        col_es = f"E/S{suf}"
        col_dept = f"Deptº{suf}"
        col_acao = f"Ação{suf}"

        # garante que as colunas existam
        for c in [col_data, col_es, col_dept, col_acao]:
            if c not in df.columns:
                df[c] = None

        bloco = df[col_fixas + [col_data, col_es, col_dept, col_acao]].copy()
        bloco = bloco.rename(
            columns={
                col_data: "Data Mov",
                col_es: "E/S",
                col_dept: "Deptº",
                col_acao: "Ação",
            }
        )
        grupos.append(bloco)

    # empilha tudo (equivalente ao Table.Combine da TabelaUnida)
    tabela_unida = pd.concat(grupos, ignore_index=True)

    # remove linhas totalmente vazias (NaN ou string vazia em todas as colunas)
    t_aux = tabela_unida.replace({None: pd.NA}).fillna("")
    mask_nao_vazia = t_aux.apply(
        lambda row: any(v not in ("", None) for v in row.values), axis=1
    )
    tabela_unida = tabela_unida[mask_nao_vazia].copy()

    # tipagem básica
    tabela_unida["Linha"] = tabela_unida["Linha"].astype(str)

    for col in [
        "Finalizado",
        "Processo",
        "Requisitante",
        "Objeto",
        "Modalidade",
        "Número",
        "Não concluído",
        "E/S",
        "Deptº",
        "Ação",
    ]:
        if col in tabela_unida.columns:
            tabela_unida[col] = tabela_unida[col].astype("string")

    if "Valor inicial" in tabela_unida.columns:
        tabela_unida["Valor inicial"] = pd.to_numeric(
            tabela_unida["Valor inicial"], errors="coerce"
        )

    # datas como datetime
    for col in ["Entrada na DCC", "Data Mov"]:
        if col in tabela_unida.columns:
            tabela_unida[col] = pd.to_datetime(
                tabela_unida[col], errors="coerce", dayfirst=True
            )

    tabela_unida["Finalizado"] = tabela_unida["Finalizado"].fillna("")

    return tabela_unida


df_status = carregar_dados_status()

dropdown_style = {
    "color": "black",
    "width": "100%",
    "marginBottom": "6px",
    "whiteSpace": "normal",
}

# monta opções de Processo na ordem da coluna Linha (descendente)
df_proc_opts = df_status[["Processo", "Linha"]].copy()
df_proc_opts = df_proc_opts.dropna(subset=["Processo"])
df_proc_opts["Linha_num"] = pd.to_numeric(df_proc_opts["Linha"], errors="coerce")
df_proc_opts = df_proc_opts.sort_values("Linha_num", ascending=False)
df_proc_opts = df_proc_opts.drop_duplicates(subset=["Processo"], keep="first")

processo_options = [
    {"label": row["Processo"], "value": row["Processo"]}
    for _, row in df_proc_opts.iterrows()
]


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_status",
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
                            style={"minWidth": "220px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Processo (digitação)"),
                                dcc.Input(
                                    id="filtro_processo_texto",
                                    type="text",
                                    placeholder="Digite parte do processo",
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
                                html.Label("Processo (seleção)"),
                                dcc.Dropdown(
                                    id="filtro_processo",
                                    options=processo_options,
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
                                html.Label("Requisitante"),
                                dcc.Dropdown(
                                    id="filtro_requisitante",
                                    options=[
                                        {"label": r, "value": r}
                                        for r in sorted(
                                            df_status["Requisitante"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(r) != ""
                                    ],
                                    value=None,
                                    placeholder="Todos",
                                    clearable=True,
                                    style=dropdown_style,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"minWidth": "260px", "flex": "2 1 320px"},
                            children=[
                                html.Label("Objeto"),
                                dcc.Dropdown(
                                    id="filtro_objeto",
                                    options=[
                                        {"label": o, "value": o}
                                        for o in sorted(
                                            df_status["Objeto"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(o) != ""
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
                                html.Label("Modalidade"),
                                dcc.Dropdown(
                                    id="filtro_modalidade",
                                    options=[
                                        {"label": m, "value": m}
                                        for m in sorted(
                                            df_status["Modalidade"]
                                            .dropna()
                                            .unique()
                                        )
                                        if str(m) != ""
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
                            id="btn_limpar_filtros_status",
                            n_clicks=0,
                            className="filtros-button",
                        ),
                        html.Button(
                            "Baixar Relatório PDF",
                            id="btn_download_relatorio_status",
                            n_clicks=0,
                            className="filtros-button",
                            style={"marginLeft": "10px"},
                        ),
                        dcc.Download(id="download_relatorio_status"),
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
                        html.H4("Dados do Processo"),
                        dash_table.DataTable(
                            id="tabela_status_esquerda",
                            columns=[
                                {"name": "Processo", "id": "Processo"},
                                {"name": "Requisitante", "id": "Requisitante"},
                                {"name": "Objeto", "id": "Objeto"},
                                {"name": "Modalidade", "id": "Modalidade"},
                                {"name": "Linha", "id": "Linha"},
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
                        html.H4("Movimentações"),
                        dash_table.DataTable(
                            id="tabela_status_direita",
                            columns=[
                                {"name": "Data Mov", "id": "Data Mov"},
                                {"name": "E/S", "id": "E/S"},
                                {"name": "Ação", "id": "Ação"},
                                {"name": "Deptº", "id": "Deptº"},
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
        dcc.Store(id="store_dados_status"),
    ]
)


# --------------------------------------------------
# Callbacks principais
# --------------------------------------------------
@dash.callback(
    Output("tabela_status_esquerda", "data"),
    Output("tabela_status_direita", "data"),
    Output("store_dados_status", "data"),
    Input("filtro_processo_texto", "value"),
    Input("filtro_processo", "value"),
    Input("filtro_requisitante", "value"),
    Input("filtro_objeto", "value"),
    Input("filtro_modalidade", "value"),
)
def atualizar_tabelas(
    proc_texto,
    proc_select,
    requisitante,
    objeto,
    modalidade,
):
    dff = df_status.copy()

    if proc_texto and str(proc_texto).strip():
        termo = str(proc_texto).strip()
        dff = dff[
            dff["Processo"]
            .astype(str)
            .str.contains(termo, case=False, na=False)
        ]

    if proc_select:
        dff = dff[dff["Processo"] == proc_select]

    if requisitante:
        dff = dff[dff["Requisitante"] == requisitante]

    if objeto:
        dff = dff[dff["Objeto"] == objeto]

    if modalidade:
        dff = dff[dff["Modalidade"] == modalidade]

    # ordena por Linha
    try:
        dff["Linha_ordenacao"] = pd.to_numeric(dff["Linha"], errors="coerce")
    except Exception:
        dff["Linha_ordenacao"] = dff["Linha"]
    dff = dff.sort_values("Linha_ordenacao", ascending=False)

    # tabela esquerda: 1 linha por processo
    mask_proc_valido = dff["Processo"].astype(str).str.strip().ne("")
    dff_esq = dff[mask_proc_valido].copy()
    dff_esq = dff_esq.drop_duplicates(subset=["Processo"], keep="first")

    dados_esquerda = dff_esq[
        ["Processo", "Requisitante", "Objeto", "Modalidade", "Linha"]
    ].to_dict("records")

    # tabela direita: todas as movimentações do filtro
    dff_dir = dff.copy()
    mask_acao_valida = dff_dir["Ação"].astype(str).str.strip().ne("")
    dff_dir = dff_dir[mask_acao_valida].copy()

    dff_dir["Data Mov"] = pd.to_datetime(
        dff_dir["Data Mov"], errors="coerce"
    )
    dff_dir["Data Mov"] = dff_dir["Data Mov"].dt.strftime("%d/%m/%Y").fillna("")

    dados_direita = dff_dir[
        ["Data Mov", "E/S", "Ação", "Deptº"]
    ].to_dict("records")

    # para o PDF, pode-se optar por dff_dir (só movimentações) ou dff completo
    return dados_esquerda, dados_direita, dff_dir.to_dict("records")


# --------------------------------------------------
# Callback: limpar filtros
# --------------------------------------------------
@dash.callback(
    Output("filtro_processo_texto", "value"),
    Output("filtro_processo", "value"),
    Output("filtro_requisitante", "value"),
    Output("filtro_objeto", "value"),
    Output("filtro_modalidade", "value"),
    Input("btn_limpar_filtros_status", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_status(n):
    return None, None, None, None, None


# --------------------------------------------------
# Callback: gerar relatório PDF
# --------------------------------------------------
wrap_style_status = ParagraphStyle(
    name="wrap_status",
    fontSize=8,
    leading=10,
    spaceAfter=4,
)


def wrap_text_status(text):
    return Paragraph(str(text), wrap_style_status)


@dash.callback(
    Output("download_relatorio_status", "data"),
    Input("btn_download_relatorio_status", "n_clicks"),
    State("store_dados_status", "data"),
    prevent_initial_call=True,
)
def gerar_pdf_status(n, dados_status):
    if not n or not dados_status:
        return None

    df = pd.DataFrame(dados_status)

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
        "Relatório de Status do Processo",
        ParagraphStyle(
            "titulo_status",
            fontSize=16,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0b2b57"),
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

    story.append(Paragraph(f"Total de registros: {len(df)}", styles["Normal"]))
    story.append(Spacer(1, 0.15 * inch))

    cols = [
        "Processo",
        "Requisitante",
        "Objeto",
        "Modalidade",
        "Linha",
        "Data Mov",
        "E/S",
        "Deptº",
        "Ação",
    ]
    cols = [c for c in cols if c in df.columns]

    df_pdf = df.copy()
    if "Data Mov" in df_pdf.columns:
        df_pdf["Data Mov"] = pd.to_datetime(
            df_pdf["Data Mov"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    header = cols
    table_data = [header]
    for _, row in df_pdf[cols].iterrows():
        table_data.append([wrap_text_status(row[c]) for c in cols])

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
    return dcc.send_bytes(buffer.getvalue(), "status_processos_paisagem.pdf")
