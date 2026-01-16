import dash
from dash import html, dcc, dash_table, Input, Output, State, callback
import pandas as pd
from io import BytesIO
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image,
)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from datetime import datetime
from pytz import timezone
import os


dash.register_page(
    __name__,
    path="/extrato-contrato",
    name="Extrato de Contrato",
    title="Extrato de Contrato",
)


URL_BI_EXTRATO = (
    "https://docs.google.com/spreadsheets/d/"
    "17nBhvSoCeK3hNgCj2S57q3pF2Uxj6iBpZDvCX481KcU/"
    "gviz/tq?tqx=out:csv&sheet=BI%20Extrato"
)

button_style = {
    "backgroundColor": "#0b2b57",
    "color": "white",
    "padding": "8px 16px",
    "border": "none",
    "borderRadius": "4px",
    "cursor": "pointer",
    "fontSize": "14px",
    "fontWeight": "bold",
}


def conv_moeda_br(v):
    if isinstance(v, str):
        v = v.strip()
        if v in ("", "-"):
            return None
        v = v.replace("R$", "").replace(" ", "")
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
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# --------------------------------------------------
# Carga e tratamento de dados
# --------------------------------------------------
def carregar_dados_extrato():
    df = pd.read_csv(URL_BI_EXTRATO, header=0)
    df.columns = [c.strip() for c in df.columns]

    base_cols = [
        "Contrato",
        "Processo",
        "Modalidade",
        "Vigência - de",
        "Vigência - até",
        "Prazo",
        "Contratada",
        "CNPJ",
        "Objeto",
        "Comprasnet",
        "Valor original",
        "Acrésc/Supressões",
        "Valor atualizado",
        "Tipo de garantia",
        "Instituição",
        "Vigência contrato",
        "Vigência + 90 dias",
        "Vigência contratada",
        "Base de cálculo",
        "Percentual",
        "Cobertura",
        "Valor contratado",
        "Valor inicial",
    ]
    for c in base_cols:
        if c not in df.columns:
            df[c] = None

    cols_moeda = [
        "Valor original",
        "Acrésc/Supressões",
        "Valor atualizado",
        "Base de cálculo",
        "Cobertura",
        "Valor contratado",
        "Valor inicial",
    ]
    for c in cols_moeda:
        if c in df.columns:
            df[c] = df[c].apply(conv_moeda_br)

    for i in range(0, 13):
        suf = "" if i == 0 else f".{i}"
        col_valor = f"Valor{suf}"
        col_valor_atualizado = f"Valor Atualizado{suf}"
        if col_valor in df.columns:
            df[col_valor] = df[col_valor].apply(conv_moeda_br)
        if col_valor_atualizado in df.columns:
            df[col_valor_atualizado] = df[col_valor_atualizado].apply(conv_moeda_br)

    cols_texto = [
        "Contrato",
        "Processo",
        "Modalidade",
        "Prazo",
        "Contratada",
        "CNPJ",
        "Objeto",
        "Comprasnet",
        "Tipo de garantia",
        "Instituição",
    ]
    for c in cols_texto:
        if c in df.columns:
            df[c] = df[c].astype("string")

    cols_data = [
        "Vigência - de",
        "Vigência - até",
        "Vigência contrato",
        "Vigência + 90 dias",
        "Vigência contratada",
    ]
    for c in cols_data:
        if c in df.columns:
            df[c] = df[c].astype("string")

    for i in range(0, 13):
        suf = "" if i == 0 else f".{i}"
        col_vig = f"Vigência{suf}"
        if col_vig in df.columns:
            df[col_vig] = df[col_vig].astype("string")

    return df


df_extrato_base = carregar_dados_extrato()

# --------------------------------------------------
# Definição das colunas de cada tabela
# --------------------------------------------------
cols_contrato_info = [
    "Processo",
    "Modalidade",
    "Vigência - de",
    "Vigência - até",
    "Prazo",
    "Contratada",
    "CNPJ",
    "Comprasnet",
]
for c in cols_contrato_info:
    if c not in df_extrato_base.columns:
        df_extrato_base[c] = None

cols_contrato_valores = [
    "Valor original",
    "Acrésc/Supressões",
    "Valor atualizado",
]
for c in cols_contrato_valores:
    if c not in df_extrato_base.columns:
        df_extrato_base[c] = None

cols_garantia = [
    "Tipo de garantia",
    "Instituição",
    "Vigência contrato",
    "Vigência + 90 dias",
    "Vigência contratada",
    "Base de cálculo",
    "Percentual",
    "Cobertura",
    "Valor contratado",
]
for c in cols_garantia:
    if c not in df_extrato_base.columns:
        df_extrato_base[c] = None


def gerar_grupo_fiscalizacao(df_local, indice: int) -> pd.DataFrame:
    if indice == 0:
        col_fisc = "Fiscalização"
        col_serv = "Servidor"
        col_fisc_subst = "Fiscalização - subst."
        col_serv_subst = "Servidor.1"
    else:
        col_fisc = f"Fiscalização.{indice}"
        col_serv = f"Servidor.{indice*2}"
        col_fisc_subst = f"Fiscalização - subst..{indice}"
        col_serv_subst = f"Servidor.{indice*2 + 1}"

    colunas_originais = [col_fisc, col_serv, col_fisc_subst, col_serv_subst]
    for c in colunas_originais:
        if c not in df_local.columns:
            df_local[c] = None

    tabela_sel = df_local[colunas_originais].copy()
    tabela_ren = tabela_sel.rename(
        columns={
            col_fisc: "Fiscalização",
            col_serv: "Servidor",
            col_fisc_subst: "Fiscalização_subst",
            col_serv_subst: "Servidor_subst",
        }
    )
    return tabela_ren


def gerar_grupo_alteracao(df_local, indice: int) -> pd.DataFrame:
    suf = "" if indice == 1 else f".{indice-1}"
    col_tipo = f"Tipo{suf}"
    col_vig = f"Vigência{suf}"
    col_valor = f"Valor{suf}"
    col_valor_at = f"Valor Atualizado{suf}"

    colunas_originais = [col_tipo, col_vig, col_valor, col_valor_at]
    for c in colunas_originais:
        if c not in df_local.columns:
            df_local[c] = None

    tabela_sel = df_local[colunas_originais].copy()
    tabela_ren = tabela_sel.rename(
        columns={
            col_tipo: "Tipo",
            col_vig: "Vigência",
            col_valor: "Valor",
            col_valor_at: "Valor Atualizado",
        }
    )
    return tabela_ren


# --------------------------------------------------
# Layout
# --------------------------------------------------
layout = html.Div(
    children=[
        html.Div(
            id="barra_filtros_extrato",
            className="filtros-sticky",
            style={
                "position": "relative",
                "zIndex": 1100,
                "backgroundColor": "white",
            },
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
                            style={"minWidth": "240px", "flex": "1 1 260px"},
                            children=[
                                html.Label("Contrato"),
                                dcc.Input(
                                    id="filtro_contrato_extrato",
                                    type="text",
                                    placeholder="Digite o número do contrato",
                                    style={
                                        "width": "100%",
                                        "marginBottom": "6px",
                                        "height": "32px",
                                    },
                                ),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    style={
                        "display": "flex",
                        "alignItems": "center",
                        "justifyContent": "space-between",
                        "marginTop": "6px",
                        "flexWrap": "wrap",
                        "gap": "10px",
                    },
                    children=[
                        html.Div(
                            children=[
                                html.Button(
                                    "Limpar filtros",
                                    id="btn_limpar_filtros_extrato",
                                    n_clicks=0,
                                    style={**button_style, "marginRight": "10px"},
                                ),
                                html.Button(
                                    "Baixar Relatório PDF",
                                    id="btn_download_relatorio_extrato",
                                    n_clicks=0,
                                    style=button_style,
                                ),
                                dcc.Download(id="download_relatorio_extrato"),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        # CARTÃO CONTRATO
        html.Div(
            id="card_numero_contrato",
            style={
                "marginTop": "10px",
                "padding": "10px 0",
                "backgroundColor": "#0b2b57",
                "border": "2px solid #0b2b57",
                "borderRadius": "4px",
                "width": "100%",
                "textAlign": "center",
                "color": "white",
            },
            children=[
                html.Span(
                    "Contrato ",
                    style={"fontSize": "14px", "marginRight": "5px"},
                ),
                html.Span(
                    id="valor_numero_contrato",
                    style={"fontSize": "22px", "fontWeight": "bold"},
                ),
            ],
        ),
        # INFORMAÇÕES
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                dash_table.DataTable(
                    id="tabela_extrato_info",
                    columns=[{"name": col, "id": col} for col in cols_contrato_info],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "280px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # VALORES (agora vem ANTES do Objeto)
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                dash_table.DataTable(
                    id="tabela_extrato_valores",
                    columns=[{"name": col, "id": col} for col in cols_contrato_valores],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "280px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # OBJETO (logo DEPOIS dos valores)
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                dash_table.DataTable(
                    id="tabela_extrato_objeto",
                    columns=[{"name": "Objeto", "id": "Objeto"}],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "280px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "8px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                        "minWidth": "100%",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # COMPRASNET – tabela abaixo do Objeto
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                dash_table.DataTable(
                    id="tabela_extrato_comprasnet",
                    columns=[
                        {
                            "name": "Comprasnet",
                            "id": "Comprasnet_link",
                            "presentation": "markdown",
                        }
                    ],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "80px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # FISCALIZAÇÃO
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                html.H4(
                    "Equipe de Fiscalização do Contrato",
                    style={
                        "textAlign": "center",
                        "backgroundColor": "#0b2b57",
                        "color": "white",
                        "padding": "6px 0",
                        "margin": "0",
                    },
                ),
                dash_table.DataTable(
                    id="tabela_extrato_fiscalizacao",
                    columns=[
                        {"name": "Fiscalização", "id": "Fiscalização"},
                        {"name": "Servidor", "id": "Servidor"},
                        {"name": "Fiscalização (subst.)", "id": "Fiscalização_subst"},
                        {"name": "Servidor (subst.)", "id": "Servidor_subst"},
                    ],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "320px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # GARANTIA
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "0px",
            },
            children=[
                html.H4(
                    "Garantia de Execução Contratual",
                    style={
                        "textAlign": "center",
                        "backgroundColor": "#0b2b57",
                        "color": "white",
                        "padding": "6px 0",
                        "marginTop": "10px",
                        "marginBottom": "0",
                    },
                ),
                dash_table.DataTable(
                    id="tabela_extrato_garantia",
                    columns=[{"name": col, "id": col} for col in cols_garantia],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "320px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        # EVOLUÇÃO
        html.Div(
            style={
                "flex": "1 1 100%",
                "minWidth": "300px",
                "position": "relative",
                "zIndex": 1,
                "marginTop": "10px",
            },
            children=[
                html.H4(
                    "Evolução do Contrato",
                    style={
                        "textAlign": "center",
                        "backgroundColor": "#0b2b57",
                        "color": "white",
                        "padding": "6px 0",
                        "marginTop": "0",
                        "marginBottom": "0",
                    },
                ),
                dash_table.DataTable(
                    id="tabela_extrato_evolucao",
                    columns=[
                        {"name": "Tipo", "id": "Tipo"},
                        {"name": "Vigência", "id": "Vigência"},
                        {"name": "Valor", "id": "Valor_fmt"},
                        {"name": "Valor Atualizado", "id": "Valor Atualizado_fmt"},
                    ],
                    data=[],
                    fixed_rows={"headers": True},
                    style_table={
                        "overflowX": "auto",
                        "overflowY": "auto",
                        "maxHeight": "360px",
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    style_cell={
                        "textAlign": "center",
                        "padding": "6px",
                        "fontSize": "12px",
                        "whiteSpace": "normal",
                    },
                    style_header={
                        "fontWeight": "bold",
                        "backgroundColor": "#d9d9d9",
                        "color": "black",
                        "textAlign": "center",
                    },
                ),
            ],
        ),
        dcc.Store(id="store_dados_extrato"),
    ]
)

# --------------------------------------------------
# Cabeçalho padrão do PDF
# --------------------------------------------------
def adicionar_cabecalho_relatorio(story, num_contrato, styles):
    logo_esq = (
        Image("assets/brasaobrasil.png", 1.0 * inch, 1.0 * inch)
        if os.path.exists("assets/brasaobrasil.png")
        else ""
    )
    logo_dir = (
        Image("assets/simbolo_RGB.png", 1.0 * inch, 1.0 * inch)
        if os.path.exists("assets/simbolo_RGB.png")
        else ""
    )

    texto_instituicao = (
        "<b><font color='#0b2b57' size=12>Ministério da Educação</font></b><br/>"
        "<b><font color='#0b2b57' size=12>Universidade Federal de Itajubá</font></b><br/>"
        "<font color='#0b2b57' size=10>Diretoria de Compras e Contratos</font>"
    )

    instituicao = Paragraph(
        texto_instituicao,
        ParagraphStyle(
            "instituicao",
            alignment=TA_CENTER,
            leading=14,
        ),
    )

    cabecalho = Table(
        [[logo_esq, instituicao, logo_dir]],
        colWidths=[1.2 * inch, 6.5 * inch, 1.2 * inch],
    )
    cabecalho.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    story.append(cabecalho)
    story.append(Spacer(1, 0.15 * inch))

    titulo = Paragraph(
        f"EXTRATO DO CONTRATO {num_contrato}",
        ParagraphStyle(
            "titulo_status",
            alignment=TA_CENTER,
            fontSize=9,
            leading=12,
            textColor=colors.black,
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.1 * inch))

    data_geracao = datetime.now(timezone("America/Sao_Paulo")).strftime(
        "%d/%m/%Y às %H:%M:%S"
    )
    story.append(
        Paragraph(
            f"<font size=8>Data de geração: {data_geracao}</font>",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 0.08 * inch))


# --------------------------------------------------
# Geração do PDF em modo PAISAGEM
# --------------------------------------------------
def gerar_pdf_relatorio_extrato(
    df_info, df_objeto, df_valores, df_fiscalizacao, df_garantia, df_evolucao, num_contrato
):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        rightMargin=0.4 * inch,
    )

    story = []
    styles = getSampleStyleSheet()

    adicionar_cabecalho_relatorio(story, num_contrato, styles)

    # Info contrato
    if not df_info.empty:
        story.append(
            Paragraph(
                "<b><font color='white' size=9>INFORMAÇÕES DO CONTRATO</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        dados = [df_info.columns.tolist()] + df_info.values.tolist()
        tabela = Table(dados, colWidths=[1.0 * inch] * len(df_info.columns))
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 8),
                    ("FONTSIZE", (0, 1), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                    ("TOPPADDING", (0, 0), (-1, 0), 4),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(tabela)

    # Objeto
    if not df_objeto.empty:
        story.append(
            Paragraph(
                "<b><font color='white' size=9>OBJETO DO CONTRATO</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        dados = [df_objeto.columns.tolist()] + df_objeto.values.tolist()
        tabela = Table(dados, colWidths=[8.9 * inch])
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 8),
                    ("FONTSIZE", (0, 1), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                    ("TOPPADDING", (0, 0), (-1, 0), 4),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(tabela)

    # Valores
    if not df_valores.empty:
        story.append(
            Paragraph(
                "<b><font color='white' size=9>VALORES DO CONTRATO</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        df_valores_fmt = df_valores.copy()
        for col in df_valores_fmt.columns:
            df_valores_fmt[col] = df_valores_fmt[col].apply(formatar_moeda)
        dados = [df_valores_fmt.columns.tolist()] + df_valores_fmt.values.tolist()
        tabela = Table(dados, colWidths=[2.8 * inch] * len(df_valores_fmt.columns))
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 8),
                    ("FONTSIZE", (0, 1), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                    ("TOPPADDING", (0, 0), (-1, 0), 4),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(tabela)

    # Fiscalização
    if not df_fiscalizacao.empty:
        story.append(
            Paragraph(
                "<b><font color='white' size=9>EQUIPE DE FISCALIZAÇÃO DO CONTRATO</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        dados = [df_fiscalizacao.columns.tolist()] + df_fiscalizacao.values.tolist()
        tabela = Table(dados, colWidths=[2.2 * inch] * len(df_fiscalizacao.columns))
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 7),
                    ("FONTSIZE", (0, 1), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
                    ("TOPPADDING", (0, 0), (-1, 0), 3),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(tabela)

    # Garantia
    if not df_garantia.empty:
        story.append(
            Paragraph(
                "<b><font color='white' size=9>GARANTIA DE EXECUÇÃO CONTRATUAL</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        df_garantia_fmt = df_garantia.copy()
        for col in ["Base de cálculo", "Cobertura", "Valor contratado"]:
            if col in df_garantia_fmt.columns:
                df_garantia_fmt[col] = df_garantia_fmt[col].apply(formatar_moeda)
        dados = [df_garantia_fmt.columns.tolist()] + df_garantia_fmt.values.tolist()
        tabela = Table(dados, colWidths=[0.95 * inch] * len(df_garantia_fmt.columns))
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 7),
                    ("FONTSIZE", (0, 1), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
                    ("TOPPADDING", (0, 0), (-1, 0), 3),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(tabela)

    # Evolução
    if not df_evolucao.empty:
        story.append(PageBreak())
        story.append(
            Paragraph(
                "<b><font color='white' size=9>EVOLUÇÃO DO CONTRATO</font></b>",
                ParagraphStyle(
                    "secao_titulo",
                    alignment=TA_CENTER,
                    backgroundColor=colors.HexColor("#0b2b57"),
                    textColor=colors.white,
                    fontSize=9,
                    leading=12,
                    spaceAfter=0,
                ),
            )
        )
        df_evolucao_fmt = df_evolucao.copy()
        for col in ["Valor", "Valor Atualizado"]:
            if col in df_evolucao_fmt.columns:
                df_evolucao_fmt[col] = df_evolucao_fmt[col].apply(formatar_moeda)
        dados = [df_evolucao_fmt.columns.tolist()] + df_evolucao_fmt.values.tolist()
        tabela = Table(dados, colWidths=[2.2 * inch] * len(df_evolucao_fmt.columns))
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 8),
                    ("FONTSIZE", (0, 1), (-1, -1), 7),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                    ("TOPPADDING", (0, 0), (-1, 0), 4),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.white, colors.HexColor("#f0f0f0")],
                    ),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(tabela)

    doc.build(story)
    buffer.seek(0)
    return buffer


# --------------------------------------------------
# Callbacks
# --------------------------------------------------
@callback(
    Output("tabela_extrato_info", "data"),
    Output("tabela_extrato_objeto", "data"),
    Output("tabela_extrato_valores", "data"),
    Output("tabela_extrato_fiscalizacao", "data"),
    Output("tabela_extrato_garantia", "data"),
    Output("tabela_extrato_evolucao", "data"),
    Output("tabela_extrato_comprasnet", "data"),
    Output("valor_numero_contrato", "children"),
    Input("filtro_contrato_extrato", "value"),
)
def atualizar_tabelas_extrato_cb(contrato):
    if not contrato:
        return [], [], [], [], [], [], [], ""
    dff = df_extrato_base[
        df_extrato_base["Contrato"].astype(str)
        .str.contains(str(contrato).strip(), case=False, na=False)
    ]
    if dff.empty:
        return [], [], [], [], [], [], [], ""

    dff_sorted = dff.copy()

    df_info = dff_sorted[cols_contrato_info].head(1)

    # VALORES -> moeda
    df_valores = dff_sorted[cols_contrato_valores].head(1).copy()
    for col in ["Valor original", "Acrésc/Supressões", "Valor atualizado"]:
        if col in df_valores.columns:
            df_valores[col] = df_valores[col].apply(formatar_moeda)

    df_objeto = dff_sorted[["Objeto"]].head(1)

    # COMPRASNET como link (Markdown)
    df_comp = dff_sorted[["Comprasnet"]].head(1).copy()
    df_comp["Comprasnet_link"] = df_comp["Comprasnet"].apply(
        lambda x: f"[{x}]({x})" if isinstance(x, str) and x.strip() else ""
    )

    df_fisc = gerar_grupo_fiscalizacao(dff_sorted, 0)

    df_garan = dff_sorted[cols_garantia].head(1).copy()
    for col in ["Base de cálculo", "Cobertura", "Valor contratado"]:
        if col in df_garan.columns:
            df_garan[col] = df_garan[col].apply(formatar_moeda)

    lista_evol = []
    for i in range(1, 13):
        df_alt = gerar_grupo_alteracao(dff_sorted, i)
        lista_evol.append(df_alt)
    df_evol_all = pd.concat(lista_evol, ignore_index=True)
    df_evol_all = df_evol_all[df_evol_all["Tipo"].astype(str).str.strip().ne("")]
    df_evol_all["Valor_fmt"] = df_evol_all["Valor"].apply(formatar_moeda)
    df_evol_all["Valor Atualizado_fmt"] = df_evol_all["Valor Atualizado"].apply(
        formatar_moeda
    )

    return (
        df_info.to_dict("records"),
        df_objeto.to_dict("records"),
        df_valores.to_dict("records"),
        df_fisc.to_dict("records"),
        df_garan.to_dict("records"),
        df_evol_all[["Tipo", "Vigência", "Valor_fmt", "Valor Atualizado_fmt"]].to_dict(
            "records"
        ),
        df_comp[["Comprasnet_link"]].to_dict("records"),
        dff_sorted["Contrato"].iloc[0],
    )


@callback(
    Output("filtro_contrato_extrato", "value"),
    Input("btn_limpar_filtros_extrato", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_extrato(n_clicks):
    return ""


@callback(
    Output("download_relatorio_extrato", "data"),
    Input("btn_download_relatorio_extrato", "n_clicks"),
    State("filtro_contrato_extrato", "value"),
    prevent_initial_call=True,
)
def download_relatorio_pdf(n_clicks, filtro_contrato):
    from dash import dcc

    if not n_clicks or not filtro_contrato:
        return dash.no_update

    dff = df_extrato_base[
        df_extrato_base["Contrato"].astype(str)
        .str.contains(str(filtro_contrato).strip(), case=False, na=False)
    ]
    if dff.empty:
        return dash.no_update

    dff_sorted = dff.copy()
    num_contrato = dff_sorted["Contrato"].iloc[0]

    df_info = dff_sorted[cols_contrato_info].head(1)
    df_objeto = dff_sorted[["Objeto"]].head(1)
    df_valores = dff_sorted[cols_contrato_valores].head(1)
    df_fisc = gerar_grupo_fiscalizacao(dff_sorted, 0)
    df_garan = dff_sorted[cols_garantia].head(1)

    lista_evol = []
    for i in range(1, 13):
        df_alt = gerar_grupo_alteracao(dff_sorted, i)
        lista_evol.append(df_alt)
    df_evol_all = pd.concat(lista_evol, ignore_index=True)
    df_evol_all = df_evol_all[df_evol_all["Tipo"].astype(str).str.strip().ne("")]

    pdf_buffer = gerar_pdf_relatorio_extrato(
        df_info, df_objeto, df_valores, df_fisc, df_garan, df_evol_all, num_contrato
    )

    return dcc.send_bytes(pdf_buffer.getvalue(), f"Extrato_{num_contrato}.pdf")
