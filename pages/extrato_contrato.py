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
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
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

# ===== FUNÇÕES AUXILIARES =====
def conv_moeda_br(v):
    """Converte string de moeda brasileira para float"""
    if isinstance(v, str):
        v = v.strip()
        if v in ("", "-"):
            return None
        v = v.replace("R$", "").replace(" ", "")
        v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def formatar_moeda(v):
    """Formata float para string de moeda brasileira"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_fiscalizacao_para_html(fiscalizacao, servidor):
    """Formata a fiscalização para HTML com negrito e quebra de linha"""
    if not fiscalizacao and not servidor:
        return ""
    
    fiscalizacao = str(fiscalizacao).strip() if fiscalizacao else ""
    servidor = str(servidor).strip() if servidor else ""
    
    # Remover "nan"
    fiscalizacao = fiscalizacao.replace("nan", "").strip()
    servidor = servidor.replace("nan", "").strip()
    
    if not fiscalizacao and not servidor:
        return ""
    
    # Formatar como HTML com negrito e centralizado
    html_content = f'<div style="text-align: center; line-height: 1.2;">'
    if fiscalizacao:
        html_content += f'<div style="font-weight: bold; font-size: 11px;">{fiscalizacao}</div>'
    if servidor:
        # Adicionar quebra de linha automática para nomes longos
        servidor_formatado = servidor.replace(" ", "&#8203; ")
        html_content += f'<div style="font-size: 10px;">{servidor_formatado}</div>'
    html_content += '</div>'
    
    return html_content

def formatar_fiscalizacao_para_pdf(fiscalizacao, servidor):
    """Formata a fiscalização para PDF com negrito e quebra de linha"""
    if not fiscalizacao and not servidor:
        return ""
    
    fiscalizacao = str(fiscalizacao).strip() if fiscalizacao else ""
    servidor = str(servidor).strip() if servidor else ""
    
    # Remover "nan"
    fiscalizacao = fiscalizacao.replace("nan", "").strip()
    servidor = servidor.replace("nan", "").strip()
    
    if not fiscalizacao and not servidor:
        return ""
    
    # Criar parágrafo com estilo
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import ParagraphStyle
    
    texto = ""
    if fiscalizacao:
        texto += f"<b>{fiscalizacao}</b><br/>"
    if servidor:
        texto += servidor
    
    estilo = ParagraphStyle(
        'fiscalizacao_pdf',
        alignment=TA_CENTER,
        fontSize=7,
        leading=9,
        spaceAfter=2,
    )
    
    return Paragraph(texto, estilo)

# ===== CARREGAMENTO E TRATAMENTO DE DADOS =====
def carregar_dados_extrato():
    """Carrega dados do Google Sheets e realiza conversões de tipos"""
    df = pd.read_csv(URL_BI_EXTRATO, header=0)
    df.columns = [c.strip() for c in df.columns]
    
    base_cols = [
        "Contrato", "Processo", "Modalidade", "Vigência - de", "Vigência - até", "Prazo",
        "Contratada", "CNPJ", "Objeto", "Comprasnet", "Valor original", "Acrésc/Supressões",
        "Valor atualizado", "Tipo de garantia", "Instituição", "Vigência contrato",
        "Vigência + 90 dias", "Vigência contratada", "Base de cálculo", "Percentual",
        "Cobertura", "Valor contratado", "Valor inicial",
    ]
    for c in base_cols:
        if c not in df.columns:
            df[c] = None

    # Conversão de colunas monetárias
    cols_moeda = [
        "Valor original", "Acrésc/Supressões", "Valor atualizado",
        "Base de cálculo", "Cobertura", "Valor contratado", "Valor inicial",
    ]
    for c in cols_moeda:
        if c in df.columns:
            df[c] = df[c].apply(conv_moeda_br)

    # Conversão de valores dinâmicos (1-12)
    for i in range(0, 13):
        suf = "" if i == 0 else f".{i}"
        col_valor = f"Valor{suf}"
        col_valor_atualizado = f"Valor Atualizado{suf}"
        if col_valor in df.columns:
            df[col_valor] = df[col_valor].apply(conv_moeda_br)
        if col_valor_atualizado in df.columns:
            df[col_valor_atualizado] = df[col_valor_atualizado].apply(conv_moeda_br)

    # Conversão para string
    cols_texto = [
        "Contrato", "Processo", "Modalidade", "Prazo", "Contratada", "CNPJ",
        "Objeto", "Comprasnet", "Tipo de garantia", "Instituição",
    ]
    for c in cols_texto:
        if c in df.columns:
            df[c] = df[c].astype("string")

    # Conversão de datas
    cols_data = [
        "Vigência - de", "Vigência - até", "Vigência contrato",
        "Vigência + 90 dias", "Vigência contratada",
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

# ===== DEFINIÇÃO DAS COLUNAS =====
cols_contrato_info = [
    "Processo", "Modalidade", "Vigência - de", "Vigência - até", 
    "Prazo", "Contratada", "CNPJ",
]
cols_contrato_valores = [
    "Valor original", "Acrésc/Supressões", "Valor atualizado",
]
cols_garantia = [
    "Tipo de garantia", "Instituição", "Vigência contrato", 
    "Vigência + 90 dias", "Vigência contratada", "Base de cálculo",
    "Percentual", "Cobertura", "Valor contratado",
]

for c in cols_contrato_info + cols_contrato_valores + cols_garantia:
    if c not in df_extrato_base.columns:
        df_extrato_base[c] = None

def gerar_grupo_fiscalizacao_otimizado(df_local, indice):
    """Gera dados de fiscalização otimizados para um índice específico"""
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
    
    # Extrair os dados
    fisc_titular = df_local[col_fisc].iloc[0] if not df_local.empty else ""
    serv_titular = df_local[col_serv].iloc[0] if not df_local.empty else ""
    fisc_subst = df_local[col_fisc_subst].iloc[0] if not df_local.empty else ""
    serv_subst = df_local[col_serv_subst].iloc[0] if not df_local.empty else ""
    
    # Retornar dicionário com dados separados
    return {
        "fiscalizacao_titular": fisc_titular,
        "servidor_titular": serv_titular,
        "fiscalizacao_substituto": fisc_subst,
        "servidor_substituto": serv_subst
    }

# ===== LAYOUT DASH =====
layout = html.Div(
    children=[
        # Barra de filtros
        html.Div(
            id="barra_filtros_extrato",
            className="filtros-sticky",
            style={
                "position": "relative",
                "zIndex": 1100,
                "backgroundColor": "white",
                "padding": "10px",
                "borderBottom": "1px solid #ddd",
            },
            children=[
                html.Div(
                    style={
                        "display": "flex",
                        "flexWrap": "wrap",
                        "gap": "15px",
                        "alignItems": "flex-start",
                    },
                    children=[
                        # Filtro por Contrato
                        html.Div(
                            style={"minWidth": "250px", "flex": "1 1 250px"},
                            children=[
                                html.Label("Contrato", style={"fontWeight": "bold", "marginBottom": "5px"}),
                                dcc.Dropdown(
                                    id="filtro_contrato_extrato",
                                    options=[
                                        {"label": str(contrato), "value": str(contrato)}
                                        for contrato in sorted(df_extrato_base["Contrato"].dropna().unique())
                                    ],
                                    placeholder="Selecione um contrato...",
                                    clearable=True,
                                    style={"width": "100%"},
                                ),
                            ],
                        ),
                        
                        # Filtro por Objeto
                        html.Div(
                            style={"minWidth": "300px", "flex": "2 1 300px"},
                            children=[
                                html.Label("Objeto", style={"fontWeight": "bold", "marginBottom": "5px"}),
                                dcc.Dropdown(
                                    id="filtro_objeto_extrato",
                                    options=[
                                        {"label": str(objeto)[:100] + "..." if len(str(objeto)) > 100 else str(objeto), 
                                         "value": str(objeto)}
                                        for objeto in sorted(df_extrato_base["Objeto"].dropna().unique())
                                    ],
                                    placeholder="Busque pelo objeto do contrato...",
                                    clearable=True,
                                    style={"width": "100%"},
                                ),
                            ],
                        ),
                    ],
                ),
                
                # Botões de ação
                html.Div(
                    style={
                        "display": "flex",
                        "alignItems": "center",
                        "justifyContent": "space-between",
                        "marginTop": "15px",
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
                        html.Div(
                            id="info_filtros",
                            style={
                                "fontSize": "12px",
                                "color": "#666",
                                "fontStyle": "italic",
                            },
                        ),
                    ],
                ),
            ],
        ),

        # Conteúdo das tabelas
        html.Div(
            style={
                "display": "flex",
                "flexDirection": "column",
                "gap": "20px",
                "marginTop": "10px",
                "padding": "0 10px",
            },
            children=[
                # Cartão Contrato
                html.Div(
                    id="card_numero_contrato",
                    style={
                        "marginTop": "10px",
                        "padding": "15px 0",
                        "backgroundColor": "#0b2b57",
                        "border": "2px solid #0b2b57",
                        "borderRadius": "6px",
                        "width": "100%",
                        "textAlign": "center",
                        "color": "white",
                    },
                    children=[
                        html.Span(
                            "Contrato ",
                            style={"fontSize": "16px", "marginRight": "5px"},
                        ),
                        html.Span(
                            id="valor_numero_contrato",
                            style={"fontSize": "24px", "fontWeight": "bold"},
                        ),
                    ],
                ),

                # INFORMAÇÕES DO CONTRATO
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "INFORMAÇÕES DO CONTRATO",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_info",
                            columns=[{"name": col, "id": col} for col in cols_contrato_info],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "8px",
                                "fontSize": "13px",
                                "whiteSpace": "normal",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                # VALORES DO CONTRATO
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "VALORES DO CONTRATO",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_valores",
                            columns=[{"name": col, "id": col} for col in cols_contrato_valores],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "8px",
                                "fontSize": "13px",
                                "whiteSpace": "normal",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                # OBJETO
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "OBJETO",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_objeto",
                            columns=[{"name": "Objeto", "id": "Objeto"}],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "left",
                                "padding": "10px",
                                "fontSize": "13px",
                                "whiteSpace": "normal",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                # COMPRASNET
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "DOCUMENTOS DA CONTRATAÇÃO DISPONÍVEIS PARA CONSULTA NO PORTAL COMPRASNET",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                                "fontSize": "14px",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_comprasnet",
                            columns=[{
                                "name": "",
                                "id": "Comprasnet_link",
                                "presentation": "markdown",
                            }],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "10px",
                                "fontSize": "13px",
                                "whiteSpace": "normal",
                                "height": "60px",
                                "verticalAlign": "middle",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                # FISCALIZAÇÃO
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "EQUIPE DE FISCALIZAÇÃO DO CONTRATO",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_fiscalizacao",
                            columns=[],  # Serão preenchidas dinamicamente
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "10px 4px",
                                "fontSize": "12px",
                                "whiteSpace": "normal",
                                "height": "70px",
                                "verticalAlign": "middle",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "textAlign": "center",
                                "padding": "10px",
                            },
                            style_data_conditional=[
                                {
                                    "if": {"row_index": 0},
                                    "backgroundColor": "#f0f0f0",
                                },
                                {
                                    "if": {"row_index": 1},
                                    "backgroundColor": "#ffffff",
                                },
                            ],
                            css=[{
                                'selector': '.dash-cell div',
                                'rule': 'text-align: center; line-height: 1.2;'
                            }]
                        ),
                    ],
                ),

                # GARANTIA
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "GARANTIA DE EXECUÇÃO CONTRATUAL",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_garantia",
                            columns=[{"name": col, "id": col} for col in cols_garantia],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "8px",
                                "fontSize": "12px",
                                "whiteSpace": "normal",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                # EVOLUÇÃO
                html.Div(
                    style={
                        "width": "100%",
                        "position": "relative",
                        "zIndex": 1,
                    },
                    children=[
                        html.H4(
                            "EVOLUÇÃO DO CONTRATO",
                            style={
                                "textAlign": "center",
                                "backgroundColor": "#0b2b57",
                                "color": "white",
                                "padding": "8px 0",
                                "margin": "0",
                                "borderRadius": "4px 4px 0 0",
                            },
                        ),
                        html.Div(
                            id="valor_original_label",
                            style={
                                "textAlign": "right",
                                "padding": "8px",
                                "backgroundColor": "#fff",
                                "border": "1px solid #ddd",
                                "borderBottom": "none",
                                "fontSize": "13px",
                                "fontWeight": "bold",
                            },
                        ),
                        dash_table.DataTable(
                            id="tabela_extrato_evolucao",
                            columns=[
                                {"name": "Alteração", "id": "Alteração"},
                                {"name": "Tipo", "id": "Tipo"},
                                {"name": "Vigência", "id": "Vigência"},
                                {"name": "Valor", "id": "Valor_fmt"},
                                {"name": "Valor Atualizado", "id": "Valor Atualizado_fmt"},
                            ],
                            data=[],
                            style_table={
                                "overflowX": "auto",
                                "width": "100%",
                                "border": "1px solid #ddd",
                                "borderRadius": "0 0 4px 4px",
                            },
                            style_cell={
                                "textAlign": "center",
                                "padding": "6px",
                                "fontSize": "12px",
                                "whiteSpace": "normal",
                            },
                            style_cell_conditional=[
                                {"if": {"column_id": "Alteração"}, "width": "15%"},
                                {"if": {"column_id": "Tipo"}, "width": "25%"},
                                {"if": {"column_id": "Vigência"}, "width": "15%"},
                                {"if": {"column_id": "Valor_fmt"}, "width": "22.5%"},
                                {"if": {"column_id": "Valor Atualizado_fmt"}, "width": "22.5%"},
                            ],
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f0f0f0",
                                "color": "black",
                                "textAlign": "center",
                            },
                        ),
                    ],
                ),

                dcc.Store(id="store_dados_extrato"),
            ],
        ),
    ]
)

# ===== FUNÇÕES AUXILIARES PARA PDF =====
def wrap_header(text):
    """Cria header com quebra de linha para PDF"""
    return Paragraph(
        f"<b>{text}</b>",
        ParagraphStyle(
            "header_pdf",
            fontSize=7,
            alignment=TA_CENTER,
            textColor=colors.black,
            leading=10,
        ),
    )

def wrap_header_azul(text):
    """Cria header com fundo azul para PDF"""
    return Paragraph(
        f"<b>{text}</b>",
        ParagraphStyle(
            "header_pdf_azul",
            fontSize=8,
            alignment=TA_CENTER,
            textColor=colors.white,
            leading=10,
        ),
    )

def wrap_data(text, align=TA_CENTER):
    """Cria célula de dados com quebra de linha para PDF"""
    if pd.isna(text) or text is None or text == "" or str(text).strip() == "":
        return Paragraph(
            "",
            ParagraphStyle(
                "data_pdf",
                fontSize=7,
                alignment=align,
                leading=9,
            ),
        )
    return Paragraph(
        str(text),
        ParagraphStyle(
            "data_pdf",
            fontSize=7,
            alignment=align,
            leading=9,
        ),
    )

def wrap_data_left(text):
    """Cria célula de dados com alinhamento à esquerda para PDF"""
    if pd.isna(text) or text is None or text == "" or str(text).strip() == "":
        return Paragraph(
            "",
            ParagraphStyle(
                "data_pdf_left",
                fontSize=7,
                alignment=TA_LEFT,
                leading=9,
            ),
        )
    return Paragraph(
        str(text),
        ParagraphStyle(
            "data_pdf_left",
            fontSize=7,
            alignment=TA_LEFT,
            leading=9,
        ),
    )

def wrap_data_right(text):
    """Cria célula de dados com alinhamento à direita para PDF"""
    if pd.isna(text) or text is None or text == "" or str(text).strip() == "":
        return Paragraph(
            "",
            ParagraphStyle(
                "data_pdf_right",
                fontSize=7,
                alignment=TA_RIGHT,
                leading=9,
            ),
        )
    return Paragraph(
        str(text),
        ParagraphStyle(
            "data_pdf_right",
            fontSize=7,
            alignment=TA_RIGHT,
            leading=9,
        ),
    )

# ===== FUNÇÕES DE PDF =====
def adicionar_cabecalho_relatorio(story, num_contrato):
    """Adiciona cabeçalho com logos e instituição ao PDF"""
    tz_brasilia = timezone("America/Sao_Paulo")
    data_hora = datetime.now(tz_brasilia).strftime("%d/%m/%Y %H:%M:%S")
    pagesize = A4

    # Data / Hora (topo direito)
    story.append(
        Table(
            [[Paragraph(
                data_hora,
                ParagraphStyle(
                    "data_topo",
                    fontSize=9,
                    alignment=TA_RIGHT,
                    textColor="#333333",
                ),
            )]],
            colWidths=[pagesize[0] - 0.6 * inch],
        )
    )
    story.append(Spacer(1, 0.1 * inch))

    # Logos e instituição
    logo_esq = (
        Image("assets/brasaobrasil.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/brasaobrasil.png") else ""
    )
    logo_dir = (
        Image("assets/simbolo_RGB.png", 1.2 * inch, 1.2 * inch)
        if os.path.exists("assets/simbolo_RGB.png") else ""
    )

    texto_instituicao = (
        "<b><font color='#0b2b57' size='13'>Ministério da Educação</font></b><br/>"
        "<b><font color='#0b2b57' size='13'>Universidade Federal de Itajubá</font></b><br/>"
        "<font color='#0b2b57' size='11'>Diretoria de Compras e Contratos</font>"
    )
    instituicao = Paragraph(
        texto_instituicao,
        ParagraphStyle(
            "instituicao",
            alignment=TA_CENTER,
            leading=16,
        ),
    )

    cabecalho = Table(
        [[logo_esq, instituicao, logo_dir]],
        colWidths=[1.4 * inch, 4.2 * inch, 1.4 * inch],
    )
    cabecalho.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ])
    )
    story.append(cabecalho)
    story.append(Spacer(1, 0.15 * inch))

    # Título com número do contrato
    titulo = Paragraph(
        f"<b>Contrato {num_contrato}</b>",
        ParagraphStyle(
            "titulo_contrato",
            alignment=TA_CENTER,
            fontSize=13,
            leading=16,
            textColor=colors.HexColor("#0b2b57"),
        ),
    )
    story.append(titulo)
    story.append(Spacer(1, 0.2 * inch))

def criar_tabela_pdf(story, titulo, dados_header, dados_linhas, colWidths, alinhamentos=None, larguras_personalizadas=None):
    """Cria tabela formatada no PDF com título azul e header cinza claro"""
    LARGURA_PADRAO = 6.7 * inch
    table_data = [dados_header] + dados_linhas
    
    # Se larguras personalizadas fornecidas, use-as
    if larguras_personalizadas:
        # Normalizar larguras para somar LARGURA_PADRAO
        total_especificado = sum(larguras_personalizadas)
        colWidths = [largura/total_especificado * LARGURA_PADRAO for largura in larguras_personalizadas]
    
    tbl = Table(table_data, colWidths=colWidths, repeatRows=1)
    
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
    ]
    
    # Aplicar alinhamentos personalizados se fornecidos
    if alinhamentos:
        for col_idx, align in enumerate(alinhamentos):
            if align == "LEFT":
                table_styles.append(("ALIGN", (col_idx, 1), (col_idx, -1), "LEFT"))
            elif align == "RIGHT":
                table_styles.append(("ALIGN", (col_idx, 1), (col_idx, -1), "RIGHT"))
    
    tbl.setStyle(TableStyle(table_styles))

    # Título com mesma largura da tabela
    titulo_table = Table(
        [[Paragraph(
            f"<b>{titulo}</b>",
            ParagraphStyle(
                "titulo_secao",
                alignment=TA_CENTER,
                fontSize=9,
                textColor=colors.white,
                leading=10,
            ),
        )]],
        colWidths=[LARGURA_PADRAO],
    )
    titulo_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))

    story.append(titulo_table)
    story.append(tbl)

def criar_tabela_com_header_azul(story, titulo, dados_header, dados_linhas, colWidths, alinhamentos=None):
    """Cria tabela com header azul (para Objeto e ComprasNet)"""
    LARGURA_PADRAO = 6.7 * inch
    table_data = [dados_header] + dados_linhas
    
    tbl = Table(table_data, colWidths=colWidths, repeatRows=1)
    
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
    ]
    
    # Aplicar alinhamentos personalizados se fornecidos
    if alinhamentos:
        for col_idx, align in enumerate(alinhamentos):
            if align == "LEFT":
                table_styles.append(("ALIGN", (col_idx, 1), (col_idx, -1), "LEFT"))
            elif align == "RIGHT":
                table_styles.append(("ALIGN", (col_idx, 1), (col_idx, -1), "RIGHT"))
    
    tbl.setStyle(TableStyle(table_styles))
    
    story.append(tbl)

def criar_tabela_fiscalizacao_pdf(story, dados_equipes):
    """Cria tabela de fiscalização no PDF no formato otimizado"""
    if not dados_equipes:
        return
    
    LARGURA_PADRAO = 6.7 * inch
    num_equipes = len(dados_equipes)
    
    # Calcular largura das colunas proporcionalmente
    col_width = LARGURA_PADRAO / num_equipes if num_equipes > 0 else LARGURA_PADRAO
    
    # Criar cabeçalhos vazios
    headers = [wrap_header("") for _ in range(num_equipes)]
    
    # Criar linhas de dados
    linhas = []
    # Linha 1: Titulares (fundo cinza)
    linha_titular = []
    for i in range(num_equipes):
        paragraph = formatar_fiscalizacao_para_pdf(
            dados_equipes[i]["fiscalizacao_titular"],
            dados_equipes[i]["servidor_titular"]
        )
        linha_titular.append(paragraph if paragraph else wrap_data(""))
    linhas.append(linha_titular)
    
    # Linha 2: Substitutos (fundo branco)
    linha_substituto = []
    for i in range(num_equipes):
        paragraph = formatar_fiscalizacao_para_pdf(
            dados_equipes[i]["fiscalizacao_substituto"],
            dados_equipes[i]["servidor_substituto"]
        )
        linha_substituto.append(paragraph if paragraph else wrap_data(""))
    linhas.append(linha_substituto)
    
    # Criar tabela
    table_data = [headers] + linhas
    tbl = Table(table_data, colWidths=[col_width] * num_equipes, repeatRows=1)
    
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#f0f0f0")),  # Titular - fundo cinza
        ("BACKGROUND", (0, 2), (-1, 2), colors.white),  # Substituto - fundo branco
    ]
    
    tbl.setStyle(TableStyle(table_styles))

    # Título
    titulo_table = Table(
        [[Paragraph(
            "<b>EQUIPE DE FISCALIZAÇÃO DO CONTRATO</b>",
            ParagraphStyle(
                "titulo_secao",
                alignment=TA_CENTER,
                fontSize=9,
                textColor=colors.white,
                leading=10,
            ),
        )]],
        colWidths=[LARGURA_PADRAO],
    )
    titulo_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))

    story.append(titulo_table)
    story.append(tbl)

def gerar_pdf_relatorio_extrato(
    df_info, df_objeto, df_valores, df_fiscalizacao, df_garantia, df_evolucao, num_contrato, df_comprasnet, equipes_dados
):
    """Gera PDF do relatório em modo retrato"""
    buffer = BytesIO()
    pagesize = A4
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.4 * inch,
        leftMargin=0.4 * inch,
        topMargin=0.3 * inch,
        bottomMargin=0.4 * inch,
    )
    story = []

    adicionar_cabecalho_relatorio(story, num_contrato)

    # INFORMAÇÕES DO CONTRATO - COM LARGURAS PERSONALIZADAS
    if not df_info.empty:
        header = [wrap_header(col) for col in df_info.columns]
        linhas = []
        for _, row in df_info.iterrows():
            linha = [wrap_data(row[col]) for col in df_info.columns]
            linhas.append(linha)
        
        # LARGURAS PERSONALIZADAS: Processo menor, CNPJ maior
        # Valores em proporções relativas
        # 7 colunas: Processo, Modalidade, Vigência - de, Vigência - até, Prazo, Contratada, CNPJ
        larguras_personalizadas = [1.2, 1.1, 0.6, 0.6, 0.5, 1.4, 1.0]  # Proporções relativas (CNPJ maior)
        col_widths = None  # Será calculado automaticamente
        
        criar_tabela_pdf(story, "INFORMAÇÕES DO CONTRATO", header, linhas, col_widths, larguras_personalizadas=larguras_personalizadas)

    # VALORES DO CONTRATO
    if not df_valores.empty:
        df_valores_fmt = df_valores.copy()
        for col in df_valores_fmt.columns:
            df_valores_fmt[col] = df_valores_fmt[col].apply(formatar_moeda)
        header = [wrap_header(col) for col in df_valores_fmt.columns]
        linhas = []
        for _, row in df_valores_fmt.iterrows():
            linha = [wrap_data(row[col], TA_CENTER) for col in df_valores_fmt.columns]
            linhas.append(linha)
        col_widths = [2.233 * inch] * 3
        criar_tabela_pdf(story, "VALORES DO CONTRATO", header, linhas, col_widths)

    # OBJETO - COM HEADER AZUL
    if not df_objeto.empty:
        objeto_text = df_objeto["Objeto"].iloc[0] if not df_objeto.empty else ""
        
        # Header azul com "OBJETO"
        header_objeto = [wrap_header_azul("OBJETO")]
        linhas_obj = [[wrap_data_left(objeto_text)]]  # Usando wrap_data_left em vez de wrap_data_justify
        
        col_widths_obj = [6.7 * inch]
        
        criar_tabela_com_header_azul(story, "", header_objeto, linhas_obj, col_widths_obj, alinhamentos=["LEFT"])
        
        story.append(Spacer(1, 0.1 * inch))

    # DOCUMENTOS COMPRASNET - COM HEADER AZUL
    comprasnet_url = df_comprasnet["Comprasnet"].iloc[0] if not df_comprasnet.empty else ""
    
    if comprasnet_url:
        # Header azul com o texto completo
        header_comprasnet = [wrap_header_azul("DOCUMENTOS DA CONTRATAÇÃO DISPONÍVEIS PARA CONSULTA NO PORTAL COMPRASNET")]
        link_paragraph = Paragraph(
            f"<a href='{comprasnet_url}'>{comprasnet_url}</a>",
            ParagraphStyle("link_pdf", fontSize=8, alignment=TA_CENTER, textColor=colors.blue, leading=10)
        )
        linhas_comprasnet = [[link_paragraph]]
        
        col_widths_comprasnet = [6.7 * inch]
        
        criar_tabela_com_header_azul(story, "", header_comprasnet, linhas_comprasnet, col_widths_comprasnet, alinhamentos=["CENTER"])
        
        story.append(Spacer(1, 0.1 * inch))

    # FISCALIZAÇÃO - NOVO FORMATO OTIMIZADO
    if equipes_dados:
        criar_tabela_fiscalizacao_pdf(story, equipes_dados)

    # GARANTIA
    if not df_garantia.empty:
        df_garantia_fmt = df_garantia.copy()
        for col in ["Base de cálculo", "Cobertura", "Valor contratado"]:
            if col in df_garantia_fmt.columns:
                df_garantia_fmt[col] = df_garantia_fmt[col].apply(formatar_moeda)
        header = [wrap_header(col) for col in df_garantia_fmt.columns]
        linhas = []
        for _, row in df_garantia_fmt.iterrows():
            linha = [wrap_data(row[col]) for col in df_garantia_fmt.columns]
            linhas.append(linha)
        col_widths = [0.744 * inch] * 9
        criar_tabela_pdf(story, "GARANTIA DE EXECUÇÃO CONTRATUAL", header, linhas, col_widths)

    # EVOLUÇÃO DO CONTRATO - COM VALOR ORIGINAL À DIREITA
    if not df_evolucao.empty:
        df_evolucao_fmt = df_evolucao.copy()
        for col in ["Valor", "Valor Atualizado"]:
            if col in df_evolucao_fmt.columns:
                df_evolucao_fmt[col] = df_evolucao_fmt[col].apply(formatar_moeda)
        
        # Incluir as colunas Alteração e Tipo se existirem
        cols_to_display = []
        if "Alteração" in df_evolucao_fmt.columns and "Tipo" in df_evolucao_fmt.columns:
            cols_to_display = ["Alteração", "Tipo", "Vigência", "Valor", "Valor Atualizado"]
        else:
            cols_to_display = ["Tipo", "Vigência", "Valor", "Valor Atualizado"]
        
        header = [wrap_header(col) for col in cols_to_display]
        linhas = []
        for _, row in df_evolucao_fmt.iterrows():
            linha = [wrap_data(row[col]) for col in cols_to_display]
            linhas.append(linha)
        
        # LARGURAS PROPORCIONAIS
        if "Alteração" in df_evolucao_fmt.columns and "Tipo" in df_evolucao_fmt.columns:
            # 5 colunas: 15%, 30%, 15%, 20%, 20%
            col_widths = [1.005 * inch, 2.01 * inch, 1.005 * inch, 1.34 * inch, 1.34 * inch]
            alinhamentos = ["CENTER", "CENTER", "CENTER", "RIGHT", "RIGHT"]
        else:
            # 4 colunas: 30%, 20%, 25%, 25%
            col_widths = [2.01 * inch, 1.34 * inch, 1.675 * inch, 1.675 * inch]
            alinhamentos = ["CENTER", "CENTER", "RIGHT", "RIGHT"]
        
        # Adicionar título e linha com "Valor original:" alinhado à direita
        titulo_evolucao_texto = f"EVOLUÇÃO DO CONTRATO"
        titulo_table_evol = Table(
            [[
                Paragraph(
                    f"<b>{titulo_evolucao_texto}</b>",
                    ParagraphStyle(
                        "titulo_secao",
                        alignment=TA_CENTER,
                        fontSize=9,
                        textColor=colors.white,
                        leading=10,
                    ),
                )
            ]],
            colWidths=[6.7 * inch],
        )
        titulo_table_evol.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b2b57")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, 0), 4),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ]))
        story.append(titulo_table_evol)
        
        # Adicionar linha com "Valor original:" alinhado à DIREITA
        valor_original = df_valores["Valor original"].iloc[0] if not df_valores.empty else ""
        valor_original_fmt = formatar_moeda(valor_original) if valor_original else ""
        
        linha_valor_original = Table(
            [[
                Paragraph(
                    f"<b>Valor original:</b> {valor_original_fmt}",
                    ParagraphStyle(
                        "valor_original",
                        alignment=TA_RIGHT,
                        fontSize=8,
                        leading=10,
                    ),
                )
            ]],
            colWidths=[6.7 * inch],
        )
        linha_valor_original.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.white),
            ("TOPPADDING", (0, 0), (-1, 0), 3),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
            ("LEFTPADDING", (0, 0), (-1, 0), 4),
            ("RIGHTPADDING", (0, 0), (-1, 0), 4),
        ]))
        story.append(linha_valor_original)
        
        # Adicionar a tabela de evolução
        tbl_evol = Table([header] + linhas, colWidths=col_widths, repeatRows=1)
        table_styles_evol = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9d9d9")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
        ]
        
        # Aplicar alinhamentos personalizados
        for col_idx, align in enumerate(alinhamentos):
            if align == "LEFT":
                table_styles_evol.append(("ALIGN", (col_idx, 1), (col_idx, -1), "LEFT"))
            elif align == "RIGHT":
                table_styles_evol.append(("ALIGN", (col_idx, 1), (col_idx, -1), "RIGHT"))
        
        tbl_evol.setStyle(TableStyle(table_styles_evol))
        story.append(tbl_evol)

    doc.build(story)
    buffer.seek(0)
    return buffer

# ===== CALLBACKS =====
@callback(
    Output("filtro_contrato_extrato", "options"),
    Output("filtro_objeto_extrato", "options"),
    Input("filtro_contrato_extrato", "value"),
    Input("filtro_objeto_extrato", "value"),
)
def atualizar_filtros_cascata(contrato_selecionado, objeto_selecionado):
    """Atualiza os filtros em cascata"""
    df_filtrado = df_extrato_base.copy()
    
    # Aplicar filtros em cascata
    if contrato_selecionado:
        df_filtrado = df_filtrado[df_filtrado["Contrato"] == contrato_selecionado]
    
    if objeto_selecionado:
        df_filtrado = df_filtrado[df_filtrado["Objeto"] == objeto_selecionado]
    
    # Gerar opções para Contrato
    opcoes_contrato = [
        {"label": str(contrato), "value": str(contrato)}
        for contrato in sorted(df_filtrado["Contrato"].dropna().unique())
    ]
    
    # Gerar opções para Objeto (com truncamento para nomes muito longos)
    opcoes_objeto = []
    for objeto in sorted(df_filtrado["Objeto"].dropna().unique()):
        objeto_str = str(objeto)
        label = objeto_str[:100] + "..." if len(objeto_str) > 100 else objeto_str
        opcoes_objeto.append({"label": label, "value": objeto_str})
    
    return opcoes_contrato, opcoes_objeto

@callback(
    Output("info_filtros", "children"),
    Input("filtro_contrato_extrato", "value"),
    Input("filtro_objeto_extrato", "value"),
)
def atualizar_info_filtros(contrato, objeto):
    """Atualiza a informação sobre filtros aplicados"""
    filtros = []
    if contrato:
        filtros.append(f"Contrato: {contrato}")
    if objeto:
        # Truncar objeto muito longo
        objeto_display = objeto[:50] + "..." if len(objeto) > 50 else objeto
        filtros.append(f"Objeto: {objeto_display}")
    
    if filtros:
        return f"Filtros aplicados: {' | '.join(filtros)}"
    return "Nenhum filtro aplicado"

@callback(
    Output("tabela_extrato_info", "data"),
    Output("tabela_extrato_objeto", "data"),
    Output("tabela_extrato_valores", "data"),
    Output("tabela_extrato_fiscalizacao", "data"),
    Output("tabela_extrato_fiscalizacao", "columns"),
    Output("tabela_extrato_garantia", "data"),
    Output("tabela_extrato_evolucao", "data"),
    Output("valor_original_label", "children"),
    Output("tabela_extrato_comprasnet", "data"),
    Output("valor_numero_contrato", "children"),
    Input("filtro_contrato_extrato", "value"),
    Input("filtro_objeto_extrato", "value"),
)
def atualizar_tabelas_extrato_cb(contrato, objeto):
    """Callback principal: atualiza todas as tabelas ao filtrar"""
    # Se não houver nenhum filtro selecionado, limpar tudo
    if not contrato and not objeto:
        return [], [], [], [], [], [], [], "", [], ""
    
    # Aplicar filtros
    dff = df_extrato_base.copy()
    
    if contrato:
        dff = dff[dff["Contrato"] == contrato]
    
    if objeto:
        dff = dff[dff["Objeto"] == objeto]
    
    if dff.empty:
        return [], [], [], [], [], [], [], "", [], ""
    
    dff_sorted = dff.copy()
    
    # Pegar o primeiro contrato se houver múltiplos
    contrato_display = dff_sorted["Contrato"].iloc[0]
    
    df_info = dff_sorted[cols_contrato_info].head(1)
    
    # VALORES formatados
    df_valores = dff_sorted[cols_contrato_valores].head(1).copy()
    for col in ["Valor original", "Acrésc/Supressões", "Valor atualizado"]:
        if col in df_valores.columns:
            df_valores[col] = df_valores[col].apply(lambda x: formatar_moeda(x))
    
    # Valor original para o label
    valor_original_raw = dff_sorted["Valor original"].iloc[0] if not dff_sorted.empty else None
    valor_original_fmt = formatar_moeda(valor_original_raw) if valor_original_raw else ""
    
    df_objeto = dff_sorted[["Objeto"]].head(1).copy()
    
    # COMPRASNET como link Markdown
    df_comp = dff_sorted[["Comprasnet"]].head(1).copy()
    df_comp["Comprasnet_link"] = df_comp["Comprasnet"].apply(
        lambda x: f"[{x}]({x})" if isinstance(x, str) and x.strip() else ""
    )
    
    # FISCALIZAÇÃO: NOVO FORMATO OTIMIZADO - COLUNAS DINÂMICAS
    equipes_dados = []
    for i in range(10):  # Verificar até 10 equipes
        dados_equipe = gerar_grupo_fiscalizacao_otimizado(dff_sorted, i)
        # Verificar se há dados na equipe
        if (dados_equipe["fiscalizacao_titular"] or dados_equipe["servidor_titular"] or 
            dados_equipe["fiscalizacao_substituto"] or dados_equipe["servidor_substituto"]):
            equipes_dados.append(dados_equipe)
    
    # Criar DataFrame para a tabela - 2 linhas (titular e substituto)
    if equipes_dados:
        dados_finais = []
        colunas_tabela = []
        
        # Para cada linha (titular e substituto)
        for linha_idx in range(2):
            linha_dict = {}
            for col_idx in range(len(equipes_dados)):
                col_id = f"Equipe_{col_idx}"
                if linha_idx == 0:  # Titular
                    html_content = formatar_fiscalizacao_para_html(
                        equipes_dados[col_idx]["fiscalizacao_titular"],
                        equipes_dados[col_idx]["servidor_titular"]
                    )
                    linha_dict[col_id] = html_content
                else:  # Substituto
                    html_content = formatar_fiscalizacao_para_html(
                        equipes_dados[col_idx]["fiscalizacao_substituto"],
                        equipes_dados[col_idx]["servidor_substituto"]
                    )
                    linha_dict[col_id] = html_content
            dados_finais.append(linha_dict)
        
        # Criar colunas com nome vazio
        for col_idx in range(len(equipes_dados)):
            colunas_tabela.append({
                "name": "",  # NOME VAZIO
                "id": f"Equipe_{col_idx}",
            })
        
        df_fisc = pd.DataFrame(dados_finais)
    else:
        df_fisc = pd.DataFrame()
        colunas_tabela = []
    
    df_garan = dff_sorted[cols_garantia].head(1).copy()
    for col in ["Base de cálculo", "Cobertura", "Valor contratado"]:
        if col in df_garan.columns:
            df_garan[col] = df_garan[col].apply(lambda x: formatar_moeda(x))
    
    # EVOLUÇÃO: SEPARAR EM DUAS COLUNAS (ALTERAÇÃO E TIPO)
    lista_evol = []
    
    # Mapeamento CORRETO baseado nos dados reais
    alteracoes = [
        {"nome": "1ª Alteração", "alt_col": "1ª Alteração", "tipo_col": "Tipo", "vig_col": "Vigência", "valor_col": "Valor", "valor_at_col": "Valor Atualizado"},
        {"nome": "2ª Alteração", "alt_col": "2ª Alteração", "tipo_col": "Tipo.1", "vig_col": "Vigência.1", "valor_col": "Valor.1", "valor_at_col": "Valor Atualizado.1"},
        {"nome": "3ª Alteração", "alt_col": "3ª Alteração", "tipo_col": "Tipo.2", "vig_col": "Vigência.2", "valor_col": "Valor.2", "valor_at_col": "Valor Atualizado.2"},
        {"nome": "4ª Alteração", "alt_col": "4ª Alteração", "tipo_col": "Tipo.3", "vig_col": "Vigência.3", "valor_col": "Valor.3", "valor_at_col": "Valor Atualizado.3"},
        {"nome": "5ª Alteração", "alt_col": "5ª Alteração", "tipo_col": "Tipo.4", "vig_col": "Vigência.4", "valor_col": "Valor.4", "valor_at_col": "Valor Atualizado.4"},
        {"nome": "6ª Alteração", "alt_col": "6ª Alteração", "tipo_col": "Tipo.5", "vig_col": "Vigência.5", "valor_col": "Valor.5", "valor_at_col": "Valor Atualizado.5"},
        {"nome": "7ª Alteração", "alt_col": "7ª Alteração", "tipo_col": "Tipo.6", "vig_col": "Vigência.6", "valor_col": "Valor.6", "valor_at_col": "Valor Atualizado.6"},
        {"nome": "8ª Alteração", "alt_col": "8ª Alteração", "tipo_col": "Tipo.7", "vig_col": "Vigência.7", "valor_col": "Valor.7", "valor_at_col": "Valor Atualizado.7"},
        {"nome": "9ª Alteração", "alt_col": "9ª Alteração", "tipo_col": "Tipo.8", "vig_col": "Vigência.8", "valor_col": "Valor.8", "valor_at_col": "Valor Atualizado.8"},
        {"nome": "10ª Alteração", "alt_col": "10ª Alteração", "tipo_col": "Tipo.9", "vig_col": "Vigência.9", "valor_col": "Valor.9", "valor_at_col": "Valor Atualizado.9"},
        {"nome": "11ª Alteração", "alt_col": "11ª Alteração", "tipo_col": "Tipo.10", "vig_col": "Vigência.10", "valor_col": "Valor.10", "valor_at_col": "Valor Atualizado.10"},
        {"nome": "12ª Alteração", "alt_col": "12ª Alteração", "tipo_col": "Tipo.11", "vig_col": "Vigência.11", "valor_col": "Valor.11", "valor_at_col": "Valor Atualizado.11"},
    ]
    
    for alt in alteracoes:
        alteracao_valor = dff_sorted[alt["alt_col"]].iloc[0] if alt["alt_col"] in dff_sorted.columns and not pd.isna(dff_sorted[alt["alt_col"]].iloc[0]) else ""
        tipo_valor = dff_sorted[alt["tipo_col"]].iloc[0] if alt["tipo_col"] in dff_sorted.columns else ""
        vig_valor = dff_sorted[alt["vig_col"]].iloc[0] if alt["vig_col"] in dff_sorted.columns else ""
        valor_valor = dff_sorted[alt["valor_col"]].iloc[0] if alt["valor_col"] in dff_sorted.columns else ""
        valor_at_valor = dff_sorted[alt["valor_at_col"]].iloc[0] if alt["valor_at_col"] in dff_sorted.columns else ""
        
        has_data = any([
            pd.notna(alteracao_valor) and str(alteracao_valor).strip() != "",
            pd.notna(tipo_valor) and str(tipo_valor).strip() != "",
            pd.notna(vig_valor) and str(vig_valor).strip() != "",
            pd.notna(valor_valor) and str(valor_valor).strip() != "",
            pd.notna(valor_at_valor) and str(valor_at_valor).strip() != ""
        ])
        
        if has_data:
            df_alt = pd.DataFrame({
                "Alteração": [alteracao_valor],
                "Tipo": [tipo_valor],
                "Vigência": [vig_valor],
                "Valor": [valor_valor],
                "Valor Atualizado": [valor_at_valor]
            })
            lista_evol.append(df_alt)
    
    # Concatenar apenas se houver dados
    if lista_evol:
        df_evol_all = pd.concat(lista_evol, ignore_index=True)
        df_evol_all["Valor_fmt"] = df_evol_all["Valor"].apply(lambda x: formatar_moeda(x))
        df_evol_all["Valor Atualizado_fmt"] = df_evol_all["Valor Atualizado"].apply(lambda x: formatar_moeda(x))
        df_evol_display = df_evol_all[["Alteração", "Tipo", "Vigência", "Valor_fmt", "Valor Atualizado_fmt"]].copy()
    else:
        df_evol_display = pd.DataFrame()

    return (
        df_info.to_dict("records"),
        df_objeto.to_dict("records"),
        df_valores.to_dict("records"),
        df_fisc.to_dict("records"),
        colunas_tabela,
        df_garan.to_dict("records"),
        df_evol_display.to_dict("records"),
        f"Valor original: {valor_original_fmt}",
        df_comp[["Comprasnet_link"]].to_dict("records"),
        str(contrato_display),
    )

@callback(
    Output("filtro_contrato_extrato", "value", allow_duplicate=True),
    Output("filtro_objeto_extrato", "value", allow_duplicate=True),
    Input("btn_limpar_filtros_extrato", "n_clicks"),
    prevent_initial_call=True,
)
def limpar_filtros_extrato(n_clicks):
    """Limpa os filtros de contrato e objeto"""
    return "", ""

@callback(
    Output("download_relatorio_extrato", "data"),
    Input("btn_download_relatorio_extrato", "n_clicks"),
    State("filtro_contrato_extrato", "value"),
    State("filtro_objeto_extrato", "value"),
    prevent_initial_call=True,
)
def download_relatorio_pdf(n_clicks, filtro_contrato, filtro_objeto):
    """Gera e baixa o PDF do relatório"""
    from dash import dcc
    import dash
    
    if not n_clicks or (not filtro_contrato and not filtro_objeto):
        return dash.no_update
    
    # Aplicar os mesmos filtros usados na visualização
    dff = df_extrato_base.copy()
    
    if filtro_contrato:
        dff = dff[dff["Contrato"] == filtro_contrato]
    
    if filtro_objeto:
        dff = dff[dff["Objeto"] == filtro_objeto]
    
    if dff.empty:
        return dash.no_update

    dff_sorted = dff.copy()
    num_contrato = dff_sorted["Contrato"].iloc[0]
    
    df_info = dff_sorted[cols_contrato_info].head(1)
    df_objeto = dff_sorted[["Objeto"]].head(1)
    df_comprasnet = dff_sorted[["Comprasnet"]].head(1)
    df_valores = dff_sorted[cols_contrato_valores].head(1)
    
    # FISCALIZAÇÃO: Preparar dados no formato otimizado para PDF
    equipes_dados_pdf = []
    for i in range(10):  # Verificar até 10 equipes
        dados_equipe = gerar_grupo_fiscalizacao_otimizado(dff_sorted, i)
        # Verificar se há dados na equipe
        if (dados_equipe["fiscalizacao_titular"] or dados_equipe["servidor_titular"] or 
            dados_equipe["fiscalizacao_substituto"] or dados_equipe["servidor_substituto"]):
            equipes_dados_pdf.append(dados_equipe)
    
    # Para o PDF, precisamos de um DataFrame vazio (os dados serão processados pela função específica)
    df_fisc = pd.DataFrame()
    
    df_garan = dff_sorted[cols_garantia].head(1)
    
    # Evolução: SEPARAR EM DUAS COLUNAS (ALTERAÇÃO E TIPO)
    lista_evol = []
    
    alteracoes = [
        {"nome": "1ª Alteração", "alt_col": "1ª Alteração", "tipo_col": "Tipo", "vig_col": "Vigência", "valor_col": "Valor", "valor_at_col": "Valor Atualizado"},
        {"nome": "2ª Alteração", "alt_col": "2ª Alteração", "tipo_col": "Tipo.1", "vig_col": "Vigência.1", "valor_col": "Valor.1", "valor_at_col": "Valor Atualizado.1"},
        {"nome": "3ª Alteração", "alt_col": "3ª Alteração", "tipo_col": "Tipo.2", "vig_col": "Vigência.2", "valor_col": "Valor.2", "valor_at_col": "Valor Atualizado.2"},
        {"nome": "4ª Alteração", "alt_col": "4ª Alteração", "tipo_col": "Tipo.3", "vig_col": "Vigência.3", "valor_col": "Valor.3", "valor_at_col": "Valor Atualizado.3"},
        {"nome": "5ª Alteração", "alt_col": "5ª Alteração", "tipo_col": "Tipo.4", "vig_col": "Vigência.4", "valor_col": "Valor.4", "valor_at_col": "Valor Atualizado.4"},
        {"nome": "6ª Alteração", "alt_col": "6ª Alteração", "tipo_col": "Tipo.5", "vig_col": "Vigência.5", "valor_col": "Valor.5", "valor_at_col": "Valor Atualizado.5"},
        {"nome": "7ª Alteração", "alt_col": "7ª Alteração", "tipo_col": "Tipo.6", "vig_col": "Vigência.6", "valor_col": "Valor.6", "valor_at_col": "Valor Atualizado.6"},
        {"nome": "8ª Alteração", "alt_col": "8ª Alteração", "tipo_col": "Tipo.7", "vig_col": "Vigência.7", "valor_col": "Valor.7", "valor_at_col": "Valor Atualizado.7"},
        {"nome": "9ª Alteração", "alt_col": "9ª Alteração", "tipo_col": "Tipo.8", "vig_col": "Vigência.8", "valor_col": "Valor.8", "valor_at_col": "Valor Atualizado.8"},
        {"nome": "10ª Alteração", "alt_col": "10ª Alteração", "tipo_col": "Tipo.9", "vig_col": "Vigência.9", "valor_col": "Valor.9", "valor_at_col": "Valor Atualizado.9"},
        {"nome": "11ª Alteração", "alt_col": "11ª Alteração", "tipo_col": "Tipo.10", "vig_col": "Vigência.10", "valor_col": "Valor.10", "valor_at_col": "Valor Atualizado.10"},
        {"nome": "12ª Alteração", "alt_col": "12ª Alteração", "tipo_col": "Tipo.11", "vig_col": "Vigência.11", "valor_col": "Valor.11", "valor_at_col": "Valor Atualizado.11"},
    ]
    
    for alt in alteracoes:
        alteracao_valor = dff_sorted[alt["alt_col"]].iloc[0] if alt["alt_col"] in dff_sorted.columns and not pd.isna(dff_sorted[alt["alt_col"]].iloc[0]) else ""
        tipo_valor = dff_sorted[alt["tipo_col"]].iloc[0] if alt["tipo_col"] in dff_sorted.columns else ""
        vig_valor = dff_sorted[alt["vig_col"]].iloc[0] if alt["vig_col"] in dff_sorted.columns else ""
        valor_valor = dff_sorted[alt["valor_col"]].iloc[0] if alt["valor_col"] in dff_sorted.columns else ""
        valor_at_valor = dff_sorted[alt["valor_at_col"]].iloc[0] if alt["valor_at_col"] in dff_sorted.columns else ""
        
        has_data = any([
            pd.notna(alteracao_valor) and str(alteracao_valor).strip() != "",
            pd.notna(tipo_valor) and str(tipo_valor).strip() != "",
            pd.notna(vig_valor) and str(vig_valor).strip() != "",
            pd.notna(valor_valor) and str(valor_valor).strip() != "",
            pd.notna(valor_at_valor) and str(valor_at_valor).strip() != ""
        ])
        
        if has_data:
            df_alt = pd.DataFrame({
                "Alteração": [alteracao_valor],
                "Tipo": [tipo_valor],
                "Vigência": [vig_valor],
                "Valor": [valor_valor],
                "Valor Atualizado": [valor_at_valor]
            })
            lista_evol.append(df_alt)
    
    if lista_evol:
        df_evol_all = pd.concat(lista_evol, ignore_index=True)
    else:
        df_evol_all = pd.DataFrame()

    pdf_buffer = gerar_pdf_relatorio_extrato(
        df_info, df_objeto, df_valores, df_fisc, df_garan, df_evol_all, num_contrato, df_comprasnet, equipes_dados_pdf
    )
    
    return dcc.send_bytes(pdf_buffer.getvalue(), f"Extrato_{num_contrato}.pdf")