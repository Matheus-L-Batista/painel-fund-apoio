import dash
from dash import Dash, html, dcc

app = Dash(
    __name__,
    use_pages=True,
    suppress_callback_exceptions=True,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
)
server = app.server

menu_links = [
    {"label": "Processos de Compras", "href": "/processos-de-compras"},
    {"label": "Status do Processo", "href": "/statusdoprocesso"}, 
    {"label": "Fracionamento de Despesas CATSER", "href": "/fracionamento_catser"},
    {"label": "Fracionamento de Despesas PDM", "href": "/fracionamento_pdm"},
    {"label": "Portarias Agente de Compras/ Contratos Tipo Empenho", "href": "/portarias_agentedecompras"},
    {"label": "Portarias de Planejamento da Contratação", "href": "/portarias_planejamento"},
    {"label": "Contratos", "href": "/contratos"},
    {"label": "Fiscais", "href": "/fiscais"}, 
    {"label": "Plano de Contratação Anual", "href": "/pca"},
    {"label": "Controle de Atas", "href": "/atas"}, 
    #{"label": "Consultar tabela", "href": "/consultartabelas"},
]

app.layout = html.Div(
    className="app-root",
    children=[
        dcc.Location(id="url"),

        dcc.Interval(
            id="interval-atualizacao",
            interval=60 * 60 * 1000,
            n_intervals=0,
        ),

        html.Div(
            className="app-container",
            children=[
                # SIDEBAR
                html.Div(
                    className="sidebar",
                    children=[
                        html.Div(
                            className="sidebar-header",
                            children=[
                                html.Img(
                                    src="/assets/logo_unifei.png",
                                    className="sidebar-logo",
                                ),
                                html.H2(
                                    "Painéis",
                                    className="sidebar-title",
                                ),
                            ],
                        ),
                        html.Div(
                            id="sidebar-menu",
                            className="sidebar-menu",
                        ),
                    ],
                ),

                # CONTEÚDO PRINCIPAL
                html.Div(
                    className="main-content",
                    children=html.Div(
                        className="page-wrapper",
                        children=dash.page_container,
                    ),
                ),
            ],
        ),
    ],
)

@app.callback(
    dash.Output("sidebar-menu", "children"),
    dash.Input("url", "pathname"),
)
def atualizar_menu(pathname):
    itens = []
    for m in menu_links:
        class_name = (
            "sidebar-button sidebar-button-active"
            if pathname == m["href"]
            else "sidebar-button"
        )
        itens.append(
            dcc.Link(
                m["label"],
                href=m["href"],
                className=class_name,
            )
        )
    return itens

if __name__ == "__main__":
    app.run(debug=True)