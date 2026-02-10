import dash
from dash import Dash, html, dcc, callback, Input, Output


app = Dash(
    __name__,
    use_pages=True,
    suppress_callback_exceptions=True,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
)
server = app.server


app.layout = html.Div(
    className="app-root",
    children=[
        dcc.Location(id="url"),

        dcc.Interval(
            id="interval-atualizacao",
            interval=60 * 60 * 1000,  # 1 hora
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


@callback(
    Output("sidebar-menu", "children"),
    Input("url", "pathname"),
)
def atualizar_menu(pathname):
    ativo = pathname == "/contratos"

    contratos_item = dcc.Link(
        "Contratos com Fundações",
        href="/contratos",
        className="sidebar-button sidebar-button-active" if ativo else "sidebar-button",
    )

    return [contratos_item]


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=8050,
        debug=False,
    )
