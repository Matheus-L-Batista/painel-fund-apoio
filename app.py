import dash
from dash import Dash, html, dcc

app = Dash(__name__, use_pages=True)
server = app.server

app.layout = html.Div(
    className="app-root",
    children=[
        html.Div(
            className="app-container",
            children=[
                html.Div(
                    className="sidebar",
                    children=[
                        html.Img(
                            src="/assets/logo_unifei.png",
                            className="sidebar-logo"
                        ),

                        html.H2("Painéis", style={"color": "white"}),

                        dcc.Link(
                            "Passagens DCF",
                            href="/passagens-dcf",
                            refresh=True
                        ),
                        html.Br(),

                        dcc.Link(
                            "Pagamentos Efetivados",
                            href="/pagamentos",
                            refresh=True
                        ),
                        html.Br(),

                        dcc.Link(
                            "Dotação Atualizada",
                            href="/dotacao",
                            refresh=True
                        ),
                        html.Br(),

                        dcc.Link(
                            "Execução Orçamento UNIFEI",
                            href="/execucao-orcamento-unifei",
                            refresh=True
                        ),
                        html.Br(),

                        dcc.Link(
                            "Naturezas Despesa 2024",
                            href="/natureza-despesa-2024",
                            refresh=True
                        ),
                        html.Br(),

                        dcc.Link(
                            "Execução TED",
                            href="/execucao-ted",
                            refresh=True
                        ),
                    ],
                ),

                html.Div(
                    className="main-content",
                    children=dash.page_container,
                ),
            ],
        )
    ],
)

if __name__ == "__main__":
    app.run(debug=True)
