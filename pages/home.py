import dash
from dash import html

dash.register_page(
    __name__,
    path="/",
    name="Início",
)

layout = html.Div(
    className="home-container",
    children=[
        html.Div(className="home-overlay"),
        # aqui você pode colocar conteúdo por cima da imagem
        # html.Div(
        #     children=[
        #         html.H1("Diretoria de Compras e Contratos"),
        #         html.P("Texto descritivo..."),
        #     ]
        # ),
    ],
)
