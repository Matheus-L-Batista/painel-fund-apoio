import dash
from dash import Dash, html, dcc, callback, Input, Output


app = Dash(
    __name__,
    use_pages=True,
    suppress_callback_exceptions=True,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
)
server = app.server


# LINKS NORMAIS (sem Processos de Compras e Status do Processo,
# pois eles ficam dentro da caixinha "Processos")
menu_links = [
    {"label": "Contratos", "href": "/contratos"},
    {"label": "Fiscais", "href": "/fiscais"},
    {"label": "Plano de Contratação Anual", "href": "/pca"},
    {"label": "Controle de Atas", "href": "/atas"},
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


@callback(
    Output("sidebar-menu", "children"),
    Input("url", "pathname"),
)
def atualizar_menu(pathname):
    itens = []

    # =========================
    # 0) Caixa Processos
    # =========================
    processos_paths = [
        "/processos-de-compras",   # ajuste se o path da página for diferente
        "/statusdoprocesso",
    ]
    processos_ativo = pathname in processos_paths

    pr_btn_classes = "processos-toggle"
    pr_content_classes = "processos-content"
    if processos_ativo:
        pr_btn_classes += " active"
        pr_content_classes += " expanded"

    processos_box = html.Div(
        className="processos-container",
        children=[
            html.Div(
                "Processos",
                id="btn-processos",
                className=pr_btn_classes,
            ),
            html.Div(
                id="box-processos",
                className=pr_content_classes,
                children=[
                    dcc.Link(
                        "Processos de Compras",
                        href="/processos-de-compras",
                        className=(
                            "processos-subbutton processos-subbutton-active"
                            if pathname == "/processos-de-compras"
                            else "processos-subbutton"
                        ),
                    ),
                    dcc.Link(
                        "Status do Processo",
                        href="/statusdoprocesso",
                        className=(
                            "processos-subbutton processos-subbutton-active"
                            if pathname == "/statusdoprocesso"
                            else "processos-subbutton"
                        ),
                    ),
                ],
            ),
        ],
    )
    itens.append(processos_box)

    # =========================
    # 1) Caixa Fracionamento
    # =========================
    fracionamento_ativo = pathname in ["/fracionamento_pdm", "/fracionamento_catser"]

    fr_btn_classes = "fracionamento-toggle"
    fr_content_classes = "fracionamento-content"
    if fracionamento_ativo:
        fr_btn_classes += " active"
        fr_content_classes += " expanded"

    fracionamento_box = html.Div(
        className="fracionamento-container",
        children=[
            html.Div(
                "Fracionamento de Despesas",
                id="btn-fracionamento",
                className=fr_btn_classes,
            ),
            html.Div(
                id="box-fracionamento",
                className=fr_content_classes,
                children=[
                    dcc.Link(
                        "Fracionamento de Despesas PDM (Material)",
                        href="/fracionamento_pdm",
                        className=(
                            "fracionamento-subbutton fracionamento-subbutton-active"
                            if pathname == "/fracionamento_pdm"
                            else "fracionamento-subbutton"
                        ),
                    ),
                    dcc.Link(
                        "Fracionamento de Despesas CATSER (Serviço)",
                        href="/fracionamento_catser",
                        className=(
                            "fracionamento-subbutton fracionamento-subbutton-active"
                            if pathname == "/fracionamento_catser"
                            else "fracionamento-subbutton"
                        ),
                    ),
                ],
            ),
        ],
    )
    itens.append(fracionamento_box)

    # =========================
    # 2) Caixa Portarias
    # =========================
    portarias_paths = [
        "/portarias_agentedecompras",
        "/portarias_planejamento",
    ]
    portarias_ativa = pathname in portarias_paths

    pt_btn_classes = "portarias-toggle"
    pt_content_classes = "portarias-content"
    if portarias_ativa:
        pt_btn_classes += " active"
        pt_content_classes += " expanded"

    portarias_box = html.Div(
        className="portarias-container",
        children=[
            html.Div(
                "Portarias",
                id="btn-portarias",
                className=pt_btn_classes,
            ),
            html.Div(
                id="box-portarias",
                className=pt_content_classes,
                children=[
                    dcc.Link(
                        "Portarias Agente de Compras/ Contratos Tipo Empenho",
                        href="/portarias_agentedecompras",
                        className=(
                            "portarias-subbutton portarias-subbutton-active"
                            if pathname == "/portarias_agentedecompras"
                            else "portarias-subbutton"
                        ),
                    ),
                    dcc.Link(
                        "Portarias de Planejamento da Contratação",
                        href="/portarias_planejamento",
                        className=(
                            "portarias-subbutton portarias-subbutton-active"
                            if pathname == "/portarias_planejamento"
                            else "portarias-subbutton"
                        ),
                    ),
                ],
            ),
        ],
    )
    itens.append(portarias_box)

    # =========================
    # 3) Demais itens normais
    # =========================
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


# Abre/fecha Processos
@callback(
    Output("btn-processos", "className"),
    Output("box-processos", "className"),
    Input("btn-processos", "n_clicks"),
    prevent_initial_call=True,
)
def toggle_processos(n):
    base_btn = "processos-toggle"
    base_box = "processos-content"
    if n and n % 2 == 1:
        return base_btn + " active", base_box + " expanded"
    return base_btn, base_box


# Abre/fecha Fracionamento
@callback(
    Output("btn-fracionamento", "className"),
    Output("box-fracionamento", "className"),
    Input("btn-fracionamento", "n_clicks"),
    prevent_initial_call=True,
)
def toggle_fracionamento(n):
    base_btn = "fracionamento-toggle"
    base_box = "fracionamento-content"
    if n and n % 2 == 1:
        return base_btn + " active", base_box + " expanded"
    return base_btn, base_box


# Abre/fecha Portarias
@callback(
    Output("btn-portarias", "className"),
    Output("box-portarias", "className"),
    Input("btn-portarias", "n_clicks"),
    prevent_initial_call=True,
)
def toggle_portarias(n):
    base_btn = "portarias-toggle"
    base_box = "portarias-content"
    if n and n % 2 == 1:
        return base_btn + " active", base_box + " expanded"
    return base_btn, base_box


if __name__ == "__main__":
    app.run(debug=True)
