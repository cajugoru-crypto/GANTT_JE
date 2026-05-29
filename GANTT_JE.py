from dash import Dash, html, dcc, Input, Output, State, callback_context, no_update
import locale

try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except:
    pass
import dash_bootstrap_components as dbc
import dash_ag_grid as dag
import plotly.graph_objects as go
import pandas as pd
import uuid
import os
from datetime import datetime
import logging

# =========================================================
# CONFIG
# =========================================================
ARCHIVO = "Gantt.xlsx"

COLUMNAS = [
    "ID",
    "Tarea",
    "FechaInicio",
    "FechaFin",
    "Duracion",
    "Responsable",
    "Estado",
    "Prioridad",
    "Recurrencia"
]

ANIO_ACTUAL = datetime.now().year

# =========================================================
# COLORES NOVAVENTA + NUTRESA
# =========================================================
COLOR_PRIMARIO = "#278236"
COLOR_SECUNDARIO = "#0C7803"
COLOR_ROSA = "#E31C79"
COLOR_FONDO = "#F5FBF7"
COLOR_CARD = "#FFFFFF"
COLOR_BORDE = "#E8DDF1"
COLOR_TEXTO = "#2B1B35"

# =========================================================
# CREAR BASE
# =========================================================
def crear_base():

    df = pd.DataFrame(columns=COLUMNAS)

    df.to_excel(
        ARCHIVO,
        index=False,
        engine="openpyxl"
    )

    return df

# =========================================================
# CARGAR DATOS
# =========================================================
def cargar_datos():

    if not os.path.exists(ARCHIVO):

        return crear_base()

    try:

        df = pd.read_excel(
            ARCHIVO,
            engine="openpyxl"
        )

    except:

        return crear_base()

    for col in COLUMNAS:

        if col not in df.columns:

            df[col] = ""

    df = df[COLUMNAS]

    df = df.fillna("")

    return df

# =========================================================
# GUARDAR DATOS
# =========================================================
def guardar_datos(df):

    try:

        df.to_excel(
            ARCHIVO,
            index=False,
            engine="openpyxl"
        )

    except Exception as e:

        print("ERROR:", e)

# =========================================================
# GANTT
# =========================================================
def crear_gantt(data):

    fig = go.Figure()

    if len(data) == 0:

        fig.update_layout(

            plot_bgcolor="white",
            paper_bgcolor="white",

            xaxis=dict(visible=False),
            yaxis=dict(visible=False)
        )

        return fig

    data = data.copy()

    data["FechaInicio"] = pd.to_datetime(
        data["FechaInicio"],
        errors="coerce"
    )

    data["FechaFin"] = pd.to_datetime(
        data["FechaFin"],
        errors="coerce"
    )

    data = data.dropna(
        subset=["FechaInicio", "FechaFin"]
    )

    data = data[
        data["FechaInicio"].dt.year == ANIO_ACTUAL
    ]

    data = data.sort_values(
        by="FechaInicio"
    )

    colores = {

        "Alta": "#FB6E6E",
        "Media": "#373C79",
        "Baja": "#388E3C"
    }

    for _, row in data.iterrows():

        inicio = row["FechaInicio"]

        fin = row["FechaFin"]

        fig.add_trace(

            go.Scatter(

                x=[inicio, fin],

                y=[row["Tarea"], row["Tarea"]],

                mode="lines",

                line=dict(

                    color=colores.get(
                        row["Prioridad"],
                        COLOR_PRIMARIO
                    ),

                    width=14
                ),

                hovertemplate=
                f"""
                <b>{row['Tarea']}</b><br>
                Inicio: {inicio.strftime('%Y-%m-%d')}<br>
                Fin: {fin.strftime('%Y-%m-%d')}<br>
                Responsable: {row['Responsable']}<br>
                Estado: {row['Estado']}<br>
                Prioridad: {row['Prioridad']}<extra></extra>
                """
            )
        )

    inicio_anio = pd.Timestamp(f"{ANIO_ACTUAL}-01-01")
    fin_anio = pd.Timestamp(f"{ANIO_ACTUAL}-12-31")

    fig.update_layout(

        height=max(
            450,
            len(data) * 38
        ),

        plot_bgcolor="white",

        paper_bgcolor="white",

        showlegend=False,

        margin=dict(
            l=20,
            r=20,
            t=20,
            b=20
        ),

        font=dict(
            family="Segoe UI",
            size=15,
            color=COLOR_TEXTO
        ),

        xaxis=dict(

            title=f"Cronograma {ANIO_ACTUAL}",

            type="date",

            range=[inicio_anio, fin_anio],

            tickformat="%B",

            dtick="M1",

            showgrid=True,

            gridcolor="#ACF5B7"
        ),

        yaxis=dict(

            autorange="reversed",

            showgrid=False
        )
    )

    return fig

# =========================================================
# APP
# =========================================================
app = Dash(

    __name__,

    external_stylesheets=[
        dbc.themes.BOOTSTRAP
    ],

    update_title=None
)

# =========================================================
# QUITAR DEBUG MENU
# =========================================================
app.index_string = '''

<!DOCTYPE html>
<html>

    <head>

        {%metas%}

        <title>Gantt Novaventa</title>

        {%favicon%}

        {%css%}

        <style>

            .dash-debug-menu__outer,
            .dash-debug-menu,
            .dash-debug-menu__button {

                display: none !important;
            }

        </style>

    </head>

    <body>

        {%app_entry%}

        <footer>

            {%config%}
            {%scripts%}
            {%renderer%}

        </footer>

    </body>

</html>

'''

server = app.server

# =========================================================
# POSICION LOGO SIDEBAR (AJUSTABLE)
# =========================================================
LOGO_SIDEBAR_BOTTOM = "20px"     # subir / bajar
LOGO_SIDEBAR_LEFT = "15px"       # mover izquierda
LOGO_SIDEBAR_RIGHT = "15px"      # mover derecha

LOGO_SIDEBAR_WIDTH = "80px"     # tamaño ancho
LOGO_SIDEBAR_HEIGHT = "auto"     # tamaño alto


# =========================================================
# SIDEBAR
# =========================================================
sidebar = html.Div(

    id="sidebar",

    children=[

        # TITULO
        html.H2(

            "",

            style={

                "fontWeight": "900",

                "marginBottom": "5px",

                "fontSize": "30px",

                "color": "white"

            }
        ),

        # SUBTITULO
        html.Div(

            "",

            style={

                "color": "rgba(255,255,255,0.7)",

                "marginBottom": "40px"

            }
        ),

        # BOTON DASHBOARD
        dbc.Button(

            "📋 DASHBOARD",

            id="btn-dashboard",

            color="light",

            className="mb-3",

            style={

                "width": "100%",

                "height": "50px",

                "borderRadius": "10px",

                "fontWeight": "700",

                "color": COLOR_PRIMARIO

            }
        ),

        # BOTON OCULTO PARA QUE NO FALLE EL CALLBACK
        html.Div(

            id="btn-calendario",

            style={

                "display": "none"

            }
        ),

        # LOGO SIDEBAR
        html.Img(

            src="/assets/logo nutresa.png",

            style={

                "position": "absolute",

                "bottom": LOGO_SIDEBAR_BOTTOM,

                "left": LOGO_SIDEBAR_LEFT,

                "right": LOGO_SIDEBAR_RIGHT,

                "width": LOGO_SIDEBAR_WIDTH,

                "height": LOGO_SIDEBAR_HEIGHT,

                "objectFit": "contain",

                "background": "transparent",

                "margin": "auto",

                "display": "block",

                "zIndex": "999"

            }
        )

    ],

    style={

        "position": "fixed",

        "top": 0,

        "left": "0px",

        "bottom": 0,

        "width": "260px",

        "padding": "30px",

        "background": f"linear-gradient(180deg,{COLOR_PRIMARIO},{COLOR_SECUNDARIO})",

        "zIndex": 999,

        "transition": "all 0.3s ease"

    }

)
# =========================================================
# KPI
# =========================================================
def crear_kpi(titulo, id_texto, id_boton):

    return dbc.Button(

        [

            html.Div(

                titulo,

                style={
                    "fontSize": "14px",
                    "fontWeight": "700",
                    "color": "#4B5563"
                }
            ),

            html.H5(

                id=id_texto,

                style={
                    "fontWeight": "900",
                    "marginTop": "6px",
                    "marginBottom": "0px",
                    "fontSize": "26px",
                    "color": COLOR_TEXTO
                }
            )

        ],

        id=id_boton,

        color="light",

        style={

            "width": "100%",

            "height": "78px",

            "borderRadius": "20px",

            "border": "none",

            "background": "linear-gradient(145deg,#FFFFFF,#F4F7F4)",

            "color": COLOR_TEXTO,

            "boxShadow": "0 8px 22px rgba(0,0,0,0.10)",

            "transition": "all 0.25s ease",

            "padding": "12px"

        }
    )


# =========================================================
# POSICION LOGO (AJUSTABLE)
# =========================================================
LOGO_TOP = "-30px"
LOGO_RIGHT = "20px"
LOGO_WIDTH = "260px"
LOGO_HEIGHT = "auto"


# =========================================================
# DASHBOARD
# =========================================================
dashboard = html.Div(

    id="vista-dashboard",

    children=[

        # LOGO
        html.Img(

            src="/assets/logo.png",

            style={

                "position": "absolute",

                "top": LOGO_TOP,

                "right": LOGO_RIGHT,

                "width": LOGO_WIDTH,

                "height": LOGO_HEIGHT,

                "objectFit": "contain",

                "background": "transparent",

                "zIndex": "999"

            }
        ),

        # TITULO
        html.H1(

            f"PLAN DE TRABAJO {ANIO_ACTUAL}",

            style={

                "fontWeight": "900",

                "fontSize": "42px",

                "color": COLOR_TEXTO,

                "marginBottom": "5px"

            }
        ),

        # SUBTITULO
        html.P(

            "Vista general del cronograma operativo",

            style={

                "color": "#074607",

                "marginBottom": "25px",

                "fontSize": "17px",

                "fontWeight": "500"

            }
        ),

        # KPIS
        dbc.Row([

            dbc.Col(crear_kpi("Total", "kpi-total", "btn-kpi-total"), width=1),

            dbc.Col(crear_kpi("Alta", "kpi-alta", "btn-kpi-alta"), width=1),

            dbc.Col(crear_kpi("Media", "kpi-media", "btn-kpi-media"), width=1),

            dbc.Col(crear_kpi("Baja", "kpi-baja", "btn-kpi-baja"), width=1),

            dbc.Col(crear_kpi("Pendiente", "kpi-pendiente", "btn-kpi-pendiente"), width=2),

            dbc.Col(crear_kpi("En progreso", "kpi-progreso", "btn-kpi-progreso"), width=2),

            dbc.Col(crear_kpi("Finalizado", "kpi-finalizado", "btn-kpi-finalizado"), width=2)

        ], className="mb-4 g-3"),

        # GRAFICO
        dbc.Card(

            dbc.CardBody([

                dcc.Graph(

                    id="grafico-dashboard",

                    config={

                        "displayModeBar": False

                    }

                )

            ]),

            style={

                "borderRadius": "28px",

                "border": "1px solid rgba(255,255,255,0.5)",

                "boxShadow": "0 12px 35px rgba(0,0,0,0.12)",

                "background": "linear-gradient(145deg,#FFFFFF,#F3F8F3)",

                "backdropFilter": "blur(12px)",

                "padding": "10px"

            }

        )

    ],

    style={

        "padding": "35px",

        "position": "relative",

        "backgroundColor": COLOR_FONDO,

        "minHeight": "100vh"

    }

)
# =========================================================
# CALENDARIO
# =========================================================
calendario = html.Div(

    id="vista-calendario",

    children=[

        html.H1(

            "Calendario Operativo",

            style={
                "fontWeight": "800",
                "color": COLOR_TEXTO
            }
        ),

        html.Br(),

        dbc.Row([

            dbc.Col([

                dbc.Card(

                    dbc.CardBody([

                        html.H4(
                            "Nueva tarea",
                            className="mb-4"
                        ),

                        dbc.Input(
                            id="input-tarea",
                            placeholder="Nombre tarea",
                            className="mb-3"
                        ),

                        dcc.DatePickerSingle(
                            id="input-fecha"
                        ),

                        html.Br(),
                        html.Br(),

                        dbc.Input(
                            id="input-duracion",
                            type="number",
                            value=1
                        ),

                        html.Br(),

                        dbc.Input(
                            id="input-responsable",
                            value="Oscar"
                        ),

                        html.Br(),

                        dcc.Dropdown(

                            id="input-estado",

                            options=[
                                {"label": "Pendiente", "value": "Pendiente"},
                                {"label": "En progreso", "value": "En progreso"},
                                {"label": "Finalizado", "value": "Finalizado"}
                            ],

                            value="Pendiente"
                        ),

                        html.Br(),

                        dcc.Dropdown(

                            id="input-prioridad",

                            options=[
                                {"label": "Alta", "value": "Alta"},
                                {"label": "Media", "value": "Media"},
                                {"label": "Baja", "value": "Baja"}
                            ],

                            value="Media"
                        ),

                        html.Br(),

                        dcc.Dropdown(

                            id="input-recurrencia",

                            options=[
                                {"label": "No recurrente", "value": "NO"},
                                {"label": "Mensual", "value": "MENSUAL"}
                            ],

                            value="NO"
                        ),

                        html.Br(),

                        dbc.Button(

                            "Guardar tarea",

                            id="btn-guardar",

                            style={
                                "width": "100%",
                                "height": "50px",
                                "backgroundColor": COLOR_PRIMARIO,
                                "border": "none",
                                "fontWeight": "700"
                            }
                        )

                    ]),

                    style={
                        "borderRadius": "22px",
                        "border": "none",
                        "boxShadow": "0 8px 24px rgba(111,44,145,0.08)"
                    }
                )

            ], width=4),

            dbc.Col([

                dbc.Card(

                    dbc.CardBody([

                        dag.AgGrid(

                            id="tabla-tareas",

                            rowData=[],

                            columnDefs=[

                                {
                                    "field": "ID",
                                    "hide": True
                                },

                                {
                                    "field": "Tarea",
                                    "checkboxSelection": True,
                                    "headerCheckboxSelection": True,
                                    "flex": 2
                                },

                                {"field": "FechaInicio"},
                                {"field": "FechaFin"},
                                {"field": "Duracion"},
                                {"field": "Responsable"},
                                {"field": "Estado"},
                                {"field": "Prioridad"},
                                {"field": "Recurrencia"}

                            ],

                            defaultColDef={
                                "sortable": True,
                                "filter": True,
                                "resizable": True
                            },

                            dashGridOptions={

                                "rowSelection": "multiple",

                                "animateRows": True,

                                "getRowId": {
                                    "function": "params.data.ID"
                                }
                            },

                            className="ag-theme-alpine",

                            style={
                                "height": "320px"
                            }
                        ),

                        html.Br(),

                        dbc.Button(

                            "🗑️ Eliminar tarea",

                            id="btn-eliminar",

                            color="danger"
                        ),

                        html.Br(),
                        html.Br(),

                        dcc.Graph(

                            id="grafico-gantt",

                            config={
                                "displayModeBar": False
                            }
                        )

                    ]),

                    style={
                        "borderRadius": "22px",
                        "border": "none",
                        "boxShadow": "0 8px 24px rgba(111,44,145,0.08)"
                    }
                )

            ], width=8)

        ])

    ],

    style={
        "padding": "35px",
        "display": "none"
    }
)

# =========================================================
# LAYOUT
# =========================================================
app.layout = html.Div([

    dcc.Store(
        id="store-refresh",
        data=0
    ),

    dcc.Store(
        id="store-filtro",
        data="TODOS"
    ),

    # BOTON MENU
    html.Button(

        "☰",

        id="btn-menu",

        n_clicks=0,

        style={

            "position": "fixed",

            "top": "15px",

            "left": "15px",

            "zIndex": "2000",

            "border": "none",

            "background": COLOR_PRIMARIO,

            "color": "white",

            "width": "45px",

            "height": "45px",

            "borderRadius": "10px",

            "fontSize": "22px",

            "fontWeight": "bold",

            "boxShadow": "0 4px 12px rgba(0,0,0,0.25)"

        }
    ),

    sidebar,

    html.Div(

        id="contenido-principal",

        children=[

            dashboard,
            calendario

        ],

        style={

            "marginLeft": "260px",

            "background": "linear-gradient(180deg,#F8F5FB,#F3EDF8)",

            "minHeight": "100vh",

            "transition": "all 0.3s ease"

        }
    )
])
# =========================================================
# CAMBIAR VISTA
# =========================================================
@app.callback(

    Output("vista-dashboard", "style"),
    Output("vista-calendario", "style"),

    Input("btn-dashboard", "n_clicks"),
    Input("btn-calendario", "n_clicks")
)
def cambiar_vista(a, b):

    ctx = callback_context

    dashboard_style = {
        "padding": "35px",
        "display": "none",
        "position": "relative",
        "backgroundColor": COLOR_FONDO
    }

    calendario_style = {
        "padding": "35px",
        "display": "none"
    }

    if not ctx.triggered:

        dashboard_style["display"] = "block"

        return dashboard_style, calendario_style

    boton = ctx.triggered[0]["prop_id"].split(".")[0]

    if boton == "btn-calendario":

        calendario_style["display"] = "block"

    else:

        dashboard_style["display"] = "block"

    return dashboard_style, calendario_style


# =========================================================
# FILTRO KPI
# =========================================================
@app.callback(

    Output("store-filtro", "data"),

    Input("btn-kpi-total", "n_clicks"),
    Input("btn-kpi-alta", "n_clicks"),
    Input("btn-kpi-media", "n_clicks"),
    Input("btn-kpi-baja", "n_clicks"),
    Input("btn-kpi-pendiente", "n_clicks"),
    Input("btn-kpi-progreso", "n_clicks"),
    Input("btn-kpi-finalizado", "n_clicks"),

    prevent_initial_call=True
)
def cambiar_filtro(a, b, c, d, e, f, g):

    ctx = callback_context

    if not ctx.triggered:
        return "TODOS"

    boton = ctx.triggered[0]["prop_id"].split(".")[0]

    mapa = {

        "btn-kpi-alta": "Alta",
        "btn-kpi-media": "Media",
        "btn-kpi-baja": "Baja",

        "btn-kpi-pendiente": "Pendiente",
        "btn-kpi-progreso": "En progreso",
        "btn-kpi-finalizado": "Finalizado"
    }

    return mapa.get(boton, "TODOS")


# =========================================================
# ACTUALIZAR
# =========================================================
@app.callback(

    Output("kpi-total", "children"),
    Output("kpi-alta", "children"),
    Output("kpi-media", "children"),
    Output("kpi-baja", "children"),
    Output("kpi-pendiente", "children"),
    Output("kpi-progreso", "children"),
    Output("kpi-finalizado", "children"),
    Output("grafico-dashboard", "figure"),
    Output("grafico-gantt", "figure"),
    Output("tabla-tareas", "rowData"),

    Input("store-filtro", "data"),
    Input("store-refresh", "data")
)
def actualizar(filtro, refresh):

    df_actual = cargar_datos()

    total = len(df_actual)

    alta = len(df_actual[df_actual["Prioridad"] == "Alta"])
    media = len(df_actual[df_actual["Prioridad"] == "Media"])
    baja = len(df_actual[df_actual["Prioridad"] == "Baja"])

    pendiente = len(df_actual[df_actual["Estado"] == "Pendiente"])
    progreso = len(df_actual[df_actual["Estado"] == "En progreso"])
    finalizado = len(df_actual[df_actual["Estado"] == "Finalizado"])

    df_filtrado = df_actual.copy()

    if filtro in ["Alta", "Media", "Baja"]:

        df_filtrado = df_filtrado[
            df_filtrado["Prioridad"] == filtro
        ]

    elif filtro in [

        "Pendiente",
        "En progreso",
        "Finalizado"

    ]:

        df_filtrado = df_filtrado[
            df_filtrado["Estado"] == filtro
        ]

    gantt = crear_gantt(df_filtrado)

    return (

        total,
        alta,
        media,
        baja,
        pendiente,
        progreso,
        finalizado,

        gantt,
        gantt,

        df_filtrado.to_dict("records")
    )


# =========================================================
# AGREGAR
# =========================================================
@app.callback(

    Output("store-refresh", "data", allow_duplicate=True),

    Input("btn-guardar", "n_clicks"),

    State("input-tarea", "value"),
    State("input-fecha", "date"),
    State("input-duracion", "value"),
    State("input-responsable", "value"),
    State("input-estado", "value"),
    State("input-prioridad", "value"),
    State("input-recurrencia", "value"),

    State("store-refresh", "data"),

    prevent_initial_call=True
)
def agregar(

    n,
    tarea,
    fecha,
    duracion,
    responsable,
    estado,
    prioridad,
    recurrencia,
    refresh
):

    if not tarea:
        return no_update

    df_actual = cargar_datos()

    fecha_inicio = pd.to_datetime(fecha)

    tareas_nuevas = []

    if recurrencia == "NO":

        fecha_fin = fecha_inicio + pd.Timedelta(
            weeks=float(duracion)
        )

        tareas_nuevas.append({

            "ID": str(uuid.uuid4()),

            "Tarea": tarea,

            "FechaInicio": fecha_inicio.strftime("%Y-%m-%d"),

            "FechaFin": fecha_fin.strftime("%Y-%m-%d"),

            "Duracion": duracion,

            "Responsable": responsable,

            "Estado": estado,

            "Prioridad": prioridad,

            "Recurrencia": recurrencia
        })

    elif recurrencia == "MENSUAL":

        for i in range(12):

            nueva_fecha = fecha_inicio + pd.DateOffset(
                months=i
            )

            fecha_fin = nueva_fecha + pd.Timedelta(
                weeks=float(duracion)
            )

            tareas_nuevas.append({

                "ID": str(uuid.uuid4()),

                "Tarea": f"{tarea} ({nueva_fecha.strftime('%B %Y')})",

                "FechaInicio": nueva_fecha.strftime("%Y-%m-%d"),

                "FechaFin": fecha_fin.strftime("%Y-%m-%d"),

                "Duracion": duracion,

                "Responsable": responsable,

                "Estado": estado,

                "Prioridad": prioridad,

                "Recurrencia": recurrencia
            })

    df_nuevo = pd.DataFrame(tareas_nuevas)

    df_actual = pd.concat(
        [df_actual, df_nuevo],
        ignore_index=True
    )

    guardar_datos(df_actual)

    return refresh + 1


# =========================================================
# ELIMINAR
# =========================================================
@app.callback(

    Output("store-refresh", "data", allow_duplicate=True),

    Input("btn-eliminar", "n_clicks"),

    State("tabla-tareas", "selectedRows"),

    State("store-refresh", "data"),

    prevent_initial_call=True
)
def eliminar(n, seleccionadas, refresh):

    if not seleccionadas:
        return no_update

    df_actual = cargar_datos()

    ids_eliminar = [

        str(fila["ID"])

        for fila in seleccionadas
    ]

    df_actual["ID"] = df_actual["ID"].astype(str)

    df_actual = df_actual[
        ~df_actual["ID"].isin(ids_eliminar)
    ]

    guardar_datos(df_actual)

    return refresh + 1

# =========================================================
# MOSTRAR / OCULTAR MENU
# =========================================================
@app.callback(

    Output("sidebar", "style"),
    Output("contenido-principal", "style"),

    Input("btn-menu", "n_clicks")
)
def toggle_sidebar(n):

    visible = n % 2 == 0

    sidebar_style = {

        "position": "fixed",

        "top": 0,

        "left": "0px" if visible else "-260px",

        "bottom": 0,

        "width": "260px",

        "padding": "30px",

        "background": f"linear-gradient(180deg,{COLOR_PRIMARIO},{COLOR_SECUNDARIO})",

        "zIndex": 999,

        "transition": "all 0.3s ease"

    }

    contenido_style = {

        "marginLeft": "260px" if visible else "0px",

        "background": "linear-gradient(180deg,#F8F5FB,#F3EDF8)",

        "minHeight": "100vh",

        "transition": "all 0.3s ease"

    }

    return sidebar_style, contenido_style
# =========================================================
# LIMPIAR LOGS
# =========================================================
log = logging.getLogger('werkzeug')
log.setLevel(logging.ERROR)


# =========================================================
# RUN
# =========================================================
if __name__ == "__main__":

    app.run(
        debug=False,
        use_reloader=False
    )
