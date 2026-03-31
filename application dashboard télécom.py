import dash
from dash import html, dcc, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import base64
import io
import json # Pour la sérialisation/désérialisation des DataFrames
import plotly.express as px # Importation de Plotly Express pour les graphiques
import plotly.graph_objects as go # Import for more complex graph options if needed
from sklearn.linear_model import LinearRegression # For potential regression, add if actually used
from sklearn.model_selection import train_test_split # For regression, add if actually used
from sklearn.metrics import mean_squared_error # For regression, add if actually used
import random # For simulating data if necessary
import unicodedata # Pour la normalisation Unicode
import re # Pour les expressions régulières

# Initialisation de l'application Dash avec un thème Bootstrap personnalisé
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
app.title = "Tableau de Bord Réseau Télécom"
app.config.suppress_callback_exceptions = True

# Définition de la capacité par port par défaut (sera ajustable par l'utilisateur)
DEFAULT_PORT_CAPACITY = 64 # Capacité par défaut, peut être modifiée via l'interface

# Mappage des types de splitter à leur capacité correspondante
# Mise à jour basée sur l'explication de l'utilisateur pour "4:32"
SPLITTER_CAPACITY_MAP = {
    "1:2": 2,
    "1:4": 4,
    "1:8": 8,
    "1:16": 16,
    "1:64": 64,
    "4:32": 256, # Capacité calculée comme 4 * 8 * 8 = 256 selon l'explication de l'utilisateur
    "Par défaut (64)": 64 # Option pour revenir à la capacité par défaut
}

# Définition du layout de l'application
app.layout = dbc.Container([
    # Composants dcc.Store pour stocker les DataFrames importés
    dcc.Store(id='stored-df1', data=pd.DataFrame().to_json(orient='split'), storage_type='local'),
    dcc.Store(id='stored-df2', data=pd.DataFrame().to_json(orient='split'), storage_type='local'),
    # dcc.Store pour stocker la capacité de port personnalisée globale
    dcc.Store(id='custom-port-capacity-store', data=DEFAULT_PORT_CAPACITY, storage_type='local'),
    # Nouveau dcc.Store pour stocker les capacités spécifiques par port (dictionnaire)
    # Note: Pour un test propre, l'utilisateur pourrait devoir vider le Local Storage de son navigateur
    dcc.Store(id='per-port-capacity-store', data={}, storage_type='local'),
    # Nouveau dcc.Store pour stocker l'OLT et le Port cliqués pour la modale
    dcc.Store(id='clicked-port-info-store', data={}),


    dbc.NavbarSimple(
        brand="📊 Bienvenue sur votre Tableau de Bord Réseau Télécom !",
        brand_style={"color": "#FFFFFF", "font-weight": "bold"},
        color="#007bff",
        dark=True,
        className="mb-4"
    ),

    dbc.Tabs([
        dbc.Tab(label="Affichage des Données", tab_id="tab-data",
                labelClassName="text-primary", activeTabClassName="text-primary border-primary"),
        dbc.Tab(label="Taux de Saturation", tab_id="tab-saturation",
                labelClassName="text-primary", activeTabClassName="text-primary border-primary"),
        dbc.Tab(label="Statistiques", tab_id="tab-stats",
                labelClassName="text-primary", activeTabClassName="text-primary border-primary"),
    ], id="tabs", active_tab="tab-data", className="my-3"),

    html.Div(id="content"),

    # Modale pour la saisie du type de splitter
    dbc.Modal(
        [
            dbc.ModalHeader(dbc.ModalTitle("Configurer le Splitter pour le Port")),
            dbc.ModalBody([
                html.P(id="modal-port-label"), # Pour afficher le port sélectionné
                html.Label("Sélectionnez le type de Splitter :"),
                dcc.Dropdown(
                    id='splitter-type-dropdown',
                    options=[{'label': k, 'value': k} for k in SPLITTER_CAPACITY_MAP.keys()],
                    placeholder="Sélectionnez un type de splitter",
                    clearable=False
                )
            ]),
            dbc.ModalFooter(
                dbc.Button("Appliquer", id="apply-splitter-btn", className="ms-auto")
            ),
        ],
        id="splitter-modal",
        is_open=False,
    )

], fluid=True, style={"background-color": "#f8f9fa", "padding": "20px", "border-radius": "8px"})

# Callback pour rendre le contenu des onglets principaux (Affichage des Données, Taux de Saturation, Statistiques)
@app.callback(
    Output("content", "children"),
    Input("tabs", "active_tab")
)
def render_tab_content(tab):
    if tab == "tab-data":
        return html.Div([
            html.H3("🚀 Gérez et Visualisez vos Données Réseau en toute Simplicité", className="text-primary mb-4", style={"text-align": "center"}),
            dbc.Row([
                dbc.Col([
                    html.H5("📦 Base Ports / Slots / Équipements", className="text-secondary mb-2"),
                    dcc.Upload(
                        id="upload-data-1",
                        children=dbc.Button(
                            [html.I(className="bi bi-cloud-arrow-up-fill me-2"), "Upload File"], # Icône et texte
                            color="primary",
                            className="fw-bold shadow-sm rounded-pill px-4 py-2" # Adjusted button size
                        ),
                        style={
                            "width": "100%", "height": "60px", "lineHeight": "60px",
                            "borderWidth": "2px", "borderStyle": "dashed", "borderRadius": "8px",
                            "borderColor": "#007bff", # A primary color border
                            "backgroundColor": "#e9f5ff", # Light blue background
                            "textAlign": "center", "marginBottom": "10px", "cursor": "pointer",
                            "display": "flex", "alignItems": "center", "justifyContent": "center", # For centering the button
                            "padding": "5px" # Add some padding
                        },
                        multiple=False
                    )
                ], md=6),
                dbc.Col([
                    html.H5("📇 Base Infos Abonnés", className="text-secondary mb-2"),
                    dcc.Upload(
                        id="upload-data-2",
                        children=dbc.Button(
                            [html.I(className="bi bi-cloud-arrow-up-fill me-2"), "Upload File"], # Icône et texte
                            color="primary",
                            className="fw-bold shadow-sm rounded-pill px-4 py-2" # Adjusted button size
                        ),
                        style={
                            "width": "100%", "height": "60px", "lineHeight": "60px",
                            "borderWidth": "2px", "borderStyle": "dashed", "borderRadius": "8px",
                            "borderColor": "#007bff", # A primary color border
                            "backgroundColor": "#e9f5ff", # Light blue background
                            "textAlign": "center", "marginBottom": "10px", "cursor": "pointer",
                            "display": "flex", "alignItems": "center", "justifyContent": "center", # For centering the button
                            "padding": "5px" # Add some padding
                        },
                        multiple=False
                    )
                ], md=6)
            ]),
            # Ce div sera mis à jour par un callback séparé pour afficher les tableaux
            html.Div(id="tabs-datatables", className="mt-4")
        ])

    elif tab == "tab-saturation":
        # Le contenu de cet onglet inclut le champ d'entrée pour la capacité du port globale
        return html.Div([
            html.H3("📶 Taux de Saturation par Port et OLT", className="text-primary mb-4", style={"text-align": "center"}),
            dbc.Row([
                dbc.Col([
                    html.Label("Capacité par Port GPON (nombre d'abonnés maximum) :", className="mb-2"),
                    dcc.Input(
                        id="port-capacity-input",
                        type="number",
                        value=DEFAULT_PORT_CAPACITY, # Valeur par défaut
                        min=1,
                        step=1,
                        style={"width": "150px", "margin-right": "10px"}
                    ),
                    html.Small("Ex: 32 pour un splitter 1:32, 64 pour un 1:64", className="text-muted")
                ], width=12, className="mb-4")
            ]),
            html.Div(id="saturation-content", className="mt-4")
        ])

    elif tab == "tab-stats":
        return html.Div([
            html.H3("📈 Statistiques et Analyse du Réseau", className="text-primary mb-4", style={"text-align": "center"}),
            dbc.Row([
                dbc.Col([
                    html.H5("🌍 Carte des Utilisateurs par Région", className="text-secondary mb-2"),
                    dcc.Graph(id='region-choropleth-map', style={'height': '500px'})
                ], md=6),
                dbc.Col([
                    html.H5("📊 Distribution des Débits", className="text-secondary mb-2"),
                    dcc.Graph(id='debit-distribution-chart', style={'height': '500px'})
                ], md=6)
            ]),
            dbc.Row([
                dbc.Col([
                    html.H5("💡 Analyse des Offres Souscrites", className="text-secondary mb-2"),
                    dcc.Graph(id='offer-distribution-chart', style={'height': '500px'})
                ], md=6),
                dbc.Col([
                    html.H5("🚀 Prédiction d'Offre (Tendances)", className="text-secondary mb-2"),
                    html.Div(id='offer-prediction-output', className="alert alert-info mt-3")
                ], md=6)
            ])
        ])

    return html.Div("Sélectionnez un onglet.")

# Callback pour gérer l'upload des fichiers et les stocker dans dcc.Store
@app.callback(
    [Output("stored-df1", "data"),
     Output("stored-df2", "data")],
    [Input("upload-data-1", "contents"),
     Input("upload-data-2", "contents")],
    [State("upload-data-1", "filename"),
     State("upload-data-2", "filename"),
     State("stored-df1", "data"), # Add current stored data as State
     State("stored-df2", "data")], # Add current stored data as State
    prevent_initial_call=True
)
def handle_file_upload_and_store(content1, content2, filename1, filename2, current_df1_data, current_df2_data):
    # DEBUG: Check if contents and filenames are received
    print(f"DEBUG: handle_file_upload_and_store - content1: {'Received' if content1 else 'None'}, filename1: {filename1}")
    print(f"DEBUG: handle_file_upload_and_store - content2: {'Received' if content2 else 'None'}, filename2: {filename2}")
    print(f"DEBUG: handle_file_upload_and_store - current_df1_data length: {len(current_df1_data)}")
    print(f"DEBUG: handle_file_upload_and_store - current_df2_data length: {len(current_df2_data)}")


    def parse_file(contents, filename):
        if contents is None:
            return pd.DataFrame(), None # Retourne un DataFrame vide si pas de contenu
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        ext = filename.split('.')[-1].lower()

        try:
            if ext == 'xlsx':
                df = pd.read_excel(io.BytesIO(decoded), engine='openpyxl')
            elif ext == 'xls':
                df = pd.read_excel(io.BytesIO(decoded), engine='xlrd')
            elif ext == 'csv':
                df = pd.read_csv(io.BytesIO(decoded))
            elif ext in ['txt', 'tsv', 'dat']:
                df = pd.read_csv(io.BytesIO(decoded), sep=None, engine='python')
            else:
                return pd.DataFrame(), f"Format non supporté : {ext}. Veuillez importer un fichier Excel, CSV ou TXT."
        except Exception as e:
            print(f"ERROR: parse_file - Error parsing {filename}: {e}")
            return pd.DataFrame(), str(e)
        
        print(f"DEBUG: parse_file - Successfully parsed {filename}. Shape: {df.shape}")
        return df, None

    new_df1_data = current_df1_data
    new_df2_data = current_df2_data

    if content1:
        df1, err1 = parse_file(content1, filename1)
        if err1:
            print(f"ERROR: handle_file_upload_and_store - Error with df1: {err1}")
        else:
            new_df1_data = df1.to_json(orient='split')
    
    if content2:
        df2, err2 = parse_file(content2, filename2)
        if err2:
            print(f"ERROR: handle_file_upload_and_store - Error with df2: {err2}")
        else:
            new_df2_data = df2.to_json(orient='split')

    print(f"DEBUG: handle_file_upload_and_store - new_df1_data length: {len(new_df1_data)}")
    print(f"DEBUG: handle_file_upload_and_store - new_df2_data length: {len(new_df2_data)}")

    return new_df1_data, new_df2_data

# Nouveau callback pour afficher les tableaux dans l'onglet "Affichage des Données"
@app.callback(
    Output("tabs-datatables", "children"),
    [Input("tabs", "active_tab"),
     Input("stored-df1", "data"),
     Input("stored-df2", "data")]
)
def render_data_tab_tables(active_tab, stored_df1_json, stored_df2_json):
    # Ne rien faire si l'onglet actif n'est pas "Affichage des Données"
    if active_tab != "tab-data":
        raise dash.exceptions.PreventUpdate

    tabs = []

    # Styles pour les tableaux Dash DataTable (réutilisés)
    header_style = {
        "backgroundColor": "#ADD8E6",
        "color": "#333333",
        "fontWeight": "bold",
        "border": "1px solid #B0E0B6"
    }
    cell_style = {
        "backgroundColor": "#E0FFFF",
        "color": "#333333",
        "textAlign": "left",
        "padding": "10px",
        "border": "1px solid #B0E0B6",
        "whiteSpace": "normal",
        "minWidth": "100px", "width": "150px", "maxWidth": "300px"
    }
    data_conditional_style = [
        {"if": {"row_index": "odd"},
         "backgroundColor": "#AFEEEE"}
    ]

    df1_loaded = pd.DataFrame()
    df2_loaded = pd.DataFrame()

    # Tenter de charger depuis stored-df1
    if stored_df1_json:
        print(f"DEBUG: render_data_tab_tables - stored_df1_json length: {len(stored_df1_json)}")
        try:
            df1_loaded = pd.read_json(io.StringIO(stored_df1_json), orient='split')
            print(f"DEBUG: render_data_tab_tables - df1_loaded after read_json: empty={df1_loaded.empty}, shape={df1_loaded.shape}")
            if not df1_loaded.empty:
                tabs.append(dbc.Tab(
                    label="📦 Ports / Slots / Équipements",
                    children=html.Div([
                        html.H5("📁 Fichier : Ports / Slots / Équipements (Chargé)", className="text-primary my-3"),
                        dash_table.DataTable(
                            data=df1_loaded.to_dict('records'),
                            columns=[{"name": i, "id": i} for i in df1_loaded.columns],
                            page_size=10,
                            filter_action="native",
                            sort_action="native",
                            style_table={"overflowX": "auto", "margin": "10px", "border-radius": "8px", "box-shadow": "0 4px 8px rgba(0,0,0,0.1)"},
                            style_header=header_style,
                            style_cell=cell_style,
                            style_data_conditional=data_conditional_style
                        )
                    ]),
                    tab_id="tab1"
                ))
            else:
                tabs.append(dbc.Tab(
                    label="📦 Ports / Slots / Équipements (Vide)",
                    children=html.Div("Le fichier 'Base Ports / Slots / Équipements' est vide ou n'a pas été chargé.", className="alert alert-info"),
                    tab_id="tab1-empty"
                ))
        except Exception as e:
            print(f"ERROR: render_data_tab_tables - Error loading df1 from store: {e}")
            tabs.append(dbc.Tab(
                label="📦 Ports / Slots / Équipements (Erreur Chargement)",
                children=html.Div(f"Erreur lors du chargement du fichier 1 depuis le stockage : {e}", className="alert alert-danger"),
                tab_id="tab1-load-error"
            ))

    # Tenter de charger depuis stored-df2
    if stored_df2_json:
        print(f"DEBUG: render_data_tab_tables - stored_df2_json length: {len(stored_df2_json)}")
        try:
            df2_loaded = pd.read_json(io.StringIO(stored_df2_json), orient='split')
            print(f"DEBUG: render_data_tab_tables - df2_loaded after read_json: empty={df2_loaded.empty}, shape={df2_loaded.shape}")
            if not df2_loaded.empty:
                tabs.append(dbc.Tab(
                    label="📇 Infos Abonnés",
                    children=html.Div([
                        html.H5("📁 Fichier : Infos Abonnés (Chargé)", className="text-primary my-3"),
                        dash_table.DataTable(
                            data=df2_loaded.to_dict('records'),
                            columns=[{"name": i, "id": i} for i in df2_loaded.columns],
                            page_size=10,
                            filter_action="native",
                            sort_action="native",
                            style_table={"overflowX": "auto", "margin": "10px", "border-radius": "8px", "box-shadow": "0 4px 8px rgba(0,0,0,0.1)"},
                            style_header=header_style,
                            style_cell=cell_style,
                            style_data_conditional=data_conditional_style
                        )
                    ]),
                    tab_id="tab2"
                ))
            else:
                tabs.append(dbc.Tab(
                    label="📇 Infos Abonnés (Vide)",
                    children=html.Div("Le fichier 'Base Infos Abonnés' est vide ou n'a pas été chargé.", className="alert alert-info"),
                    tab_id="tab2-empty"
                ))
        except Exception as e:
            print(f"ERROR: render_data_tab_tables - Error loading df2 from store: {e}")
            tabs.append(dbc.Tab(
                label="📇 Infos Abonnés (Erreur Chargement)",
                children=html.Div(f"Erreur lors du chargement du fichier 2 depuis le stockage : {e}", className="alert alert-danger"),
                tab_id="tab2-load-error"
            ))

    if not tabs:
        return html.Div("⚠️ Veuillez importer au moins un fichier pour afficher les données.", className="alert alert-warning text-center")

    return dbc.Tabs(tabs, id="tabs-content-display", className="mt-4")


# Callback pour mettre à jour la capacité de port globale dans le dcc.Store
@app.callback(
    Output("custom-port-capacity-store", "data"),
    Input("port-capacity-input", "value")
)
def update_custom_port_capacity(value):
    if value is None or value <= 0:
        return DEFAULT_PORT_CAPACITY # Retourne la valeur par défaut si l'entrée est vide ou invalide
    return value

# Callback pour ouvrir la modale et afficher les infos du port cliqué
@app.callback(
    [Output("splitter-modal", "is_open"),
     Output("modal-port-label", "children"),
     Output("clicked-port-info-store", "data"),
     Output("splitter-type-dropdown", "value")], # Pré-remplir le dropdown si une capacité existe déjà
    [Input("sunburst-graph", "clickData")],
    [State("splitter-modal", "is_open"),
     State("per-port-capacity-store", "data")] # État des capacités par port
)
def open_splitter_modal(clickData, is_open, per_port_capacity_data):
    if not clickData or not clickData['points']:
        raise dash.exceptions.PreventUpdate

    point = clickData['points'][0]

    # Pour Sunburst, 'id' contient le chemin complet (ex: "OLT1/Port 0<br>Actifs:...")
    # 'label' contient le nom du segment cliqué (ex: "Port 0<br>Actifs:..." ou "OLT1")
    # Nous voulons ouvrir la modale uniquement si un port est cliqué (un nœud feuille).
    # On s'assure que le 'id' contient un '/' (pour les nœuds imbriqués) et que le 'label' contient "Port"
    # pour identifier un segment de port.
    if 'id' in point and '/' in point['id'] and 'Port' in str(point.get('label', '')):
        full_path_id = point['id']
        path_segments = full_path_id.split('/')

        # L'identifiant OLT est le premier segment du chemin
        clicked_olt_id = path_segments[0]
        # Le libellé du port est le deuxième segment, qui contient "Port X<br>..."
        clicked_port_label_full = path_segments[1]
        # Nous devons extraire juste le "X" de "Port X<br>..." pour la clé.
        clicked_port_id_clean = clicked_port_label_full.split('<br>')[0].replace('Port ', '')

        port_key = f"{clicked_olt_id}-{clicked_port_id_clean}" # Clé unique pour ce port

        # Trouver le type de splitter actuel si déjà défini pour ce port
        current_splitter_type = None
        if port_key in per_port_capacity_data:
            current_capacity = per_port_capacity_data[port_key]
            for splitter_type, capacity in SPLITTER_CAPACITY_MAP.items():
                if current_capacity == capacity:
                    current_splitter_type = splitter_type
                    break
        # Si pas de capacité spécifique pour ce port, vérifier si la capacité *globale par défaut*
        # correspond à un type de splitter connu pour pré-remplir le dropdown
        elif DEFAULT_PORT_CAPACITY == SPLITTER_CAPACITY_MAP.get("Par défaut (64)", -1):
            current_splitter_type = "Par défaut (64)"

        return True, f"Configurer le Splitter pour {clicked_olt_id} - {clicked_port_id_clean}", {"olt": clicked_olt_id, "port": clicked_port_id_clean}, current_splitter_type
    else:
        # Si le clic n'est pas sur un port (par exemple, sur un nœud OLT pour zoomer), ne pas ouvrir la modale
        raise dash.exceptions.PreventUpdate

# Callback pour appliquer la capacité du splitter et fermer la modale
@app.callback(
    [Output("splitter-modal", "is_open", allow_duplicate=True), # Permettre les sorties dupliquées
     Output("per-port-capacity-store", "data")],
    [Input("apply-splitter-btn", "n_clicks")],
    [State("splitter-type-dropdown", "value"),
     State("clicked-port-info-store", "data"),
     State("per-port-capacity-store", "data")],
    prevent_initial_call=True
)
def apply_splitter_capacity(n_clicks, selected_splitter_type, clicked_port_info, current_per_port_capacity_data):
    if not n_clicks or not selected_splitter_type or not clicked_port_info:
        raise dash.exceptions.PreventUpdate

    olt = clicked_port_info.get("olt")
    port = clicked_port_info.get("port")

    if not olt or not port:
        raise dash.exceptions.PreventUpdate

    # Récupérer la capacité correspondant au type de splitter sélectionné
    new_capacity = SPLITTER_CAPACITY_MAP.get(selected_splitter_type, DEFAULT_PORT_CAPACITY)

    # Mettre à jour le dictionnaire des capacités par port
    updated_per_port_capacity = current_per_port_capacity_data.copy()
    port_key = f"{olt}-{port}"
    updated_per_port_capacity[port_key] = new_capacity

    print(f"DEBUG: apply_splitter_capacity - OLT: {olt}, Port: {port}, Selected Type: {selected_splitter_type}, New Capacity: {new_capacity}")
    print(f"DEBUG: updated_per_port_capacity: {updated_per_port_capacity}")

    return False, updated_per_port_capacity # Fermer la modale et mettre à jour le store

# Callback pour calculer et afficher le taux de saturation (y compris le Sunburst)
@app.callback(
    Output("saturation-content", "children"),
    [Input("tabs", "active_tab"),
     Input("stored-df1", "data"),
     Input("stored-df2", "data"), # df2 est toujours une entrée pour déclencher le callback
     Input("custom-port-capacity-store", "data"),
     Input("per-port-capacity-store", "data")],
    [State("clicked-port-info-store", "data")] # Ajout de clicked_port_info comme State
)
def calculate_and_display_saturation(active_tab, stored_df1_json, stored_df2_json, custom_port_capacity, per_port_capacity_data, clicked_port_info):
    if active_tab != "tab-saturation":
        raise dash.exceptions.PreventUpdate

    df_ports = pd.DataFrame()
    df_subscribers = pd.DataFrame() # df_subscribers est chargé mais non utilisé pour la saturation ici

    global_port_capacity = custom_port_capacity if custom_port_capacity is not None else DEFAULT_PORT_CAPACITY
    if not isinstance(global_port_capacity, (int, float)) or global_port_capacity <= 0:
        global_port_capacity = DEFAULT_PORT_CAPACITY

    if stored_df1_json:
        try:
            df_ports = pd.read_json(io.StringIO(stored_df1_json), orient='split')
            # print(f"DEBUG: calculate_and_display_saturation - df_ports loaded. Columns: {df_ports.columns.tolist()}, Shape: {df_ports.shape}")
        except ValueError:
            df_ports = pd.DataFrame()
            print(f"ERROR: calculate_and_display_saturation - Failed to load df_ports from JSON.")

    if df_ports.empty:
        # Si df1 est vide, on peut tenter de simuler df1 pour que la saturation puisse s'afficher
        simulated_df1_data = []
        num_olts = 3
        ports_per_olt = 5
        for i in range(num_olts):
            olt_name = f"OLT_{i+1}"
            for j in range(ports_per_olt):
                port_name = f"{j}" # Simuler le port comme un simple numéro
                num_onus = random.randint(10, 70)
                for k in range(num_onus):
                    status = random.choice(['Online', 'Offline'])
                    simulated_df1_data.append({
                        'OLT type': olt_name,
                        'Port': port_name,
                        'Running Status': status,
                        'ONU ID': f"ONU_{olt_name}_{port_name}_{k}",
                        'Device Name': olt_name,
                        'Region': random.choice(['Nord', 'Sud', 'Est', 'Ouest'])
                    })
        df_ports = pd.DataFrame(simulated_df1_data)
        if df_ports.empty:
            return html.Div("Veuillez importer la 'Base Ports / Slots / Équipements' pour calculer le taux de saturation.", className="alert alert-warning")


    # Déterminer la colonne principale pour l'identification de l'OLT
    primary_olt_col_ports = 'Device Name'
    if 'OLT type' in df_ports.columns:
        primary_olt_col_ports = 'OLT type'

    # Assurer que les colonnes nécessaires existent dans df_ports et gérer les NaNs
    required_cols_df1 = [primary_olt_col_ports, 'Port', 'Running Status']
    for col in required_cols_df1:
        if col not in df_ports.columns:
            return html.Div(f"Le fichier 'Base Ports / Slots / Équipements' doit contenir les colonnes '{primary_olt_col_ports}', 'Port' et 'Running Status'. La colonne '{col}' est manquante.", className="alert alert-danger")
        df_ports[col] = df_ports[col].fillna('Unknown')

    df_ports[primary_olt_col_ports] = df_ports[primary_olt_col_ports].astype(str)
    # FIX: Clean the 'Port' column to ensure it's just the number, matching the key stored in per_port_capacity_store
    # This is crucial for consistency between the stored `per_port_capacity_data` keys and the keys generated during calculation.
    df_ports['Port'] = df_ports['Port'].astype(str).str.replace('Port ', '')
    df_ports['Running Status'] = df_ports['Running Status'].astype(str)

    # Dériver 'Activity_Status' de 'Running Status' pour les calculs
    df_ports['Activity_Status'] = df_ports['Running Status'].apply(lambda x: 'Active' if x.lower() == 'online' else 'Inactive')

    saturation_data = []
    # Grouper par OLT et Port directement à partir de df_ports (df1)
    for group_keys, group in df_ports.groupby([primary_olt_col_ports, 'Port']):
        current_olt_identifier = group_keys[0]
        current_port = group_keys[1]

        # Compter les utilisateurs "Online" (connectés)
        connected_users_group = group[group['Running Status'].str.lower() == 'online']
        
        active_connected = connected_users_group[connected_users_group['Activity_Status'] == 'Active'].shape[0]
        inactive_connected = connected_users_group[connected_users_group['Activity_Status'] == 'Inactive'].shape[0]
        total_connected = active_connected + inactive_connected

        # Logique de capacité par port ou globale
        port_key = f"{current_olt_identifier}-{current_port}" # Cette clé est maintenant cohérente
        port_specific_capacity = per_port_capacity_data.get(port_key)
        effective_port_capacity = port_specific_capacity if port_specific_capacity is not None else global_port_capacity

        print(f"DEBUG: Saturation calculation - OLT: {current_olt_identifier}, Port: {current_port}, Port Key: {port_key}, Stored Capacity: {port_specific_capacity}, Effective Capacity: {effective_port_capacity}")


        if not isinstance(effective_port_capacity, (int, float)) or effective_port_capacity <= 0:
            effective_port_capacity = DEFAULT_PORT_CAPACITY

        saturation_rate = (total_connected / effective_port_capacity) * 100 if effective_port_capacity > 0 else 0

        row_data = {
            "OLT": current_olt_identifier,
            "Port": current_port,
            "Utilisateurs Actifs Connectés": active_connected,
            "Utilisateurs Inactifs Connectés": inactive_connected,
            "Total Utilisateurs Connectés": total_connected,
            "Taux de Saturation (%)": saturation_rate
        }
        saturation_data.append(row_data)

    df_saturation = pd.DataFrame(saturation_data)
    # print(f"DEBUG: df_saturation created. Columns: {df_saturation.columns.tolist()}, Shape: {df_saturation.shape}")
    # print(f"DEBUG: df_saturation head:\n{df_saturation.head()}")


    if df_saturation.empty:
        return html.Div("Aucune donnée de saturation à afficher. Vérifiez vos fichiers importés et les colonnes utilisées.", className="alert alert-info")

    numeric_cols_for_sunburst = ['Total Utilisateurs Connectés', 'Taux de Saturation (%)']
    for col in numeric_cols_for_sunburst:
        if col not in df_saturation.columns:
            return html.Div(f"Erreur: La colonne '{col}' est manquante dans les données de saturation.", className="alert alert-danger")
        df_saturation[col] = pd.to_numeric(df_saturation[col], errors='coerce').fillna(0)

    df_saturation_filtered = df_saturation[
        (df_saturation['Total Utilisateurs Connectés'] > 0) &
        (df_saturation['Taux de Saturation (%)'].notna())
    ].copy()
    # print(f"DEBUG: df_saturation_filtered created. Shape: {df_saturation_filtered.shape}")
    # print(f"DEBUG: df_saturation_filtered head:\n{df_saturation_filtered.head()}")

    # --- NOUVEAU DEBUG PRINT POUR LE PORT MODIFIÉ ---
    if clicked_port_info and clicked_port_info.get("olt") and clicked_port_info.get("port"):
        modified_olt = clicked_port_info["olt"]
        modified_port = clicked_port_info["port"]
        
        specific_port_data = df_saturation_filtered[
            (df_saturation_filtered['OLT'] == modified_olt) &
            (df_saturation_filtered['Port'] == modified_port)
        ]
        if not specific_port_data.empty:
            print(f"DEBUG: Saturation for modified port ({modified_olt}-{modified_port}):\n{specific_port_data[['Total Utilisateurs Connectés', 'Taux de Saturation (%)']].to_string()}")
        else:
            print(f"DEBUG: Modified port ({modified_olt}-{modified_port}) not found in df_saturation_filtered after update (might be filtered out).")
    # --- FIN NOUVEAU DEBUG PRINT ---


    if df_saturation_filtered.empty:
        return html.Div("Aucune donnée de saturation valide (utilisateurs connectés > 0 et taux de saturation valide) à afficher pour le graphique Sunburst. Veuillez vérifier vos fichiers importés.", className="alert alert-info")

    output_columns_order = ["OLT", "Port", "Utilisateurs Actifs Connectés", "Utilisateurs Inactifs Connectés", "Total Utilisateurs Connectés", "Taux de Saturation (%)"]

    # Maintenir le format "Port X" pour l'affichage dans le Sunburst, mais la clé interne est juste "X"
    df_saturation_filtered['Port_Label_With_Details'] = df_saturation_filtered.apply(
        lambda row: f"Port {row['Port']}<br>"
                    f"Actifs: <b>{row['Utilisateurs Actifs Connectés']}</b><br>"
                    f"Inactifs: <b>{row['Utilisateurs Inactifs Connectés']}</b><br>"
                    f"Saturation: {row['Taux de Saturation (%)']:.2f}%",
        axis=1
    )

    fig_sunburst = go.Figure()

    try:
        # Calculer la saturation maximale pour ajuster l'échelle de couleur
        max_saturation_value = df_saturation_filtered['Taux de Saturation (%)'].max()
        # Définir la limite supérieure de l'échelle de couleur.
        # Elle doit être au moins 100, mais s'étendre si la saturation dépasse 100 pour montrer les nuances.
        # On peut ajouter un petit facteur (e.g., 1.1) pour ne pas coller exactement au max,
        # et une limite supérieure absolue (e.g., 400) pour éviter des échelles trop grandes.
        color_range_upper_limit = max(100, min(max_saturation_value * 1.1, 400)) # Ajustez 400 si nécessaire

        fig_sunburst = px.sunburst(
            df_saturation_filtered,
            path=['OLT', 'Port_Label_With_Details'],
            values='Total Utilisateurs Connectés',
            color='Taux de Saturation (%)',
            color_continuous_scale=px.colors.sequential.Bluyl,
            range_color=[0, color_range_upper_limit], # Utilisation de la plage de couleurs dynamique
            title='Taux de Saturation par Port et OLT (Vue Hiérarchique)',
            labels={'OLT': 'Nom OLT', 'Port_Label_With_Details': 'Port et Détails', 'Taux de Saturation (%)': 'Taux de Saturation (%)'},
            hover_data=['Utilisateurs Actifs Connectés', 'Utilisateurs Inactifs Connectés', 'Total Utilisateurs Connectés', 'Taux de Saturation (%)']
        )

        fig_sunburst.update_traces(
            hovertemplate='<b>OLT: %{parent}</b><br>' +
                          'Port: <b>%{label}</b><br>' +
                          'Total Connectés: <b>%{customdata[2]}</b><br>' +
                          'Saturation: %{customdata[3]:.2f}%<extra></extra>'
        )

        fig_sunburst.update_layout(
            plot_bgcolor="#f8f9fa",
            paper_bgcolor="#f8f9fa",
            font_color="#333333",
            title_x=0.5,
            margin=dict(l=10, r=10, t=40, b=10)
        )
    except Exception as e:
        print(f"ERROR: Failed to create Sunburst chart: {e}")
        fig_sunburst = go.Figure().add_annotation(
            text=f"Erreur lors de la création du graphique Sunburst: {e}. Vérifiez vos données.",
            xref="paper", yref="paper", showarrow=False,
            font=dict(size=16, color="red")
        )
        fig_sunburst.update_layout(xaxis={"visible": False}, yaxis={"visible": False})


    header_style = {
        "backgroundColor": "#ADD8E6",
        "color": "#333333",
        "fontWeight": "bold",
        "border": "1px solid #B0E0B6"
    }
    cell_style = {
        "backgroundColor": "#E0FFFF",
        "color": "#333333",
        "textAlign": "left",
        "padding": "10px",
        "border": "1px solid #B0E0B6",
        "whiteSpace": "normal",
        "minWidth": "100px", "width": "150px", "maxWidth": "300px"
    }
    data_conditional_style = [
        {"if": {"row_index": "odd"},
         "backgroundColor": "#AFEEEE"}
    ]

    df_saturation_display = df_saturation.copy()
    df_saturation_display['Taux de Saturation (%)'] = df_saturation_display['Taux de Saturation (%)'].apply(lambda x: f"{x:.2f}%")

    return html.Div([
        dcc.Graph(id='sunburst-graph', figure=fig_sunburst, style={'height': '600px'}),
        html.Hr(className="my-4"),
        dash_table.DataTable(
            data=df_saturation_display[output_columns_order].to_dict('records'),
            columns=[{"name": i, "id": i} for i in output_columns_order],
            page_size=10,
            filter_action="native",
            sort_action="native",
            style_table={"overflowX": "auto", "margin": "10px", "border-radius": "8px", "box-shadow": "0 4px 8px rgba(0,0,0,0.1)"},
            style_header=header_style,
            style_cell=cell_style,
            style_data_conditional=data_conditional_style
        )
    ])


# --- Callbacks for the 'Statistiques' tab ---
@app.callback(
    [Output('region-choropleth-map', 'figure'),
     Output('debit-distribution-chart', 'figure'),
     Output('offer-distribution-chart', 'figure'),
     Output('offer-prediction-output', 'children')],
    [Input('tabs', 'active_tab'),
     Input('stored-df1', 'data'),
     Input('stored-df2', 'data')]
)
def update_statistics_tab(active_tab, stored_df1_json, stored_df2_json):
    if active_tab != "tab-stats":
        raise dash.exceptions.PreventUpdate

    df1_loaded = pd.DataFrame()
    df2_loaded = pd.DataFrame()

    if stored_df1_json:
        try:
            df1_loaded = pd.read_json(io.StringIO(stored_df1_json), orient='split')
        except ValueError as e:
            pass

    if stored_df2_json:
        try:
            df2_loaded = pd.read_json(io.StringIO(stored_df2_json), orient='split')
        except ValueError as e:
            pass

    # Initialize figures and prediction output
    fig_map = go.Figure()
    fig_debit = go.Figure()
    fig_offer = go.Figure()
    prediction_output = html.Div("Veuillez importer les données pour afficher les statistiques.", className="alert alert-warning")

    # --- Carte des Utilisateurs par Région (df1) ---
    current_df1_for_stats = df1_loaded
    if current_df1_for_stats.empty:
        # Si df1 est vide, nous utilisons une simulation pour les stats
        simulated_df1_data = []
        num_olts = 3
        ports_per_olt = 5
        for i in range(num_olts):
            olt_name = f"OLT_{i+1}"
            for j in range(ports_per_olt):
                port_name = f"{j}"
                num_onus = random.randint(10, 70)
                for k in range(num_onus):
                    status = random.choice(['Online', 'Offline'])
                    simulated_df1_data.append({
                        'OLT type': olt_name,
                        'Port': port_name,
                        'Running Status': status,
                        'ONU ID': f"ONU_{olt_name}_{port_name}_{k}",
                        'Device Name': olt_name,
                        'Region': random.choice(['Nord', 'Sud', 'Est', 'Ouest'])
                    })
        current_df1_for_stats = pd.DataFrame(simulated_df1_data)

    if not current_df1_for_stats.empty:
        # Normaliser les noms de colonnes pour une recherche robuste
        normalized_df1_columns_map = {
            unicodedata.normalize('NFKD', col).encode('ascii', 'ignore').decode('utf-8').strip().lower(): col
            for col in current_df1_for_stats.columns
        }
        
        # DEBUG: Afficher les noms de colonnes normalisés pour le débogage
        print(f"DEBUG: Noms de colonnes normalisés dans df1: {list(normalized_df1_columns_map.keys())}")

        region_col_name_key_search = 'region' # La clé que nous allons rechercher après normalisation
        region_original_col_name = None

        if region_col_name_key_search in normalized_df1_columns_map:
            region_original_col_name = normalized_df1_columns_map[region_col_name_key_search]
        
        if region_original_col_name:
            region_counts = current_df1_for_stats[region_original_col_name].value_counts().reset_index()
            region_counts.columns = ['Region', 'Count']
            
            # DEBUG: Print region_counts to verify data for the map
            print(f"DEBUG: Data for Region Map:\n{region_counts.to_string()}")

            # Créer des coordonnées factices pour la visualisation des cercles
            # Cela place les cercles sur une ligne horizontale avec un espacement
            region_counts['x_pos'] = region_counts.index * 10
            region_counts['y_pos'] = 0 # Tous les cercles sur la même ligne y

            fig_map = px.scatter(region_counts,
                                 x='x_pos',
                                 y='y_pos',
                                 size='Count',
                                 color='Region', # Chaque région a une couleur distincte
                                 hover_name='Region',
                                 hover_data={'Count': True, 'x_pos': False, 'y_pos': False}, # Afficher le compte au survol
                                 title='Concentration d\'Utilisateurs par Région',
                                 labels={'x_pos': '', 'y_pos': ''}, # Cacher les étiquettes des axes
                                 size_max=100, # Taille maximale des bulles, ajustez si nécessaire
                                 template='plotly_white', # Utiliser un template propre
                                 color_discrete_sequence=px.colors.sequential.Plasma) # Utilisation d'une palette plus foncée
            fig_map.update_layout(showlegend=True,
                                  xaxis={'visible': False, 'showgrid': False, 'zeroline': False},
                                  yaxis={'visible': False, 'showgrid': False, 'zeroline': False},
                                  plot_bgcolor='rgba(0,0,0,0)', # Fond transparent
                                  paper_bgcolor='rgba(0,0,0,0)', # Fond du papier transparent
                                  title_x=0.5)

            # Ajouter des annotations pour les noms de région directement sur les cercles
            for i, row in region_counts.iterrows():
                fig_map.add_annotation(
                    x=row['x_pos'],
                    y=row['y_pos'],
                    text=row['Region'],
                    showarrow=False,
                    font=dict(size=12, color="black"),
                    yanchor="middle",
                    xanchor="center"
                )
        else:
            # Message d'erreur plus précis si la colonne 'Region' est introuvable
            error_message = (
                f"La colonne 'Region' est manquante dans le fichier 'Base Ports / Slots / Équipements'. "
                f"Aucune colonne ne correspond à '{region_col_name_key_search}' après normalisation. "
                f"Colonnes originales disponibles : {', '.join(current_df1_for_stats.columns.tolist())}. "
                f"Veuillez vérifier le nom de la colonne et son encodage."
            )
            fig_map = go.Figure().add_annotation(text=error_message,
                                                  xref="paper", yref="paper", showarrow=False,
                                                  font=dict(size=16, color="red"))
            fig_map.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                                  annotations=[{"text": error_message, "showarrow": False, "font": {"size": 16, "color": "red"}}])
    else:
        fig_map = go.Figure().add_annotation(text="Fichier 1 (Base Ports / Slots / Équipements) vide ou non chargé.",
                                              xref="paper", yref="paper", showarrow=False,
                                              font=dict(size=16, color="grey"))
        fig_map.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                              annotations=[{"text": "Fichier 1 (Base Ports / Slots / Équipements) vide ou non chargé.", "showarrow": False, "font": {"size": 16, "color": "grey"}}])


    # --- Distribution des Débits (df2) ---
    current_df2_for_stats = df2_loaded
    if current_df2_for_stats.empty:
        # Simuler df2 pour la distribution des débits si le fichier n'est pas chargé
        simulated_df2_data = []
        for _ in range(200): # 200 abonnés simulés
            simulated_df2_data.append({
                'Débit': random.randint(10, 1000), # Débits entre 10 et 1000 Mbps
                'Offre': random.choice(['Offre A', 'Offre B', 'Offre C', 'Offre D'])
            })
        current_df2_for_stats = pd.DataFrame(simulated_df2_data)

    if not current_df2_for_stats.empty:
        # Normaliser les noms de colonnes pour df2
        normalized_df2_columns_map = {
            unicodedata.normalize('NFKD', col).encode('ascii', 'ignore').decode('utf-8').strip().lower(): col
            for col in current_df2_for_stats.columns
        }
        
        print(f"DEBUG: Noms de colonnes normalisés dans df2: {list(normalized_df2_columns_map.keys())}")

        debit_col_name_key_search = 'debit'
        debit_original_col_name = None

        if debit_col_name_key_search in normalized_df2_columns_map:
            debit_original_col_name = normalized_df2_columns_map[debit_col_name_key_search]

        if debit_original_col_name:
            # DEBUG: Print raw data before cleaning
            print(f"DEBUG: Raw data for '{debit_original_col_name}' (first 5 rows before cleaning):\n{current_df2_for_stats[debit_original_col_name].head()}")
            print(f"DEBUG: Dtype of '{debit_original_col_name}' before cleaning: {current_df2_for_stats[debit_original_col_name].dtype}")

            # Nettoyage robuste : supprimer tout ce qui n'est pas un chiffre ou un point
            cleaned_debit_series = current_df2_for_stats[debit_original_col_name].astype(str).apply(lambda x: re.sub(r'[^\d.]', '', x))
            
            # Attempt to convert to numeric and drop NaNs
            debit_data_cleaned = pd.to_numeric(cleaned_debit_series, errors='coerce').dropna()
            
            print(f"DEBUG: Data for '{debit_original_col_name}' (first 5 rows after cleaning):\n{debit_data_cleaned.head()}")
            print(f"DEBUG: Dtype of '{debit_original_col_name}' after cleaning: {debit_data_cleaned.dtype}")
            print(f"DEBUG: Is '{debit_original_col_name}' column empty after cleaning? {debit_data_cleaned.empty}")

            if not debit_data_cleaned.empty:
                # Utilisation d'un Histogramme (bar plot pour distribution)
                fig_debit = px.histogram(debit_data_cleaned, x=debit_data_cleaned.name,
                                         title='Distribution des Débits Internet',
                                         labels={debit_data_cleaned.name: 'Débit (Mbps)', 'count': 'Nombre d\'Abonnés'},
                                         color_discrete_sequence=px.colors.sequential.Viridis) # Utilisation d'une palette plus foncée
                fig_debit.update_layout(title_x=0.5)
                # Forcing marker color to ensure visibility
                fig_debit.update_traces(marker_color=px.colors.sequential.Viridis[3]) # Choose a specific color from the palette
            else:
                fig_debit = go.Figure().add_annotation(text=(
                    f"La colonne '{debit_original_col_name}' dans le fichier 2 est vide ou ne contient pas de données numériques valides "
                    f"après nettoyage. Veuillez vérifier le contenu de cette colonne."
                ), xref="paper", yref="paper", showarrow=False,
                                                      font=dict(size=16, color="grey"))
                fig_debit.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                                      annotations=[{"text": (
                                          f"La colonne '{debit_original_col_name}' dans le fichier 2 est vide ou ne contient pas de données numériques valides "
                                          f"après nettoyage. Veuillez vérifier le contenu de cette colonne."
                                      ), "showarrow": False, "font": {"size": 16, "color": "grey"}}])
        else:
            error_message_debit = (
                f"La colonne 'Débit' est manquante dans le fichier 'Base Infos Abonnés'. "
                f"Aucune colonne ne correspond à '{debit_col_name_key_search}' après normalisation. "
                f"Colonnes originales disponibles : {', '.join(current_df2_for_stats.columns.tolist())}. "
                f"Veuillez vérifier le nom de la colonne et son encodage."
            )
            fig_debit = go.Figure().add_annotation(text=error_message_debit,
                                                  xref="paper", yref="paper", showarrow=False,
                                                  font=dict(size=16, color="red"))
            fig_debit.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                                  annotations=[{"text": error_message_debit, "showarrow": False, "font": {"size": 16, "color": "red"}}])
    else:
        fig_debit = go.Figure().add_annotation(text="Fichier 2 (Base Infos Abonnés) vide ou non chargé.",
                                              xref="paper", yref="paper", showarrow=False,
                                              font=dict(size=16, color="grey"))
        fig_debit.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                              annotations=[{"text": "Fichier 2 (Base Infos Abonnés) vide ou non chargé.", "showarrow": False, "font": {"size": 16, "color": "grey"}}])


    # --- Analyse des Offres Souscrites (df2) ---
    # Renormaliser les noms de colonnes pour df2 pour cette section aussi
    normalized_df2_columns_map = {
        unicodedata.normalize('NFKD', col).encode('ascii', 'ignore').decode('utf-8').strip().lower(): col
        for col in df2_loaded.columns
    }
    offre_col_name_key_search = 'offre'
    offre_original_col_name = None

    if offre_col_name_key_search in normalized_df2_columns_map:
        offre_original_col_name = normalized_df2_columns_map[offre_col_name_key_search]

    if not df2_loaded.empty and offre_original_col_name:
        # DEBUG: Print raw data for 'Offre' before processing
        print(f"DEBUG: Raw data for '{offre_original_col_name}' (first 5 rows before processing):\n{df2_loaded[offre_original_col_name].head()}")
        print(f"DEBUG: Dtype of '{offre_original_col_name}' before processing: {df2_loaded[offre_original_col_name].dtype}")

        # Ensure 'Offre' column is not entirely empty after potential NaN values
        if not df2_loaded[offre_original_col_name].dropna().empty:
            offer_counts = df2_loaded[offre_original_col_name].value_counts().reset_index()
            offer_counts.columns = ['Offre', 'Count']
            fig_offer = px.pie(offer_counts, values='Count', names='Offre',
                               title='Répartition des Offres Souscrites',
                               color_discrete_sequence=px.colors.sequential.RdBu)
            fig_offer.update_traces(textposition='inside', textinfo='percent+label')
            fig_offer.update_layout(title_x=0.5)
        else:
            fig_offer = go.Figure().add_annotation(text=f"La colonne '{offre_original_col_name}' est vide dans le fichier 2. Vérifiez le contenu.",
                                                  xref="paper", yref="paper", showarrow=False,
                                                  font=dict(size=16, color="grey"))
            fig_offer.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                                  annotations=[{"text": f"La colonne '{offre_original_col_name}' est vide dans le fichier 2. Vérifiez le contenu.", "showarrow": False, "font": {"size": 16, "color": "grey"}}])
    else:
        error_message_offre = (
            f"Colonne 'Offre' manquante dans le fichier 2. Aucune colonne ne correspond à '{offre_col_name_key_search}' après normalisation. "
            f"Veuillez vérifier le nom de la colonne et son encodage."
        )
        fig_offer = go.Figure().add_annotation(text=error_message_offre,
                                              xref="paper", yref="paper", showarrow=False,
                                              font=dict(size=16, color="grey"))
        fig_offer.update_layout(xaxis={"visible": False}, yaxis={"visible": False},
                              annotations=[{"text": error_message_offre, "showarrow": False, "font": {"size": 16, "color": "grey"}}])

    # --- Prédiction d'Offre (Tendances) ---
    # Utiliser les noms de colonnes normalisés pour la prédiction
    debit_col_name_key_search = 'debit'
    offre_col_name_key_search = 'offre'
    debit_original_col_name = None
    offre_original_col_name = None

    if not df2_loaded.empty:
        normalized_df2_columns_map = {
            unicodedata.normalize('NFKD', col).encode('ascii', 'ignore').decode('utf-8').strip().lower(): col
            for col in df2_loaded.columns
        }
        if debit_col_name_key_search in normalized_df2_columns_map:
            debit_original_col_name = normalized_df2_columns_map[debit_col_name_key_search]
        if offre_col_name_key_search in normalized_df2_columns_map:
            offre_original_col_name = normalized_df2_columns_map[offre_col_name_key_search]

    if not df2_loaded.empty and debit_original_col_name and offre_original_col_name:
        df_for_prediction = df2_loaded.copy()

        # Apply robust cleaning to 'Débit' column
        df_for_prediction[debit_original_col_name] = df_for_prediction[debit_original_col_name].astype(str).apply(lambda x: re.sub(r'[^\d.]', '', x))
        df_for_prediction[debit_original_col_name] = pd.to_numeric(df_for_prediction[debit_original_col_name], errors='coerce')
        
        # DEBUG: Print data for prediction after cleaning
        print(f"DEBUG: df_for_prediction['{debit_original_col_name}'] head after cleaning:\n{df_for_prediction[debit_original_col_name].head()}")
        print(f"DEBUG: df_for_prediction['{debit_original_col_name}'] dtype after cleaning: {df_for_prediction[debit_original_col_name].dtype}")
        print(f"DEBUG: df_for_prediction['{offre_original_col_name}'] head before dropna:\n{df_for_prediction[offre_original_col_name].head()}")
        print(f"DEBUG: df_for_prediction['{offre_original_col_name}'] dtype before dropna: {df_for_prediction[offre_original_col_name].dtype}")


        df_for_prediction.dropna(subset=[debit_original_col_name, offre_original_col_name], inplace=True)
        
        print(f"DEBUG: df_for_prediction shape after dropna: {df_for_prediction.shape}")

        if not df_for_prediction.empty:
            # Calculate average debit per offer
            average_debits_per_offer = df_for_prediction.groupby(offre_original_col_name)[debit_original_col_name].mean().reset_index()
            average_debits_per_offer.columns = ['Offre', 'Débit Moyen (Mbps)']
            average_debits_per_offer = average_debits_per_offer.sort_values(by='Débit Moyen (Mbps)', ascending=False)


            # Create a Bar Plot for average debit per offer
            fig_prediction = px.bar(average_debits_per_offer, x='Offre', y='Débit Moyen (Mbps)',
                                     title='Débit Moyen Internet par Offre Souscrite',
                                     labels={'Offre': 'Offre', 'Débit Moyen (Mbps)': 'Débit Moyen (Mbps)'},
                                     color='Offre', # Color by offer for distinction
                                     color_discrete_sequence=px.colors.qualitative.Pastel) # Use a qualitative palette
            fig_prediction.update_layout(title_x=0.5)

            # Generate textual explanation based on average debits
            explanation_text = html.Div([
                html.P("Ce graphique à barres représente le débit internet moyen associé à chaque offre souscrite. Une barre plus haute indique un débit moyen plus élevé pour cette offre."),
                html.P(f"Observations principales :"),
                html.Ul([
                    html.Li(f"L'offre '{average_debits_per_offer.iloc[0]['Offre']}' a le débit moyen le plus élevé ({average_debits_per_offer.iloc[0]['Débit Moyen (Mbps)']:.2f} Mbps), suggérant qu'elle est probablement la plus performante ou la plus demandée par les utilisateurs ayant des besoins importants."),
                    html.Li(f"À l'inverse, l'offre '{average_debits_per_offer.iloc[-1]['Offre']}' affiche le débit moyen le plus faible ({average_debits_per_offer.iloc[-1]['Débit Moyen (Mbps)']:.2f} Mbps), ce qui peut indiquer une offre d'entrée de gamme ou pour des usages moins gourmands en bande passante."),
                    html.Li("Ces tendances peuvent aider à comprendre le positionnement de chaque offre sur le marché et à identifier les segments d'utilisateurs ciblés.")
                ])
            ])

            prediction_output = html.Div([
                dcc.Graph(id='prediction-bar-plot', figure=fig_prediction, style={'height': '500px'}),
                html.Hr(className="my-3"),
                explanation_text
            ])
        else:
            prediction_output = html.Div(
                f"Données insuffisantes pour la prédiction d'offres. "
                f"Assurez-vous que les colonnes '{debit_original_col_name}' et '{offre_original_col_name}' "
                f"contiennent des données valides après nettoyage et qu'il n'y a pas trop de valeurs manquantes."
                f"Colonnes normalisées disponibles dans df2: {', '.join(normalized_df2_columns_map.keys())}."
                , className="alert alert-warning")
    else:
        missing_cols = []
        if not df2_loaded.empty:
            if not debit_original_col_name:
                missing_cols.append("'Débit'")
            if not offre_original_col_name:
                missing_cols.append("'Offre'")
            prediction_output = html.Div(
                f"Colonnes manquantes pour la prédiction : {', '.join(missing_cols)}. "
                f"Veuillez vérifier les noms de colonnes dans le fichier 'Base Infos Abonnés'. "
                f"Colonnes normalisées disponibles dans df2: {', '.join(normalized_df2_columns_map.keys())}."
                , className="alert alert-warning")
        else:
            prediction_output = html.Div("Fichier 2 (Base Infos Abonnés) vide ou non chargé pour la prédiction.", className="alert alert-warning")


    return fig_map, fig_debit, fig_offer, prediction_output


if __name__ == "__main__":
    app.run(debug=True)
