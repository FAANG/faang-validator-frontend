"""
Reusable components for validation tabs (Samples, Experiments, etc.)
"""
from dash import dcc, html


def create_file_upload_area(tab_type: str):
    """
    Create file upload component for a specific tab type.
    
    Args:
        tab_type: 'samples' or 'experiments' (used for unique IDs)
    
    Returns:
        HTML Div containing file upload component
    """
    return html.Div([
        html.Label("1. Upload template"),
        html.Div([
            dcc.Upload(
                id=f'upload-data-{tab_type}',
                children=html.Div([
                    html.Button('Choose File',
                                type='button',
                                className='upload-button',
                                n_clicks=0,
                                style={
                                    'backgroundColor': '#cccccc',
                                    'color': 'black',
                                    'padding': '10px 20px',
                                    'border': 'none',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                }),
                    html.Div('No file chosen', id=f'file-chosen-text-{tab_type}')
                ], style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'}),
                style={'width': 'auto', 'margin': '10px 0'},
                className='upload-area',
                multiple=False,
                accept='.xlsx,.xls'  # Explicitly accept Excel files
            ),
            html.Div(
                html.Button(
                    'Validate',
                    id=f'validate-button-{tab_type}',
                    className='validate-button',
                    disabled=True,
                    style={
                        'backgroundColor': '#4CAF50',
                        'color': 'white',
                        'padding': '10px 20px',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px',
                    }
                ),
                id=f'validate-button-container-{tab_type}',
                style={'display': 'none', 'marginLeft': '10px'}
            ),
            html.Div(
                html.Button(
                    'Reset',
                    id=f'reset-button-{tab_type}',
                    n_clicks=0,
                    className='reset-button',
                    style={
                        'backgroundColor': '#f44336',
                        'color': 'white',
                        'padding': '10px 20px',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px',
                    }
                ),
                id=f'reset-button-container-{tab_type}',
                style={'display': 'none', 'marginLeft': '10px'}
            ),
        ], style={'display': 'flex', 'alignItems': 'center'}),
        html.Div(
            dcc.RadioItems(
                id=f'biosamples-action-{tab_type}',
                options=[
                    {"label": " Submit new sample", "value": "submission"},
                    {"label": " Update existing sample", "value": "update"},
                ],
                value="submission",
                labelStyle={"marginRight": "24px"},
                style={"marginTop": "12px"}
            )
        ),
        html.Div(id=f'selected-file-display-{tab_type}', style={'display': 'none'}),
    ], style={'margin': '20px 0'})


def create_biosamples_form(tab_type: str):
    """
    Create BioSamples submission form for a specific tab type.
    
    Args:
        tab_type: 'samples' or 'experiments' (used for unique IDs)
    
    Returns:
        HTML Div containing BioSamples form
    """
    return html.Div(
        [
            html.H2(f"Submit data to {tab_type}", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id=f"biosamples-username-{tab_type}",
                type="text",
                placeholder="Webin username",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 4px"
                }
            ),
            html.Div(
                ["Please use Webin ",
                 html.A("service", href="https://www.ebi.ac.uk/ena/submit/webin/login", target="_blank"),
                 " to get your credentials"
                 ],
                style={"color": "#64748b", "marginBottom": "12px"}
            ),

            html.Label("Password", style={"fontWeight": 600}),
            dcc.Input(
                id=f"biosamples-password-{tab_type}",
                type="password",
                placeholder="Password",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            dcc.RadioItems(
                id=f"biosamples-env-{tab_type}",
                options=[{"label": " Test server", "value": "test"},
                         {"label": " Production server", "value": "prod"}],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "16px"}
            ),

            html.Div(id=f"biosamples-status-banner-{tab_type}",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id=f"biosamples-submit-btn-{tab_type}", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id=f"biosamples-submit-msg-{tab_type}", style={"marginTop": "10px"}),
        ],
        id=f"biosamples-form-{tab_type}",
        style={"display": "none", "marginTop": "16px"},
    )


def create_ena_form(tab_type: str):
    """
    Create BioSamples submission form for a specific tab type.

    Args:
        tab_type: 'samples' or 'experiments' (used for unique IDs)

    Returns:
        HTML Div containing BioSamples form
    """
    return html.Div(
        [
            html.H2("Submit data to ENA", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id=f"biosamples-username-{tab_type}",
                type="text",
                placeholder="Webin username",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 4px"
                }
            ),
            html.Div(
                ["Please use Webin ",
                 html.A("service", href="https://www.ebi.ac.uk/ena/submit/webin/login", target="_blank"),
                 " to get your credentials"
                 ],
                style={"color": "#64748b", "marginBottom": "12px"}
            ),

            html.Label("Password", style={"fontWeight": 600}),
            dcc.Input(
                id=f"biosamples-password-{tab_type}",
                type="password",
                placeholder="Password",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            dcc.RadioItems(
                id=f"biosamples-env-{tab_type}",
                options=[{"label": " Test server", "value": "test"},
                         {"label": " Production server", "value": "prod"}],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "16px"}
            ),

            html.Div(id=f"biosamples-status-banner-{tab_type}",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id=f"biosamples-submit-btn-{tab_type}", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id=f"biosamples-submit-msg-{tab_type}", style={"marginTop": "10px"}),
        ],
        id=f"biosamples-form-{tab_type}",
        style={"display": "none", "marginTop": "16px"},
    )


def create_validation_results_area(tab_type: str):
    """
    Create validation results display area for a specific tab type.
    
    Args:
        tab_type: 'samples' or 'experiments' (used for unique IDs)
    
    Returns:
        HTML Div containing validation results area
    """
    return html.Div([
        dcc.Loading(
            id=f"loading-validation-{tab_type}",
            type="circle",
            children=html.Div(id=f'output-data-upload-{tab_type}')
        ),
        html.Div(id=f"biosamples-results-table-{tab_type}")
    ])


def create_tab_content(tab_type: str):
    """
    Create complete tab content with all components.
    
    Args:
        tab_type: 'samples' or 'experiments'
    
    Returns:
        HTML Div containing all tab components
    """
    # Use experiments-specific BioSamples form for experiments tab
    if tab_type == 'experiments':
        from experiments_tab import create_biosamples_form_experiments
        biosamples_form = create_biosamples_form_experiments()
    else:
        biosamples_form = create_biosamples_form(tab_type)
    
    return html.Div([
        create_file_upload_area(tab_type),
        create_validation_results_area(tab_type),
        biosamples_form,
    ])

