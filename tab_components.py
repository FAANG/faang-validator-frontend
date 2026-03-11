"""
Reusable components for validation tabs (Samples, Experiments, etc.)
"""
from dash import dcc, html

sample_metadata_template_with_examples = '../../assets/with_examples/faang_sample.xlsx'
sample_biosample_update_template = '../../assets/with_examples/faang_update_sample.xlsx'
experiment_metadata_template_with_examples = '../../assets/with_examples/faang_experiment.xlsx'
analysis_metadata_template_with_examples = '../../assets/with_examples/faang_analysis.xlsx'
trackhubs_template_with_examples = '../../assets/with_examples/trackhubs.xlsx'

sample_metadata_template_without_examples = '../../assets/empty/faang_sample.xlsx'
experiment_metadata_template_without_examples = '../../assets/empty/faang_experiment.xlsx'
analysis_metadata_template_without_examples = '../../assets/empty/faang_analysis.xlsx'
trackhubs_template_without_examples = '../../assets/empty/trackhubs.xlsx'

tooltipUpdate = ('• This action will update the sample details with the provided metadata. \n '
                 '• Please ensure that each entry in the submitted spreadsheet contains the correct Biosample ID. \n'
                 '• The relationship columns (e.g \'Derived From\' column) should also contain the Biosample ID of the related sample. \n' +
                 '• Note that in the UPDATE spreadsheet, the column \'Sample Name\' has been replaced with \'Biosample ID\'. See provided example for updates.')




def create_file_upload_area(tab_type: str):
    """
    Create file upload component for a specific tab type.
    
    Args:
        tab_type: 'samples' or 'experiments' (used for unique IDs)
    
    Returns:
        HTML Div containing file upload component
    """
    # Extra helper buttons are only shown on the Samples / Experiments / Analysis tabs
    extra_buttons = None
    if tab_type == "samples":
        extra_buttons = html.Div(
            [
                html.A(
                    "Download example template",
                    id="samples-download-example-template-btn",
                    href=sample_metadata_template_with_examples.replace('../../assets', '/assets'),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Download empty template",
                    id="samples-download-empty-template-btn",
                    href=sample_metadata_template_without_examples.replace('../../assets', '/assets'),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Download example template for UPDATE",
                    id="samples-download-update-template-btn",
                    href=sample_biosample_update_template.replace('../../assets', '/assets'),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Upload protocol",
                    id="samples-upload-protocol-btn",
                    href="https://data.faang.org/upload_protocol?from=samples",
                    target="_blank",
                    style={
                        "backgroundColor": "green",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Submission guideline",
                    id="samples-submission-guideline-btn",
                    href="https://dcc-documentation.readthedocs.io/en/latest/sample/biosamples_template/",
                    target="_blank",
                    style={
                        "backgroundColor": "yellow",
                        "color": "black",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
            ],
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "8px",
                "marginTop": "8px",
            },
        )
    elif tab_type == "experiments":
        extra_buttons = html.Div(
            [
                html.A(
                    "Download example template",
                    id="experiments-download-example-template-btn",
                    href=experiment_metadata_template_with_examples.replace("../../assets", "/assets"),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Download empty template",
                    id="experiments-download-empty-template-btn",
                    href=experiment_metadata_template_without_examples.replace("../../assets", "/assets"),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Upload protocol",
                    id="experiments-upload-protocol-btn",
                    href="https://data.faang.org/upload_protocol?from=experiments",
                    target="_blank",
                    style={
                        "backgroundColor": "green",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Submission guideline",
                    id="experiments-submission-guideline-btn",
                    href="https://dcc-documentation.readthedocs.io/en/latest/experiment/ena_template/",
                    target="_blank",
                    style={
                        "backgroundColor": "yellow",
                        "color": "black",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
            ],
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "8px",
                "marginTop": "8px",
            },
        )
    elif tab_type == "analysis":
        extra_buttons = html.Div(
            [
                html.A(
                    "Download example template",
                    id="analysis-download-example-template-btn",
                    href=analysis_metadata_template_with_examples.replace("../../assets", "/assets"),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Download empty template",
                    id="analysis-download-empty-template-btn",
                    href=analysis_metadata_template_without_examples.replace("../../assets", "/assets"),
                    target="_blank",
                    style={
                        "backgroundColor": "#2563eb",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Upload protocol",
                    id="analysis-upload-protocol-btn",
                    href="https://data.faang.org/upload_protocol?from=analyses",
                    target="_blank",
                    style={
                        "backgroundColor": "green",
                        "color": "white",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
                html.A(
                    "Submission guideline",
                    id="analysis-submission-guideline-btn",
                    href="https://dcc-documentation.readthedocs.io/en/latest/analysis/analysis_index/",
                    target="_blank",
                    style={
                        "backgroundColor": "yellow",
                        "color": "black",
                        "padding": "8px 12px",
                        "border": "none",
                        "borderRadius": "4px",
                        "cursor": "pointer",
                        "fontSize": "13px",
                        "textDecoration": "none",
                        "display": "inline-block",
                    },
                ),
            ],
            style={
                "display": "flex",
                "flexWrap": "wrap",
                "gap": "8px",
                "marginTop": "8px",
            },
        )

    return html.Div(
        [
            extra_buttons,
            html.Label("1. Upload template", style={"marginTop": "12px", "display": "block"}),
            html.Div(
                [
                    dcc.Upload(
                        id=f"upload-data-{tab_type}",
                        children=html.Div(
                            [
                                html.Button(
                                    "Choose File",
                                    type="button",
                                    className="upload-button",
                                    n_clicks=0,
                                    style={
                                        "backgroundColor": "#cccccc",
                                        "color": "black",
                                        "padding": "10px 20px",
                                        "border": "none",
                                        "borderRadius": "4px",
                                        "cursor": "pointer",
                                    },
                                ),
                                html.Div(
                                    "No file chosen",
                                    id=f"file-chosen-text-{tab_type}",
                                ),
                            ],
                            style={
                                "display": "flex",
                                "alignItems": "center",
                                "gap": "10px",
                            },
                        ),
                        style={"width": "auto", "margin": "10px 0"},
                        className="upload-area",
                        multiple=False,
                        accept=".xlsx,.xls",  # Explicitly accept Excel files
                    ),
                    html.Div(
                        html.Button(
                            "Validate",
                            id=f"validate-button-{tab_type}",
                            className="validate-button",
                            disabled=True,
                            style={
                                "backgroundColor": "#4CAF50",
                                "color": "white",
                                "padding": "10px 20px",
                                "border": "none",
                                "borderRadius": "4px",
                                "cursor": "pointer",
                                "fontSize": "16px",
                            },
                        ),
                        id=f"validate-button-container-{tab_type}",
                        style={"display": "none", "marginLeft": "10px"},
                    ),
                    html.Div(
                        html.Button(
                            "Reset",
                            id=f"reset-button-{tab_type}",
                            n_clicks=0,
                            className="reset-button",
                            style={
                                "backgroundColor": "#f44336",
                                "color": "white",
                                "padding": "10px 20px",
                                "border": "none",
                                "borderRadius": "4px",
                                "cursor": "pointer",
                                "fontSize": "16px",
                            },
                        ),
                        id=f"reset-button-container-{tab_type}",
                        style={"display": "none", "marginLeft": "10px"},
                    ),
                ],
                style={"display": "flex", "alignItems": "center"},
            ),
            # Custom radio groups per tab so we can have fast CSS tooltips on the circles
            (
                # Samples tab custom radio
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Submit new samples"),
                            ],
                            id="biosamples-action-samples-option-submission",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip": "This action will make a new sample submission to Biosample."
                            },
                        ),
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Update existing sample"),
                            ],
                            id="biosamples-action-samples-option-update",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip":
                                    "This action will update the sample details with the provided metadata.\n"
                                    "• Please ensure that each entry in the submitted spreadsheet contains "
                                    "the correct Biosample ID.\n"
                                    "• The relationship columns (e.g. 'Derived From' column) should also "
                                    "contain the Biosample ID of the related sample.\n"
                                    "• Note that in the UPDATE spreadsheet, the column 'Sample Name' has been "
                                    "replaced with 'Biosample ID'. See provided example for updates."

                            },
                        ),
                        dcc.Store(id="biosamples-action-samples", data="submission"),
                    ],
                    className="custom-radio-group",
                    style={"marginTop": "12px"},
                )
                if tab_type == "samples"
                # Experiments tab custom radio
                else
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Submit new experiments"),
                            ],
                            id="biosamples-action-experiments-option-submission",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip": (
                                    "• This action will make a new experiment submission to ENA.\n"
                                    "• The alias used for the submitted object should be unique for the object's "
                                    "type within the submission account."
                                )
                            },
                        ),
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Update existing experiments"),
                            ],
                            id="biosamples-action-experiments-option-update",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip": (
                                    "• This action will update the experiment details with the provided metadata.\n"
                                    "• Please ensure that the submitted spreadsheet contains the original "
                                    "alias used during initial submission.\n"
                                    "• Runs cannot be updated to point to different data files."
                                )
                            },
                        ),
                        dcc.Store(
                            id="biosamples-action-experiments", data="submission"
                        ),
                    ],
                    className="custom-radio-group",
                    style={"marginTop": "12px"},
                )
                if tab_type == "experiments"
                # Analysis tab custom radio
                else
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Submit new analysis"),
                            ],
                            id="biosamples-action-analysis-option-submission",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip": (
                                    "• This action will make a new analysis submission to ENA.\n"
                                    "• The alias used for the submitted object should be unique for the object's "
                                    "type within the submission account."
                                )
                            },
                        ),
                        html.Div(
                            [
                                html.Div(className="custom-radio-circle"),
                                html.Span("Update existing analysis"),
                            ],
                            id="biosamples-action-analysis-option-update",
                            className="custom-radio-option tooltip-target",
                            **{
                                "data-tooltip": (
                                    "• This action will update the analysis details with the provided metadata.\n"
                                    "• Please ensure that the submitted spreadsheet contains the original alias "
                                    "used during initial submission.\n"
                                    "• Analysis entries cannot be updated to point to different data files."
                                )
                            },
                        ),
                        dcc.Store(
                            id="biosamples-action-analysis", data="submission"
                        ),
                    ],
                    className="custom-radio-group",
                    style={"marginTop": "12px"},
                )
            ),
            html.Div(
                id=f"selected-file-display-{tab_type}",
                style={"display": "none"},
            ),
        ],
        style={"margin": "20px 0"},
    )


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
                value="",
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
                value="",
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

            dcc.Loading(
                id=f"loading-submit-{tab_type}",
                type="circle",
                children=html.Div([
                    html.Button(
                        "Submit", id=f"biosamples-submit-btn-{tab_type}", n_clicks=0,
                        style={
                            "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                            "border": "none", "borderRadius": "8px", "cursor": "pointer",
                            "fontSize": "16px", "width": "140px"
                        }
                    ),
                    html.Div(id=f"biosamples-submit-msg-{tab_type}", style={"marginTop": "10px"}),
                ])
            ),
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
                value="",
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
                value="",
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

            dcc.Loading(
                id=f"loading-submit-{tab_type}",
                type="circle",
                children=html.Div([
                    html.Button(
                        "Submit", id=f"biosamples-submit-btn-{tab_type}", n_clicks=0,
                        style={
                            "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                            "border": "none", "borderRadius": "8px", "cursor": "pointer",
                            "fontSize": "16px", "width": "140px"
                        }
                    ),
                    html.Div(id=f"biosamples-submit-msg-{tab_type}", style={"marginTop": "10px"}),
                ])
            ),
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
        )
    ])


def create_tab_content(tab_type: str):
    """
    Create complete tab content with all components.
    
    Args:
        tab_type: 'samples', 'experiments', or 'analysis'
    
    Returns:
        HTML Div containing all tab components
    """
    # Use experiments/analysis-specific forms where needed
    if tab_type == "experiments":
        from experiments_tab import create_experiments
        biosamples_form = create_experiments()
    elif tab_type == "analysis":
        from analysis_tab import create_biosamples_form_analysis
        biosamples_form = create_biosamples_form_analysis()
    else:
        biosamples_form = create_biosamples_form(tab_type)

    # Only Samples tab needs a generic submission-results panel here;
    # Experiments and Analysis define their own panels in their forms.
    submission_panel = None
    if tab_type == "samples":
        submission_panel = html.Div(id="samples-submission-results-panel")

    return html.Div(
        [
            create_file_upload_area(tab_type),
            create_validation_results_area(tab_type),
            biosamples_form,
            submission_panel,
            html.Div(
                id=f"biosamples-results-table-{tab_type}"
            ),  # Results table at the end
        ]
    )

