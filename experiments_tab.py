"""
Experiments tab callbacks and UI components for FAANG Validator.
This module contains all experiments-specific functionality.
"""
import json
import base64
import io
import pandas as pd
import requests
from dash import dcc, html, dash_table
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH
from dash.exceptions import PreventUpdate
import dash
import os

from file_processor import process_headers, build_json_data


def create_biosamples_form_experiments():
    """
    Create BioSamples submission form specifically for experiments tab.
    
    Returns:
        HTML Div containing BioSamples form for experiments
    """
    return html.Div(
        [
            html.H2("Submit data to ENA", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id="biosamples-username-ena",
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
                id="biosamples-password-ena",
                type="password",
                placeholder="Password",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            # dcc.RadioItems(
            #     id="biosamples-env-ena",
            #     options=[{"label": " Test server", "value": "test"},
            #              {"label": " Production server", "value": "prod"}],
            #     value="test",
            #     labelStyle={"marginRight": "18px"},
            #     style={"marginBottom": "16px"}
            # ),

            html.Div(id="biosamples-status-banner-ena",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            dcc.Loading(
                id="loading-submit-ena",
                type="circle",
                children=html.Div([
                    html.Button(
                        "Submit", id="biosamples-submit-btn-ena", n_clicks=0,
                        style={
                            "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                            "border": "none", "borderRadius": "8px", "cursor": "pointer",
                            "fontSize": "16px", "width": "140px"
                        }
                    ),
                    html.Div(id="biosamples-submit-msg-ena", style={"marginTop": "10px"}),
                ])
            ),
        ],
        id="biosamples-form-ena",
        style={"display": "none", "marginTop": "16px"},
    )

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'http://localhost:8000')


def get_all_errors_and_warnings(record):
    """Extract all errors and warnings from a validation record."""
    import re
    errors = {}
    warnings = {}

    # Build a mapping from identifiers to field names using data/model fields
    # For experiments, the backend returns identifiers like "identifier" 
    # but Excel columns use names like "Identifier"
    identifier_to_field_name = {}
    data = record.get('data', {}) or {}
    model = record.get('model', {}) or {}
    
    # Use model first (has the actual field names), then data as fallback
    source_dict = model if model else data
    
    # Create reverse mapping: lowercase identifier -> actual field name
    for field_name in source_dict.keys():
        if field_name:
            # Create identifier versions (lowercase, underscore, etc.)
            identifier_lower = field_name.lower().replace(' ', '_')
            identifier_to_field_name[identifier_lower] = field_name
            # Also map the original if it's already an identifier
            identifier_to_field_name[field_name.lower()] = field_name
            identifier_to_field_name[field_name] = field_name  # Direct match

    def map_identifier_to_field_name(identifier):
        """Map backend identifier to Excel field name."""
        if not identifier:
            return identifier
        
        # Direct match first
        if identifier in identifier_to_field_name:
            return identifier_to_field_name[identifier]
        
        # Try lowercase match
        identifier_lower = identifier.lower()
        if identifier_lower in identifier_to_field_name:
            return identifier_to_field_name[identifier_lower]
        
        # Try replacing underscores with spaces and capitalizing
        identifier_spaced = identifier.replace('_', ' ').title()
        if identifier_spaced in source_dict:
            return identifier_spaced
        
        # Try exact match in source_dict (case-insensitive)
        for field_name in source_dict.keys():
            if field_name.lower() == identifier_lower:
                return field_name
        
        # If no match found, return original (might be a special field like "Cell Type.0")
        return identifier

    if 'errors' in record and record['errors']:
        if 'field_errors' in record['errors']:
            for field, messages in record['errors']['field_errors'].items():
                # Map identifier to field name
                mapped_field = map_identifier_to_field_name(field)
                errors[mapped_field] = messages
        if 'relationship_errors' in record['errors']:
            for message in record['errors']['relationship_errors']:
                field_to_blame = 'general'
                if 'Child Of' in source_dict and source_dict.get('Child Of'):
                    field_to_blame = 'Child Of'
                elif 'Derived From' in source_dict and source_dict.get('Derived From'):
                    field_to_blame = 'Derived From'
                if field_to_blame not in errors:
                    errors[field_to_blame] = []
                errors[field_to_blame].append(message)

    if 'field_warnings' in record and record['field_warnings']:
        for field, messages in record['field_warnings'].items():
            # Map identifier to field name
            mapped_field = map_identifier_to_field_name(field)
            warnings[mapped_field] = messages

    if 'ontology_warnings' in record and record['ontology_warnings']:
        for message in record['ontology_warnings']:
            match = re.search(r"in field '([^']*)'", message)
            if match:
                field = match.group(1)
                # Map identifier to field name
                mapped_field = map_identifier_to_field_name(field)
                if mapped_field not in warnings:
                    warnings[mapped_field] = []
                warnings[mapped_field].append(message)
            else:
                if 'general' not in warnings:
                    warnings['general'] = []
                warnings['general'].append(message)

    if 'relationship_errors' in record and record['relationship_errors']:
        field_to_blame = 'general'
        if 'Child Of' in source_dict and source_dict.get('Child Of'):
            field_to_blame = 'Child Of'
        elif 'Derived From' in source_dict and source_dict.get('Derived From'):
            field_to_blame = 'Derived From'
        if field_to_blame not in warnings:
            warnings[field_to_blame] = []
        warnings[field_to_blame].extend(record['relationship_errors'])

    return errors, warnings


def _resolve_col(field, cols):
    """Resolve field name to column name (case-insensitive)."""
    if not field:
        return None
    for c in cols:
        if c.lower() == field.lower():
            return c
    return field if field in cols else None


def _valid_invalid_experiments_counts(v):
    """Get valid/invalid counts for experiments using experiment_summary"""
    try:
        s = v.get("results", {}).get("experiment_summary", {}) or {}
        return int(s.get("valid", 0)), int(s.get("invalid", 0))
    except Exception:
        return 0, 0


def register_experiments_callbacks(app):
    """
    Register all experiments tab callbacks with the Dash app.
    This function should be called from dash_app.py after the app is created.
    """
    
    # File upload callback for Experiments tab
    @app.callback(
        [Output('stored-file-data-experiments', 'data'),
         Output('stored-filename-experiments', 'data'),
         Output('file-chosen-text-experiments', 'children'),
         Output('selected-file-display-experiments', 'children'),
         Output('selected-file-display-experiments', 'style'),
         Output('output-data-upload-experiments', 'children'),
         Output('stored-all-sheets-data-experiments', 'data'),
         Output('stored-sheet-names-experiments', 'data'),
         Output('stored-parsed-json-experiments', 'data'),
         Output('active-sheet-experiments', 'data')],
        [Input('upload-data-experiments', 'contents')],
        [State('upload-data-experiments', 'filename')]
    )
    def store_file_data_experiments(contents, filename):
        """Store uploaded file data for Experiments tab"""
        if contents is None:
            return None, None, "No file chosen", [], {'display': 'none'}, [], None, None, None, None

        try:
            # Handle case where contents might not have comma (shouldn't happen but safety check)
            if ',' not in contents:
                raise ValueError("Invalid file format: missing content separator")
            
            content_type, content_string = contents.split(',', 1)  # Split only on first comma

            # Validate file type
            if not filename or not (filename.endswith('.xlsx') or filename.endswith('.xls')):
                raise ValueError(f"Invalid file type. Please upload an Excel file (.xlsx or .xls). Got: {filename}")

            # Parse Excel file to JSON immediately
            try:
                decoded = base64.b64decode(content_string)
            except Exception as e:
                raise ValueError(f"Error decoding file: {str(e)}")
            
            try:
                excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")
            except Exception as e:
                raise ValueError(f"Error reading Excel file: {str(e)}. Please ensure the file is a valid Excel file.")
            sheet_names = excel_file.sheet_names
            all_sheets_data = {}
            parsed_json_data = {}

            sheets_with_data = []

            for sheet in sheet_names:
                # Skip faang_field_values sheet
                if sheet.lower() == "faang_field_values":
                    continue
                
                df_sheet = excel_file.parse(sheet, dtype=str)
                df_sheet = df_sheet.fillna("")

                if df_sheet.empty or len(df_sheet) == 0:
                    continue

                sheet_records = df_sheet.to_dict("records")
                all_sheets_data[sheet] = sheet_records

                original_headers = [str(col) for col in df_sheet.columns]
                processed_headers = process_headers(original_headers)

                # Use vectorized conversion instead of iterrows() for much better performance
                rows = df_sheet.values.tolist()

                parsed_json_records = build_json_data(processed_headers, rows, sheet)
                parsed_json_data[sheet] = parsed_json_records
                sheets_with_data.append(sheet)

            active_sheet = sheets_with_data[0] if sheets_with_data else None
            sheet_names = sheets_with_data

            file_selected_display = html.Div([
                html.H3("File Selected", id='original-file-heading-experiments'),
                html.P(f"File: {filename}", style={'fontWeight': 'bold'})
            ])

            if len(sheets_with_data) == 0:
                output_data_upload_children = html.Div([
                    html.P("No data found in any sheet. Please upload a file with data.",
                           style={'color': 'orange', 'fontWeight': 'bold', 'margin': '20px 0'})
                ], style={'margin': '20px 0'})
            elif len(sheets_with_data) > 1:
                output_data_upload_children = html.Div([
                    dcc.Tabs(
                        id='uploaded-sheets-tabs-experiments',
                        value=active_sheet,
                        children=[],
                        style={'margin': '20px 0', 'border': 'none'},
                        colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"}
                    )
                ], style={'margin': '20px 0'})
            else:
                output_data_upload_children = html.Div([
                    html.P("Click 'Validate' to send data to backend for validation.",
                           style={'marginTop': '20px', 'fontStyle': 'italic', 'color': '#666'})
                ], style={'margin': '20px 0'})

            return (contents, filename, filename, file_selected_display,
                    {'display': 'block', 'margin': '20px 0'},
                    output_data_upload_children,
                    all_sheets_data, sheet_names, parsed_json_data, active_sheet)

        except Exception as e:
            error_display = html.Div([
                html.H5(filename),
                html.P(f"Error processing file: {str(e)}", style={'color': 'red'})
            ])
            return contents, filename, filename, error_display, {'display': 'block',
                                                                 'margin': '20px 0'}, [], None, None, None, None

    # Enable/disable validate button for Experiments tab
    @app.callback(
        [Output('validate-button-experiments', 'disabled'),
         Output('validate-button-container-experiments', 'style'),
         Output('reset-button-container-experiments', 'style')],
        [Input('stored-file-data-experiments', 'data')]
    )
    def show_and_enable_buttons_experiments(file_data):
        """Show and enable buttons for Experiments tab"""
        if file_data is None:
            return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
        else:
            return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}

    # Validation callback for Experiments tab
    @app.callback(
        [Output('output-data-upload-experiments', 'children', allow_duplicate=True),
         Output('stored-json-validation-results-experiments', 'data')],
        [Input('validate-button-experiments', 'n_clicks')],
        [State('stored-file-data-experiments', 'data'),
         State('stored-filename-experiments', 'data'),
         State('output-data-upload-experiments', 'children'),
         State('stored-all-sheets-data-experiments', 'data'),
         State('stored-sheet-names-experiments', 'data'),
         State('stored-parsed-json-experiments', 'data')],
        prevent_initial_call=True
    )
    def validate_data_experiments(n_clicks, contents, filename, current_children, all_sheets_data, sheet_names, parsed_json):
        """Validate data for Experiments tab with robust response handling"""
        if n_clicks is None or parsed_json is None:
            return current_children if current_children else html.Div([]), None

        error_data = []
        all_sheets_validation_data = {}
        print(json.dumps(parsed_json))
        try:
            try:
                response = requests.post(
                    f'{BACKEND_API_URL}/validate-data',
                    json={"data": parsed_json,
                          "data_type": "experiment" },
                    headers={'accept': 'application/json', 'Content-Type': 'application/json'}
                )
                if response.status_code != 200:
                    raise Exception(f"JSON endpoint returned {response.status_code}")
            except Exception as json_err:
                # Fallback: if JSON endpoint doesn't exist, send as file
                print(f"JSON endpoint failed: {json_err}")
            if response.status_code == 200:
                response_json = response.json()
            else:
                raise Exception(f"Error {response.status_code}: {response.text}")

            if isinstance(response_json, dict) and 'results' in response_json:
                json_validation_results = response_json
                validation_results = response_json['results']
                experiment_summary = validation_results.get('experiment_summary', {})
                total_summary = validation_results.get('total_summary', {})
                valid_count = experiment_summary.get('valid', 0)
                invalid_count = experiment_summary.get('invalid', 0)

            elif isinstance(response_json, dict):
                # Backend might return data directly without 'results' wrapper
                # Try to extract data from various possible structures
                if 'experiment_summary' in response_json or 'experiment_types_processed' in response_json:
                    # Data is at top level, wrap it in 'results'
                    json_validation_results = {"results": response_json}
                    validation_results = response_json
                    experiment_summary = validation_results.get('experiment_summary', {})
                    total_summary = validation_results.get('total_summary', {})

                    valid_count = experiment_summary.get('valid', 0)
                    invalid_count = experiment_summary.get('invalid', 0)

                else:
                    # Unknown format, try to extract what we can
                    validation_data = response_json
                    valid_count = validation_data.get('valid', 0)
                    invalid_count = validation_data.get('invalid', 0)
                    error_data = validation_data.get('warnings', [])

                    # Try to build a minimal structure
                    json_validation_results = {
                        "results": {
                            "total_summary": {
                                "valid_samples": valid_count,
                                "invalid_samples": invalid_count
                            },
                            "experiment_results": {},
                            "experiment_types_processed": []
                        }
                    }
            else:
                # Old format (list) - convert to new format for consistency
                validation_data = response_json[0] if isinstance(response_json, list) else response_json
                records = validation_data.get('validation_result', [])
                valid_count = validation_data.get('valid_experiments', 0)
                invalid_count =  validation_data.get('invalid_experiments', 0)
                error_data = validation_data.get('warnings', [])
                all_sheets_validation_data = validation_data.get('all_sheets_data', {})

                if not all_sheets_validation_data and sheet_names:
                    first_sheet = sheet_names[0]
                    all_sheets_validation_data = {first_sheet: records}

                # Convert old format to new format structure
                json_validation_results = {
                    "results": {
                        "total_summary": {
                            "valid_samples": valid_count,
                            "invalid_samples": invalid_count
                        },
                        "experiment_results": all_sheets_validation_data,
                        "experiment_types_processed": list(all_sheets_validation_data.keys()) if all_sheets_validation_data else []
                    }
                }

        except Exception as e:
            error_div = html.Div([
                html.H5(filename),
                html.P(f"Error connecting to backend API: {str(e)}", style={'color': 'red'})
            ])
            return html.Div(current_children + [error_div] if isinstance(current_children, list) else [current_children,
                                                                                                       error_div]), None

        validation_components = [
            dcc.Store(id='stored-error-data-experiments', data=error_data),
            dcc.Store(id='stored-validation-results-experiments', data={'valid_count': valid_count, 'invalid_count': invalid_count,
                                                            'all_sheets_data': all_sheets_validation_data}),
            html.H3("2. Conversion and Validation results"),

            html.Div([
                html.P("Conversion Status", style={'fontWeight': 'bold'}),
                html.P("Success", style={'color': 'green', 'fontWeight': 'bold'}),
                html.P("Validation Status", style={'fontWeight': 'bold'}),
                html.P("Finished", style={'color': 'green', 'fontWeight': 'bold'}),
            ], style={'margin': '10px 0'}),

            html.Div(id='error-table-container-experiments', style={'display': 'none'}),
            html.Div(id='validation-results-container-experiments', style={'margin': '20px 0'})
        ]

        if current_children is None:
            return html.Div(validation_components), json_validation_results if json_validation_results else {'results': {}}
        elif isinstance(current_children, list):
            return html.Div(current_children + validation_components), json_validation_results if json_validation_results else {'results': {}}
        else:
            return html.Div(validation_components + [current_children]), json_validation_results if json_validation_results else {'results': {}}

    # Reset callback for Experiments tab
    app.clientside_callback(
        """
        function(n_clicks) {
            if (n_clicks > 0) {
                window.location.reload();
            }
            return '';
        }
        """,
        Output("dummy-output-for-reset-experiments", "children"),
        [Input("reset-button-experiments", "n_clicks")],
        prevent_initial_call=True,
    )

    # BioSamples form toggle for Experiments tab
    @app.callback(
        [
            Output("biosamples-form-ena", "style"),
            Output("biosamples-status-banner-ena", "children"),
            Output("biosamples-status-banner-ena", "style"),
        ],
        Input("stored-json-validation-results-experiments", "data"),
    )
    def _toggle_biosamples_form_experiments(v):
        """Toggle BioSamples form visibility for Experiments tab"""
        base_style = {"display": "block", "marginTop": "16px"}

        if not v or "results" not in v:
            return ({"display": "none"}, "", {"display": "none"})

        valid_cnt, invalid_cnt = _valid_invalid_experiments_counts(v)
        style_ok = {
            "display": "block",
            "backgroundColor": "#e6f4ea",
            "border": "1px solid #b7e1c5",
            "color": "#137333",
            "padding": "10px 12px",
            "borderRadius": "8px",
            "marginBottom": "12px",
            "fontWeight": 500,
        }
        style_warn = {
            "display": "block",
            "backgroundColor": "#fff7e6",
            "border": "1px solid #ffd699",
            "color": "#8a6d3b",
            "padding": "10px 12px",
            "borderRadius": "8px",
            "marginBottom": "12px",
            "fontWeight": 500,
        }
        if valid_cnt > 0:
            msg_children = [
                html.Span(
                    f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid sample(s)."
                ),
                html.Br(),
            ]
            return base_style, msg_children, style_ok
        else:
            return (
                base_style,
                f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid sample(s). No valid samples to submit.",
                style_warn,
            )

    # BioSamples submit button enable/disable for Experiments tab
    @app.callback(
        Output("biosamples-submit-btn-ena", "disabled"),
        [
            Input("biosamples-username-ena", "value"),
            Input("biosamples-password-ena", "value"),
            Input("stored-json-validation-results-experiments", "data"),
        ],
    )
    def _disable_submit_experiments(u, p, v):
        """Enable/disable submit button for Experiments tab"""
        if not v or "results" not in v:
            return True
        valid_cnt, _ = _valid_invalid_experiments_counts(v)
        if valid_cnt == 0:
            return True
        return not (u and p)

    # BioSamples submission for Experiments tab
    @app.callback(
        [
            Output("biosamples-submit-msg-ena", "children"),
            Output("biosamples-results-table-experiments", "children"),
        ],
        Input("biosamples-submit-btn-ena", "n_clicks"),
        State("biosamples-username-ena", "value"),
        State("biosamples-password-ena", "value"),
        State("biosamples-env-ena", "value"),
        State("biosamples-action-experiments", "value"),
        State("stored-json-validation-results-experiments", "data"),
        prevent_initial_call=True,
    )
    def _submit_to_biosamples_experiments(n, username, password, env, action, v):
        """Submit to BioSamples for Experiments tab"""
        if not n:
            raise PreventUpdate

        if not v or "results" not in v:
            msg = html.Span(
                "No validation results available. Please validate your file first.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update

        valid_cnt, invalid_cnt = _valid_invalid_experiments_counts(v)
        if valid_cnt == 0:
            msg = html.Span(
                "No valid samples to submit. Please fix errors and re-validate.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update

        if not username or not password:
            msg = html.Span(
                "Please enter Webin username and password.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update

        validation_results = v["results"]

        body = {
            "validation_results": validation_results,
            "webin_username": username,
            "webin_password": password,
            "mode": env,
            "update_existing": action == "update",
        }

        try:
            url = f"{BACKEND_API_URL}/submit-to-biosamples"
            r = requests.post(url, json=body, timeout=600)

            if not r.ok:
                msg = html.Span(
                    f"Submission failed [{r.status_code}]: {r.text}",
                    style={"color": "#c62828", "fontWeight": 500},
                )
                return msg, dash.no_update

            data = r.json() if r.content else {}

            success = data.get("success", False)
            message = data.get("message", "No message from server")
            submitted_count = data.get("submitted_count")
            errors = data.get("errors") or []
            biosamples_ids = data.get("biosamples_ids") or {}

            color = "#388e3c" if success else "#c62828"

            msg_children = [html.Span(message, style={"fontWeight": 500})]
            if submitted_count is not None:
                msg_children += [
                    html.Br(),
                    html.Span(f"Submitted samples: {submitted_count}"),
                ]
            if errors:
                msg_children += [
                    html.Br(),
                    html.Ul(
                        [html.Li(e) for e in errors],
                        style={"marginTop": "6px", "color": "#c62828"},
                    ),
                ]

            msg = html.Div(msg_children, style={"color": color})

            if biosamples_ids:
                table_data = [
                    {"Sample Descriptor": name, "BioSample ID": acc}
                    for name, acc in biosamples_ids.items()
                ]

                for row in table_data:
                    acc = row.get("BioSample ID")
                    if acc:
                        row[
                            "BioSample ID"
                        ] = f"[{acc}](https://www.ebi.ac.uk/biosamples/samples/{acc})"

                table = dash_table.DataTable(
                    data=table_data,
                    columns=[
                        {"name": "Sample Descriptor", "id": "Sample Descriptor"},
                        {
                            "name": "BioSample ID",
                            "id": "BioSample ID",
                            "presentation": "markdown",
                        },
                    ],
                    page_size=10,
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "left"},
                )
            else:
                table = html.Div(
                    "No BioSample accessions returned.",
                    style={"marginTop": "8px", "color": "#555"},
                )

            return msg, table

        except Exception as e:
            msg = html.Span(
                f"Submission error: {e}",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update

    # Validation results display callbacks
    @app.callback(
        Output('validation-results-container-experiments', 'children'),
        [Input('stored-json-validation-results-experiments', 'data')],
        [State('stored-sheet-names-experiments', 'data'),
         State('stored-all-sheets-data-experiments', 'data')]
    )
    def populate_validation_results_tabs_experiments(validation_results, sheet_names, all_sheets_data):
        """Populate validation results tabs for experiments tab"""
        if not validation_results or 'results' not in validation_results:
            return []

        if not sheet_names:
            return []

        validation_data = validation_results['results']
        
        # Filter sheet_names to only include those in experiment_types_processed
        experiment_types_processed = validation_data.get('experiment_types_processed', []) or []
        if experiment_types_processed:
            # Only show sheets that are in experiment_types_processed
            sheet_names = [sheet for sheet in sheet_names if sheet in experiment_types_processed]

        # Calculate sheet statistics for experiments
        sheet_stats = _calculate_sheet_statistics_experiments(validation_results, all_sheets_data or {})

        sheet_tabs = []
        sheets_with_data = []

        for sheet_name in sheet_names:
            # Get statistics for this sheet
            stats = sheet_stats.get(sheet_name, {})
            errors = stats.get('error_records', 0)
            warnings = stats.get('warning_records', 0)
            valid = stats.get('valid_records', 0)
            
            # Show all sheets in experiment_types_processed, regardless of errors/warnings
            sheets_with_data.append(sheet_name)
            # Create label showing counts for THIS sheet
            label = f"{sheet_name.capitalize()} ({valid} valid / {errors} invalid)"

            sheet_tabs.append(
                dcc.Tab(
                    label=label,
                    value=sheet_name,
                    id={'type': 'sheet-validation-tab-experiments', 'sheet_name': sheet_name},
                    style={
                        'border': 'none',
                        'padding': '12px 24px',
                        'marginRight': '4px',
                        'backgroundColor': '#f5f5f5',
                        'color': '#666',
                        'borderRadius': '8px 8px 0 0',
                        'fontWeight': '500',
                        'transition': 'all 0.3s ease',
                        'cursor': 'pointer'
                    },
                    selected_style={
                        'border': 'none',
                        'borderBottom': '3px solid #4CAF50',
                        'backgroundColor': '#ffffff',
                        'color': '#4CAF50',
                        'padding': '12px 24px',
                        'marginRight': '4px',
                        'borderRadius': '8px 8px 0 0',
                        'fontWeight': 'bold',
                        'boxShadow': '0 -2px 4px rgba(0,0,0,0.1)'
                    },
                    children=[html.Div(id={'type': 'sheet-validation-content-experiments', 'index': sheet_name})]
                )
            )

        if not sheet_tabs:
            return html.Div([
                html.P(
                    "The provided data has been validated successfully with no errors or warnings. You may proceed with submission.",
                    style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
            ])

        tabs = dcc.Tabs(
            id='sheet-validation-tabs-experiments',
            value=sheets_with_data[0] if sheets_with_data else None,
            children=sheet_tabs,
            style={
                'border': 'none',
                'borderBottom': '2px solid #e0e0e0',
                'marginBottom': '20px'
            },
            colors={
                "border": "transparent",
                "primary": "#4CAF50",
                "background": "#f5f5f5"
            }
        )

        header_bar = html.Div(
            [
                html.Div(),
                html.Button(
                    "Download annotated template",
                    id="download-errors-btn-experiments",
                    n_clicks=0,
                    style={
                        'backgroundColor': '#ffd740',
                        'color': 'black',
                        'padding': '10px 20px',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px'
                    }
                ),
            ],
            style={
                'display': 'flex',
                'justifyContent': 'space-between',
                'alignItems': 'center',
                'marginBottom': '10px'
            }
        )

        return html.Div([header_bar, tabs], style={"marginTop": "8px"})

    # Clientside callback to style tab labels with colors
    app.clientside_callback(
        """
        function(validation_results) {
            if (!validation_results) {
                if (typeof window !== 'undefined' && window.dash_clientside) {
                    return window.dash_clientside.no_update;
                }
                return null;
            }
            
            if (typeof window === 'undefined' || !window.dash_clientside) {
                return null;
            }
            
            try {
                function styleTabLabels() {
                    try {
                        const tabContainer = document.getElementById('sheet-validation-tabs-experiments');
                        if (!tabContainer) {
                            return;
                        }

                        let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
                        
                        if (tabLabels.length === 0) {
                            tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"]');
                        }
                        
                        if (tabLabels.length === 0) {
                            tabLabels = tabContainer.querySelectorAll('div[role="tab"], button[role="tab"]');
                        }

                        tabLabels.forEach((tab) => {
                            try {
                                if (tab.querySelector && tab.querySelector('span[style*="color"]')) {
                                    return;
                                }

                                let originalText = tab.textContent || tab.innerText || '';
                                
                                if (!originalText || (!originalText.includes('valid') && !originalText.includes('invalid'))) {
                                    return;
                                }

                                const match = originalText.match(/\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/);
                                if (match) {
                                    const validCount = match[1];
                                    const invalidCount = match[2];
                                    
                                    const styled = originalText.replace(
                                        /\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/,
                                        '(<span style="color: #4CAF50 !important; font-weight: bold !important;">' + validCount + ' valid</span> / <span style="color: #f44336 !important; font-weight: bold !important;">' + invalidCount + ' invalid</span>)'
                                    );

                                    if (styled !== originalText) {
                                        tab.innerHTML = styled;
                                    }
                                } else {
                                    let styled = originalText.replace(
                                        /(\\d+)\\s+valid/g, 
                                        '<span style="color: #4CAF50 !important; font-weight: bold !important;">$1 valid</span>'
                                    );
                                    styled = styled.replace(
                                        /(\\d+)\\s+invalid/g, 
                                        '<span style="color: #f44336 !important; font-weight: bold !important;">$1 invalid</span>'
                                    );

                                    if (styled !== originalText && styled.includes('<span')) {
                                        tab.innerHTML = styled;
                                    }
                                }
                            } catch (e) {
                                console.error('[Tab Styling Experiments] Error processing tab:', e);
                            }
                        });
                    } catch (e) {
                        console.error('[Tab Styling Experiments] Error in styleTabLabels:', e);
                    }
                }

                // Run with multiple delays to catch tabs that render at different times
                setTimeout(styleTabLabels, 100);
                setTimeout(styleTabLabels, 300);
                setTimeout(styleTabLabels, 500);
                setTimeout(styleTabLabels, 1000);
                setTimeout(styleTabLabels, 2000);
                
                // Also set up a MutationObserver to watch for when tabs are added
                try {
                    const tabContainer = document.getElementById('sheet-validation-tabs-experiments');
                    if (tabContainer) {
                        const observer = new MutationObserver(function(mutations) {
                            setTimeout(styleTabLabels, 50);
                        });
                        observer.observe(tabContainer, {
                            childList: true,
                            subtree: true,
                            characterData: true
                        });
                    }
                } catch (e) {
                    console.error('[Tab Styling Experiments] Error setting up observer:', e);
                }
            } catch (e) {
                console.error('[Tab Styling Experiments] Error:', e);
            }
            
            return window.dash_clientside.no_update;
        }
        """,
        Output('dummy-output-tab-styling-experiments', 'children'),
        Input('stored-json-validation-results-experiments', 'data'),
        prevent_initial_call=False
    )

    # Callback to populate sheet content when tab is selected for experiments
    @app.callback(
        Output({'type': 'sheet-validation-content-experiments', 'index': MATCH}, 'children'),
        [Input('sheet-validation-tabs-experiments', 'value')],
        [State('stored-json-validation-results-experiments', 'data'),
         State('stored-all-sheets-data-experiments', 'data')]
    )
    def populate_sheet_validation_content_experiments(selected_sheet_name, validation_results, all_sheets_data):
        """Populate sheet validation content for experiments tab"""
        if validation_results is None or selected_sheet_name is None:
            return []

        if not all_sheets_data or selected_sheet_name not in all_sheets_data:
            return html.Div("No data available for this sheet.")

        return make_sheet_validation_panel_experiments(selected_sheet_name, validation_results, all_sheets_data)

    # Download annotated template callback for Experiments tab
    @app.callback(
        Output('download-table-csv-experiments', 'data'),
        Input('download-errors-btn-experiments', 'n_clicks'),
        [State('stored-json-validation-results-experiments', 'data'),
         State('stored-all-sheets-data-experiments', 'data'),
         State('stored-sheet-names-experiments', 'data')],
        prevent_initial_call=True
    )
    def download_annotated_xlsx_experiments(n_clicks, validation_results, all_sheets_data, sheet_names):
        """Download annotated Excel file with Error and Warning columns for Experiments tab"""
        if not n_clicks:
            raise PreventUpdate

        if not all_sheets_data or not sheet_names:
            raise PreventUpdate

        # Build a mapping of experiment identifiers to their field-level errors/warnings
        # Structure: {identifier_normalized: {"errors": {field: [msgs]}, "warnings": {field: [msgs]}}}
        identifier_to_field_errors = {}

        # Helper function to map backend field names to Excel column names
        def _map_field_to_column_excel(field_name, columns):
            if not field_name:
                return None

            # 1) Special case for Health Status term errors
            if field_name.startswith("Health Status") and ".term" in field_name:
                try:
                    parts = field_name.split(".")
                    idx = int(parts[1])
                except Exception:
                    idx = 0

                def _clean_col_name(col):
                    col_str = str(col)
                    if '.' in col_str:
                        return col_str.split('.')[0]
                    return col_str

                # Find all Health Status columns in order
                health_status_cols = []
                for i, col in enumerate(columns):
                    col_str = str(col)
                    if "Health Status" in col_str and "Term Source ID" not in col_str:
                        health_status_cols.append((i, col))

                # For each Health Status column, find the Term Source ID column immediately after it
                term_cols_after_health_status = []
                for hs_idx, hs_col in health_status_cols:
                    next_idx = hs_idx + 1
                    if next_idx < len(columns):
                        next_col = columns[next_idx]
                        cleaned_next = _clean_col_name(next_col)
                        if cleaned_next == "Term Source ID":
                            term_cols_after_health_status.append(next_col)

                if term_cols_after_health_status:
                    if 0 <= idx < len(term_cols_after_health_status):
                        return term_cols_after_health_status[idx]
                    return term_cols_after_health_status[-1] if term_cols_after_health_status else None

            # 2) Special case for Cell Type - map to Term Source ID column after Cell Type
            field_name_lower = field_name.lower().replace("_", " ").replace("-", " ")
            if "cell type" in field_name_lower:
                def _clean_col_name(col):
                    col_str = str(col)
                    if '.' in col_str:
                        return col_str.split('.')[0]
                    return col_str

                # Find Cell Type column
                cell_type_col_idx = None
                for i, col in enumerate(columns):
                    col_str = str(col)
                    col_str_normalized = col_str.lower().replace("_", " ").replace("-", " ")
                    if "cell type" in col_str_normalized and "term source id" not in col_str_normalized:
                        cell_type_col_idx = i
                        break

                # If Cell Type column found, find Term Source ID column immediately after it
                if cell_type_col_idx is not None:
                    next_idx = cell_type_col_idx + 1
                    if next_idx < len(columns):
                        next_col = columns[next_idx]
                        cleaned_next = _clean_col_name(next_col)
                        cleaned_next_normalized = cleaned_next.lower().replace("_", " ").replace("-", " ")
                        if cleaned_next_normalized == "term source id":
                            return next_col

            # 3) Try direct match (case-insensitive)
            direct = _resolve_col(field_name, columns)
            if direct:
                return direct

            # 4) Special handling for generic "Term Source ID"
            if field_name == "Term Source ID" or field_name.lower() == "term source id":
                for col in columns:
                    col_str = str(col)
                    if col_str.lower() == "term source id":
                        return col
                for col in columns:
                    col_str = str(col)
                    col_lower = col_str.lower()
                    if col_lower.startswith("term source id."):
                        suffix = col_lower[len("term source id."):]
                        if suffix and suffix.isdigit():
                            if col_str[:len("Term Source ID")].lower() == "term source id":
                                return col

            # 5) If field has dot notation, try using only the base name
            if "." in field_name:
                base = field_name.split(".", 1)[0]
                # Try to match the base field name (e.g., "Secondary Project" from "Secondary Project.0")
                base_match = _resolve_col(base, columns)
                if base_match:
                    return base_match
                # Also try matching columns that start with the base name (case-insensitive)
                base_lower = base.lower()
                for col in columns:
                    col_str = str(col)
                    col_lower = col_str.lower()
                    # Check if column name starts with base name (handles "Secondary Project", "Secondary Project.1", etc.)
                    if col_lower.startswith(base_lower):
                        # Check if it's the exact base or a numbered version
                        if col_lower == base_lower or (col_lower.startswith(base_lower + ".") and col_lower[len(base_lower)+1:].isdigit()):
                            return col
                        # Also match if column name equals base (without number)
                        if "." in col_str:
                            col_base = col_str.split(".", 1)[0]
                            if col_base.lower() == base_lower:
                                return col

            return None

        if validation_results and 'results' in validation_results:
            validation_data = validation_results['results']
            results_by_type = validation_data.get('experiment_results', {}) or {}
            experiment_types = validation_data.get('experiment_types_processed', []) or []

            for experiment_type in experiment_types:
                et_data = results_by_type.get(experiment_type, {}) or {}
                et_key = experiment_type.replace(' ', '_')
                invalid_key = f"invalid_{et_key}s"
                if invalid_key.endswith('ss'):
                    invalid_key = invalid_key[:-1]
                valid_key = f"valid_{et_key}s"

                # Process invalid records with errors
                invalid_records = et_data.get(invalid_key, [])
                for record in invalid_records:
                    # Try to get identifier from various possible fields
                    identifier = record.get("identifier", "") or record.get("Identifier", "")
                    if not identifier:
                        continue

                    identifier_normalized = str(identifier).strip().lower()
                    errors, warnings = get_all_errors_and_warnings(record)

                    if errors or warnings:
                        if identifier_normalized not in identifier_to_field_errors:
                            identifier_to_field_errors[identifier_normalized] = {"errors": {}, "warnings": {}}
                        if errors:
                            identifier_to_field_errors[identifier_normalized]["errors"] = errors
                        if warnings:
                            identifier_to_field_errors[identifier_normalized]["warnings"] = warnings

                # Process valid records with warnings
                valid_records = et_data.get(valid_key, [])
                for record in valid_records:
                    # Try to get identifier from various possible fields
                    identifier = record.get("identifier", "") or record.get("Identifier", "")
                    if not identifier:
                        continue

                    identifier_normalized = str(identifier).strip().lower()
                    errors, warnings = get_all_errors_and_warnings(record)

                    if errors or warnings:
                        if identifier_normalized not in identifier_to_field_errors:
                            identifier_to_field_errors[identifier_normalized] = {"errors": {}, "warnings": {}}
                        if errors:
                            # Merge errors if they exist
                            for field, msgs in errors.items():
                                if field in identifier_to_field_errors[identifier_normalized]["errors"]:
                                    existing = identifier_to_field_errors[identifier_normalized]["errors"][field]
                                    if isinstance(existing, list):
                                        existing.extend(msgs if isinstance(msgs, list) else [msgs])
                                    else:
                                        identifier_to_field_errors[identifier_normalized]["errors"][field] = [existing] + (msgs if isinstance(msgs, list) else [msgs])
                                else:
                                    identifier_to_field_errors[identifier_normalized]["errors"][field] = msgs
                        if warnings:
                            # Merge warnings if they exist
                            for field, msgs in warnings.items():
                                if field in identifier_to_field_errors[identifier_normalized]["warnings"]:
                                    existing = identifier_to_field_errors[identifier_normalized]["warnings"][field]
                                    if isinstance(existing, list):
                                        existing.extend(msgs if isinstance(msgs, list) else [msgs])
                                    else:
                                        identifier_to_field_errors[identifier_normalized]["warnings"][field] = [existing] + (msgs if isinstance(msgs, list) else [msgs])
                                else:
                                    identifier_to_field_errors[identifier_normalized]["warnings"][field] = msgs

        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            for sheet_name in sheet_names:
                if sheet_name not in all_sheets_data:
                    continue

                # Get original sheet data
                sheet_records = all_sheets_data[sheet_name]
                if not sheet_records:
                    continue

                # Convert to DataFrame
                df = pd.DataFrame(sheet_records)

                # Map field errors/warnings to columns for highlighting
                row_to_field_errors = {}  # {row_index: {"errors": {col_idx: msgs}, "warnings": {col_idx: msgs}}}
                cols_original = list(df.columns)  # Original columns

                for row_idx, record in enumerate(sheet_records):
                    # Try to find identifier in various possible column names
                    identifier = None
                    for key in ["Identifier", "identifier", "Experiment Alias", "experiment_alias"]:
                        if key in record:
                            identifier = str(record.get(key, ""))
                            break

                    if not identifier:
                        identifier = str(list(record.values())[0]) if record else ""

                    # Normalize identifier for matching
                    identifier_normalized = identifier.strip().lower() if identifier else ""

                    # Get field-level errors/warnings for this identifier
                    field_data = identifier_to_field_errors.get(identifier_normalized, {})
                    field_errors = field_data.get("errors", {})
                    field_warnings = field_data.get("warnings", {})

                    # Map field errors/warnings to column indices for highlighting
                    if field_errors or field_warnings:
                        row_to_field_errors[row_idx] = {"errors": {}, "warnings": {}}

                        # Map error fields to columns
                        for field, msgs in field_errors.items():
                            col = _map_field_to_column_excel(field, cols_original)
                            if col and col in cols_original:
                                col_idx = cols_original.index(col)
                                # Store both messages and field name for tooltip
                                row_to_field_errors[row_idx]["errors"][col_idx] = {
                                    "field": field,
                                    "messages": msgs
                                }

                        # Map warning fields to columns
                        for field, msgs in field_warnings.items():
                            col = _map_field_to_column_excel(field, cols_original)
                            if col and col in cols_original:
                                col_idx = cols_original.index(col)
                                # Store both messages and field name for tooltip
                                row_to_field_errors[row_idx]["warnings"][col_idx] = {
                                    "field": field,
                                    "messages": msgs
                                }

                # Write to Excel
                sheet_name_clean = sheet_name[:31]  # Excel sheet name limit
                df.to_excel(writer, sheet_name=sheet_name_clean, index=False)

                # Get Excel formatting objects
                book = writer.book
                fmt_red = book.add_format({"bg_color": "#FFCCCC"})
                fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})

                ws = writer.sheets[sheet_name_clean]
                cols = list(df.columns)  # Original columns only

                # Helper function to format messages for tooltip
                def format_tooltip_message(field_name, msgs, is_warning=False):
                    """Format error/warning messages for Excel comment/tooltip."""
                    msgs_list = msgs if isinstance(msgs, list) else [msgs]
                    prefix = "Warning" if is_warning else "Error"
                    # Join messages with line breaks for better readability
                    formatted = f"{prefix} - {field_name}:\n"
                    formatted += "\n".join([f"• {str(msg)}" for msg in msgs_list])
                    # Excel comments have a limit, so truncate if too long
                    max_length = 2000
                    if len(formatted) > max_length:
                        formatted = formatted[:max_length] + "..."
                    return formatted

                # Highlight specific cells with errors/warnings and add tooltips
                for row_idx, record in enumerate(sheet_records):
                    excel_row = row_idx + 1  # Excel is 1-indexed

                    if row_idx in row_to_field_errors:
                        field_data = row_to_field_errors[row_idx]

                        # Highlight error cells (red) - use original column indices
                        for col_idx, error_data in field_data.get("errors", {}).items():
                            if col_idx < len(cols_original):
                                cell_value = df.iat[row_idx, col_idx] if row_idx < len(df) else ""
                                ws.write(excel_row, col_idx, cell_value, fmt_red)
                                # Add tooltip/comment with error message
                                field_name = error_data.get("field",
                                                            cols_original[col_idx] if col_idx < len(cols_original) else "")
                                msgs = error_data.get("messages", [])
                                tooltip_text = format_tooltip_message(field_name, msgs, is_warning=False)
                                ws.write_comment(excel_row, col_idx, tooltip_text,
                                                 {"visible": False, "x_scale": 1.5, "y_scale": 1.8})

                        # Highlight warning cells (yellow) - use original column indices
                        for col_idx, warning_data in field_data.get("warnings", {}).items():
                            if col_idx < len(cols_original):
                                cell_value = df.iat[row_idx, col_idx] if row_idx < len(df) else ""
                                ws.write(excel_row, col_idx, cell_value, fmt_yellow)
                                # Add tooltip/comment with warning message
                                field_name = warning_data.get("field", cols_original[col_idx] if col_idx < len(
                                    cols_original) else "")
                                msgs = warning_data.get("messages", [])
                                tooltip_text = format_tooltip_message(field_name, msgs, is_warning=True)
                                ws.write_comment(excel_row, col_idx, tooltip_text,
                                                 {"visible": False, "x_scale": 1.5, "y_scale": 1.8})

        buffer.seek(0)
        return dcc.send_bytes(buffer.getvalue(), "annotated_template_experiments.xlsx")


def make_sheet_validation_panel_experiments(sheet_name: str, validation_results: dict, all_sheets_data: dict):
    """Create a panel showing validation results for experiments sheet"""
    import uuid
    panel_id = str(uuid.uuid4())

    # Get sheet data
    sheet_records = all_sheets_data.get(sheet_name, [])
    if not sheet_records:
        return html.Div([html.H4("No data available", style={'textAlign': 'center', 'margin': '10px 0'})])

    # Get validation data - use experiment_types_processed for experiments
    validation_data = validation_results.get('results', {})
    results_by_type = validation_data.get('experiment_results', {}) or {}
    experiment_summary = validation_data.get('experiment_summary', {})
    experiment_types = validation_data.get('experiment_types_processed', []) or []

    # Get all validation rows for this sheet

    sheet_sample_names = set()
    for record in sheet_records:
        identifier = str(record.get("Identifier", "") or record.get("identifier", "") )
        sheet_sample_names.add(identifier)


    # Collect all rows that belong to this sheet
    error_map = {}
    warning_map = {}

    for experiment_type in experiment_types:
        et_data = results_by_type.get(experiment_type, {}) or {}
        et_key = experiment_type.replace(' ', '_')
        invalid_key = f"invalid_{et_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
        valid_key = f"valid_{et_key}s"

        invalid_records = et_data.get(invalid_key, [])
        valid_records = et_data.get(valid_key, [])

        for record in invalid_records + valid_records:
            sample_descriptor =  record.get("identifier", "") or record.get("Identifier", "") or record.get("Sample Descriptor", "") or record.get("sample_descriptor", "")
            if sample_descriptor in sheet_sample_names:
                errors, warnings = get_all_errors_and_warnings(record)
                if errors:
                    # Merge errors if sample_descriptor already exists in error_map
                    if sample_descriptor in error_map:
                        # Merge the error dictionaries
                        for field, msgs in errors.items():
                            if field in error_map[sample_descriptor]:
                                # Merge messages if field already exists
                                existing_msgs = error_map[sample_descriptor][field]
                                if isinstance(existing_msgs, list):
                                    if isinstance(msgs, list):
                                        error_map[sample_descriptor][field] = existing_msgs + msgs
                                    else:
                                        error_map[sample_descriptor][field] = existing_msgs + [msgs]
                                else:
                                    if isinstance(msgs, list):
                                        error_map[sample_descriptor][field] = [existing_msgs] + msgs
                                    else:
                                        error_map[sample_descriptor][field] = [existing_msgs, msgs]
                            else:
                                error_map[sample_descriptor][field] = msgs
                    else:
                        error_map[sample_descriptor] = errors
                if warnings:
                    # Merge warnings if sample_descriptor already exists in warning_map
                    if sample_descriptor in warning_map:
                        # Merge the warning dictionaries
                        for field, msgs in warnings.items():
                            if field in warning_map[sample_descriptor]:
                                # Merge messages if field already exists
                                existing_msgs = warning_map[sample_descriptor][field]
                                if isinstance(existing_msgs, list):
                                    if isinstance(msgs, list):
                                        warning_map[sample_descriptor][field] = existing_msgs + msgs
                                    else:
                                        warning_map[sample_descriptor][field] = existing_msgs + [msgs]
                                else:
                                    if isinstance(msgs, list):
                                        warning_map[sample_descriptor][field] = [existing_msgs] + msgs
                                    else:
                                        warning_map[sample_descriptor][field] = [existing_msgs, msgs]
                            else:
                                warning_map[sample_descriptor][field] = msgs
                    else:
                        warning_map[sample_descriptor] = warnings

    # Create DataFrame from sheet records
    df_all = pd.DataFrame(sheet_records)
    if df_all.empty:
        return html.Div([html.H4("No data available", style={'textAlign': 'center', 'margin': '10px 0'})])

    # Use the same styling logic
    def _as_list(msgs):
        if isinstance(msgs, list):
            return [str(m) for m in msgs]
        return [str(msgs)]

    def _map_field_to_column(field_name, columns):
        if not field_name:
            return None
        
        # Handle Health Status fields (with or without .term)
        if "Health Status" in field_name:
            if ".term" in field_name:
                try:
                    parts = field_name.split(".")
                    idx = int(parts[1])
                except Exception:
                    idx = 0

                def _clean_col_name(col):
                    col_str = str(col)
                    if '.' in col_str:
                        return col_str.split('.')[0]
                    return col_str

                health_status_cols = []
                for i, col in enumerate(columns):
                    col_str = str(col)
                    if "Health Status" in col_str and "Term Source ID" not in col_str:
                        health_status_cols.append((i, col))
                term_cols_after_health_status = []
                for hs_idx, hs_col in health_status_cols:
                    next_idx = hs_idx + 1
                    if next_idx < len(columns):
                        next_col = columns[next_idx]
                        cleaned_next = _clean_col_name(next_col)
                        if cleaned_next == "Term Source ID":
                            term_cols_after_health_status.append(next_col)
                if term_cols_after_health_status:
                    if 0 <= idx < len(term_cols_after_health_status):
                        return term_cols_after_health_status[idx]
                    return term_cols_after_health_status[-1] if term_cols_after_health_status else None
            else:
                # Health Status without .term - find the first Health Status column
                for col in columns:
                    col_str = str(col)
                    if "Health Status" in col_str and "Term Source ID" not in col_str:
                        return col
        
        # Handle Cell Type fields (with or without .term)
        if "Cell Type" in field_name:
            if ".term" in field_name:
                try:
                    parts = field_name.split(".")
                    idx = int(parts[1])
                except Exception:
                    idx = 0

                def _clean_col_name(col):
                    col_str = str(col)
                    if '.' in col_str:
                        return col_str.split('.')[0]
                    return col_str

                cell_type_cols = []
                for i, col in enumerate(columns):
                    col_str = str(col)
                    if "Cell Type" in col_str and "Term Source ID" not in col_str:
                        cell_type_cols.append((i, col))
                term_cols_after_cell_type = []
                for ct_idx, ct_col in cell_type_cols:
                    next_idx = ct_idx + 1
                    if next_idx < len(columns):
                        next_col = columns[next_idx]
                        cleaned_next = _clean_col_name(next_col)
                        if cleaned_next == "Term Source ID":
                            term_cols_after_cell_type.append(next_col)
                if term_cols_after_cell_type:
                    if 0 <= idx < len(term_cols_after_cell_type):
                        return term_cols_after_cell_type[idx]
                    return term_cols_after_cell_type[-1] if term_cols_after_cell_type else None
            else:
                # Cell Type without .term - find the first Cell Type column
                for col in columns:
                    col_str = str(col)
                    if "Cell Type" in col_str and "Term Source ID" not in col_str:
                        return col
        
        # Handle Term Source ID fields - check if field starts with "Term Source ID"
        if field_name and field_name.lower().startswith("term source id"):
            # Find the first "Term Source ID" column (without numbered suffix first, then with suffix)
            for col in columns:
                col_str = str(col)
                col_lower = col_str.lower()
                # Try exact match first
                if col_lower == "term source id":
                    return col
            # If no exact match, try numbered versions (Term Source ID.1, Term Source ID.2, etc.)
            for col in columns:
                col_str = str(col)
                col_lower = col_str.lower()
                if col_lower.startswith("term source id."):
                    # This is a numbered Term Source ID column - return the first one found
                    return col
            # If still no match, try columns that contain "Term Source ID" (case insensitive)
            for col in columns:
                col_str = str(col)
                if "term source id" in col_str.lower():
                    return col
        
        # Try direct match first
        direct = _resolve_col(field_name, columns)
        if direct:
            return direct
        
        # Handle fields with dot notation (e.g., "Secondary Project.0", "Field.1", etc.)
        if "." in field_name:
            base = field_name.split(".", 1)[0]
            # Try to match the base field name (e.g., "Secondary Project" from "Secondary Project.0")
            base_match = _resolve_col(base, columns)
            if base_match:
                return base_match
            # Also try matching columns that start with the base name (case-insensitive)
            base_lower = base.lower()
            for col in columns:
                col_str = str(col)
                col_lower = col_str.lower()
                # Check if column name starts with base name (handles "Secondary Project", "Secondary Project.1", etc.)
                if col_lower.startswith(base_lower):
                    # Check if it's the exact base or a numbered version
                    if col_lower == base_lower or (col_lower.startswith(base_lower + ".") and col_lower[len(base_lower)+1:].isdigit()):
                        return col
                    # Also match if column name equals base (without number)
                    if "." in col_str:
                        col_base = col_str.split(".", 1)[0]
                        if col_base.lower() == base_lower:
                            return col
        
        return None

    # Build cell styles and tooltips
    cell_styles = []
    tooltip_data = []

    for i, row in df_all.iterrows():
        # Try to match by Identifier or Sample Descriptor
        identifier = str(row.get("Identifier", "") or row.get("identifier", "") or row.get("Sample Descriptor", "") or row.get("sample_descriptor", ""))

        match_key = identifier
        
        tips = {}
        row_styles = []

        # Try exact match first
        field_errors = {}
        if match_key in error_map:
            field_errors = error_map[match_key] or {}
        else:
            # Try case-insensitive match
            match_key_lower = match_key.lower() if match_key else ""
            for key in error_map.keys():
                if str(key).strip().lower() == match_key_lower:
                    field_errors = error_map[key] or {}
                    break
        
        # Process errors if any found
        if field_errors:
            for field, msgs in field_errors.items():
                msgs_list = _as_list(msgs)
                lower_msgs = [m.lower() for m in msgs_list]
                lower_field = field.lower()
                
                # Check if error messages or field name mention "Health Status", "Cell Type", or "Term Source ID"
                # and map to those columns even if field name doesn't match exactly
                col = None
                if "health status" in lower_field or any("health status" in lm for lm in lower_msgs):
                    col = _map_field_to_column("Health Status", df_all.columns)
                elif "cell type" in lower_field or any("cell type" in lm for lm in lower_msgs):
                    col = _map_field_to_column("Cell Type", df_all.columns)
                elif lower_field.startswith("term source id") or any(lm.startswith("term source id") for lm in lower_msgs):
                    # Map to first Term Source ID column
                    col = _map_field_to_column("Term Source ID", df_all.columns)
                
                # If not found via message check, use normal mapping (this will handle "Secondary Project.0" etc.)
                if not col:
                    col = _map_field_to_column(field, df_all.columns)
                
                # If still not found, try to extract base name from dot notation
                if not col and "." in field:
                    base_field = field.split(".", 1)[0]
                    col = _map_field_to_column(base_field, df_all.columns)
                
                # Last resort: use field name as-is (but this should rarely happen)
                if not col:
                    col = field

                col_id = None
                # First try exact match
                if col in df_all.columns:
                    col_id = col
                else:
                    # Try case-insensitive match and partial match
                    col_str = str(col)
                    col_lower = col_str.lower().strip()
                    for df_col in df_all.columns:
                        df_col_str = str(df_col).strip()
                        df_col_lower = df_col_str.lower()
                        
                        # Exact match (case-insensitive)
                        if df_col_lower == col_lower:
                            col_id = df_col
                            break
                        
                        # Check if column name equals base name (for fields like "Secondary Project.0" matching "Secondary Project")
                        if "." in df_col_str:
                            df_col_base = df_col_str.split(".", 1)[0].strip().lower()
                            if df_col_base == col_lower:
                                col_id = df_col
                                break
                        
                        # Check if column name starts with the field name (for numbered columns like "Secondary Project.1")
                        if df_col_lower.startswith(col_lower + "."):
                            # Check if what follows is a number
                            remainder = df_col_lower[len(col_lower)+1:]
                            if remainder and (remainder.split('.')[0].isdigit() or remainder.split('.')[0] == ''):
                                col_id = df_col
                                break
                
                if not col_id:
                    # Debug: print what we're looking for vs what's available (only for first few to avoid spam)
                    if i < 3:
                        print(f"DEBUG Experiments: Could not find column '{col}' for field '{field}'. Available columns: {list(df_all.columns)[:15]}")
                    continue

                is_extra = any("extra inputs are not permitted" in lm for lm in lower_msgs)
                is_warning_like = any("warning" in lm for lm in lower_msgs)
                
                # Format error message for tooltip
                # Use markdown format with bold error label
                if is_extra or is_warning_like:
                    prefix = "**Warning**: "
                    bg_color = '#fff4cc'  # Yellow for warnings
                else:
                    prefix = "**Error**: "
                    bg_color = '#ffcccc'  # Red for errors
                
                # Format the message - show field name and error messages
                # Clean field name for display (remove .0, .1, etc.)
                field_display = field
                if '.' in field:
                    # Remove numbered suffixes like .0, .1, .2, etc.
                    parts = field.split('.')
                    if len(parts) > 1 and parts[-1].isdigit():
                        field_display = '.'.join(parts[:-1])
                
                # Escape any special characters that might break markdown
                msg_text = prefix + field_display + " — " + " | ".join(msgs_list)
                
                # Combine with existing tooltip if column already has one
                # Use pipe separator like in analysis tab for consistency
                if col_id in tips:
                    existing = tips[col_id].get("value", "")
                    combined = f"{existing} | {msg_text}" if existing else msg_text
                else:
                    combined = msg_text
                
                # Apply styling and tooltip
                row_styles.append({'if': {'row_index': i, 'column_id': col_id}, 'backgroundColor': bg_color})
                tips[col_id] = {'value': combined, 'type': 'markdown'}

        # Try exact match first for warnings
        field_warnings = {}
        if match_key in warning_map:
            field_warnings = warning_map[match_key] or {}
        else:
            # Try case-insensitive match
            match_key_lower = match_key.lower() if match_key else ""
            for key in warning_map.keys():
                if str(key).strip().lower() == match_key_lower:
                    field_warnings = warning_map[key] or {}
                    break
        
        if field_warnings:
            for field, msgs in field_warnings.items():
                msgs_list = _as_list(msgs)
                lower_msgs = [m.lower() for m in msgs_list]
                lower_field = field.lower()
                
                # Check if warning messages or field name mention "Health Status", "Cell Type", or "Term Source ID"
                # and map to those columns even if field name doesn't match exactly
                col = None
                if "health status" in lower_field or any("health status" in lm for lm in lower_msgs):
                    col = _map_field_to_column("Health Status", df_all.columns)
                elif "cell type" in lower_field or any("cell type" in lm for lm in lower_msgs):
                    col = _map_field_to_column("Cell Type", df_all.columns)
                elif lower_field.startswith("term source id") or any(lm.startswith("term source id") for lm in lower_msgs):
                    # Map to first Term Source ID column
                    col = _map_field_to_column("Term Source ID", df_all.columns)
                
                # If not found via message check, use normal mapping (this will handle "Secondary Project.0" etc.)
                if not col:
                    col = _map_field_to_column(field, df_all.columns)
                
                # If still not found, try to extract base name from dot notation
                if not col and "." in field:
                    base_field = field.split(".", 1)[0]
                    col = _map_field_to_column(base_field, df_all.columns)
                
                # Last resort: use field name as-is (but this should rarely happen)
                if not col:
                    col = field

                col_id = None
                # First try exact match
                if col in df_all.columns:
                    col_id = col
                else:
                    # Try case-insensitive match and partial match
                    col_str = str(col)
                    col_lower = col_str.lower().strip()
                    for df_col in df_all.columns:
                        df_col_str = str(df_col).strip()
                        df_col_lower = df_col_str.lower()
                        
                        # Exact match (case-insensitive)
                        if df_col_lower == col_lower:
                            col_id = df_col
                            break
                        
                        # Check if column name equals base name (for fields like "Secondary Project.0" matching "Secondary Project")
                        if "." in df_col_str:
                            df_col_base = df_col_str.split(".", 1)[0].strip().lower()
                            if df_col_base == col_lower:
                                col_id = df_col
                                break
                        
                        # Check if column name starts with the field name (for numbered columns like "Secondary Project.1")
                        if df_col_lower.startswith(col_lower + "."):
                            # Check if what follows is a number
                            remainder = df_col_lower[len(col_lower)+1:]
                            if remainder and (remainder.split('.')[0].isdigit() or remainder.split('.')[0] == ''):
                                col_id = df_col
                                break
                
                if not col_id:
                    # Debug: print what we're looking for vs what's available (only for first few to avoid spam)
                    if i < 3:
                        print(f"DEBUG Experiments: Could not find column '{col}' for field '{field}'. Available columns: {list(df_all.columns)[:15]}")
                    continue
                warn_text = "**Warning**: " + (field if field else 'General') + " — " + " | ".join(msgs_list)
                if col_id in tips:
                    existing = tips[col_id].get("value", "")
                    combined = f"{existing} | {warn_text}" if existing else warn_text
                else:
                    combined = warn_text
                tips[col_id] = {'value': combined, 'type': 'markdown'}
                row_styles.append({'if': {'row_index': i, 'column_id': col_id}, 'backgroundColor': '#fff4cc'})

        cell_styles.extend(row_styles)
        tooltip_data.append(tips)

    # Ensure tooltip_data has the same length as the number of rows
    # Also ensure all entries are dictionaries (not None or other types)
    while len(tooltip_data) < len(df_all):
        tooltip_data.append({})
    tooltip_data = tooltip_data[:len(df_all)]
    
    # Clean tooltip_data to ensure all values are valid dictionaries
    for i, tip_dict in enumerate(tooltip_data):
        if not isinstance(tip_dict, dict):
            tooltip_data[i] = {}
        else:
            # Ensure all tooltip values are properly formatted dictionaries
            cleaned_dict = {}
            for col_id, tip_value in tip_dict.items():
                if isinstance(tip_value, dict):
                    # Already a dict, just ensure value is a string if present
                    if 'value' in tip_value and tip_value.get('value'):
                        cleaned_dict[col_id] = {
                            'value': str(tip_value['value']),
                            'type': tip_value.get('type', 'markdown')
                        }
                # If tip_value is not a dict, skip it (shouldn't happen with our code)
            tooltip_data[i] = cleaned_dict

    def clean_header_name(header):
        if '.' in header:
            return header.split('.')[0]
        return header

    columns = [{"name": clean_header_name(c), "id": c} for c in df_all.columns]

    # Calculate statistics for this sheet
    total_records = len(sheet_records)
    error_records = len([k for k in sheet_sample_names if k in error_map])
    warning_records = len([k for k in sheet_sample_names if k in warning_map and k not in error_map])
    valid_records = total_records - error_records

    # Count errors and warnings by field
    error_fields_count = {}
    for sample_descriptor, field_errors in error_map.items():
        for field in field_errors.keys():
            error_fields_count[field] = error_fields_count.get(field, 0) + 1

    warning_fields_count = {}
    for sample_descriptor, field_warnings in warning_map.items():
        for field in field_warnings.keys():
            warning_fields_count[field] = warning_fields_count.get(field, 0) + 1

    # Build report sections
    report_sections = []

    # Errors by field
    if error_fields_count:
        error_items = sorted(error_fields_count.items(), key=lambda x: x[1], reverse=True)
        report_sections.append(
            html.Div([
                html.H5("Errors by Field", style={
                    'marginTop': '20px',
                    'marginBottom': '15px',
                    'color': '#f44336',
                    'borderBottom': '2px solid #f44336',
                    'paddingBottom': '8px'
                }),
                html.Ul([
                    html.Li([
                        html.Span(f"{field}: ", style={'fontWeight': 'bold'}),
                        html.Span(f"{count} error(s)", style={'color': '#666'})
                    ], style={'marginBottom': '6px'})
                    for field, count in error_items
                ], style={'padding': '10px', 'backgroundColor': '#ffebee', 'borderRadius': '6px',
                          'listStylePosition': 'inside'})
            ])
        )

    # Warnings by field
    if warning_fields_count:
        warning_items = sorted(warning_fields_count.items(), key=lambda x: x[1], reverse=True)
        report_sections.append(
            html.Div([
                html.H5("Warnings by Field", style={
                    'marginTop': '20px',
                    'marginBottom': '15px',
                    'color': '#ff9800',
                    'borderBottom': '2px solid #ff9800',
                    'paddingBottom': '8px'
                }),
                html.Ul([
                    html.Li([
                        html.Span(f"{field}: ", style={'fontWeight': 'bold'}),
                        html.Span(f"{count} warning(s)", style={'color': '#666'})
                    ], style={'marginBottom': '6px'})
                    for field, count in warning_items
                ], style={'padding': '10px', 'backgroundColor': '#fff3e0', 'borderRadius': '6px',
                          'listStylePosition': 'inside'})
            ])
        )

    # Total Summary
    if experiment_summary:
        summary_items = []
        for key, value in experiment_summary.items():
            if isinstance(value, (int, float, str)):
                display_key = key.replace('_', ' ').title()
                summary_items.append(
                    html.Div([
                        html.Span(f"{display_key}: ", style={'fontWeight': 'bold'}),
                        html.Span(str(value), style={'color': '#666'})
                    ], style={'marginBottom': '8px'})
                )
        if summary_items:
            report_sections.append(
                html.Div([
                    html.H5("Total Summary", style={
                        'marginTop': '20px',
                        'marginBottom': '15px',
                        'color': '#2196F3',
                        'borderBottom': '2px solid #2196F3',
                        'paddingBottom': '8px'
                    }),
                    html.Div(
                        summary_items,
                        style={'padding': '10px', 'backgroundColor': '#e3f2fd', 'borderRadius': '6px'}
                    )
                ])
            )

    zebra = [{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]

    blocks = [
        html.H4(f"Validation Results - {sheet_name.capitalize()}", style={'textAlign': 'center', 'margin': '10px 0'}),
        html.Div([
            DataTable(
                id={"type": "sheet-result-table-experiments", "sheet_name": sheet_name, "panel_id": panel_id},
                data=df_all.to_dict("records"),
                columns=columns,
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left", "padding": "6px"},
                style_header={"fontWeight": "bold", "backgroundColor": "rgb(230, 230, 230)"},
                style_data_conditional=zebra + cell_styles,
                tooltip_data=tooltip_data,
                tooltip_duration=None
            )
        ], id={"type": "sheet-table-container-experiments", "sheet_name": sheet_name, "panel_id": panel_id},
            style={'display': 'block'}),
        # Add validation report after the table
        html.Div(
            report_sections,
            style={
                'marginTop': '30px',
                'padding': '20px',
                'backgroundColor': '#ffffff',
                'border': '1px solid #e0e0e0',
                'borderRadius': '8px',
                'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'
            }
        )
    ]

    return html.Div(blocks)


def _calculate_sheet_statistics_experiments(validation_results, all_sheets_data):
    """Calculate errors and warnings count for each Excel sheet for experiments"""
    sheet_stats = {}

    if not validation_results or 'results' not in validation_results:
        return sheet_stats

    validation_data = validation_results['results']
    results_by_type = validation_data.get('experiment_results', {}) or {}
    experiment_types = validation_data.get('experiment_types_processed', []) or []

    # Initialize sheet stats
    if all_sheets_data:
        for sheet_name in all_sheets_data.keys():
            sheet_stats[sheet_name] = {
                'total_records': len(all_sheets_data[sheet_name]) if all_sheets_data[sheet_name] else 0,
                'valid_records': 0,
                'error_records': 0,
                'warning_records': 0,
                'sample_status': {}
            }

    # Process each experiment type and map to sheets
    for experiment_type in experiment_types:
        et_data = results_by_type.get(experiment_type, {}) or {}

        invalid_key = "invalid"
        valid_key = "valid"
        
        invalid_records = et_data.get(invalid_key, [])
        valid_records = et_data.get(valid_key, [])

        all_records = invalid_records + valid_records

        for record in all_records:
            sample_descriptor = record.get("identifier", "") or record.get("Identifier", "") or record.get("Sample Descriptor", "") or record.get("sample_descriptor", "")
            if not sample_descriptor:
                continue

            errors, warnings = get_all_errors_and_warnings(record)

            # Find which sheet contains this sample
            # For experiments, try "Identifier" first, then "Sample Descriptor"
            for sheet_name, sheet_records in (all_sheets_data or {}).items():
                if not sheet_records:
                    continue
                sheet_sample_names = set()
                for r in sheet_records:
                    identifier = str(r.get("Identifier", "") or r.get("identifier", ""))

                    if identifier:
                        sheet_sample_names.add(identifier)

                
                if sample_descriptor in sheet_sample_names:
                    if sheet_name not in sheet_stats:
                        sheet_stats[sheet_name] = {
                            'total_records': len(sheet_records),
                            'valid_records': 0,
                            'error_records': 0,
                            'warning_records': 0,
                            'sample_status': {}
                        }

                    if sample_descriptor not in sheet_stats[sheet_name]['sample_status']:
                        if errors:
                            sheet_stats[sheet_name]['error_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_descriptor] = 'error'
                        elif warnings:
                            sheet_stats[sheet_name]['warning_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_descriptor] = 'warning'
                        else:
                            sheet_stats[sheet_name]['valid_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_descriptor] = 'valid'
                    break

    # Correct valid counts
    for sheet_name in sheet_stats:
        stats = sheet_stats[sheet_name]
        stats['valid_records'] = stats['total_records'] - stats['error_records']

    return sheet_stats

