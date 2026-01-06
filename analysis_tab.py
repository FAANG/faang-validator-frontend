"""
Analysis tab callbacks and UI components for FAANG Validator.
This module contains all analysis-specific functionality.
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


def create_biosamples_form_analysis():
    """
    Create BioSamples submission form specifically for analysis tab.
    
    Returns:
        HTML Div containing BioSamples form for analysis
    """
    return html.Div(
        [
            html.H2("Prepare data for submission", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id="biosamples-username-ena-analysis",
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
                id="biosamples-password-ena-analysis",
                type="password",
                placeholder="Password",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            # dcc.RadioItems(
            #     id="biosamples-env-ena-analysis",
            #     options=[{"label": " Test server", "value": "test"},
            #              {"label": " Production server", "value": "prod"}],
            #     value="test",
            #     labelStyle={"marginRight": "18px"},
            #     style={"marginBottom": "16px"}
            # ),

            html.Div(id="biosamples-status-banner-ena-analysis",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id="biosamples-submit-btn-ena-analysis", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id="biosamples-submit-msg-ena-analysis", style={"marginTop": "10px"}),
        ],
        id="biosamples-form-ena-analysis",
        style={"display": "none", "marginTop": "16px"},
    )

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'https://faang-validator-backend-service-964531885708.europe-west2.run.app')


def get_all_errors_and_warnings(record):
    """Extract all errors and warnings from a validation record."""
    import re
    errors = {}
    warnings = {}

    if 'errors' in record and record['errors']:
        if 'field_errors' in record['errors']:
            for field, messages in record['errors']['field_errors'].items():
                errors[field] = messages
        if 'relationship_errors' in record['errors']:
            for message in record['errors']['relationship_errors']:
                field_to_blame = 'general'
                data = record.get('data', {})
                if 'Child Of' in data and data.get('Child Of'):
                    field_to_blame = 'Child Of'
                elif 'Derived From' in data and data.get('Derived From'):
                    field_to_blame = 'Derived From'
                if field_to_blame not in errors:
                    errors[field_to_blame] = []
                errors[field_to_blame].append(message)

    if 'field_warnings' in record and record['field_warnings']:
        for field, messages in record['field_warnings'].items():
            warnings[field] = messages

    if 'ontology_warnings' in record and record['ontology_warnings']:
        for message in record['ontology_warnings']:
            match = re.search(r"in field '([^']*)'", message)
            if match:
                field = match.group(1)
                if field not in warnings:
                    warnings[field] = []
                warnings[field].append(message)
            else:
                if 'general' not in warnings:
                    warnings['general'] = []
                warnings['general'].append(message)

    if 'relationship_errors' in record and record['relationship_errors']:
        field_to_blame = 'general'
        data = record.get('data', {})
        if 'Child Of' in data and data.get('Child Of'):
            field_to_blame = 'Child Of'
        elif 'Derived From' in data and data.get('Derived From'):
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


def _valid_invalid_analysis_counts(v):
    """Get valid/invalid counts for analysis using analysis_summary"""
    try:
        s = v.get("results", {}).get("analysis_summary", {}) or {}
        return int(s.get("valid_analyses", 0)), int(s.get("invalid_analyses", 0))
    except Exception:
        return 0, 0


def register_analysis_callbacks(app):
    """
    Register all analysis tab callbacks with the Dash app.
    This function should be called from dash_app.py after the app is created.
    """
    
    # File upload callback for Analysis tab
    @app.callback(
        [Output('stored-file-data-analysis', 'data'),
         Output('stored-filename-analysis', 'data'),
         Output('file-chosen-text-analysis', 'children'),
         Output('selected-file-display-analysis', 'children'),
         Output('selected-file-display-analysis', 'style'),
         Output('output-data-upload-analysis', 'children'),
         Output('stored-all-sheets-data-analysis', 'data'),
         Output('stored-sheet-names-analysis', 'data'),
         Output('stored-parsed-json-analysis', 'data'),
         Output('active-sheet-analysis', 'data')],
        [Input('upload-data-analysis', 'contents')],
        [State('upload-data-analysis', 'filename')]
    )
    def store_file_data_analysis(contents, filename):
        """Store uploaded file data for Analysis tab"""
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
                df_sheet = excel_file.parse(sheet, dtype=str)
                df_sheet = df_sheet.fillna("")

                if df_sheet.empty or len(df_sheet) == 0:
                    continue

                sheet_records = df_sheet.to_dict("records")
                all_sheets_data[sheet] = sheet_records

                original_headers = [str(col) for col in df_sheet.columns]
                processed_headers = process_headers(original_headers)

                rows = []
                for _, row in df_sheet.iterrows():
                    row_list = [row[col] for col in df_sheet.columns]
                    rows.append(row_list)

                parsed_json_records = build_json_data(processed_headers, rows, sheet_name=sheet)
                parsed_json_data[sheet] = parsed_json_records
                sheets_with_data.append(sheet)

            active_sheet = sheets_with_data[0] if sheets_with_data else None
            sheet_names = sheets_with_data

            file_selected_display = html.Div([
                html.H3("File Selected", id='original-file-heading-analysis'),
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
                        id='uploaded-sheets-tabs-analysis',
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

    # Enable/disable validate button for Analysis tab
    @app.callback(
        [Output('validate-button-analysis', 'disabled'),
         Output('validate-button-container-analysis', 'style'),
         Output('reset-button-container-analysis', 'style')],
        [Input('stored-file-data-analysis', 'data')]
    )
    def show_and_enable_buttons_analysis(file_data):
        """Show and enable buttons for Analysis tab"""
        if file_data is None:
            return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
        else:
            return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}

    # Validation callback for Analysis tab
    @app.callback(
        [Output('output-data-upload-analysis', 'children', allow_duplicate=True),
         Output('stored-json-validation-results-analysis', 'data')],
        [Input('validate-button-analysis', 'n_clicks')],
        [State('stored-file-data-analysis', 'data'),
         State('stored-filename-analysis', 'data'),
         State('output-data-upload-analysis', 'children'),
         State('stored-all-sheets-data-analysis', 'data'),
         State('stored-sheet-names-analysis', 'data'),
         State('stored-parsed-json-analysis', 'data')],
        prevent_initial_call=True
    )
    def validate_data_analysis(n_clicks, contents, filename, current_children, all_sheets_data, sheet_names, parsed_json):
        """Validate data for Analysis tab with robust response handling"""
        if n_clicks is None or parsed_json is None:
            return current_children if current_children else html.Div([]), None

        error_data = []
        all_sheets_validation_data = {}
        valid_count = 0
        invalid_count = 0
        json_validation_results = None
        print(json.dumps(parsed_json))

        try:
            try:
                response = requests.post(
                    f'{BACKEND_API_URL}/validate-data',
                    json={"data": parsed_json},
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
                analysis_summary = validation_results.get('analysis_summary', {})
                total_summary = validation_results.get('total_summary', {})
                # Try analysis_summary first, fallback to total_summary
                if analysis_summary:
                    valid_count = analysis_summary.get('valid_analyses', 0)
                    invalid_count = analysis_summary.get('invalid_analyses', 0)
                else:
                    valid_count = total_summary.get('valid_samples', 0)
                    invalid_count = total_summary.get('invalid_samples', 0)
            elif isinstance(response_json, dict):
                # Backend might return data directly without 'results' wrapper
                # Try to extract data from various possible structures
                if 'total_summary' in response_json or 'analysis_summary' in response_json or 'analysis_types_processed' in response_json:
                    # Data is at top level, wrap it in 'results'
                    json_validation_results = {"results": response_json}
                    validation_results = response_json
                    analysis_summary = validation_results.get('analysis_summary', {})
                    total_summary = validation_results.get('total_summary', {})
                    if analysis_summary:
                        valid_count = analysis_summary.get('valid_analyses', 0)
                        invalid_count = analysis_summary.get('invalid_analyses', 0)
                    else:
                        valid_count = total_summary.get('valid_samples', 0)
                        invalid_count = total_summary.get('invalid_samples', 0)
                else:
                    # Unknown format, try to extract what we can
                    validation_data = response_json
                    valid_count = validation_data.get('valid_samples', 0) or validation_data.get('valid_analyses', 0) or validation_data.get('valid_submissions', 0)
                    invalid_count = validation_data.get('invalid_samples', 0) or validation_data.get('invalid_analyses', 0) or validation_data.get('invalid_submissions', 0)
                    error_data = validation_data.get('warnings', [])

                    # Try to build a minimal structure
                    json_validation_results = {
                        "results": {
                            "total_summary": {
                                "valid_samples": valid_count,
                                "invalid_samples": invalid_count
                            },
                            "results_by_type": {},
                            "analysis_types_processed": []
                        }
                    }
            else:
                # Old format (list) - convert to new format for consistency
                validation_data = response_json[0] if isinstance(response_json, list) else response_json
                records = validation_data.get('validation_result', [])
                valid_count = validation_data.get('valid_samples', 0) or validation_data.get('valid_analyses', 0)
                invalid_count = validation_data.get('invalid_samples', 0) or validation_data.get('invalid_analyses', 0)
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
                        "results_by_type": all_sheets_validation_data,
                        "analysis_types_processed": list(all_sheets_validation_data.keys()) if all_sheets_validation_data else []
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
            dcc.Store(id='stored-error-data-analysis', data=error_data),
            dcc.Store(id='stored-validation-results-analysis', data={'valid_count': valid_count, 'invalid_count': invalid_count,
                                                            'all_sheets_data': all_sheets_validation_data}),
            html.H3("2. Conversion and Validation results"),

            html.Div([
                html.P("Conversion Status", style={'fontWeight': 'bold'}),
                html.P("Success", style={'color': 'green', 'fontWeight': 'bold'}),
                html.P("Validation Status", style={'fontWeight': 'bold'}),
                html.P("Finished", style={'color': 'green', 'fontWeight': 'bold'}),
            ], style={'margin': '10px 0'}),

            html.Div(id='error-table-container-analysis', style={'display': 'none'}),
            html.Div(id='validation-results-container-analysis', style={'margin': '20px 0'})
        ]

        if current_children is None:
            return html.Div(validation_components), json_validation_results if json_validation_results else {'results': {}}
        elif isinstance(current_children, list):
            return html.Div(current_children + validation_components), json_validation_results if json_validation_results else {'results': {}}
        else:
            return html.Div(validation_components + [current_children]), json_validation_results if json_validation_results else {'results': {}}

    # Reset callback for Analysis tab
    app.clientside_callback(
        """
        function(n_clicks) {
            if (n_clicks > 0) {
                window.location.reload();
            }
            return '';
        }
        """,
        Output("dummy-output-for-reset-analysis", "children"),
        [Input("reset-button-analysis", "n_clicks")],
        prevent_initial_call=True,
    )

    # BioSamples form toggle for Analysis tab
    @app.callback(
        [
            Output("biosamples-form-ena-analysis", "style"),
            Output("biosamples-status-banner-ena-analysis", "children"),
            Output("biosamples-status-banner-ena-analysis", "style"),
        ],
        Input("stored-json-validation-results-analysis", "data"),
    )
    def _toggle_biosamples_form_analysis(v):
        """Toggle BioSamples form visibility for Analysis tab"""
        base_style = {"display": "block", "marginTop": "16px"}

        if not v or "results" not in v:
            return ({"display": "none"}, "", {"display": "none"})

        valid_cnt, invalid_cnt = _valid_invalid_analysis_counts(v)
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
                    f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid analysis/analyses."
                ),
                html.Br(),
            ]
            return base_style, msg_children, style_ok
        else:
            return (
                base_style,
                f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid analysis/analyses. No valid analyses to submit.",
                style_warn,
            )

    # BioSamples submit button enable/disable for Analysis tab
    @app.callback(
        Output("biosamples-submit-btn-ena-analysis", "disabled"),
        [
            Input("biosamples-username-ena-analysis", "value"),
            Input("biosamples-password-ena-analysis", "value"),
            Input("stored-json-validation-results-analysis", "data"),
        ],
    )
    def _disable_submit_analysis(u, p, v):
        """Enable/disable submit button for Analysis tab"""
        if not v or "results" not in v:
            return True
        valid_cnt, _ = _valid_invalid_analysis_counts(v)
        if valid_cnt == 0:
            return True
        return not (u and p)

    # BioSamples submission for Analysis tab
    @app.callback(
        [
            Output("biosamples-submit-msg-ena-analysis", "children"),
            Output("biosamples-results-table-analysis", "children"),
        ],
        Input("biosamples-submit-btn-ena-analysis", "n_clicks"),
        State("biosamples-username-ena-analysis", "value"),
        State("biosamples-password-ena-analysis", "value"),
        State("biosamples-env-ena-analysis", "value"),
        State("biosamples-action-analysis", "value"),
        State("stored-json-validation-results-analysis", "data"),
        prevent_initial_call=True,
    )
    def _submit_to_biosamples_analysis(n, username, password, env, action, v):
        """Submit to BioSamples for Analysis tab"""
        if not n:
            raise PreventUpdate

        if not v or "results" not in v:
            msg = html.Span(
                "No validation results available. Please validate your file first.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update

        valid_cnt, invalid_cnt = _valid_invalid_analysis_counts(v)
        if valid_cnt == 0:
            msg = html.Span(
                "No valid analyses to submit. Please fix errors and re-validate.",
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
            "mode": env or "test",  # Default to test if env is None
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
                    html.Span(f"Submitted analyses: {submitted_count}"),
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
                    {"Analysis Name": name, "BioSample ID": acc}
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
                        {"name": "Analysis Name", "id": "Analysis Name"},
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
        Output('validation-results-container-analysis', 'children'),
        [Input('stored-json-validation-results-analysis', 'data')],
        [State('stored-sheet-names-analysis', 'data'),
         State('stored-all-sheets-data-analysis', 'data')]
    )
    def populate_validation_results_tabs_analysis(validation_results, sheet_names, all_sheets_data):
        """Populate validation results tabs for analysis tab"""
        if not validation_results or 'results' not in validation_results:
            return []

        if not sheet_names:
            return []

        validation_data = validation_results['results']
        
        # Filter sheet_names to only include those in analysis_types_processed
        analysis_types_processed = validation_data.get('analysis_types_processed', []) or []
        if analysis_types_processed:
            # Only show sheets that are in analysis_types_processed
            sheet_names = [sheet for sheet in sheet_names if sheet in analysis_types_processed]

        # Calculate sheet statistics for analysis
        sheet_stats = _calculate_sheet_statistics_analysis(validation_results, all_sheets_data or {})

        sheet_tabs = []
        sheets_with_data = []

        for sheet_name in sheet_names:
            # Get statistics for this sheet
            stats = sheet_stats.get(sheet_name, {})
            errors = stats.get('error_records', 0)
            warnings = stats.get('warning_records', 0)
            valid = stats.get('valid_records', 0)
            
            # Show all sheets in analysis_types_processed, regardless of errors/warnings
            sheets_with_data.append(sheet_name)
            # Create label showing counts for THIS sheet
            label = f"{sheet_name} ({valid} valid / {errors} invalid)"

            sheet_tabs.append(
                dcc.Tab(
                    label=label,
                    value=sheet_name,
                    id={'type': 'sheet-validation-tab-analysis', 'sheet_name': sheet_name},
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
                    children=[html.Div(id={'type': 'sheet-validation-content-analysis', 'index': sheet_name})]
                )
            )

        if not sheet_tabs:
            return html.Div([
                html.P(
                    "The provided data has been validated successfully with no errors or warnings. You may proceed with submission.",
                    style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
            ])

        tabs = dcc.Tabs(
            id='sheet-validation-tabs-analysis',
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
                    id="download-errors-btn-analysis",
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

    # Callback to populate sheet content when tab is selected for analysis
    @app.callback(
        Output({'type': 'sheet-validation-content-analysis', 'index': MATCH}, 'children'),
        [Input('sheet-validation-tabs-analysis', 'value')],
        [State('stored-json-validation-results-analysis', 'data'),
         State('stored-all-sheets-data-analysis', 'data')]
    )
    def populate_sheet_validation_content_analysis(selected_sheet_name, validation_results, all_sheets_data):
        """Populate sheet validation content for analysis tab"""
        if validation_results is None or selected_sheet_name is None:
            return []

        if not all_sheets_data or selected_sheet_name not in all_sheets_data:
            return html.Div("No data available for this sheet.")

        return make_sheet_validation_panel_analysis(selected_sheet_name, validation_results, all_sheets_data)


def make_sheet_validation_panel_analysis(sheet_name: str, validation_results: dict, all_sheets_data: dict):
    """Create a panel showing validation results for analysis sheet"""
    import uuid
    panel_id = str(uuid.uuid4())

    # Get sheet data
    sheet_records = all_sheets_data.get(sheet_name, [])
    if not sheet_records:
        return html.Div([html.H4("No data available", style={'textAlign': 'center', 'margin': '10px 0'})])

    # Get validation data - use analysis_types_processed for analysis
    validation_data = validation_results.get('results', {})
    results_by_type = validation_data.get('results_by_type', {}) or {}
    analysis_summary = validation_data.get('analysis_summary', {})
    analysis_types = validation_data.get('analysis_types_processed', []) or []

    # Get all validation rows for this sheet
    # For analysis, try "Analysis Alias" first, then "Sample Name"
    sheet_sample_names = set()
    for record in sheet_records:
        analysis_alias = str(record.get("Analysis Alias", ""))
        sample_name = str(record.get("Sample Name", ""))
        if analysis_alias:
            sheet_sample_names.add(analysis_alias)
        if sample_name:
            sheet_sample_names.add(sample_name)

    # Collect all rows that belong to this sheet
    error_map = {}
    warning_map = {}

    for analysis_type in analysis_types:
        at_data = results_by_type.get(analysis_type, {}) or {}
        at_key = analysis_type.replace(' ', '_')
        invalid_key = f"invalid_{at_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
        valid_key = f"valid_{at_key}s"

        invalid_records = at_data.get(invalid_key, [])
        valid_records = at_data.get(valid_key, [])

        for record in invalid_records + valid_records:
            sample_name = record.get("sample_name", "") or record.get("analysis_alias", "")
            if sample_name in sheet_sample_names:
                errors, warnings = get_all_errors_and_warnings(record)
                if errors:
                    error_map[sample_name] = errors
                if warnings:
                    warning_map[sample_name] = warnings

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
        direct = _resolve_col(field_name, columns)
        if direct:
            return direct
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
        if "." in field_name:
            base = field_name.split(".", 1)[0]
            base_match = _resolve_col(base, columns)
            if base_match:
                return base_match
        return None

    # Build cell styles and tooltips
    cell_styles = []
    tooltip_data = []

    for i, row in df_all.iterrows():
        # Try to match by Analysis Alias or Sample Name
        analysis_alias = str(row.get("Analysis Alias", ""))
        sample_name = str(row.get("Sample Name", ""))
        match_key = analysis_alias if analysis_alias else sample_name
        
        tips = {}
        row_styles = []

        if match_key in error_map:
            field_errors = error_map[match_key] or {}
            for field, msgs in field_errors.items():
                col = _map_field_to_column(field, df_all.columns)
                if not col:
                    col = field

                col_id = None
                if col in df_all.columns:
                    col_id = col
                else:
                    col_str = str(col)
                    for df_col in df_all.columns:
                        if str(df_col) == col_str:
                            col_id = df_col
                            break
                if not col_id:
                    continue

                msgs_list = _as_list(msgs)
                lower_msgs = [m.lower() for m in msgs_list]
                is_extra = any("extra inputs are not permitted" in lm for lm in lower_msgs)
                is_warning_like = any("warning" in lm for lm in lower_msgs)
                prefix = "**Warning**: " if (is_extra or is_warning_like) else "**Error**: "
                msg_text = prefix + field + " — " + " | ".join(msgs_list)
                if col_id in tips:
                    existing = tips[col_id].get("value", "")
                    combined = f"{existing} | {msg_text}" if existing else msg_text
                else:
                    combined = msg_text
                if is_extra or is_warning_like:
                    row_styles.append({'if': {'row_index': i, 'column_id': col_id}, 'backgroundColor': '#fff4cc'})
                    tips[col_id] = {'value': combined, 'type': 'markdown'}
                else:
                    row_styles.append({'if': {'row_index': i, 'column_id': col_id}, 'backgroundColor': '#ffcccc'})
                    tips[col_id] = {'value': combined, 'type': 'markdown'}

        if match_key in warning_map:
            field_warnings = warning_map[match_key] or {}
            for field, msgs in field_warnings.items():
                col = _map_field_to_column(field, df_all.columns)
                if not col:
                    col = field

                col_id = None
                if col in df_all.columns:
                    col_id = col
                else:
                    col_str = str(col)
                    for df_col in df_all.columns:
                        if str(df_col) == col_str:
                            col_id = df_col
                            break
                if not col_id:
                    continue
                msgs_list = _as_list(msgs)
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
    for sample_name, field_errors in error_map.items():
        for field in field_errors.keys():
            error_fields_count[field] = error_fields_count.get(field, 0) + 1

    warning_fields_count = {}
    for sample_name, field_warnings in warning_map.items():
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
    if analysis_summary:
        summary_items = []
        for key, value in analysis_summary.items():
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
        html.H4(f"Validation Results - {sheet_name}", style={'textAlign': 'center', 'margin': '10px 0'}),
        html.Div([
            DataTable(
                id={"type": "sheet-result-table-analysis", "sheet_name": sheet_name, "panel_id": panel_id},
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
        ], id={"type": "sheet-table-container-analysis", "sheet_name": sheet_name, "panel_id": panel_id},
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


def _calculate_sheet_statistics_analysis(validation_results, all_sheets_data):
    """Calculate errors and warnings count for each Excel sheet for analysis"""
    sheet_stats = {}

    if not validation_results or 'results' not in validation_results:
        return sheet_stats

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {}) or {}
    analysis_types = validation_data.get('analysis_types_processed', []) or []

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

    # Process each analysis type and map to sheets
    for analysis_type in analysis_types:
        at_data = results_by_type.get(analysis_type, {}) or {}
        at_key = analysis_type.replace(' ', '_')

        invalid_key = f"invalid_{at_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
        valid_key = f"valid_{at_key}s"

        invalid_records = at_data.get(invalid_key, [])
        valid_records = at_data.get(valid_key, [])

        all_records = invalid_records + valid_records

        for record in all_records:
            sample_name = record.get("sample_name", "") or record.get("analysis_alias", "")
            if not sample_name:
                continue

            errors, warnings = get_all_errors_and_warnings(record)

            # Find which sheet contains this sample
            # For analysis, try "Analysis Alias" first, then "Sample Name"
            for sheet_name, sheet_records in (all_sheets_data or {}).items():
                if not sheet_records:
                    continue
                sheet_sample_names = set()
                for r in sheet_records:
                    analysis_alias = str(r.get("Analysis Alias", ""))
                    samp_name = str(r.get("Sample Name", ""))
                    if analysis_alias:
                        sheet_sample_names.add(analysis_alias)
                    if samp_name:
                        sheet_sample_names.add(samp_name)
                
                if sample_name in sheet_sample_names:
                    if sheet_name not in sheet_stats:
                        sheet_stats[sheet_name] = {
                            'total_records': len(sheet_records),
                            'valid_records': 0,
                            'error_records': 0,
                            'warning_records': 0,
                            'sample_status': {}
                        }

                    if sample_name not in sheet_stats[sheet_name]['sample_status']:
                        if errors:
                            sheet_stats[sheet_name]['error_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_name] = 'error'
                        elif warnings:
                            sheet_stats[sheet_name]['warning_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_name] = 'warning'
                        else:
                            sheet_stats[sheet_name]['valid_records'] += 1
                            sheet_stats[sheet_name]['sample_status'][sample_name] = 'valid'
                    break

    # Correct valid counts
    for sheet_name in sheet_stats:
        stats = sheet_stats[sheet_name]
        stats['valid_records'] = stats['total_records'] - stats['error_records']

    return sheet_stats

