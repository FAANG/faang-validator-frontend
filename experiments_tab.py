"""
Experiments tab callbacks and UI components for FAANG Validator.
This module contains all experiments-specific functionality.
"""
import json
import base64
import io
import re

import pandas as pd
import requests
from dash import dcc, html, dash_table
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH
from dash.exceptions import PreventUpdate
import dash
import os

from file_processor import process_headers, build_json_data


def create_experiments():
    """
    Create experiments submission form specifically for experiments tab.
    
    Returns:
        HTML Div containing experiments form for experiments
    """
    return html.Div(
        [
            html.H2("Submit data to ENA", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id="experiments-username-ena",
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
                id="experiments-password-ena",
                type="password",
                placeholder="Password",
                value="",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            html.Div(id="experiments-status-banner-ena",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            dcc.Loading(
                id="loading-submit-ena",
                type="circle",
                children=html.Div([
                    html.Button(
                        "Submit", id="experiments-submit-btn-ena", n_clicks=0,
                        style={
                            "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                            "border": "none", "borderRadius": "8px", "cursor": "pointer",
                            "fontSize": "16px", "width": "140px"
                        }
                    ),
                    html.Div(id="experiments-submit-msg-ena", style={"marginTop": "10px"}),
                ])
            ),
            # Submission results panel (shown after successful ENA submission)
            html.Div(
                id="experiments-submission-results-panel",
                style={
                    "display": "none",
                    "marginTop": "20px",
                    "padding": "16px",
                    "borderRadius": "8px",
                    "border": "1px solid #cbd5e1",
                    "backgroundColor": "#f8fafc",
                },
            ),
            # Store for raw submission_results XML (for download)
            dcc.Store(id="experiments-submission-results-store"),
            # Download component for XML receipt
            dcc.Download(id="experiments-submission-results-xml-download"),
        ],
        id="experiments-form-ena",
        style={"display": "none", "marginTop": "16px"},
    )

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'https://faang-validator-backend-service-964531885708.europe-west2.run.app')


def get_all_errors_and_warnings(record):
    errors = {}
    warnings = {}

    # From 'errors' object
    if 'errors' in record and record['errors']:
        # Handle errors.errors array (e.g., "Geographic Location: Field required")
        if 'errors' in record['errors'] and isinstance(record['errors']['errors'], list):
            for error_msg in record['errors']['errors']:
                # Parse messages like "Geographic Location: Field required"
                if ':' in error_msg:
                    parts = error_msg.split(':', 1)
                    field = parts[0].strip()
                    message = parts[1].strip() if len(parts) > 1 else error_msg
                    if field not in errors:
                        errors[field] = []
                    errors[field].append(message)
                else:
                    # If no field name found, add to 'general'
                    if 'general' not in errors:
                        errors['general'] = []
                    errors['general'].append(error_msg)

        if 'field_errors' in record['errors']:
            for field, messages in record['errors']['field_errors'].items():
                if field not in errors:
                    errors[field] = []
                # Ensure messages is a list
                if isinstance(messages, list):
                    errors[field].extend(messages)
                else:
                    errors[field].append(messages)
        if 'relationship_errors' in record['errors']:
            for message in record['errors']['relationship_errors']:
                field_to_blame = 'general'
                data = record.get('data', {})
                if 'Child Of' in data and data.get('Child Of'):
                    field_to_blame = 'Child Of'
                elif 'Derived From' in data and data.get('Derived From'):
                    field_to_blame = 'Derived From'

                if field_to_blame not in warnings:
                    warnings[field_to_blame] = []
                warnings[field_to_blame].append(message)

    # From 'field_warnings'
    if 'field_warnings' in record and record['field_warnings']:
        for field, messages in record['field_warnings'].items():
            warnings[field] = messages

    # From 'ontology_warnings'
    if 'ontology_warnings' in record and record['ontology_warnings']:
        for message in record['ontology_warnings']:
            match = re.search(r"in field '([^']*)'", message)
            if match:
                field = match.group(1)
                if field not in warnings:
                    warnings[field] = []
                warnings[field].append(message)
            else:
                # Generic warning if field not found
                if 'general' not in warnings:
                    warnings['general'] = []
                warnings['general'].append(message)

    # From 'relationship_errors' (top level - treat as warnings, yellow highlighting)
    if 'relationship_errors' in record and record['relationship_errors']:
        # Try to associate with 'Child Of' or 'Derived From'
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


def _warnings_by_field(warnings_list):
    by_field = {}
    for w in warnings_list or []:
        m = re.search(r"Field '([^']*)'", str(w))
        field = m.group(1) if m else None
        by_field.setdefault(field, []).append(str(w))
    return by_field


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
        return int(s.get("valid_experiments", 0)), int(s.get("invalid_experiments", 0))
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
        """
        Store uploaded file data for Experiments tab.
        This callback is optimized for performance by only storing the file
        content and name, deferring all parsing to the validation step.
        """
        if contents is None:
            return None, None, "No file chosen", [], {'display': 'none'}, html.Div(), None, None, None, None

        try:
            # Handle case where filename might be None
            filename_display = filename if filename else "Unknown file"
            
            # Display a simple message that the file has been selected
            file_selected_display = html.Div([
                html.H3("File Selected", id='original-file-heading-experiments'),
                html.P(f"File: {filename_display}", style={'fontWeight': 'bold'})
            ])

            # Return file content and name to be stored, and update UI
            # Set sheet data and parsed JSON to None since parsing is deferred
            # Return empty div for output-data-upload-experiments (will be populated on validation)
            return (contents, filename, filename_display, file_selected_display,
                    {'display': 'block', 'margin': '20px 0'},
                    html.Div(), None, None, None, None)
        except Exception as e:
            # If there's an error, display it and still return the file data
            error_display = html.Div([
                html.H3("File Selected", id='original-file-heading-experiments'),
                html.P(f"File: {filename if filename else 'Unknown'}", style={'fontWeight': 'bold'}),
                html.P(f"Note: {str(e)}", style={'color': 'orange', 'fontSize': '12px'})
            ])
            return (contents, filename, filename if filename else "Unknown file", error_display,
                    {'display': 'block', 'margin': '20px 0'},
                    html.Div(),
                    None, None, None, None)

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
         Output('stored-json-validation-results-experiments', 'data'),
         Output('stored-all-sheets-data-experiments', 'data', allow_duplicate=True),
         Output('stored-sheet-names-experiments', 'data', allow_duplicate=True),
         Output('stored-parsed-json-experiments', 'data', allow_duplicate=True)],
        [Input('validate-button-experiments', 'n_clicks')],
        [State('stored-file-data-experiments', 'data'),
         State('stored-filename-experiments', 'data'),
         State('output-data-upload-experiments', 'children')],
        prevent_initial_call=True
    )
    def validate_data_experiments(n_clicks, contents, filename, current_children):
        """
        Validate data for Experiments tab.
        This callback now handles file parsing and JSON conversion, which was
        moved from the file upload callback for performance reasons.
        """
        global json_validation_results
        if n_clicks is None or contents is None:
            return current_children or html.Div([]), None, None, None, None

        def create_output(components):
            if current_children is None:
                return html.Div(components)
            if isinstance(current_children, list):
                return html.Div(current_children + components)
            return html.Div([current_children] + components)

        try:
            # Decode file content and parse Excel file
            content_type, content_string = contents.split(',', 1)
            decoded = base64.b64decode(content_string)
            excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")

            sheet_names = excel_file.sheet_names
            all_sheets_data = {}
            parsed_json = {}
            sheets_with_data = []

            # Process each sheet in the Excel file
            for sheet in sheet_names:
                if sheet.lower() == "faang_field_values":
                    continue

                df_sheet = excel_file.parse(sheet, dtype=str).fillna("")

                # Store ALL sheets in all_sheets_data (including empty ones) for download functionality
                # This ensures all sheets are available for download, even if they're empty
                # For empty sheets, preserve column structure by storing columns info
                if df_sheet.empty:
                    # Store empty records but preserve column structure
                    # Store as dict with columns key for empty sheets
                    all_sheets_data[sheet] = {
                        "_empty": True,
                        "_columns": list(df_sheet.columns),
                        "records": []
                    }
                else:
                    # Store normal records for non-empty sheets
                    all_sheets_data[sheet] = df_sheet.to_dict("records")

                # Only process non-empty sheets for validation
                if df_sheet.empty:
                    continue

                sheets_with_data.append(sheet)

                # Convert sheet to JSON for validation
                original_headers = [str(col) for col in df_sheet.columns]
                processed_headers = process_headers(original_headers)
                rows = df_sheet.values.tolist()
                parsed_json[sheet] = build_json_data(processed_headers, rows, sheet)

            if not parsed_json:
                # Handle case where no data was found in any sheet
                no_data_msg = html.P("No data found in the uploaded file.", style={'color': 'orange'})
                return create_output([no_data_msg]), None, None, None, None

        except Exception as e:
            # Handle file parsing errors
            error_div = html.Div([
                html.H5(filename),
                html.P(f"Error processing file: {str(e)}", style={'color': 'red'})
            ])
            return create_output([error_div]), None, None, None, None

        # Send data to backend for validation
        print("file uploaded!!")
        print(json.dumps(parsed_json))
        try:
            response = requests.post(
                f'{BACKEND_API_URL}/validate-data',
                json={"data": parsed_json, "data_type": "experiment"},
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

        # Process validation response
        # (The logic for handling different response formats remains the same)
        if isinstance(response_json, dict) and 'results' in response_json:
            json_validation_results = response_json
        elif isinstance(response_json, dict):
            json_validation_results = {"results": response_json}
        else:
            # Handle legacy list format if necessary
            validation_data = response_json[0] if isinstance(response_json, list) else response_json
            valid_count = validation_data.get('valid_experiments', 0)
            invalid_count = validation_data.get('invalid_experiments', 0)
            all_sheets_validation_data = validation_data.get('all_sheets_data', {})

        # Decide what to render based on whether any experiment types were processed.
        experiment_types_processed = json_validation_results.get("results", {}).get("experiment_types_processed", []) or []

        if not experiment_types_processed:
            # Wrong template for Experiments: keep only the container div so that
            # the populate_validation_results_tabs_experiments callback can show
            # the "provided file is not Experiments" message. Do NOT show the
            # "2. Conversion and Validation results" header or status panel.
            validation_components = [
                html.Div(id='validation-results-container-experiments', style={'margin': '20px 0'})
            ]
        else:
            # Normal case: show full Conversion and Validation panel + container.
            validation_components = [
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

        # Update the UI with validation results
        output_children = create_output(validation_components)
        return output_children, json_validation_results, all_sheets_data, sheets_with_data, parsed_json

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

    # experiments form toggle for Experiments tab
    @app.callback(
        [
            Output("experiments-form-ena", "style"),
            Output("experiments-status-banner-ena", "children"),
            Output("experiments-status-banner-ena", "style"),
        ],
        Input("stored-json-validation-results-experiments", "data"),
    )
    def _toggle_experiments_form_experiments(v):
        """Toggle experiments form visibility for Experiments tab"""
        base_style = {"display": "block", "marginTop": "16px"}

        if not v or "results" not in v:
            return ({"display": "none"}, "", {"display": "none"})

        validation_data = v.get("results", {})
        experiment_types_processed = validation_data.get("experiment_types_processed", []) or []

        # If the uploaded file did not produce any experiment sheets,
        # hide the ENA submit panel entirely (wrong template for this tab)
        if not experiment_types_processed:
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

    # experiments submit button enable/disable for Experiments tab
    @app.callback(
        [
            Output("experiments-submit-btn-ena", "disabled"),
            Output("experiments-submit-btn-ena", "style"),
        ],
        [
            Input("experiments-username-ena", "value"),
            Input("experiments-password-ena", "value"),
            Input("stored-json-validation-results-experiments", "data"),
        ],
    )
    def _disable_submit_experiments(u, p, v):
        """Enable/disable submit button for Experiments tab"""
        # Default enabled style
        enabled_style = {
            "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
            "border": "none", "borderRadius": "8px", "cursor": "pointer",
            "fontSize": "16px", "width": "140px"
        }
        # Disabled style (grayed out)
        disabled_style = {
            "backgroundColor": "#cccccc", "color": "#666666", "padding": "10px 18px",
            "border": "none", "borderRadius": "8px", "cursor": "not-allowed",
            "fontSize": "16px", "width": "140px", "opacity": "0.6"
        }

        if not v or "results" not in v:
            return True, disabled_style

        validation_data = v.get("results", {})
        experiment_types_processed = validation_data.get("experiment_types_processed", []) or []

        # If the uploaded file is not an Experiments template, keep the button disabled
        if not experiment_types_processed:
            return True, disabled_style

        valid_cnt, invalid_cnt = _valid_invalid_experiments_counts(v)
        # Disable if there are any invalid experiments
        if invalid_cnt > 0:
            return True, disabled_style  # Disable if there are invalid experiments
        # All experiments are valid, enable if username and password are provided
        if valid_cnt == 0:
            return True, disabled_style
        is_enabled = u and p
        return not is_enabled, enabled_style if is_enabled else disabled_style

    # experiments submission for Experiments tab
    @app.callback(
        [
            Output("experiments-submit-msg-ena", "children"),
            Output("biosamples-results-table-experiments", "children"),
            Output("experiments-submission-results-panel", "children"),
            Output("experiments-submission-results-panel", "style"),
            Output("experiments-submission-results-store", "data"),
        ],
        Input("experiments-submit-btn-ena", "n_clicks"),
        State("experiments-username-ena", "value"),
        State("experiments-password-ena", "value"),
        State("biosamples-action-experiments", "value"),
        State("stored-json-validation-results-experiments", "data"),
        State("stored-parsed-json-experiments", "data"),
        prevent_initial_call=True,
    )
    def _submit_experiments(n, username, password, action, v, original_data):
        """Submit to experiments for Experiments tab"""
        if not n:
            raise PreventUpdate

        hidden_style = {"display": "none"}

        if not v or "results" not in v:
            msg = html.Span(
                "No validation results available. Please validate your file first.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update, dash.no_update, hidden_style, None

        valid_cnt, invalid_cnt = _valid_invalid_experiments_counts(v)
        if valid_cnt == 0:
            msg = html.Span(
                "No valid samples to submit. Please fix errors and re-validate.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update, dash.no_update, hidden_style, None

        if not username or not password:
            msg = html.Span(
                "Please enter Webin username and password.",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update, dash.no_update, hidden_style, None

        validation_results = v["results"]

        body = {
            "validation_results": validation_results,
            "original_data": original_data,
            "webin_username": username,
            "webin_password": password,
            "mode": "test",  # Default to test server
            "action": action,
        }

        try:
            url = f"{BACKEND_API_URL}/submit-experiment"
            r = requests.post(url, json=body, timeout=600)

            data = r.json() if r.content else {}

            success = data.get("success", False)
            message = data.get("message", "No message from server")
            submitted_count = data.get("submitted_count")
            errors = data.get("errors") or []
            info_messages = data.get("info_messages") or []
            submission_results_xml = data.get("submission_results") or ""
            biosamples_ids = data.get("biosamples_ids") or {}

            color = "#388e3c" if success else "#c62828"

            msg_children = [html.Span(message, style={"fontWeight": 500})]
            if submitted_count is not None:
                msg_children += [
                    html.Br(),
                    html.Span(f"Submitted samples: {submitted_count}"),
                ]


            if not biosamples_ids:
                msg_children = [
                    html.Br(),
                ] + msg_children

            msg = html.Div(msg_children, style={"color": color})



            table = html.Div()


            panel_children = []
            if errors or info_messages or submission_results_xml:
                panel_children = [
                    html.Div(
                        [
                            html.H3(
                                "Submission Results",
                                style={"marginBottom": "0"},
                            ),
                            html.Button(
                                "Download submission results",
                                id="experiments-download-submission-xml-btn",
                                n_clicks=0,
                                style={
                                    "backgroundColor": "#ffd740",
                                    "color": "black",
                                    "padding": "8px 16px",
                                    "border": "none",
                                    "borderRadius": "6px",
                                    "cursor": "pointer",
                                    "fontSize": "14px",
                                    "marginLeft": "16px",
                                },
                            ),
                        ],
                        style={
                            "display": "flex",
                            "alignItems": "center",
                            "justifyContent": "space-between",
                            "gap": "16px",
                            "marginBottom": "8px",
                        },
                    ),
                ]

                if info_messages:
                    panel_children.append(
                        html.Div(
                            [
                                html.Ul(
                                    [html.Li(m) for m in info_messages],
                                    style={
                                        "marginLeft": "20px",
                                        "color": "#475569",
                                    },
                                ),
                            ]
                        )
                    )

                # Detailed submission errors moved into the panel, after info (no heading)
                if errors:
                    panel_children.append(
                        html.Div(
                            [
                                html.Ul(
                                    [html.Li(e) for e in errors],
                                    style={
                                        "marginLeft": "20px",
                                        "color": "#b91c1c",
                                    },
                                ),
                            ]
                        )
                    )

            panel_style = {
                "display": "block" if panel_children else "none",
                "marginTop": "20px",
                "padding": "16px",
                "borderRadius": "8px",
                "backgroundColor": "#f8fafc",
            }

            return msg, table, panel_children, panel_style, submission_results_xml

        except Exception as e:
            msg = html.Span(
                f"Submission error: {e}",
                style={"color": "#c62828", "fontWeight": 500},
            )
            return msg, dash.no_update, dash.no_update, hidden_style, None

    # Download callback for ENA submission results XML (experiments tab)
    @app.callback(
        Output("experiments-submission-results-xml-download", "data"),
        Input("experiments-download-submission-xml-btn", "n_clicks"),
        State("experiments-submission-results-store", "data"),
        prevent_initial_call=True,
    )
    def _download_experiments_submission_xml(n_clicks, xml_text):
        """Trigger download of ENA submission_results XML for experiments."""
        if not n_clicks or not xml_text:
            raise PreventUpdate
        # Use send_string so that text encoding is handled automatically
        return dcc.send_string(xml_text, "experiments_submission_results.xml")

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

        # Filter sheet_names to only include those in experiment_types_processed.
        # If nothing was processed, this means the file is not an Experiments template.
        experiment_types_processed = validation_data.get('experiment_types_processed', []) or []
        if not experiment_types_processed:
            return html.Div([
                html.P(
                    "TThe uploaded file could not be processed because it does not contain sample information. Please upload a valid experiments file.",
                    style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
            ])

        # Only show sheets that are in experiment_types_processed
        sheet_names = [sheet for sheet in sheet_names if sheet in experiment_types_processed]
        # Also filter all_sheets_data to only include processed sheets for better performance
        if all_sheets_data:
            all_sheets_data = {k: v for k, v in all_sheets_data.items() if k in experiment_types_processed}

        # Calculate sheet statistics for experiments (now only processes filtered sheets)
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
                        'borderTop': 'none',
                        'borderRight': 'none',
                        'borderBottom': 'none',
                        'borderLeft': 'none',
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
                        'borderTop': 'none',
                        'borderRight': 'none',
                        'borderLeft': 'none',
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
                    "The uploaded file could not be processed because it does not contain experiment information. Please upload a valid experiments file.",
                    style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
            ])

        tabs = dcc.Tabs(
            id='sheet-validation-tabs-experiments',
            value=sheets_with_data[0] if sheets_with_data else None,
            children=sheet_tabs,
            style={
                'borderTop': 'none',
                'borderRight': 'none',
                'borderLeft': 'none',
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
                    disabled=False,
                    style={
                        'backgroundColor': '#ffd740',
                        'color': 'black',
                        'padding': '10px 20px',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px',
                        'pointerEvents': 'auto',
                        'position': 'relative',
                        'zIndex': 1
                    }
                ),
            ],
            style={
                'display': 'flex',
                'justifyContent': 'space-between',
                'alignItems': 'center',
                'marginBottom': '10px',
                'position': 'relative',
                'zIndex': 1
            }
        )

        return html.Div([header_bar, tabs], style={"marginTop": "8px",
                                                   "transition": "opacity 0.3s ease-in-out"})

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

        if not all_sheets_data:
            raise PreventUpdate

        # Use all sheets from all_sheets_data, not just sheet_names
        # This ensures all sheets (including empty ones) are included in the download
        all_sheet_names = list(all_sheets_data.keys())

        # Build a mapping of experiment identifiers to their field-level errors/warnings
        # Structure: {identifier_normalized: {"errors": {field: [msgs]}, "warnings": {field: [msgs]}}}
        identifier_to_field_errors = {}

        # Helper function to map backend field names to Excel column names
        # Use the EXACT same mapping function as validation results panel (_map_field_to_column)
        def _map_field_to_column_excel(field_name, columns):
            if not field_name:
                return None

            # Handle Health Status fields (with or without .term) - same as validation panel
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

            # Handle Cell Type fields (with or without .term) - same as validation panel
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

            # Handle Secondary Project fields - same as validation panel
            if field_name and "secondary project" in field_name.lower():
                # Extract index from field name if present (e.g., "Secondary Project.0" -> 0, "Secondary Project.1" -> 1)
                field_index = None
                field_lower = field_name.lower()
                if "." in field_lower:
                    parts = field_lower.split(".", 1)
                    if len(parts) > 1 and parts[1].isdigit():
                        field_index = int(parts[1])

                # Collect all Secondary Project columns in order
                secondary_project_cols = []
                for col in columns:
                    col_str = str(col)
                    col_lower = col_str.lower()
                    # Match exact "Secondary Project" or numbered versions like "Secondary Project.0", "Secondary Project.1"
                    if col_lower == "secondary project" or col_lower.startswith("secondary project."):
                        secondary_project_cols.append(col)

                if secondary_project_cols:
                    # If field has a specific index (e.g., "Secondary Project.0"), try to match that index
                    if field_index is not None and 0 <= field_index < len(secondary_project_cols):
                        # Try to find column with matching index first
                        for col in secondary_project_cols:
                            col_str = str(col)
                            col_lower = col_str.lower()
                            if "." in col_lower:
                                col_parts = col_lower.split(".", 1)
                                if len(col_parts) > 1 and col_parts[1].isdigit():
                                    if int(col_parts[1]) == field_index:
                                        return col
                        # If no exact index match, return the column at that position
                        return secondary_project_cols[field_index]
                    else:
                        # No specific index in field name, or index out of range - return first Secondary Project column
                        return secondary_project_cols[0]

                # Fallback: try any column that starts with "Secondary Project" (case insensitive)
                for col in columns:
                    col_str = str(col)
                    if col_str.lower().startswith("secondary project"):
                        return col

            # Handle Term Source ID fields - same as validation panel
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

        if validation_results and 'results' in validation_results:
            validation_data = validation_results['results']
            results_by_type = validation_data.get('experiment_results', {}) or {}
            experiment_types_processed = validation_data.get('experiment_types_processed', []) or []

            # Use the same logic as make_sheet_validation_panel_experiments
            for experiment_type in experiment_types_processed:
                et_data = results_by_type.get(experiment_type, {}) or {}
                # Use simple keys like validation panel: "invalid" and "valid"
                invalid_key = "invalid"
                valid_key = "valid"

                # Process invalid records with errors
                invalid_records = et_data.get(invalid_key, [])
                for record in invalid_records:
                    # Use same field name as validation panel
                    sample_descriptor = record.get("sample_descriptor", "")
                    if not sample_descriptor:
                        continue

                    identifier_normalized = str(sample_descriptor).strip().lower()
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
                    # Use same field name as validation panel
                    sample_descriptor = record.get("sample_descriptor", "")
                    if not sample_descriptor:
                        continue

                    identifier_normalized = str(sample_descriptor).strip().lower()
                    errors, warnings = get_all_errors_and_warnings(record)

                    if warnings:
                        if identifier_normalized not in identifier_to_field_errors:
                            identifier_to_field_errors[identifier_normalized] = {"errors": {}, "warnings": {}}
                        identifier_to_field_errors[identifier_normalized]["warnings"] = warnings

        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            # Process all sheets from all_sheets_data to ensure all sheets are included
            for sheet_name in all_sheet_names:

                # Get original sheet data
                sheet_data = all_sheets_data[sheet_name]

                # Handle empty sheets - check if it's stored with special structure
                if isinstance(sheet_data, dict) and sheet_data.get("_empty"):
                    # This is an empty sheet with preserved column structure
                    columns = sheet_data.get("_columns", [])
                    # Create DataFrame with preserved columns but no rows
                    df = pd.DataFrame(columns=columns)
                    # Write empty sheet to Excel with preserved columns
                    sheet_name_clean = sheet_name[:31]  # Excel sheet name limit
                    df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                    continue

                # Normal sheet with data
                sheet_records = sheet_data if isinstance(sheet_data, list) else []

                # Skip if somehow still empty
                if not sheet_records:
                    # Fallback: create empty DataFrame
                    df = pd.DataFrame([{}])
                    sheet_name_clean = sheet_name[:31]
                    df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                    continue

                # Convert to DataFrame
                df = pd.DataFrame(sheet_records)

                # Map field errors/warnings to columns for highlighting
                row_to_field_errors = {}  # {row_index: {"errors": {col_idx: msgs}, "warnings": {col_idx: msgs}}}
                cols_original = list(df.columns)

                for row_idx, record in enumerate(sheet_records):
                    # Use same logic as validation panel - try "Sample Descriptor" first, then "sample_descriptor"
                    sample_descriptor = str(record.get("Sample Descriptor", "") or record.get("sample_descriptor", ""))

                    # Normalize sample descriptor for matching (same as validation panel)
                    identifier_normalized = sample_descriptor.strip().lower() if sample_descriptor else ""

                    # Get field-level errors/warnings for this identifier
                    field_data = identifier_to_field_errors.get(identifier_normalized, {})
                    field_errors = field_data.get("errors", {})
                    field_warnings = field_data.get("warnings", {})

                    # Map field errors/warnings to column indices for highlighting
                    # Use the same logic as validation panel
                    if field_errors or field_warnings:
                        row_to_field_errors[row_idx] = {"errors": {}, "warnings": {}}

                        # Map error fields to columns - same logic as validation panel
                        for field, msgs in field_errors.items():
                            lower_field = field.lower()

                            # Special handling for Secondary Project: highlight ALL columns (same as validation panel)
                            if "secondary project" in lower_field:
                                # Find all Secondary Project columns
                                secondary_project_cols = [c for c in cols_original if str(c).lower().startswith("secondary project")]
                                if secondary_project_cols:
                                    # Store for ALL Secondary Project columns
                                    for sp_col in secondary_project_cols:
                                        col_idx = cols_original.index(sp_col)
                                        field_display = "Secondary Project"
                                        if col_idx not in row_to_field_errors[row_idx]["errors"]:
                                            row_to_field_errors[row_idx]["errors"][col_idx] = {
                                                "field": field_display,
                                                "messages": msgs
                                            }
                                        else:
                                            # Merge messages if multiple errors
                                            existing = row_to_field_errors[row_idx]["errors"][col_idx]["messages"]
                                            if isinstance(existing, list):
                                                existing.extend(msgs if isinstance(msgs, list) else [msgs])
                                            else:
                                                row_to_field_errors[row_idx]["errors"][col_idx]["messages"] = [existing] + (msgs if isinstance(msgs, list) else [msgs])
                                continue  # Skip normal processing for Secondary Project

                            # Normal processing for other fields
                            col = _map_field_to_column_excel(field, cols_original)
                            if col:
                                # Try to find column by exact match first
                                if col in cols_original:
                                    col_idx = cols_original.index(col)
                                else:
                                    # Try case-insensitive match
                                    col_idx = None
                                    for i, c in enumerate(cols_original):
                                        if str(c).lower() == str(col).lower():
                                            col_idx = i
                                            break

                                if col_idx is not None:
                                    # Store both messages and field name for tooltip
                                    row_to_field_errors[row_idx]["errors"][col_idx] = {
                                        "field": field,
                                        "messages": msgs
                                    }

                        # Map warning fields to columns - same logic as validation panel
                        for field, msgs in field_warnings.items():
                            lower_field = field.lower()

                            # Special handling for Secondary Project: highlight ALL columns (same as validation panel)
                            if "secondary project" in lower_field:
                                # Find all Secondary Project columns
                                secondary_project_cols = [c for c in cols_original if str(c).lower().startswith("secondary project")]
                                if secondary_project_cols:
                                    # Store for ALL Secondary Project columns
                                    for sp_col in secondary_project_cols:
                                        col_idx = cols_original.index(sp_col)
                                        field_display = "Secondary Project"
                                        if col_idx not in row_to_field_errors[row_idx]["warnings"]:
                                            row_to_field_errors[row_idx]["warnings"][col_idx] = {
                                                "field": field_display,
                                                "messages": msgs
                                            }
                                        else:
                                            # Merge messages if multiple warnings
                                            existing = row_to_field_errors[row_idx]["warnings"][col_idx]["messages"]
                                            if isinstance(existing, list):
                                                existing.extend(msgs if isinstance(msgs, list) else [msgs])
                                            else:
                                                row_to_field_errors[row_idx]["warnings"][col_idx]["messages"] = [existing] + (msgs if isinstance(msgs, list) else [msgs])
                                continue  # Skip normal processing for Secondary Project

                            # Normal processing for other fields
                            col = _map_field_to_column_excel(field, cols_original)
                            if col:
                                # Try to find column by exact match first
                                if col in cols_original:
                                    col_idx = cols_original.index(col)
                                else:
                                    # Try case-insensitive match
                                    col_idx = None
                                    for i, c in enumerate(cols_original):
                                        if str(c).lower() == str(col).lower():
                                            col_idx = i
                                            break

                                if col_idx is not None:
                                    # Store both messages and field name for tooltip
                                    row_to_field_errors[row_idx]["warnings"][col_idx] = {
                                        "field": field,
                                        "messages": msgs
                                    }

                # Clean headers to match validation results table display
                def clean_header_name(header):
                    if '.' in header:
                        return header.split('.')[0]
                    return header

                # Create a copy of the DataFrame with cleaned headers
                df_cleaned = df.copy()
                df_cleaned.columns = [clean_header_name(col) for col in df.columns]

                # Write to Excel with cleaned headers
                sheet_name_clean = sheet_name[:31]  # Excel sheet name limit
                df_cleaned.to_excel(writer, sheet_name=sheet_name_clean, index=False)

                # Get Excel formatting objects - matching validation table colors
                book = writer.book
                fmt_red = book.add_format({"bg_color": "#FFCCCC"})  # Matches #ffcccc from validation table
                fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})  # Matches #fff4cc from validation table

                ws = writer.sheets[sheet_name_clean]
                cols = list(df_cleaned.columns)  # Use cleaned column names

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
                # Errors take precedence over warnings (red highlighting if both exist)
                for row_idx, record in enumerate(sheet_records):
                    excel_row = row_idx + 1  # Excel is 1-indexed (header is row 0, data starts at row 1)

                    if row_idx in row_to_field_errors:
                        field_data = row_to_field_errors[row_idx]

                        # Get all columns that have errors or warnings
                        all_affected_cols = set()
                        all_affected_cols.update(field_data.get("errors", {}).keys())
                        all_affected_cols.update(field_data.get("warnings", {}).keys())

                        for col_idx in all_affected_cols:
                            # Use cols_original for bounds check since row_to_field_errors uses cols_original indices
                            if col_idx >= len(cols_original) or col_idx < 0:
                                continue

                            # Get cell value from the cleaned DataFrame
                            # Note: col_idx should be the same for both original and cleaned since they have same structure
                            try:
                                if row_idx < len(df_cleaned) and col_idx < len(df_cleaned.columns):
                                    cell_value = df_cleaned.iloc[row_idx, col_idx]
                                    # Handle NaN values
                                    if pd.isna(cell_value):
                                        cell_value = ""
                                else:
                                    cell_value = ""
                            except Exception:
                                cell_value = ""

                            # Check if this cell has errors (errors take precedence)
                            has_errors = col_idx in field_data.get("errors", {})
                            has_warnings = col_idx in field_data.get("warnings", {})

                            if not (has_errors or has_warnings):
                                continue

                            # Combine tooltip messages from both errors and warnings
                            tooltip_parts = []

                            if has_errors:
                                error_data = field_data["errors"][col_idx]
                                field_name = error_data.get("field",
                                                            cols_original[col_idx] if col_idx < len(cols_original) else "")
                                msgs = error_data.get("messages", [])
                                msgs_list = msgs if isinstance(msgs, list) else [msgs]
                                for msg in msgs_list:
                                    tooltip_parts.append(f"Error - {field_name}: {msg}")
                                # Highlight in red (errors take precedence) - overwrite cell with formatting
                                ws.write(excel_row, col_idx, cell_value, fmt_red)
                            elif has_warnings:
                                warning_data = field_data["warnings"][col_idx]
                                field_name = warning_data.get("field",
                                                              cols_original[col_idx] if col_idx < len(cols_original) else "")
                                msgs = warning_data.get("messages", [])
                                msgs_list = msgs if isinstance(msgs, list) else [msgs]
                                for msg in msgs_list:
                                    tooltip_parts.append(f"Warning - {field_name}: {msg}")
                                # Highlight in yellow (only warnings, no errors) - overwrite cell with formatting
                                ws.write(excel_row, col_idx, cell_value, fmt_yellow)

                            # Add warnings to tooltip even if cell is highlighted red (errors take precedence)
                            if has_errors and has_warnings:
                                warning_data = field_data["warnings"][col_idx]
                                field_name = warning_data.get("field",
                                                              cols_original[col_idx] if col_idx < len(cols_original) else "")
                                msgs = warning_data.get("messages", [])
                                msgs_list = msgs if isinstance(msgs, list) else [msgs]
                                for msg in msgs_list:
                                    tooltip_parts.append(f"Warning - {field_name}: {msg}")

                            # Add combined tooltip
                            if tooltip_parts:
                                tooltip_text = "\n".join([f"• {part}" for part in tooltip_parts])
                                # Truncate if too long
                                max_length = 2000
                                if len(tooltip_text) > max_length:
                                    tooltip_text = tooltip_text[:max_length] + "..."
                                try:
                                    ws.write_comment(excel_row, col_idx, tooltip_text,
                                                     {"visible": False, "x_scale": 1.5, "y_scale": 1.8})
                                except Exception:
                                    # If comment fails, continue without it
                                    pass

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
        sample_descriptor = str(record.get("Sample Descriptor", "") or record.get("sample_descriptor", ""))
        if sample_descriptor:
            sheet_sample_names.add(sample_descriptor)


    # Collect all rows that belong to this sheet
    error_map = {}
    warning_map = {}

    for experiment_type in experiment_types:
        et_data = results_by_type.get(experiment_type, {}) or {}
        invalid_key = "invalid"
        valid_key = "valid"

        invalid_records = et_data.get(invalid_key, [])
        valid_records = et_data.get(valid_key, [])

        for record in invalid_records + valid_records:
            sample_descriptor = record.get("sample_descriptor", "")
            if sample_descriptor in sheet_sample_names:
                errors, warnings = get_all_errors_and_warnings(record)
                if errors:
                    error_map[sample_descriptor] = errors
                if warnings:
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

        # Handle Secondary Project fields - check if field starts with "Secondary Project"
        if field_name and "secondary project" in field_name.lower():
            # Extract index from field name if present (e.g., "Secondary Project.0" -> 0, "Secondary Project.1" -> 1)
            field_index = None
            field_lower = field_name.lower()
            if "." in field_lower:
                parts = field_lower.split(".", 1)
                if len(parts) > 1 and parts[1].isdigit():
                    field_index = int(parts[1])

            # Collect all Secondary Project columns in order
            secondary_project_cols = []
            for col in columns:
                col_str = str(col)
                col_lower = col_str.lower()
                # Match exact "Secondary Project" or numbered versions like "Secondary Project.0", "Secondary Project.1"
                if col_lower == "secondary project" or col_lower.startswith("secondary project."):
                    secondary_project_cols.append(col)

            if secondary_project_cols:
                # If field has a specific index (e.g., "Secondary Project.0"), try to match that index
                if field_index is not None and 0 <= field_index < len(secondary_project_cols):
                    # Try to find column with matching index first
                    for col in secondary_project_cols:
                        col_str = str(col)
                        col_lower = col_str.lower()
                        if "." in col_lower:
                            col_parts = col_lower.split(".", 1)
                            if len(col_parts) > 1 and col_parts[1].isdigit():
                                if int(col_parts[1]) == field_index:
                                    return col
                    # If no exact index match, return the column at that position
                    return secondary_project_cols[field_index]
                else:
                    # No specific index in field name, or index out of range - return first Secondary Project column
                    return secondary_project_cols[0]

            # Fallback: try any column that starts with "Secondary Project" (case insensitive)
            for col in columns:
                col_str = str(col)
                if col_str.lower().startswith("secondary project"):
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
        # Use Sample Descriptor to match error_map keys (same as used when building error_map)
        sample_descriptor = str(row.get("Sample Descriptor", "") or row.get("sample_descriptor", ""))
        tips = {}
        row_styles = []

        if sample_descriptor in error_map:
            field_errors = error_map[sample_descriptor] or {}
            for field, msgs in field_errors.items():
                msgs_list = _as_list(msgs)
                lower_msgs = [m.lower() for m in msgs_list]
                lower_field = field.lower()

                # Special handling for Secondary Project: highlight ALL columns
                if "secondary project" in lower_field:
                    # Find all Secondary Project columns
                    secondary_project_cols = [c for c in df_all.columns if str(c).lower().startswith("secondary project")]
                    if secondary_project_cols:
                        # Apply red background to ALL Secondary Project columns
                        for sp_col in secondary_project_cols:
                            row_styles.append({'if': {'row_index': i, 'column_id': sp_col}, 'backgroundColor': '#ffcccc'})
                        # Add tooltip to first column
                        field_display = "Secondary Project"
                        msg_text = "**Error**: " + field_display + " — " + " | ".join(msgs_list)
                        if secondary_project_cols[0] in tips:
                            existing = tips[secondary_project_cols[0]].get("value", "")
                            combined = f"{existing} | {msg_text}" if existing else msg_text
                        else:
                            combined = msg_text
                        tips[secondary_project_cols[0]] = {'value': combined, 'type': 'markdown'}
                        continue  # Skip normal processing for Secondary Project

                # Normal processing for other fields
                col = _map_field_to_column(field, df_all.columns)
                if not col:
                    col = field  # Use field name if no column found

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

        if sample_descriptor in warning_map:
            field_warnings = warning_map[sample_descriptor] or {}
            for field, msgs in field_warnings.items():
                msgs_list = _as_list(msgs)
                lower_field = field.lower()

                # Special handling for Secondary Project: highlight ALL columns with yellow
                if "secondary project" in lower_field:
                    # Find all Secondary Project columns
                    secondary_project_cols = [c for c in df_all.columns if str(c).lower().startswith("secondary project")]
                    if secondary_project_cols:
                        # Apply yellow background to ALL Secondary Project columns
                        for sp_col in secondary_project_cols:
                            row_styles.append({'if': {'row_index': i, 'column_id': sp_col}, 'backgroundColor': '#fff4cc'})
                        # Add tooltip to first column
                        field_display = "Secondary Project"
                        warn_text = "**Warning**: " + field_display + " — " + " | ".join(msgs_list)
                        if secondary_project_cols[0] in tips:
                            existing = tips[secondary_project_cols[0]].get("value", "")
                            combined = f"{existing} | {warn_text}" if existing else warn_text
                        else:
                            combined = warn_text
                        tips[secondary_project_cols[0]] = {'value': combined, 'type': 'markdown'}
                        continue  # Skip normal processing for Secondary Project

                # Normal processing for other fields
                col = _map_field_to_column(field, df_all.columns)
                if not col:
                    col = field  # Use field name if no column found

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
    experiment_types_processed = validation_data.get('experiment_types_processed', []) or []

    # OPTIMIZATION: Only initialize and process sheets that were actually processed by backend
    # Filter all_sheets_data to only include sheets in experiment_types_processed
    filtered_sheets_data = {}
    if all_sheets_data and experiment_types_processed:
        filtered_sheets_data = {k: v for k, v in all_sheets_data.items() if k in experiment_types_processed}
    elif all_sheets_data:
        # Fallback: if no experiment_types_processed, use all sheets (backward compatibility)
        filtered_sheets_data = all_sheets_data

    # Initialize sheet stats only for processed sheets
    if filtered_sheets_data:
        for sheet_name in filtered_sheets_data.keys():
            sheet_stats[sheet_name] = {
                'total_records': len(filtered_sheets_data[sheet_name]) if filtered_sheets_data[sheet_name] else 0,
                'valid_records': 0,
                'error_records': 0,
                'warning_records': 0,
                'sample_status': {}
            }

    # Process each experiment type and map to sheets
    for experiment_type in experiment_types_processed:
        et_data = results_by_type.get(experiment_type, {}) or {}

        invalid_key = "invalid"
        valid_key = "valid"

        invalid_records = et_data.get(invalid_key, [])
        valid_records = et_data.get(valid_key, [])

        all_records = invalid_records + valid_records

        for record in all_records:
            sample_descriptor = record.get("Sample Descriptor", "") or record.get("sample_descriptor", "")
            if not sample_descriptor:
                continue

            errors, warnings = get_all_errors_and_warnings(record)

            # Find which sheet contains this sample
            # For experiments, try "Identifier" first, then "Sample Descriptor"
            # OPTIMIZATION: Only iterate over sheets that were processed by backend
            for sheet_name, sheet_records in (filtered_sheets_data or {}).items():
                if not sheet_records:
                    continue
                sheet_sample_names = set()
                for r in sheet_records:
                    identifier = str(r.get("Sample Descriptor", "") or r.get("sample_descriptor", ""))

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