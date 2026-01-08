"""
Helper functions for tab callbacks that can be reused for both Samples and Experiments tabs.
"""
import json
import base64
import io
import pandas as pd
from dash import dcc, html
from dash.exceptions import PreventUpdate
from typing import List, Dict, Any, Tuple, Optional


def process_file_upload(contents: str, filename: str, process_headers_func, build_json_data_func):
    """
    Process uploaded Excel file and return data structures.
    
    Args:
        contents: Base64 encoded file contents
        filename: Name of the uploaded file
        process_headers_func: Function to process headers
        build_json_data_func: Function to build JSON data
    
    Returns:
        Tuple of (file_data, filename, file_chosen_text, file_display, file_display_style,
                 output_children, all_sheets_data, sheet_names, parsed_json, active_sheet)
    """
    if contents is None:
        return None, None, "No file chosen", [], {'display': 'none'}, [], None, None, None, None

    try:
        content_type, content_string = contents.split(',')

        # Parse Excel file to JSON immediately
        decoded = base64.b64decode(content_string)
        excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")
        sheet_names = excel_file.sheet_names
        all_sheets_data = {}
        parsed_json_data = {}

        sheets_with_data = []

        for sheet in sheet_names:
            df_sheet = excel_file.parse(sheet, dtype=str)
            df_sheet = df_sheet.fillna("")

            # Skip empty sheets
            if df_sheet.empty or len(df_sheet) == 0:
                continue

            # Store as list-of-dicts (JSON serializable) for display
            sheet_records = df_sheet.to_dict("records")
            all_sheets_data[sheet] = sheet_records

            # Convert to JSON format for backend
            original_headers = [str(col) for col in df_sheet.columns]
            processed_headers = process_headers_func(original_headers)

            # Prepare rows data
            rows = []
            for _, row in df_sheet.iterrows():
                row_list = [row[col] for col in df_sheet.columns]
                rows.append(row_list)

            # Apply build_json_data rules
            parsed_json_records = build_json_data_func(processed_headers, rows)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'})
        ])

        # Display the parsed Excel data
        if len(sheets_with_data) == 0:
            output_data_upload_children = html.Div([
                html.P("No data found in any sheet. Please upload a file with data.",
                       style={'color': 'orange', 'fontWeight': 'bold', 'margin': '20px 0'})
            ], style={'margin': '20px 0'})
        elif len(sheets_with_data) > 1:
            output_data_upload_children = html.Div([
                dcc.Tabs(
                    id='uploaded-sheets-tabs',
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


def validate_data_logic(parsed_json: Dict, backend_url: str, filename: str):
    """
    Validate data by sending to backend API.
    
    Args:
        parsed_json: Parsed JSON data to validate
        backend_url: Backend API URL
        filename: Name of the file being validated
    
    Returns:
        Tuple of (validation_components, json_validation_results, error_div)
    """
    if parsed_json is None:
        return None, None, None

    error_data = []
    valid_count = 0
    invalid_count = 0
    all_sheets_validation_data = {}
    json_validation_results = None

    try:
        import requests
        response = requests.post(
            f'{backend_url}/validate-data',
            json={"data": parsed_json},
            headers={'accept': 'application/json', 'Content-Type': 'application/json'}
        )

        if response.status_code != 200:
            raise Exception(f"JSON endpoint returned {response.status_code}")

        if response.status_code == 200:
            response_json = response.json()

            if isinstance(response_json, dict) and 'results' in response_json:
                json_validation_results = response_json
                validation_results = response_json['results']
                total_summary = validation_results.get('total_summary', {})
                valid_count = total_summary.get('valid_samples', 0)
                invalid_count = total_summary.get('invalid_samples', 0)
            else:
                validation_data = response_json[0]
                valid_count = validation_data.get('valid_samples', 0)
                invalid_count = validation_data.get('invalid_samples', 0)
                error_data = validation_data.get('warnings', [])
                all_sheets_validation_data = validation_data.get('all_sheets_data', {})

    except Exception as e:
        error_div = html.Div([
            html.H5(filename),
            html.P(f"Error connecting to backend API: {str(e)}", style={'color': 'red'})
        ])
        return None, None, error_div

    validation_components = [
        dcc.Store(id='stored-error-data', data=error_data),
        dcc.Store(id='stored-validation-results', data={'valid_count': valid_count, 'invalid_count': invalid_count,
                                                        'all_sheets_data': all_sheets_validation_data}),
        html.H3("2. Conversion and Validation results"),
        html.Div([
            html.P("Conversion Status", style={'fontWeight': 'bold'}),
            html.P("Success", style={'color': 'green', 'fontWeight': 'bold'}),
            html.P("Validation Status", style={'fontWeight': 'bold'}),
            html.P("Finished", style={'color': 'green', 'fontWeight': 'bold'}),
        ], style={'margin': '10px 0'}),
        html.Div(id='error-table-container', style={'display': 'none'}),
        html.Div(id='validation-results-container', style={'margin': '20px 0'})
    ]

    return validation_components, json_validation_results, None


