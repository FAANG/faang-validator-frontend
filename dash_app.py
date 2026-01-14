import json
import os
import base64
import io
import re
import dash
import requests
from uuid import uuid4
from dash import dcc, html, dash_table
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH, ALL
import pandas as pd
from dash.exceptions import PreventUpdate
from typing import List, Dict, Any
from tab_components import create_tab_content

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'https://faang-validator-backend-service-964531885708.europe-west2.run.app')

# Initialize the Dash app
app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server  # Expose server variable for gunicorn

# Add custom CSS for tab label styling
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }
            #sheet-validation-content-wrapper {
                transition: opacity 0.3s ease-in-out;
            }
            /* Style for valid count in tab labels */
            .valid-count {
                color: #4CAF50 !important;
                font-weight: bold !important;
            }
            /* Style for invalid count in tab labels */
            .invalid-count {
                color: #f44336 !important;
                font-weight: bold !important;
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
'''''


def process_headers(headers: List[str]) -> List[str]:
    """Process headers according to the rules for duplicates."""
    new_headers = []
    i = 0
    while i < len(headers):
        h = headers[i]

        # Case 1: Header contains a period (.)
        if '.' in h and new_headers:
            # Concatenate with the previous header name
            prev_header = new_headers[-1]
            new_header = h.split('.')[0]
            new_headers.append(f"{prev_header} {new_header}")
        # Case 2: Consecutive duplicates
        elif i + 1 < len(headers) and headers[i + 1] == h:
            new_headers.append(h)
            while i + 1 < len(headers) and headers[i + 1] == h:
                i += 1
                new_headers.append(h)
        else:
            # Case 3: Non-consecutive duplicate
            if h in new_headers:
                # Concatenate with the last header name
                last_header = new_headers[-1] if new_headers else ""
                new_headers.append(f"{last_header}_{h}")
            else:
                new_headers.append(h)
        i += 1
    return new_headers


def build_json_data(headers: List[str], rows: List[List[str]]) -> List[Dict[str, Any]]:
    """
    Build JSON structure from processed headers and rows.
    Only include 'Health Status' if it exists in the headers.
    Always treat 'Child Of', 'Specimen Picture URL', and 'Derived From' as lists.
    Uses processed headers (from process_headers) which may rename duplicates.
    """
    grouped_data = []
    # Check if fields exist in processed headers (may be renamed for duplicates)
    has_health_status = any("Health Status" in h for h in headers)
    has_cell_type = any("Cell Type" in h for h in headers)
    has_child_of = any("Child Of" in h for h in headers)
    has_specimen_picture_url = any("Specimen Picture URL" in h for h in headers)
    has_derived_from = any("Derived From" in h for h in headers)

    for row in rows:
        record: Dict[str, Any] = {}
        if has_health_status:
            record["Health Status"] = []
        if has_cell_type:
            record["Cell Type"] = []
        if has_child_of:
            record["Child Of"] = []
        if has_specimen_picture_url:
            record["Specimen Picture URL"] = []
        if has_derived_from:
            record["Derived From"] = []

        i = 0
        while i < len(headers):
            col = headers[i]  # Processed header
            val = row[i] if i < len(row) else ""

            # Convert val to string, handling NaN and None
            if pd.isna(val) or val is None:
                val = ""
            else:
                val = str(val).strip()

            # ✅ Special handling if Health Status is in headers
            # Check if "Health Status" appears in the column name (handles renamed duplicates)
            if has_health_status and "Health Status" in col:
                # Check next column for Term Source ID (may also be renamed)
                if i + 1 < len(headers) and "Term Source ID" in headers[i + 1]:
                    term_val = row[i + 1] if i + 1 < len(row) else ""
                    if pd.isna(term_val) or term_val is None:
                        term_val = ""
                    else:
                        term_val = str(term_val).strip()

                    record["Health Status"].append({
                        "text": val,
                        "term": term_val
                    })
                    i += 2  # Skip both Health Status and Term Source ID columns
                else:
                    # No Term Source ID following, just use the text value
                    if val:
                        record["Health Status"].append({
                            "text": val,
                            "term": ""
                        })
                    i += 1
                continue

            # ✅ Special handling if Cell Type is in headers
            # Check if "Cell Type" appears in the column name (handles renamed duplicates)
            if has_cell_type and "Cell Type" in col:
                # Check next column for Term Source ID (may also be renamed)
                if i + 1 < len(headers) and "Term Source ID" in headers[i + 1]:
                    term_val = row[i + 1] if i + 1 < len(row) else ""
                    if pd.isna(term_val) or term_val is None:
                        term_val = ""
                    else:
                        term_val = str(term_val).strip()

                    record["Cell Type"].append({
                        "text": val,
                        "term": term_val
                    })
                    i += 2
                else:
                    if val:
                        record["Cell Type"].append({
                            "text": val,
                            "term": ""
                        })
                    i += 1
                continue

            # ✅ Special handling for Child Of headers
            # Check if "Child Of" appears in the column name (handles renamed duplicates)
            elif has_child_of and "Child Of" in col:
                if val:  # Only append non-empty values
                    record["Child Of"].append(val)
                i += 1
                continue

            # ✅ Special handling for Specimen Picture URL headers
            # Check if "Specimen Picture URL" appears in the column name (handles renamed duplicates)
            elif has_specimen_picture_url and "Specimen Picture URL" in col:
                if val:  # Only append non-empty values
                    record["Specimen Picture URL"].append(val)
                i += 1
                continue

            # ✅ Special handling for Derived From headers
            # Check if "Derived From" appears in the column name (handles renamed duplicates)
            elif has_derived_from and "Derived From" in col:
                if val:  # Only append non-empty values
                    record["Derived From"].append(val)
                i += 1
                continue

            # ✅ Normal processing for all other columns
            if col in record:
                if not isinstance(record[col], list):
                    record[col] = [record[col]]
                record[col].append(val)
            else:
                record[col] = val
            i += 1

        grouped_data.append(record)

    return grouped_data


def get_all_errors_and_warnings(record):
    errors = {}
    warnings = {}

    # From 'errors' object
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
    if not field:
        return None
    for c in cols:
        if c.lower() == field.lower():
            return c
    return field if field in cols else None


def _flatten_data_rows(rows, include_errors=False):
    flat = []
    for r in rows or []:
        base = {"Sample Name": r.get("sample_name")}
        data_fields = r.get("data", {}) or {}

        processed_fields = {}
        for key, value in data_fields.items():
            if key == "Health Status" and isinstance(value, list) and value:
                health_statuses = []
                for status in value:
                    if isinstance(status, dict):
                        text = status.get("text", "")
                        term = status.get("term", "")
                        if text and term:
                            health_statuses.append(f"{text} ({term})")
                        elif text:
                            health_statuses.append(text)
                        elif term:
                            health_statuses.append(term)
                processed_fields[key] = ", ".join(health_statuses)
            elif key == "Child Of" and isinstance(value, list):
                processed_fields[key] = ", ".join(str(item) for item in value if item)
            elif not isinstance(value, (str, int, float, bool, type(None))):
                processed_fields[key] = str(value) if value else ""
            else:
                processed_fields[key] = value

        base.update(processed_fields)

        if include_errors:
            errors, warnings = get_all_errors_and_warnings(r)
            if errors:
                base['errors'] = errors
            if warnings:
                base['warnings'] = warnings
        else:
            _, warnings = get_all_errors_and_warnings(r)
            if warnings:
                base['warnings'] = warnings

        flat.append(base)
    return flat


def _df(records):
    df = pd.DataFrame(records)
    if df.empty:
        return df
    lead = [c for c in ["Sample Name"] if c in df.columns]
    other = [c for c in df.columns if c not in lead]
    return df[lead + other]


def _collect_valid_records(v):
    out = []
    try:
        res = v.get("results", {}) or {}
        by_type = res.get("sample_results", {}) or {}
        for sample_type, st_data in by_type.items():
            st_key = sample_type.replace(" ", "_")
            valid_key = f"valid_{st_key}s"
            for rec in (st_data.get(valid_key) or []):
                out.append({"sample_type": sample_type, **rec})
    except Exception:
        pass
    return out


def _valid_invalid_counts(v):
    try:
        results = v.get("results", {}) or {}

        experiment_summary = results.get("experiment_summary", {}) or {}
        if experiment_summary:
            valid = experiment_summary.get("valid_experiments", 0)
            invalid = experiment_summary.get("invalid_experiments", 0)
            if valid > 0 or invalid > 0:
                return int(valid), int(invalid)

        analysis_summary = results.get("analysis_summary", {}) or {}
        if analysis_summary:
            valid = analysis_summary.get("valid_analyses", 0)
            invalid = analysis_summary.get("invalid_analyses", 0)
            if valid > 0 or invalid > 0:
                return int(valid), int(invalid)

        total_summary = results.get("total_summary", {}) or {}
        valid = total_summary.get("valid_samples", 0)
        invalid = total_summary.get("invalid_samples", 0)
        return int(valid), int(invalid)
    except Exception:
        return 0, 0


def _count_total_warnings(v):
    """Count total number of records with warnings across all sample types."""
    try:
        validation_data = v.get("results", {}) or {}
        results_by_type = validation_data.get("sample_results", {}) or {}
        sample_types = validation_data.get("sample_types_processed", []) or []

        total_warnings = 0
        for sample_type in sample_types:
            warning_count = _count_warnings_for_type(v, sample_type)
            total_warnings += warning_count

        return total_warnings
    except Exception:
        return 0


def _count_valid_invalid_for_type(validation_results_dict, sample_type):
    try:
        validation_data = validation_results_dict.get('results', {}) or {}

        sample_results = validation_data.get('sample_results', {}) or {}
        st_data = sample_results.get(sample_type, {}) or {}
        
        # Use summary field for counts (more reliable than counting records)
        summary = st_data.get('summary', {}) or {}
        valid_count = summary.get('valid', 0)
        invalid_count = summary.get('invalid', 0)
        
        return int(valid_count), int(invalid_count)

    except Exception:
        return 0, 0


def _count_warnings_for_type(validation_results_dict, sample_type):
    """Count the number of valid records with warnings for a sample type."""
    try:
        validation_data = validation_results_dict.get('results', {}) or {}
        results_by_type = validation_data.get('sample_results', {}) or {}
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = sample_type.replace(' ', '_')

        valid_key = f"valid_{st_key}s"
        valid_records = st_data.get(valid_key) or []

        # Count records that have warnings
        warning_count = 0
        for record in valid_records:
            _, warnings = get_all_errors_and_warnings(record)
            if warnings:
                warning_count += 1

        return warning_count
    except Exception:
        return 0


# Legacy function kept for backward compatibility (if needed)
# New code should use create_tab_content from tab_components
def biosamples_form():
    """Legacy function - use create_tab_content from tab_components instead"""
    from tab_components import create_biosamples_form
    return create_biosamples_form("samples")


app.layout = html.Div([
    html.Div([
        html.H1("FAANG Validation"),
        html.Div(id='dummy-output-for-reset'),
        html.Div(id='dummy-output-for-reset-experiments'),
        html.Div(id='dummy-output-for-reset-analysis'),
        html.Div(id='dummy-output-tab-styling-analysis', style={'display': 'none'}),
        # Stores for Samples tab
        dcc.Store(id='stored-file-data'),
        dcc.Store(id='stored-filename'),
        dcc.Store(id='stored-all-sheets-data'),
        dcc.Store(id='stored-sheet-names'),
        dcc.Store(id='stored-parsed-json'),  # Store parsed JSON from Excel for backend
        dcc.Store(id='error-popup-data', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet', data=None),
        dcc.Store(id='stored-json-validation-results', data=None),
        dcc.Store(id="submission-job-id"),
        dcc.Store(id="submission-status"),
        dcc.Store(id="submission-env"),
        dcc.Store(id="submission-room-id"),
        # Stores for Experiments tab
        dcc.Store(id='stored-file-data-experiments'),
        dcc.Store(id='stored-filename-experiments'),
        dcc.Store(id='stored-all-sheets-data-experiments'),
        dcc.Store(id='stored-sheet-names-experiments'),
        dcc.Store(id='stored-parsed-json-experiments'),
        dcc.Store(id='error-popup-data-experiments', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet-experiments', data=None),
        dcc.Store(id='stored-json-validation-results-experiments', data=None),
        # Stores for Analysis tab
        dcc.Store(id='stored-file-data-analysis'),
        dcc.Store(id='stored-filename-analysis'),
        dcc.Store(id='stored-all-sheets-data-analysis'),
        dcc.Store(id='stored-sheet-names-analysis'),
        dcc.Store(id='stored-parsed-json-analysis'),
        dcc.Store(id='error-popup-data-analysis', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet-analysis', data=None),
        dcc.Store(id='stored-json-validation-results-analysis', data=None),
        dcc.Download(id='download-table-csv'),
        dcc.Download(id='download-table-csv-analysis'),
        dcc.Interval(id="submission-poller", interval=2000, n_intervals=0, disabled=True),
        html.Div(
            id='error-popup-container',
            style={'display': 'none'},
            children=[
                html.Div(className='error-popup-overlay', id='error-popup-overlay'),
                html.Div(
                    className='error-popup',
                    children=[
                        html.Div(className='error-popup-close', id='error-popup-close', children='×'),
                        html.H3(className='error-popup-title', id='error-popup-title', children='Error Details'),
                        html.Div(className='error-popup-content', id='error-popup-content', children=[])
                    ]
                )
            ]
        ),
        dcc.Tabs([
            dcc.Tab(label='Samples', style={
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
                    }, children=[
                    create_tab_content('samples')
                ]),
            dcc.Tab(label='Experiments', style={
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
                    }, children=[
                    create_tab_content('experiments')
                ]),
            dcc.Tab(label='Analysis', style={
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
                    }, children=[
                    create_tab_content('analysis')
                ])
        ], style={'margin': '20px 0', 'border': 'none', 'borderBottom': '2px solid #e0e0e0'},
            colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"})
    ], className='container')
])


@app.callback(
    [Output('stored-file-data', 'data'),
     Output('stored-filename', 'data'),
     Output('file-chosen-text-samples', 'children'),
     Output('selected-file-display-samples', 'children'),
     Output('selected-file-display-samples', 'style'),
     Output('output-data-upload-samples', 'children'),
     Output('stored-all-sheets-data', 'data'),
     Output('stored-sheet-names', 'data'),
     Output('stored-parsed-json', 'data'),
     Output('active-sheet', 'data')],
    [Input('upload-data-samples', 'contents')],
    [State('upload-data-samples', 'filename')]
)
def store_file_data(contents, filename):
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
        # Decode base64 string to bytes
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
        parsed_json_data = {}  # Store parsed JSON for backend

        # Create tabs for each sheet (only if sheet has data)
        sheet_tabs = []
        sheets_with_data = []  # Track sheets that have data

        for sheet in sheet_names:
            df_sheet = excel_file.parse(sheet, dtype=str)
            df_sheet = df_sheet.fillna("")

            # Skip empty sheets (no rows or empty DataFrame)
            if df_sheet.empty or len(df_sheet) == 0:
                continue

            # Store as list-of-dicts (JSON serializable) for display
            sheet_records = df_sheet.to_dict("records")
            all_sheets_data[sheet] = sheet_records

            # Convert to JSON format for backend using build_json_data rules
            # Use ORIGINAL headers for JSON building (not processed headers)
            # Processed headers are only for display purposes
            original_headers = [str(col) for col in df_sheet.columns]

            # Process headers according to duplicate rules (for display only)
            processed_headers = process_headers(original_headers)

            # Prepare rows data
            rows = []
            for _, row in df_sheet.iterrows():
                row_list = [row[col] for col in df_sheet.columns]
                rows.append(row_list)

            # Apply build_json_data rules with processed headers (as per original rules)
            # Processed headers handle duplicates correctly (renaming them)
            parsed_json_records = build_json_data(processed_headers, rows)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data  # Update to only include sheets with data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'})
        ])

        # Display the parsed Excel data in tabs
        if len(sheets_with_data) == 0:
            # No sheets with data
            output_data_upload_children = html.Div([
                html.P("No data found in any sheet. Please upload a file with data.",
                       style={'color': 'orange', 'fontWeight': 'bold', 'margin': '20px 0'})
            ], style={'margin': '20px 0'})
        elif len(sheets_with_data) > 1:
            # Multiple sheets - show as tabs
            output_data_upload_children = html.Div([
                dcc.Tabs(
                    id='uploaded-sheets-tabs',
                    value=active_sheet,
                    children=sheet_tabs,
                    style={'margin': '20px 0', 'border': 'none'},
                    colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"}
                )
            ], style={'margin': '20px 0'})
        else:
            # Single sheet - show directly without tabs (but still use tab structure for consistency)
            output_data_upload_children = html.Div([
                html.Div([
                    sheet_tabs[0].children[0] if sheet_tabs else html.Div()
                ], style={'margin': '20px 0'})
                ,
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


# Callback to show and enable validate button when a file is uploaded
@app.callback(
    [Output('validate-button-samples', 'disabled'),
     Output('validate-button-container-samples', 'style'),
     Output('reset-button-container-samples', 'style')],
    [Input('stored-file-data', 'data')]
)
def show_and_enable_buttons(file_data):
    if file_data is None:
        return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
    else:
        return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}


# Callback to validate data when button is clicked
@app.callback(
    [Output('output-data-upload-samples', 'children', allow_duplicate=True),
     Output('stored-json-validation-results', 'data')],
    [Input('validate-button-samples', 'n_clicks')],
    [State('stored-file-data', 'data'),
     State('stored-filename', 'data'),
     State('output-data-upload-samples', 'children'),
     State('stored-all-sheets-data', 'data'),
     State('stored-sheet-names', 'data'),
     State('stored-parsed-json', 'data')],
    prevent_initial_call=True
)
def validate_data(n_clicks, contents, filename, current_children, all_sheets_data, sheet_names, parsed_json):
    if n_clicks is None or parsed_json is None:
        return current_children if current_children else html.Div([]), None

    error_data = []
    records = []
    valid_count = 0
    invalid_count = 0

    all_sheets_validation_data = {}
    json_validation_results = None

    print("file uploaded!!")
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
            # print(json.dumps(response_json))
        else:
            raise Exception(f"Error {response.status_code}: {response.text}")

        if isinstance(response_json, dict) and 'results' in response_json:
            json_validation_results = response_json
            validation_results = response_json['results']
            total_summary = validation_results.get('total_summary', {})
            valid_count = total_summary.get('valid_samples', 0)
            invalid_count = total_summary.get('invalid_samples', 0)
        else:
            validation_data = response_json[0]
            records = validation_data.get('validation_result', [])
            valid_count = validation_data.get('valid_samples', 0)
            invalid_count = validation_data.get('invalid_samples', 0)
            error_data = validation_data.get('warnings', [])
            all_sheets_validation_data = validation_data.get('all_sheets_data', {})

            if not all_sheets_validation_data and sheet_names:
                first_sheet = sheet_names[0]
                all_sheets_validation_data = {first_sheet: records}
            print(json.loads(response_json))
    except Exception as e:
        error_div = html.Div([
            html.H5(filename),
            html.P(f"Error connecting to backend API: {str(e)}", style={'color': 'red'})
        ])
        return html.Div(current_children + [error_div] if isinstance(current_children, list) else [current_children,
                                                                                                   error_div]), None

    validation_components = [
        dcc.Store(id='stored-error-data', data=error_data),
        dcc.Store(id='stored-validation-results', data={'valid_count': valid_count, 'invalid_count': invalid_count,
                                                        'all_sheets_data': all_sheets_validation_data}),
        html.H3("2. Conversion and Validation results"),

        html.Div([
            html.P("Conversion Status", style={'fontWeight': 'bold'})
            ,
            html.P("Success", style={'color': 'green', 'fontWeight': 'bold'})
            ,
            html.P("Validation Status", style={'fontWeight': 'bold'})
            ,
            html.P("Finished", style={'color': 'green', 'fontWeight': 'bold'}),
        ], style={'margin': '10px 0'})
        ,

        html.Div(id='error-table-container', style={'display': 'none'})
        ,
        html.Div(id='validation-results-container', style={'margin': '20px 0'})
    ]

    if current_children is None:
        return html.Div(validation_components), json_validation_results
    elif isinstance(current_children, list):
        modified_children = []

        for child in current_children:
            if isinstance(child, dict) and child.get('props'):
                props = child.get('props', {})
                children = props.get('children', [])

                if isinstance(children, list) and any(
                        isinstance(c, dict) and c.get('props', {}).get('id') == 'sheet-tabs-container'
                        for c in children
                ):
                    updated_child = child.copy()
                    updated_children = []

                    for c in children:
                        if isinstance(c, dict) and c.get('props'):
                            c_props = c.get('props', {})

                            if c_props.get('id') == 'original-file-heading':
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                updated_c['props']['style'] = {}
                                updated_children.append(updated_c)
                                continue
                            elif (
                                    isinstance(c_props.get('children'), str) and
                                    (c_props.get('children').startswith("File:") or
                                     c_props.get('children') == "Click 'Validate' to process the file and see results.")
                            ):
                                updated_children.append(c)
                                continue

                            if isinstance(c_props.get('children'), str) and c_props.get(
                                    'children') == "Original File Data":
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                updated_c['props']['style'] = {}
                                updated_children.append(updated_c)

                            elif c_props.get('id') == 'sheet-tabs-container':
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                updated_c['props']['style'] = {'margin': '20px 0'}
                                updated_children.append(updated_c)

                            else:
                                updated_children.append(c)
                        else:
                            updated_children.append(c)

                    updated_child['props']['children'] = updated_children
                    modified_children.append(updated_child)
                else:
                    modified_children.append(child)
            else:
                modified_children.append(child)

        return html.Div(modified_children + validation_components), json_validation_results
    else:
        return html.Div(validation_components + [current_children]), json_validation_results


@app.callback(
    Output('download-table-csv', 'data'),
    Input('download-errors-btn', 'n_clicks'),
    [State('stored-json-validation-results', 'data'),
     State('stored-all-sheets-data', 'data'),
     State('stored-sheet-names', 'data')],
    prevent_initial_call=True
)
def download_annotated_xlsx(n_clicks, validation_results, all_sheets_data, sheet_names):
    if not n_clicks:
        raise PreventUpdate

    if not all_sheets_data or not sheet_names:
        raise PreventUpdate

    # Build a mapping of sample names to their field-level errors/warnings
    # Structure: {sample_name_normalized: {"errors": {field: [msgs]}, "warnings": {field: [msgs]}}}
    sample_to_field_errors = {}

    # Helper function to map backend field names to Excel column names
    # Use the same mapping function as validation results table
    def _map_field_to_column_excel(field_name, columns):
        # Same mapping logic as _map_field_to_column in validation results
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

        # Special case for Breed Term Source ID - find Term Source ID column after Breed
        if "Breed" in field_name and "Term Source ID" in field_name:
            def _clean_col_name(col):
                col_str = str(col)
                if '.' in col_str:
                    return col_str.split('.')[0]
                return col_str

            # Find Breed column
            breed_col_idx = None
            for i, col in enumerate(columns):
                col_str = str(col)
                if "Breed" in col_str and "Term Source ID" not in col_str:
                    breed_col_idx = i
                    break

            # If Breed column found, find Term Source ID column immediately after it
            if breed_col_idx is not None:
                next_idx = breed_col_idx + 1
                if next_idx < len(columns):
                    next_col = columns[next_idx]
                    cleaned_next = _clean_col_name(next_col)
                    if cleaned_next == "Term Source ID":
                        return next_col

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

    if validation_results and 'results' in validation_results:
        validation_data = validation_results['results']
        results_by_type = validation_data.get('sample_results', {}) or {}
        sample_types = validation_data.get('sample_types_processed', []) or []

        for sample_type in sample_types:
            st_data = results_by_type.get(sample_type, {}) or {}
            st_key = sample_type.replace(' ', '_')

            # Note: Backend creates keys with double 's' for words ending in 's' (e.g., "specimens" -> "specimenss")
            # So we don't remove the extra 's' - keep it as is
            invalid_key = f"invalid_{st_key}s"
            valid_key = f"valid_{st_key}s"

            # Process invalid rows with errors
            invalid_rows_full = _flatten_data_rows(st_data.get(invalid_key), include_errors=True) or []
            for row in invalid_rows_full:
                sample_name = row.get("Sample Name", "")
                if not sample_name:
                    continue

                sample_name_normalized = str(sample_name).strip().lower()

                row_err = row.get("errors") or {}
                row_warn = row.get("warnings") or {}

                if row_err or row_warn:
                    if sample_name_normalized not in sample_to_field_errors:
                        sample_to_field_errors[sample_name_normalized] = {"errors": {}, "warnings": {}}
                    if row_err:
                        sample_to_field_errors[sample_name_normalized]["errors"] = row_err
                    if row_warn:
                        sample_to_field_errors[sample_name_normalized]["warnings"] = row_warn

            # Process valid rows with warnings
            valid_rows_full = _flatten_data_rows(st_data.get(valid_key)) or []
            for row in valid_rows_full:
                sample_name = row.get("Sample Name", "")
                if not sample_name:
                    continue

                sample_name_normalized = str(sample_name).strip().lower()

                warnings = row.get("warnings", [])
                if warnings:
                    if sample_name_normalized not in sample_to_field_errors:
                        sample_to_field_errors[sample_name_normalized] = {"errors": {}, "warnings": {}}
                    sample_to_field_errors[sample_name_normalized]["warnings"] = warnings

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
            cols_original = list(df.columns)

            for row_idx, record in enumerate(sheet_records):
                # Try to find sample name in various possible column names
                sample_name = None
                for key in ["Sample Name", "sample_name", "SampleName", "sampleName"]:
                    if key in record:
                        sample_name = str(record.get(key, ""))
                        break

                if not sample_name:
                    sample_name = str(list(record.values())[0]) if record else ""

                # Normalize sample name for matching
                sample_name_normalized = sample_name.strip().lower() if sample_name else ""

                # Get field-level errors/warnings for this sample
                field_data = sample_to_field_errors.get(sample_name_normalized, {})
                field_errors = field_data.get("errors", {})
                field_warnings = field_data.get("warnings", {})

                # Map field errors/warnings to column indices for highlighting
                if field_errors or field_warnings:
                    row_to_field_errors[row_idx] = {"errors": {}, "warnings": {}}

                    # Map error fields to columns (same logic as validation results table)
                    for field, msgs in field_errors.items():
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

                    # Map warning fields to columns (same logic as validation results table)
                    for field, msgs in field_warnings.items():
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
                        if col_idx >= len(cols) or col_idx < 0:
                            continue
                        
                        # Get cell value from the cleaned DataFrame
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
    return dcc.send_bytes(buffer.getvalue(), "annotated_template.xlsx")


@app.callback(
    Output('validation-results-container', 'children'),
    [Input('stored-json-validation-results', 'data')],
    [State('stored-sheet-names', 'data'),
     State('stored-all-sheets-data', 'data')]
)
def populate_validation_results_tabs(validation_results, sheet_names, all_sheets_data):
    if not validation_results or 'results' not in validation_results:
        return []

    if not sheet_names:
        return []

    validation_data = validation_results['results']

    # Get sample_types_processed to filter sheets
    sample_types_processed = validation_data.get('sample_types_processed', []) or []
    
    # Get sample_results for summary data (structure: sample_results[sheet_name] = {valid_{type}s: [], invalid_{type}s: [], summary: {}})

    sample_results = validation_data.get('sample_results', {}) or {}

    sheet_tabs = []
    sheets_with_data = []

    for sheet_name in sheet_names:
        # Only show sheets that are in sample_types_processed
        if sheet_name not in sample_types_processed:
            continue
        
        # Get summary from sample_results for this sample type (sheet_name)
        st_data = sample_results.get(sheet_name, {}) or {}
        
        # Use summary field for counts (more reliable than counting records)
        summary = st_data.get('summary', {}) or {}
        valid_count = summary.get('valid', 0)
        invalid_count = summary.get('invalid', 0)

        # Make sheet name title case (first letter of each word capital)
        sheet_name_title = sheet_name.title()
        
        # Create label using sample_results summary with inline green color for valid count
        label = f"{sheet_name_title} (<span style='color: #4CAF50; font-weight: bold;'>{valid_count} valid </span>/ {invalid_count} invalid)"

        sheets_with_data.append(sheet_name)
        sheet_tabs.append(
            dcc.Tab(
                label=label,
                value=sheet_name,
                id={'type': 'sheet-validation-tab', 'sheet_name': sheet_name},
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
                    'color': '#666',
                    'padding': '12px 24px',
                    'marginRight': '4px',
                    'borderRadius': '8px 8px 0 0',
                    'fontWeight': 'bold',
                    'boxShadow': '0 -2px 4px rgba(0,0,0,0.1)'
                },
                children=[]  # Content will be shown in wrapper below tabs
            )
        )

    if not sheet_tabs:
        return html.Div([
            html.P(
                "The provided data has been validated successfully with no errors or warnings. You may proceed with submission.",
                style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
        ])

    tabs = html.Div([
        dcc.Tabs(
        id='sheet-validation-tabs',
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
        ),
        html.Div(id='sheet-validation-content-wrapper', style={'marginTop': '20px'})
    ])

    # Add script to color the counts in tab labels
    # This will be executed after the tabs are rendered
    style_script = html.Script("""
           (function() {
               function styleTabLabels() {
                   // Find the tab container - FIXED: use correct ID
                   const tabContainer = document.getElementById('sheet-validation-tabs');
                   if (!tabContainer) {
                       console.log('[Tab Styling] Tab container not found: sheet-validation-tabs');
                       return;
                   }

                   // Dash tabs are typically rendered with role="tablist" and children with role="tab"
                   let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
                   
                   if (tabLabels.length === 0) {
                       // Try alternative selectors
                       tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"]');
                   }
                   
                   if (tabLabels.length === 0) {
                       tabLabels = tabContainer.querySelectorAll('div[role="tab"], button[role="tab"]');
                   }

                   console.log('[Tab Styling] Found', tabLabels.length, 'tab elements');

                   tabLabels.forEach((tab, index) => {
                       // Get the text content
                       let textElement = tab;
                       let originalText = tab.textContent || tab.innerText || '';

                       // If the tab has children, check them for text
                       if (tab.children && tab.children.length > 0) {
                           for (let child of Array.from(tab.children)) {
                               const childText = child.textContent || child.innerText || '';
                               if (childText && (childText.includes('valid') || childText.includes('invalid'))) {
                                   textElement = child;
                                   originalText = childText;
                                   break;
                               }
                           }
                       }

                       if (originalText && (originalText.includes('valid') || originalText.includes('invalid'))) {
                           // Check if already styled to avoid re-processing
                           if (textElement.querySelector && textElement.querySelector('span[style*="color"]')) {
                               return;
                           }

                           // Match pattern: (number valid / number invalid)
                           // FIXED: Better regex pattern that matches the actual format
                           const match = originalText.match(/\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/);
                           if (match) {
                               const validCount = match[1];
                               const invalidCount = match[2];
                               
                               // Replace with colored spans - color both number and word
                               const styled = originalText.replace(
                                   /\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/,
                                   '(<span style="color: #4CAF50 !important; font-weight: bold !important;">' + validCount + ' valid</span> / <span style="color: #f44336 !important; font-weight: bold !important;">' + invalidCount + ' invalid</span>)'
                               );

                               if (styled !== originalText) {
                                   try {
                                       textElement.innerHTML = styled;
                                       console.log('[Tab Styling] Successfully styled tab', index, ':', originalText.substring(0, 50));
                                   } catch (e) {
                                       console.error('[Tab Styling] Error styling tab', index, ':', e);
                                   }
                               }
                           } else {
                               // Fallback: try individual replacements if pattern doesn't match
                               // Color both number and word
                               const styled = originalText.replace(
                                   /(\\d+)\\s+valid/g, 
                                   '<span style="color: #4CAF50 !important; font-weight: bold !important;">$1 valid</span>'
                               ).replace(
                                   /(\\d+)\\s+invalid/g, 
                                   '<span style="color: #f44336 !important; font-weight: bold !important;">$1 invalid</span>'
                               );

                               if (styled !== originalText && styled.includes('<span')) {
                                   try {
                                       textElement.innerHTML = styled;
                                       console.log('[Tab Styling] Successfully styled tab (fallback)', index);
                                   } catch (e) {
                                       console.error('[Tab Styling] Error styling tab (fallback)', index, ':', e);
                                   }
                               }
                           }
                       }
                   });
               }

               // Multiple attempts to ensure tabs are styled
               function attemptStyle() {
                   styleTabLabels();
               }

               // Run immediately
               attemptStyle();

               // Run after short delays to catch tabs that render later
               setTimeout(attemptStyle, 100);
               setTimeout(attemptStyle, 300);
               setTimeout(attemptStyle, 500);
               setTimeout(attemptStyle, 1000);
               setTimeout(attemptStyle, 2000);

               // Use MutationObserver to watch for changes
               const container = document.getElementById('sheet-validation-tabs');
               if (container) {
                   const observer = new MutationObserver(function(mutations) {
                       setTimeout(attemptStyle, 50);
                   });
                   
                   observer.observe(container, { 
                       childList: true, 
                       subtree: true,
                       characterData: true,
                       attributes: true
                   });
               }

               // Also run on various events
               if (document.readyState === 'loading') {
                   document.addEventListener('DOMContentLoaded', attemptStyle);
               }

               window.addEventListener('load', function() {
                   setTimeout(attemptStyle, 200);
               });
           })();
       """)
    header_bar = html.Div(
        [
            html.Div(),
            html.Button(
                "Download annotated template",
                id="download-errors-btn",
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

    return html.Div([
        header_bar, 
        tabs, 
        style_script
    ], style={
        "marginTop": "8px",
        "transition": "opacity 0.3s ease-in-out"
    })


# Callback to populate sheet content when tab is selected
@app.callback(
    Output('sheet-validation-content-wrapper', 'children'),
    [Input('sheet-validation-tabs', 'value')],
    [State('stored-json-validation-results', 'data'),
     State('stored-all-sheets-data', 'data')],
    prevent_initial_call=False
)
def populate_sheet_validation_content(selected_sheet_name, validation_results, all_sheets_data):
    if validation_results is None or selected_sheet_name is None:
        return html.Div(style={'opacity': 0, 'transition': 'opacity 0.3s ease-in-out'})

    if not all_sheets_data or selected_sheet_name not in all_sheets_data:
        return html.Div("No data available for this sheet.", style={'opacity': 1, 'transition': 'opacity 0.3s ease-in-out'})

    # Wrap content with smooth transition
    content = make_sheet_validation_panel(selected_sheet_name, validation_results, all_sheets_data)
    return html.Div(
        content,
        style={
            'opacity': 1,
            'transition': 'opacity 0.3s ease-in-out',
            'animation': 'fadeIn 0.3s ease-in-out'
        }
    )


def make_sheet_validation_panel(sheet_name: str, validation_results: dict, all_sheets_data: dict):
    """Create a panel showing validation results for a specific Excel sheet with report at the end."""
    import uuid
    panel_id = str(uuid.uuid4())

    # Get sheet data
    sheet_records = all_sheets_data.get(sheet_name, [])
    if not sheet_records:
        return html.Div([html.H4("No data available", style={'textAlign': 'center', 'margin': '10px 0'})])

    # Get validation data
    validation_data = validation_results.get('results', {})
    results_by_type = validation_data.get('sample_results', {}) or {}
    total_summary = validation_data.get('total_summary', {})
    sample_types = validation_data.get('sample_types_processed', []) or []

    # Get all validation rows for this sheet
    sheet_sample_names = {str(record.get("Sample Name", "")) for record in sheet_records}

    # Collect all rows that belong to this sheet
    error_map = {}
    warning_map = {}

    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = sample_type.replace(' ', '_')
        # Note: Backend creates keys with double 's' for words ending in 's' (e.g., "specimens" -> "specimenss")
        # So we don't remove the extra 's' - keep it as is
        invalid_key = f"invalid_{st_key}s"
        valid_key = f"valid_{st_key}s"

        invalid_records = st_data.get(invalid_key, [])
        valid_records = st_data.get(valid_key, [])

        for record in invalid_records + valid_records:
            sample_name = record.get("sample_name", "")
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

    # Use the same styling logic as make_sample_type_panel
    def _as_list(msgs):
        if isinstance(msgs, list):
            return [str(m) for m in msgs]
        return [str(msgs)]

    def _map_field_to_column(field_name, columns):
        # Same mapping logic as in make_sample_type_panel
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

        # Special case for Breed Term Source ID - find Term Source ID column after Breed
        if "Breed" in field_name and "Term Source ID" in field_name:
            def _clean_col_name(col):
                col_str = str(col)
                if '.' in col_str:
                    return col_str.split('.')[0]
                return col_str

            # Find Breed column
            breed_col_idx = None
            for i, col in enumerate(columns):
                col_str = str(col)
                if "Breed" in col_str and "Term Source ID" not in col_str:
                    breed_col_idx = i
                    break

            # If Breed column found, find Term Source ID column immediately after it
            if breed_col_idx is not None:
                next_idx = breed_col_idx + 1
                if next_idx < len(columns):
                    next_col = columns[next_idx]
                    cleaned_next = _clean_col_name(next_col)
                    if cleaned_next == "Term Source ID":
                        return next_col

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
        sample_name = str(row.get("Sample Name", ""))
        tips = {}
        row_styles = []

        if sample_name in error_map:
            field_errors = error_map[sample_name] or {}
            for field, msgs in field_errors.items():
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

        if sample_name in warning_map:
            field_warnings = warning_map[sample_name] or {}
            for field, msgs in field_warnings.items():
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
    error_records = len([s for s in sheet_sample_names if s in error_map])
    warning_records = len([s for s in sheet_sample_names if s in warning_map and s not in error_map])
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
    # if error_fields_count:
    #     error_items = sorted(error_fields_count.items(), key=lambda x: x[1], reverse=True)
    #     report_sections.append(
    #         html.Div([
    #             html.H5("Errors by Field", style={
    #                 'marginTop': '20px',
    #                 'marginBottom': '15px',
    #                 'color': '#f44336',
    #                 'borderBottom': '2px solid #f44336',
    #                 'paddingBottom': '8px'
    #             }),
    #             html.Ul([
    #                 html.Li([
    #                     html.Span(f"{field}: ", style={'fontWeight': 'bold'}),
    #                     html.Span(f"{count} error(s)", style={'color': '#666'})
    #                 ], style={'marginBottom': '6px'})
    #                 for field, count in error_items
    #             ], style={'padding': '10px', 'backgroundColor': '#ffebee', 'borderRadius': '6px',
    #                       'listStylePosition': 'inside'})
    #         ])
    #     )

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
    if total_summary:
        total_summary_items = []
        for key, value in total_summary.items():
            if isinstance(value, (int, float, str)):
                display_key = key.replace('_', ' ').title()
                total_summary_items.append(
                    html.Div([
                        html.Span(f"{display_key}: ", style={'fontWeight': 'bold'}),
                        html.Span(str(value), style={'color': '#666'})
                    ], style={'marginBottom': '8px'})
                )
        if total_summary_items:
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
                        total_summary_items,
                        style={'padding': '10px', 'backgroundColor': '#e3f2fd', 'borderRadius': '6px'}
                    )
                ])
            )

    zebra = [{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]

    blocks = [
        html.Div([
            DataTable(
                id={"type": "sheet-result-table", "sheet_name": sheet_name, "panel_id": panel_id},
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
        ], id={"type": "sheet-table-container", "sheet_name": sheet_name, "panel_id": panel_id},
            style={'display': 'block'})
        ,
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


@app.callback(
    Output('sheet-tabs-container', 'children'),
    [Input('stored-sheet-names', 'data')],
    [State('active-sheet', 'data'),
     State('stored-all-sheets-data', 'data')]
)
def create_sheet_tabs(sheet_names, active_sheet, all_sheets_data):
    return create_sheet_tabs_ui(sheet_names, active_sheet, all_sheets_data)


@app.callback(
    [Output('active-sheet', 'data', allow_duplicate=True),
     Output('sheet-tabs-container', 'children', allow_duplicate=True)],
    [Input('sheet-tabs', 'value')],
    [State('stored-sheet-names', 'data'),
     State('stored-all-sheets-data', 'data'),
     State('active-sheet', 'data')],
    prevent_initial_call=True
)
def handle_sheet_tab_click(selected_tab_value, sheet_names, all_sheets_data, current_active_sheet):
    if selected_tab_value is None or selected_tab_value == current_active_sheet:
        return dash.no_update, dash.no_update

    clicked_sheet = selected_tab_value
    updated_tabs = create_sheet_tabs_ui(sheet_names, clicked_sheet, all_sheets_data)

    return clicked_sheet, updated_tabs


def create_sheet_tabs_ui(sheet_names, active_sheet, all_sheets_data=None):
    if not sheet_names or len(sheet_names) <= 1:
        return []

    start_index = 2
    if len(sheet_names) <= start_index:
        return []

    filtered_sheet_names = sheet_names[start_index:]

    if all_sheets_data:
        filtered_sheet_names = [sheet_name for sheet_name in filtered_sheet_names
                                if all_sheets_data.get(sheet_name, [])]

    active_tab_index = None
    for i, sheet_name in enumerate(filtered_sheet_names):
        if sheet_name == active_sheet:
            active_tab_index = i
            break
    tabs = html.Div([
        html.H4("Samples", style={'textAlign': 'center', 'marginTop': '30px', 'marginBottom': '15px'})
        ,
        dcc.Tabs(
            id='sheet-tabs',
            value=active_sheet if active_tab_index is not None else (
                filtered_sheet_names[0] if filtered_sheet_names else None),
            children=[
                dcc.Tab(
                    label=sheet_name,
                    value=sheet_name,
                    id={'type': 'sheet-tab', 'index': i + start_index},
                    style={'padding': '10px 20px', 'borderRadius': '4px 4px 0 0', 'border': 'none'},
                    selected_style={'backgroundColor': '#4CAF50', 'color': 'white', 'padding': '10px 20px',
                                    'borderRadius': '4px 4px 0 0', 'fontWeight': 'bold',
                                    'boxShadow': '0 2px 5px rgba(0,0,0,0.2)', 'border': 'none',
                                    'borderBottom': '2px solid blue'}
                ) for i, sheet_name in enumerate(filtered_sheet_names)
            ],
            style={'width': '100%', 'marginBottom': '20px', 'border': 'none'},
            colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"}
        )
    ], style={'marginTop': '30px', 'borderTop': '1px solid #ddd', 'paddingTop': '20px'})

    return tabs


def _calculate_sheet_statistics(validation_results, all_sheets_data):
    """Calculate errors and warnings count for each Excel sheet."""
    sheet_stats = {}

    if not validation_results or 'results' not in validation_results:
        return sheet_stats

    validation_data = validation_results['results']
    results_by_type = validation_data.get('sample_results', {}) or {}
    sample_types = validation_data.get('sample_types_processed', []) or []

    # Initialize sheet stats
    if all_sheets_data:
        for sheet_name in all_sheets_data.keys():
            sheet_stats[sheet_name] = {
                'total_records': len(all_sheets_data[sheet_name]) if all_sheets_data[sheet_name] else 0,
                'valid_records': 0,
                'error_records': 0,
                'warning_records': 0,
                'sample_status': {}  # {sample_name: 'error'|'warning'|'valid'}
            }

    # Process each sample type and map to sheets
    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = sample_type.replace(' ', '_')

        # Note: Backend creates keys with double 's' for words ending in 's' (e.g., "specimens" -> "specimenss")
        # So we don't remove the extra 's' - keep it as is
        invalid_key = f"invalid_{st_key}s"
        valid_key = f"valid_{st_key}s"

        invalid_records = st_data.get(invalid_key, [])
        valid_records = st_data.get(valid_key, [])

        all_records = invalid_records + valid_records

        for record in all_records:
            sample_name = record.get("sample_name", "")
            if not sample_name:
                continue

            errors, warnings = get_all_errors_and_warnings(record)

            # Find which sheet contains this sample
            for sheet_name, sheet_records in (all_sheets_data or {}).items():
                if not sheet_records:
                    continue
                sheet_sample_names = {str(r.get("Sample Name", "")) for r in sheet_records}
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


app.callback(
    Output("biosamples-form-mount", "children"),
    Input("stored-json-validation-results", "data"),
    prevent_initial_call=True,
)


def _mount_biosamples_form(v):
    if not v or "results" not in v:
        raise PreventUpdate

    results = v.get("results", {})
    if not results.get("sample_results"):
        raise PreventUpdate

    return biosamples_form()


@app.callback(
    [
        Output("biosamples-form-samples", "style"),
        Output("biosamples-status-banner-samples", "children"),
        Output("biosamples-status-banner-samples", "style"),
    ],
    Input("stored-json-validation-results", "data"),
)
def _toggle_biosamples_form(v):
    base_style = {"display": "block", "marginTop": "16px"}

    if not v or "results" not in v:
        return ({"display": "none"}, "", {"display": "none"})

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
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


@app.callback(
    [Output("biosamples-submit-btn-samples", "disabled"),
     Output("biosamples-submit-btn-samples", "style")],
    [
        Input("biosamples-username-samples", "value"),
        Input("biosamples-password-samples", "value"),
        Input("stored-json-validation-results", "data"),
    ],
)
def _disable_submit(u, p, v):
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
    valid_cnt, invalid = _valid_invalid_counts(v)
    # Enable submit button only when total samples == valid samples (i.e., invalid == 0)
    # This means all samples are valid
    if invalid > 0:
        return True, disabled_style  # Disable if there are invalid samples
    # All samples are valid, enable if username and password are provided
    is_enabled = u and p
    return not is_enabled, enabled_style if is_enabled else disabled_style


@app.callback(
    [
        Output("biosamples-submit-msg-samples", "children"),
        Output("biosamples-results-table-samples", "children"),
    ],
    Input("biosamples-submit-btn-samples", "n_clicks"),
    State("biosamples-username-samples", "value"),
    State("biosamples-password-samples", "value"),
    State("biosamples-env-samples", "value"),
    State("biosamples-action-samples", "value"),
    State("stored-json-validation-results", "data"),
    prevent_initial_call=True,
)
def _submit_to_biosamples(n, username, password, env, action, v):
    if not n:
        raise PreventUpdate

    if not v or "results" not in v:
        msg = html.Span(
            "No validation results available. Please validate your file first.",
            style={"color": "#c62828", "fontWeight": 500},
        )
        return msg, dash.no_update

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
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
                {"Sample Name": name, "BioSample ID": acc}
                for name, acc in biosamples_ids.items()
            ]

            for row in table_data:
                acc = row.get("BioSample ID")
                if acc:
                    row[
                        "BioSample ID"
                    ] = f"[{acc}](https://wwwdev.ebi.ac.uk/biosamples/samples/{acc})"

            table = dash_table.DataTable(
                data=table_data,
                columns=[
                    {"name": "Sample Name", "id": "Sample Name"},
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


@app.callback(
    Output("biosamples-form-samples", "style", allow_duplicate=True),
    Input("upload-data-samples", "contents"),
    prevent_initial_call=True,
)
def _clear_biosamples_form_on_new_upload(_):
    return {"display": "none", "marginTop": "16px"}


app.clientside_callback(
    """
    function(n_clicks) {
        if (n_clicks > 0) {
            window.location.reload();
        }
        return '';
    }
    """,
    Output("dummy-output-for-reset", "children"),
    [Input("reset-button-samples", "n_clicks")],
    prevent_initial_call=True,
)

# Clientside callback to style tab labels when validation results are updated
app.clientside_callback(
    """
    function(validation_results) {
        if (!validation_results) {
            if (window.dash_clientside && window.dash_clientside.no_update) {
                return window.dash_clientside.no_update;
            }
            return null;
        }

        // Function to style tab labels - FIXED: use correct ID
        function styleTabLabels() {
            const tabContainer = document.getElementById('sheet-validation-tabs');
            if (!tabContainer) {
                console.log('[Clientside] Tab container not found: sheet-validation-tabs');
                return;
            }

            // Try multiple selectors to find tab elements
            let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"]');
            }
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('div[role="tab"], button[role="tab"]');
            }

            console.log('[Clientside] Found', tabLabels.length, 'tab elements');

            tabLabels.forEach((tab) => {
                let textElement = tab;
                let originalText = tab.textContent || tab.innerText || '';

                // Check children for text
                if (tab.children && tab.children.length > 0) {
                    for (let child of Array.from(tab.children)) {
                        const childText = child.textContent || child.innerText || '';
                        if (childText && (childText.includes('valid') || childText.includes('invalid'))) {
                            textElement = child;
                            originalText = childText;
                            break;
                        }
                    }
                }

                if (originalText && (originalText.includes('valid') || originalText.includes('invalid'))) {
                    // Skip if already styled
                    if (textElement.querySelector && textElement.querySelector('span[style*="color"]')) {
                        return;
                    }

                    // Match pattern: (number valid / number invalid)
                    const match = originalText.match(/\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/);
                    if (match) {
                        const validCount = match[1];
                        const invalidCount = match[2];
                        
                        const styled = originalText.replace(
                            /\\((\\d+)\\s+valid\\s+\\/\\s+(\\d+)\\s+invalid\\)/,
                            '(<span style="color: #4CAF50 !important; font-weight: bold !important;">' + validCount + ' valid</span> / <span style="color: #f44336 !important; font-weight: bold !important;">' + invalidCount + ' invalid</span>)'
                        );

                        if (styled !== originalText) {
                            try {
                                textElement.innerHTML = styled;
                                console.log('[Clientside] Successfully styled tab');
                            } catch (e) {
                                console.error('[Clientside] Error styling tab:', e);
                            }
                        }
                    } else {
                        // Fallback: individual replacements
                        // Color both number and word
                        const styled = originalText.replace(
                            /(\\d+)\\s+valid/g, 
                            '<span style="color: #4CAF50 !important; font-weight: bold !important;">$1 valid</span>'
                        ).replace(
                            /(\\d+)\\s+invalid/g, 
                            '<span style="color: #f44336 !important; font-weight: bold !important;">$1 invalid</span>'
                        );

                        if (styled !== originalText && styled.includes('<span')) {
                            try {
                                textElement.innerHTML = styled;
                            } catch (e) {
                                console.error('[Clientside] Error styling tab:', e);
                            }
                        }
                    }
                }
            });
        }

        // Run after delays to ensure DOM is ready
        setTimeout(styleTabLabels, 100);
        setTimeout(styleTabLabels, 500);
        setTimeout(styleTabLabels, 1000);
        setTimeout(styleTabLabels, 2000);

        // Return no_update safely
        if (window.dash_clientside && window.dash_clientside.no_update) {
            return window.dash_clientside.no_update;
        }
        return null;
    }
    """,
    Output('validation-results-container', 'children', allow_duplicate=True),
    [Input('stored-json-validation-results', 'data')],
    prevent_initial_call='initial_duplicate'
)


def reset_app_state(n_clicks):
    if n_clicks > 0:
        return (
            f"Resetting app state (n_clicks={n_clicks})",
            None, None, "No file chosen", [], {'display': 'none'},
            [], None, None, None, None, None,
            True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
        )
    return dash.no_update


# ============================================================================
# EXPERIMENTS TAB CALLBACKS - MOVED TO experiments_tab.py
# ============================================================================
# All experiments tab callbacks are now in experiments_tab.py
# They are registered via register_experiments_callbacks() at the end of this file

# Register experiments tab callbacks
from experiments_tab import register_experiments_callbacks
register_experiments_callbacks(app)

# Register analysis tab callbacks
from analysis_tab import register_analysis_callbacks
register_analysis_callbacks(app)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
