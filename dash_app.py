import json
import os
import base64
import io
import re
import dash
import requests
from dash import dcc, html
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH, ALL
import pandas as pd
from dash.exceptions import PreventUpdate
from typing import List, Dict, Any
from json_converter import process_headers, build_json_data

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'http://localhost:8000')

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
            /* Style valid/invalid counts in tab labels */
            .tab-label-valid {
                color: #4CAF50;
                font-weight: bold;
            }
            .tab-label-invalid {
                color: #f44336;
                font-weight: bold;
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
                
                if field_to_blame not in errors:
                    errors[field_to_blame] = []
                errors[field_to_blame].append(message)


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

    # From 'relationship_errors'
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
            elif key == "Cell Type" and isinstance(value, list) and value:
                cell_types = []
                for cell_type in value:
                    if isinstance(cell_type, dict):
                        text = cell_type.get("text", "")
                        term = cell_type.get("term", "")
                        if text and term:
                            cell_types.append(f"{text} ({term})")
                        elif text:
                            cell_types.append(text)
                        elif term:
                            cell_types.append(term)
                processed_fields[key] = ", ".join(cell_types)
            # Experiment Target and chip target → format as "text (term)"
            elif key in {"Experiment Target", "experiment target", "chip target"} and isinstance(value, dict):
                text = value.get("text", "")
                term = value.get("term", "")
                if text and term:
                    processed_fields[key] = f"{text} ({term})"
                elif text:
                    processed_fields[key] = text
                elif term:
                    processed_fields[key] = term
                else:
                    processed_fields[key] = ""
            # Simple list fields → comma‑separated string
            elif key in {
                "Child Of",
                "Specimen Picture URL",
                "Derived From",
                "Secondary Project",
                "File Names",
                "File Types",
                "Checksum Methods",
                "Checksums",
                "Samples",
                "Experiments",
                "Runs",
            } and isinstance(value, list):
                processed_fields[key] = ", ".join(str(item) for item in value if item)
            # List of objects with `value` → join their values
            elif key in {"experiment type", "platform"} and isinstance(value, list):
                processed_fields[key] = ", ".join(
                    str(item.get("value"))
                    for item in value
                    if isinstance(item, dict) and item.get("value")
                )
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
        by_type = res.get("results_by_type", {}) or {}
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


def _get_data_type_label(v):
    try:
        results = v.get("results", {}) or {}
        
        if results.get("experiment_summary"):
            return "experiment(s)"
        elif results.get("analysis_summary"):
            return "analysis/analyses"
        else:
            return "sample(s)"
    except Exception:
        return "sample(s)"
def _count_total_warnings(v):
    """Count total number of records with warnings across all sample types."""
    try:
        validation_data = v.get("results", {}) or {}
        results_by_type = validation_data.get("results_by_type", {}) or {}
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
        results_by_type = validation_data.get('results_by_type', {}) or {}
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = sample_type.replace(' ', '_')

        valid_key = f"valid_{st_key}s"
        invalid_key = f"invalid_{st_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]

        if isinstance(st_data, list):
            v = sum(1 for item in st_data if not item.get('errors'))
            iv = sum(1 for item in st_data if item.get('errors'))
            return v, iv
        else:
            v = len(st_data.get(valid_key) or [])
            iv = len(st_data.get(invalid_key) or [])
            return v, iv
    except Exception:
        return 0, 0


def _count_warnings_for_type(validation_results_dict, sample_type):
    """Count the number of valid records with warnings for a sample type."""
    try:
        validation_data = validation_results_dict.get('results', {}) or {}
        results_by_type = validation_data.get('results_by_type', {}) or {}
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


def biosamples_form():
    return html.Div(
        [
            html.H2("Submit data to BioSamples", style={"marginBottom": "14px"}),

            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id="biosamples-username",
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
                id="biosamples-password",
                type="password",
                placeholder="Password",
                style={
                    "width": "100%", "padding": "10px", "borderRadius": "8px",
                    "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF",
                    "margin": "6px 0 16px"
                }
            ),

            dcc.RadioItems(
                id="biosamples-env",
                options=[{"label": " Test server", "value": "test"},
                         {"label": " Production server", "value": "prod"}],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "16px"}
            ),

            html.Div(id="biosamples-status-banner",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id="biosamples-submit-btn", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id="biosamples-submit-msg", style={"marginTop": "10px"}),
        ],
        id="biosamples-form",
        style={"display": "none", "marginTop": "16px"},
    )


def experiments_form():
    return html.Div(
        [
            html.H2("Submit data to Experiments", style={"marginBottom": "14px"}),

            dcc.RadioItems(
                id="experiments-env",
                options=[{"label": " Test server", "value": "test"},
                         {"label": " Production server", "value": "prod"}],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "16px"}
            ),

            html.Div(id="experiments-status-banner",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id="experiments-submit-btn", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id="experiments-submit-msg", style={"marginTop": "10px"}),
        ],
        id="experiments-form",
        style={"display": "none", "marginTop": "16px"},
    )


def analysis_form():
    return html.Div(
        [
            html.H2("Submit data to Analysis", style={"marginBottom": "14px"}),

            dcc.RadioItems(
                id="analysis-env",
                options=[{"label": " Test server", "value": "test"},
                         {"label": " Production server", "value": "prod"}],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "16px"}
            ),

            html.Div(id="analysis-status-banner",
                     style={"display": "none", "padding": "10px 12px", "borderRadius": "8px", "marginBottom": "12px"}),

            html.Button(
                "Submit", id="analysis-submit-btn", n_clicks=0,
                style={
                    "backgroundColor": "#673ab7", "color": "white", "padding": "10px 18px",
                    "border": "none", "borderRadius": "8px", "cursor": "pointer",
                    "fontSize": "16px", "width": "140px"
                }
            ),
            html.Div(id="analysis-submit-msg", style={"marginTop": "10px"}),
        ],
        id="analysis-form",
        style={"display": "none", "marginTop": "16px"},
    )


app.layout = html.Div([
    html.Div([
        html.H1("FAANG Validation"),
        html.Div(id='dummy-output-for-reset'),
        dcc.Store(id='stored-file-data'),
        dcc.Store(id='stored-filename'),
        dcc.Store(id='stored-parsed-json'),  # Store parsed JSON from Excel for backend
        dcc.Store(id='error-popup-data', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet', data=None),
        dcc.Store(id='stored-all-sheets-data'),
        dcc.Store(id='stored-sheet-names'),
        dcc.Store(id='stored-json-validation-results', data=None),
        # Stores for Experiments tab
        dcc.Store(id='stored-file-data-experiments'),
        dcc.Store(id='stored-filename-experiments'),
        dcc.Store(id='stored-all-sheets-data-experiments'),
        dcc.Store(id='stored-sheet-names-experiments'),
        dcc.Store(id='stored-parsed-json-experiments'),
        dcc.Store(id='active-sheet-experiments', data=None),
        dcc.Store(id='stored-json-validation-results-experiments', data=None),
        # Stores for Analysis tab
        dcc.Store(id='stored-file-data-analysis'),
        dcc.Store(id='stored-filename-analysis'),
        dcc.Store(id='stored-all-sheets-data-analysis'),
        dcc.Store(id='stored-sheet-names-analysis'),
        dcc.Store(id='stored-parsed-json-analysis'),
        dcc.Store(id='active-sheet-analysis', data=None),
        dcc.Store(id='stored-json-validation-results-analysis', data=None),
        dcc.Store(id="submission-job-id"),
        dcc.Store(id="submission-status"),
        dcc.Store(id="submission-env"),
        dcc.Store(id="submission-room-id"),
        dcc.Download(id='download-table-csv'),
        dcc.Download(id='download-table-csv-experiments'),
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
                    html.Div([
                        html.Label("1. Upload template"),
                        html.Div([
                            dcc.Upload(
                                id='upload-data',
                                children=html.Div([
                                    html.Button('Choose File',
                                                className='upload-button',
                                                style={
                                                    'backgroundColor': '#cccccc',
                                                    'color': 'black',
                                                    'padding': '10px 20px',
                                                    'border': 'none',
                                                    'borderRadius': '4px',
                                                    'cursor': 'pointer',
                                                }),
                                    html.Div('No file chosen', id='file-chosen-text')
                                ], style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'}),
                                style={'width': 'auto', 'margin': '10px 0'},
                                className='upload-area',
                                multiple=False
                            ),
                            html.Div(
                                html.Button(
                                    'Validate',
                                    id='validate-button',
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
                                id='validate-button-container',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                            html.Div(
                                html.Button(
                                    'Reset',
                                    id='reset-button',
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
                                id='reset-button-container',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                        ], style={'display': 'flex', 'alignItems': 'center'}),
                        html.Div(
                            dcc.RadioItems(
                                id='biosamples-action',
                                options=[
                                    {"label": " Submit new sample", "value": "submission"},
                                    {"label": " Update existing sample", "value": "update"},
                                ],
                                value="submission",
                                labelStyle={"marginRight": "24px"},
                                style={"marginTop": "12px"}
                            )
                        ),
                        html.Div(id='selected-file-display', style={'display': 'none'}),
                    ], style={'margin': '20px 0'}),
                    dcc.Loading(id="loading-validation", type="circle", children=html.Div(id='output-data-upload')),
                    biosamples_form(),
                    html.Div(id="biosamples-results-table")
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
                    html.Div([
                        html.Label("1. Upload template"),
                        html.Div([
                            dcc.Upload(
                                id='upload-data-experiments',
                                children=html.Div([
                                    html.Button('Choose File',
                                                className='upload-button',
                                                style={
                                                    'backgroundColor': '#cccccc',
                                                    'color': 'black',
                                                    'padding': '10px 20px',
                                                    'border': 'none',
                                                    'borderRadius': '4px',
                                                    'cursor': 'pointer',
                                                }),
                                    html.Div('No file chosen', id='file-chosen-text-experiments')
                                ], style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'}),
                                style={'width': 'auto', 'margin': '10px 0'},
                                className='upload-area',
                                multiple=False
                            ),
                            html.Div(
                                html.Button(
                                    'Validate',
                                    id='validate-button-experiments',
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
                                id='validate-button-container-experiments',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                            html.Div(
                                html.Button(
                                    'Reset',
                                    id='reset-button-experiments',
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
                                id='reset-button-container-experiments',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                        ], style={'display': 'flex', 'alignItems': 'center'}),
                        html.Div(
                            dcc.RadioItems(
                                id='experiments-action',
                                options=[
                                    {"label": " Submit new sample", "value": "submission"},
                                    {"label": " Update existing sample", "value": "update"},
                                ],
                                value="submission",
                                labelStyle={"marginRight": "24px"},
                                style={"marginTop": "12px"}
                            )
                        ),
                        html.Div(id='selected-file-display-experiments', style={'display': 'none'}),
                    ], style={'margin': '20px 0'}),
                    dcc.Loading(id="loading-validation-experiments", type="circle", children=html.Div(id='output-data-upload-experiments')),
                    html.Div(id="experiments-form-mount"),
                    html.Div(id="experiments-results-table")
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
                    html.Div([
                        html.Label("1. Upload template"),
                        html.Div([
                            dcc.Upload(
                                id='upload-data-analysis',
                                children=html.Div([
                                    html.Button('Choose File',
                                                className='upload-button',
                                                style={
                                                    'backgroundColor': '#cccccc',
                                                    'color': 'black',
                                                    'padding': '10px 20px',
                                                    'border': 'none',
                                                    'borderRadius': '4px',
                                                    'cursor': 'pointer',
                                                }),
                                    html.Div('No file chosen', id='file-chosen-text-analysis')
                                ], style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'}),
                                style={'width': 'auto', 'margin': '10px 0'},
                                className='upload-area',
                                multiple=False
                            ),
                            html.Div(
                                html.Button(
                                    'Validate',
                                    id='validate-button-analysis',
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
                                id='validate-button-container-analysis',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                            html.Div(
                                html.Button(
                                    'Reset',
                                    id='reset-button-analysis',
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
                                id='reset-button-container-analysis',
                                style={'display': 'none', 'marginLeft': '10px'}
                            ),
                        ], style={'display': 'flex', 'alignItems': 'center'}),
                        html.Div(
                            dcc.RadioItems(
                                id='analysis-action',
                                options=[
                                    {"label": " Submit new sample", "value": "submission"},
                                    {"label": " Update existing sample", "value": "update"},
                                ],
                                value="submission",
                                labelStyle={"marginRight": "24px"},
                                style={"marginTop": "12px"}
                            )
                        ),
                        html.Div(id='selected-file-display-analysis', style={'display': 'none'}),
                    ], style={'margin': '20px 0'}),
                    dcc.Loading(id="loading-validation-analysis", type="circle", children=html.Div(id='output-data-upload-analysis')),
                    html.Div(id="analysis-form-mount"),
                    html.Div(id="analysis-results-table")
                ])
        ], style={'margin': '20px 0', 'border': 'none', 'borderBottom': '2px solid #e0e0e0'},
            colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"})
    ], className='container')
])


@app.callback(
    [Output('stored-file-data', 'data'),
     Output('stored-filename', 'data'),
     Output('file-chosen-text', 'children'),
     Output('selected-file-display', 'children'),
     Output('selected-file-display', 'style'),
     Output('output-data-upload', 'children'),
     Output('stored-all-sheets-data', 'data'),
     Output('stored-sheet-names', 'data'),
     Output('stored-parsed-json', 'data'),
     Output('active-sheet', 'data')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def store_file_data(contents, filename):
    if contents is None:
        return None, None, "No file chosen", [], {'display': 'none'}, [], None, None, None, None

    try:
        content_type, content_string = contents.split(',')

        # Parse Excel file to JSON immediately
        # Decode base64 string to bytes
        decoded = base64.b64decode(content_string)
        excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")
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

            # Apply build_json_data rules with processed headers
            # Pass sheet name so analysis/experiment-specific logic can run
            parsed_json_records = build_json_data(processed_headers, rows, sheet_name=sheet)

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
    [Output('validate-button', 'disabled'),
     Output('validate-button-container', 'style'),
     Output('reset-button-container', 'style')],
    [Input('stored-file-data', 'data')]
)
def show_and_enable_buttons(file_data):
    if file_data is None:
        return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
    else:
        return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}


# Callback to validate data when button is clicked
@app.callback(
    [Output('output-data-upload', 'children', allow_duplicate=True),
     Output('stored-json-validation-results', 'data')],
    [Input('validate-button', 'n_clicks')],
    [State('stored-file-data', 'data'),
     State('stored-filename', 'data'),
     State('output-data-upload', 'children'),
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


# ========== EXPERIMENTS TAB CALLBACKS ==========

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
    if contents is None:
        return None, None, "No file chosen", [], {'display': 'none'}, [], None, None, None, None

    try:
        content_type, content_string = contents.split(',')

        # Parse Excel file to JSON immediately
        # Decode base64 string to bytes
        decoded = base64.b64decode(content_string)
        excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")
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
            # First, get original headers
            original_headers = [str(col) for col in df_sheet.columns]
            
            # Process headers according to duplicate rules
            processed_headers = process_headers(original_headers)
            
            # Prepare rows data
            rows = []
            for _, row in df_sheet.iterrows():
                row_list = [row[col] for col in df_sheet.columns]
                rows.append(row_list)
            
            # Apply build_json_data rules with processed headers
            parsed_json_records = build_json_data(processed_headers, rows, sheet_name=sheet)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data  # Update to only include sheets with data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading-experiments'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
        ])

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

        # 2) Try direct match (case-insensitive)
        direct = _resolve_col(field_name, columns)
        if direct:
            return direct

        # 2.5) Special handling for generic "Term Source ID"
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

        # 3) If field has dot notation, try using only the base name
        if "." in field_name:
            base = field_name.split(".", 1)[0]
            base_match = _resolve_col(base, columns)
            if base_match:
                return base_match

        return None
    
    if validation_results and 'results' in validation_results:
        validation_data = validation_results['results']
        results_by_type = validation_data.get('results_by_type', {}) or {}
        sample_types = validation_data.get('sample_types_processed', []) or []

        for sample_type in sample_types:
            st_data = results_by_type.get(sample_type, {}) or {}
            st_key = sample_type.replace(' ', '_')

            invalid_key = f"invalid_{st_key}s"
            if invalid_key.endswith('ss'):
                invalid_key = invalid_key[:-1]
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
                    id='uploaded-sheets-tabs-experiments',
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
                ], style={'margin': '20px 0'}),
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


@app.callback(
    [Output('validate-button-experiments', 'disabled'),
     Output('validate-button-container-experiments', 'style'),
     Output('reset-button-container-experiments', 'style')],
    [Input('stored-file-data-experiments', 'data')]
)
def show_and_enable_buttons_experiments(file_data):
    if file_data is None:
        return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
    else:
        return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}


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
    if n_clicks is None or parsed_json is None:
        return current_children if current_children else html.Div([]), None

    error_data = []
    all_sheets_validation_data = {}
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
            experiment_summary = validation_results.get('experiment_summary', {})
            total_summary = validation_results.get('total_summary', {})
            # Try experiment_summary first, fallback to total_summary
            if experiment_summary:
                valid_count = experiment_summary.get('valid_experiments', 0)
                invalid_count = experiment_summary.get('invalid_experiments', 0)
            else:
                valid_count = total_summary.get('valid_samples', 0)
                invalid_count = total_summary.get('invalid_samples', 0)
        elif isinstance(response_json, dict):
            # Backend might return data directly without 'results' wrapper
            # Try to extract data from various possible structures
            if 'total_summary' in response_json or 'experiment_summary' in response_json or 'experiment_types_processed' in response_json:
                # Data is at top level, wrap it in 'results'
                json_validation_results = {"results": response_json}
                validation_results = response_json
                experiment_summary = validation_results.get('experiment_summary', {})
                total_summary = validation_results.get('total_summary', {})
                if experiment_summary:
                    valid_count = experiment_summary.get('valid_experiments', 0)
                    invalid_count = experiment_summary.get('invalid_experiments', 0)
                else:
                    valid_count = total_summary.get('valid_samples', 0)
                    invalid_count = total_summary.get('invalid_samples', 0)
            else:
                # Unknown format, try to extract what we can
                validation_data = response_json
                valid_count = validation_data.get('valid_samples', 0) or validation_data.get('valid_experiments', 0) or validation_data.get('valid_submissions', 0)
                invalid_count = validation_data.get('invalid_samples', 0) or validation_data.get('invalid_experiments', 0) or validation_data.get('invalid_submissions', 0)
                error_data = validation_data.get('warnings', [])
                
                # Try to build a minimal structure
                json_validation_results = {
                    "results": {
                        "total_summary": {
                            "valid_samples": valid_count,
                            "invalid_samples": invalid_count
                        },
                        "results_by_type": {},
                        "experiment_types_processed": []
                    }
                }
        else:
            # Old format (list) - convert to new format for consistency
            validation_data = response_json[0] if isinstance(response_json, list) else response_json
            records = validation_data.get('validation_result', [])
            valid_count = validation_data.get('valid_samples', 0) or validation_data.get('valid_experiments', 0)
            invalid_count = validation_data.get('invalid_samples', 0) or validation_data.get('invalid_experiments', 0)
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
            
            # Build Error column and map field errors/warnings to columns
            error_column = []
            row_to_field_errors = {}  # {row_index: {"errors": {col_idx: msgs}, "warnings": {col_idx: msgs}}}
            cols_original = list(df.columns)  # Original columns before adding Error column
            
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
                
                # Aggregate error messages for Error column
                all_error_messages = []
                for field, msgs in field_errors.items():
                    msgs_list = msgs if isinstance(msgs, list) else [msgs]
                    for msg in msgs_list:
                        all_error_messages.append(f"{field}: {msg}")
                for field, msgs in field_warnings.items():
                    msgs_list = msgs if isinstance(msgs, list) else [msgs]
                    for msg in msgs_list:
                        all_error_messages.append(f"Warning: {field}: {msg}")
                
                error_column.append(" | ".join(all_error_messages) if all_error_messages else "")
                
                # Map field errors/warnings to column indices for highlighting
                if field_errors or field_warnings:
                    row_to_field_errors[row_idx] = {"errors": {}, "warnings": {}}
                    
                    # Map error fields to columns (use original columns before adding Error column)
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
            
            # Add Error column to DataFrame
            df["Error"] = error_column
            
            # Write to Excel
            sheet_name_clean = sheet_name[:31]  # Excel sheet name limit
            df.to_excel(writer, sheet_name=sheet_name_clean, index=False)
            
            # Get Excel formatting objects
            book = writer.book
            fmt_red = book.add_format({"bg_color": "#FFCCCC"})
            fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})
            
            ws = writer.sheets[sheet_name_clean]
            cols = list(df.columns)  # Now includes Error column
            error_col_idx = cols.index("Error")
            
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
                        if col_idx < len(cols):
                            cell_value = df.iat[row_idx, col_idx] if row_idx < len(df) else ""
                            ws.write(excel_row, col_idx, cell_value, fmt_red)
                            # Add tooltip/comment with error message
                            field_name = error_data.get("field", cols_original[col_idx] if col_idx < len(cols_original) else "")
                            msgs = error_data.get("messages", [])
                            tooltip_text = format_tooltip_message(field_name, msgs, is_warning=False)
                            ws.write_comment(excel_row, col_idx, tooltip_text, 
                                            {"visible": False, "x_scale": 1.5, "y_scale": 1.8})
                    
                    # Highlight warning cells (yellow) - use original column indices
                    for col_idx, warning_data in field_data.get("warnings", {}).items():
                        if col_idx < len(cols):
                            cell_value = df.iat[row_idx, col_idx] if row_idx < len(df) else ""
                            ws.write(excel_row, col_idx, cell_value, fmt_yellow)
                            # Add tooltip/comment with warning message
                            field_name = warning_data.get("field", cols_original[col_idx] if col_idx < len(cols_original) else "")
                            msgs = warning_data.get("messages", [])
                            tooltip_text = format_tooltip_message(field_name, msgs, is_warning=True)
                            ws.write_comment(excel_row, col_idx, tooltip_text,
                                            {"visible": False, "x_scale": 1.5, "y_scale": 1.8})
                
                # Format Error column cell and add tooltip
                error_text = error_column[row_idx] if row_idx < len(error_column) else ""
                if error_text:
                    # Check if this row has any errors (not just warnings)
                    has_errors = row_idx in row_to_field_errors and row_to_field_errors[row_idx].get("errors", {})
                    fmt = fmt_yellow if not has_errors else fmt_red
                    ws.write(excel_row, error_col_idx, error_text, fmt)
                    # Add tooltip/comment to Error column with all messages
                    # Format the tooltip nicely
                    tooltip_parts = []
                    if row_idx in row_to_field_errors:
                        field_data = row_to_field_errors[row_idx]
                        # Add error messages
                        for col_idx, error_data in field_data.get("errors", {}).items():
                            field_name = error_data.get("field", cols_original[col_idx] if col_idx < len(cols_original) else "")
                            msgs = error_data.get("messages", [])
                            msgs_list = msgs if isinstance(msgs, list) else [msgs]
                            for msg in msgs_list:
                                tooltip_parts.append(f"Error - {field_name}: {msg}")
                        # Add warning messages
                        for col_idx, warning_data in field_data.get("warnings", {}).items():
                            field_name = warning_data.get("field", cols_original[col_idx] if col_idx < len(cols_original) else "")
                            msgs = warning_data.get("messages", [])
                            msgs_list = msgs if isinstance(msgs, list) else [msgs]
                            for msg in msgs_list:
                                tooltip_parts.append(f"Warning - {field_name}: {msg}")
                    
                    if tooltip_parts:
                        tooltip_text = "\n".join([f"• {part}" for part in tooltip_parts])
                        # Truncate if too long
                        max_length = 2000
                        if len(tooltip_text) > max_length:
                            tooltip_text = tooltip_text[:max_length] + "..."
                        ws.write_comment(excel_row, error_col_idx, tooltip_text,
                                        {"visible": False, "x_scale": 1.5, "y_scale": 1.8})
                else:
                    ws.write(excel_row, error_col_idx, "")

    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "annotated_template.xlsx")

# ========== ANALYSIS TAB CALLBACKS ==========

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
    
    # Calculate sheet statistics
    sheet_stats = _calculate_sheet_statistics(validation_results, all_sheets_data or {})
    
    sheet_tabs = []
    sheets_with_data = []
    
    for sheet_name in sheet_names:
        # Get statistics for this sheet
        stats = sheet_stats.get(sheet_name, {})
        errors = stats.get('error_records', 0)
        warnings = stats.get('warning_records', 0)
        
        if errors > 0 or warnings > 0:
            sheets_with_data.append(sheet_name)
            valid = stats.get('valid_records', 0)
            # Create label showing counts for THIS sheet
            label = f"{sheet_name} ({valid} valid / {errors} invalid)"
            
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
                        'color': '#4CAF50',
                        'padding': '12px 24px',
                        'marginRight': '4px',
                        'borderRadius': '8px 8px 0 0',
                        'fontWeight': 'bold',
                        'boxShadow': '0 -2px 4px rgba(0,0,0,0.1)'
                    },
                    children=[html.Div(id={'type': 'sheet-validation-content', 'index': sheet_name})]
                )
            ], style={'margin': '20px 0'})
        else:
            # Single sheet - show directly without tabs (but still use tab structure for consistency)
            output_data_upload_children = html.Div([
                html.Div([
                    sheet_tabs[0].children[0] if sheet_tabs else html.Div()
                ], style={'margin': '20px 0'}),
                html.P("Click 'Validate' to send data to backend for validation.", 
                       style={'marginTop': '20px', 'fontStyle': 'italic', 'color': '#666'})
            ], style={'margin': '20px 0'})

    if not sheet_tabs:
        return html.Div([
            html.P("The provided data has been validated successfully with no errors or warnings. You may proceed with submission.", style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
        ])
        return contents, filename, filename, error_display, {'display': 'block',
                                                             'margin': '20px 0'}, [], None, None, None, None


@app.callback(
    [Output('validate-button-analysis', 'disabled'),
     Output('validate-button-container-analysis', 'style'),
     Output('reset-button-container-analysis', 'style')],
    [Input('stored-file-data-analysis', 'data')]
)
def show_and_enable_buttons_analysis(file_data):
    if file_data is None:
        return True, {'display': 'none', 'marginLeft': '10px'}, {'display': 'none', 'marginLeft': '10px'}
    else:
        return False, {'display': 'block', 'marginLeft': '10px'}, {'display': 'block', 'marginLeft': '10px'}


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
    if n_clicks is None or parsed_json is None:
        return current_children if current_children else html.Div([]), None

    error_data = []
    records = []
    valid_count = 0
    invalid_count = 0
    all_sheets_validation_data = {}
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
            # New format with 'results' key
            json_validation_results = response_json
            validation_results = response_json['results']
            analysis_summary = validation_results.get('analysis_summary', {})
            total_summary = validation_results.get('total_summary', {})
            # Try analysis_summary first, fallback to total_summar
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
    tabs = dcc.Tabs(
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
    )
    
    # Add a script component to style the tab labels with colors
    # This will be executed after the tabs are rendered
    style_script = html.Script(r"""
        (function() {
            function styleTabLabels() {
                // Find the tab container
                const tabContainer = document.getElementById('sheet-validation-tabs');
                if (!tabContainer) {
                    console.log('Tab container not found');
                    return;
                }
                
                // Dash tabs are typically rendered with role="tablist" and children with role="tab"
                // Try multiple selectors to find tab elements
                let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
                
                // If not found, try other common patterns
                if (tabLabels.length === 0) {
                    tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"]');
                }
                
                // Also try finding by data attributes or IDs
                if (tabLabels.length === 0) {
                    tabLabels = tabContainer.querySelectorAll('div, button, a');
                }
                
                console.log('Found', tabLabels.length, 'potential tab elements');
                
                tabLabels.forEach((tab, index) => {
                    // Try to find the actual text node or label element
                    let textElement = tab;
                    let originalText = tab.textContent || tab.innerText || '';
                    
                    // If the tab has children, try to find the label element
                    if (tab.children.length > 0) {
                        for (let child of tab.children) {
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
                        
                        // Create styled version with spans
                        const styled = originalText.replace(
                            /(\d+)\s+valid/g, 
                            '<span style="color: #4CAF50 !important; font-weight: bold !important;">$1 valid</span>'
                        ).replace(
                            /(\d+)\s+invalid/g, 
                            '<span style="color: #f44336 !important; font-weight: bold !important;">$1 invalid</span>'
                        );
                        
                        if (styled !== originalText && styled.includes('<span')) {
                            try {
                                textElement.innerHTML = styled;
                                console.log('Styled tab', index, ':', originalText.substring(0, 50));
                            } catch (e) {
                                console.error('Error styling tab', index, ':', e);
                            }
                        }
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


@app.callback(
    Output('download-table-csv', 'data'),
    Input('download-errors-btn', 'n_clicks'),
    State('stored-json-validation-results', 'data'),
    prevent_initial_call=True
)
def download_annotated_xlsx(n_clicks, validation_results):
    if not n_clicks or not validation_results or 'results' not in validation_results:
        raise PreventUpdate

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {}) or {}
    sample_types = validation_data.get('sample_types_processed', []) or []

    excel_sheets = {}

    def _pick_lists_from_st_data(st_data: dict):
        valid_key = None
        invalid_key = None
        valid_list = []
        invalid_list = []

        for k, v in (st_data or {}).items():
            if isinstance(v, list) and k.startswith("valid_") and valid_key is None:
                valid_key = k
                valid_list = v
            if isinstance(v, list) and k.startswith("invalid_") and invalid_key is None:
                invalid_key = k
                invalid_list = v

        return valid_key, invalid_key, valid_list or [], invalid_list or []

    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = str(sample_type).replace(" ", "_")

        valid_key, invalid_key, valid_list, invalid_list = _pick_lists_from_st_data(st_data)

        invalid_rows_full = _flatten_data_rows(invalid_list, include_errors=True) or []
        rows_for_df_err = []
        for r in invalid_rows_full:
            rc = r.copy()
            rc.pop('errors', None)
            rc.pop('warnings', None)
            rows_for_df_err.append(rc)
        df_err = _df(rows_for_df_err)

        valid_rows_full = _flatten_data_rows(valid_list) or []
        warning_rows_full = [r for r in valid_rows_full if r.get('warnings')]
        rows_for_df_warn = []
        for r in warning_rows_full:
            rc = r.copy()
            rc.pop('warnings', None)
            rows_for_df_warn.append(rc)
        df_warn = _df(rows_for_df_warn)

        if not df_err.empty:
            excel_sheets[f"{st_key[:25]}"] = {"df": df_err, "rows": invalid_rows_full, "mode": "error"}

        if not df_warn.empty:
            excel_sheets[f"{st_key[:24]}"] = {"df": df_warn, "rows": warning_rows_full, "mode": "warn"}

        if df_err.empty and df_warn.empty:
            rows_for_df_all = []
            for r in valid_rows_full:
                rc = r.copy()
                rc.pop('errors', None)
                rc.pop('warnings', None)
                rows_for_df_all.append(rc)

            df_all = _df(rows_for_df_all)
            if not df_all.empty:
                excel_sheets[f"{st_key[:28]}"] = {"df": df_all, "rows": valid_rows_full, "mode": "all"}
            else:
                print(f"[{sample_type}] df_all is empty (no rows to export)")

    if not excel_sheets:
        raise PreventUpdate

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, payload in excel_sheets.items():
            payload["df"].to_excel(writer, sheet_name=sheet_name[:31], index=False)

        book = writer.book
        fmt_red = book.add_format({"bg_color": "#FFCCCC"})
        fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})

        comment_opts = {"visible": False, "x_scale": 1.5, "y_scale": 1.8}

        def _short(kind, field, msgs):
            return (f"{kind}: {field} — " + " | ".join(str(m) for m in msgs)).replace(" | ", "\n• ")

        def _add_prompt(ws, r, c, title, full_text):
            ws.data_validation(
                r, c, r, c,
                {"validate": "any", "input_title": title[:32], "input_message": full_text[:3000], "show_input": True}
            )

        for sheet_name, payload in excel_sheets.items():
            mode = payload["mode"]
            if mode not in ("error", "warn"):
                continue

            ws = writer.sheets[sheet_name[:31]]
            df = payload["df"]
            cols = list(df.columns)
            rows_full = payload["rows"]

            for i, raw in enumerate(rows_full, start=1):
                if mode == "error":
                    row_err = raw.get("errors") or {}
                    if isinstance(row_err, dict) and "field_errors" in row_err:
                        row_err = row_err["field_errors"]

                    for field, msgs in (row_err or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)

                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        only_extra = all("extra inputs are not permitted" in str(m).lower() for m in msgs_list)

                        if only_extra:
                            ws.write(i, c, value, fmt_yellow)
                            kind = "Warning"
                        else:
                            ws.write(i, c, value, fmt_red)
                            kind = "Error"

                        text = _short(kind, field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        long_text = f"{kind}: {field} — " + " | ".join(str(m) for m in msgs_list)
                        if len(long_text) > 800:
                            _add_prompt(ws, i, c, field, long_text)

                else:
                    by_field = _warnings_by_field(raw.get("warnings", []))
                    for field, msgs in (by_field or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)
                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        ws.write(i, c, value, fmt_yellow)

                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        text = _short("Warning", field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        if len(text) > 800:
                            _add_prompt(ws, i, c, field, text)

    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "annotated.xlsx")


@app.callback(
    Output('download-table-csv-experiments', 'data'),
    Input('download-errors-btn-experiments', 'n_clicks'),
    State('stored-json-validation-results-experiments', 'data'),
    prevent_initial_call=True
)
def download_annotated_xlsx_experiments(n_clicks, validation_results):
    if not n_clicks or not validation_results or 'results' not in validation_results:
        raise PreventUpdate

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {}) or {}
    sample_types = validation_data.get('experiment_types_processed', []) or []

    excel_sheets = {}

    def _pick_lists_from_st_data(st_data: dict):
        valid_key = None
        invalid_key = None
        valid_list = []
        invalid_list = []

        for k, v in (st_data or {}).items():
            if isinstance(v, list) and k.startswith("valid_") and valid_key is None:
                valid_key = k
                valid_list = v
            if isinstance(v, list) and k.startswith("invalid_") and invalid_key is None:
                invalid_key = k
                invalid_list = v

        return valid_key, invalid_key, valid_list or [], invalid_list or []

    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = str(sample_type).replace(" ", "_")

        valid_key, invalid_key, valid_list, invalid_list = _pick_lists_from_st_data(st_data)

        invalid_rows_full = _flatten_data_rows(invalid_list, include_errors=True) or []
        rows_for_df_err = []
        for r in invalid_rows_full:
            rc = r.copy()
            rc.pop('errors', None)
            rc.pop('warnings', None)
            rows_for_df_err.append(rc)
        df_err = _df(rows_for_df_err)

        valid_rows_full = _flatten_data_rows(valid_list) or []
        warning_rows_full = [r for r in valid_rows_full if r.get('warnings')]
        rows_for_df_warn = []
        for r in warning_rows_full:
            rc = r.copy()
            rc.pop('warnings', None)
            rows_for_df_warn.append(rc)
        df_warn = _df(rows_for_df_warn)

        if not df_err.empty:
            excel_sheets[f"{st_key[:25]}"] = {"df": df_err, "rows": invalid_rows_full, "mode": "error"}

        if not df_warn.empty:
            excel_sheets[f"{st_key[:24]}"] = {"df": df_warn, "rows": warning_rows_full, "mode": "warn"}

        if df_err.empty and df_warn.empty:
            rows_for_df_all = []
            for r in valid_rows_full:
                rc = r.copy()
                rc.pop('errors', None)
                rc.pop('warnings', None)
                rows_for_df_all.append(rc)

            df_all = _df(rows_for_df_all)
            if not df_all.empty:
                excel_sheets[f"{st_key[:28]}"] = {"df": df_all, "rows": valid_rows_full, "mode": "all"}
            else:
                print(f"[{sample_type}] df_all is empty (no rows to export)")

    if not excel_sheets:
        raise PreventUpdate

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, payload in excel_sheets.items():
            payload["df"].to_excel(writer, sheet_name=sheet_name[:31], index=False)

        book = writer.book
        fmt_red = book.add_format({"bg_color": "#FFCCCC"})
        fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})

        comment_opts = {"visible": False, "x_scale": 1.5, "y_scale": 1.8}

        def _short(kind, field, msgs):
            return (f"{kind}: {field} — " + " | ".join(str(m) for m in msgs)).replace(" | ", "\n• ")

        def _add_prompt(ws, r, c, title, full_text):
            ws.data_validation(
                r, c, r, c,
                {"validate": "any", "input_title": title[:32], "input_message": full_text[:3000], "show_input": True}
            )

        for sheet_name, payload in excel_sheets.items():
            mode = payload["mode"]
            if mode not in ("error", "warn"):
                continue

            ws = writer.sheets[sheet_name[:31]]
            df = payload["df"]
            cols = list(df.columns)
            rows_full = payload["rows"]

            for i, raw in enumerate(rows_full, start=1):
                if mode == "error":
                    row_err = raw.get("errors") or {}
                    if isinstance(row_err, dict) and "field_errors" in row_err:
                        row_err = row_err["field_errors"]

                    for field, msgs in (row_err or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)

                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        only_extra = all("extra inputs are not permitted" in str(m).lower() for m in msgs_list)

                        if only_extra:
                            ws.write(i, c, value, fmt_yellow)
                            kind = "Warning"
                        else:
                            ws.write(i, c, value, fmt_red)
                            kind = "Error"

                        text = _short(kind, field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        long_text = f"{kind}: {field} — " + " | ".join(str(m) for m in msgs_list)
                        if len(long_text) > 800:
                            _add_prompt(ws, i, c, field, long_text)

                else:
                    by_field = _warnings_by_field(raw.get("warnings", []))
                    for field, msgs in (by_field or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)
                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        ws.write(i, c, value, fmt_yellow)

                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        text = _short("Warning", field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        if len(text) > 800:
                            _add_prompt(ws, i, c, field, text)

    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "annotated.xlsx")


@app.callback(
    Output('download-table-csv-analysis', 'data'),
    Input('download-errors-btn-analysis', 'n_clicks'),
    State('stored-json-validation-results-analysis', 'data'),
    prevent_initial_call=True
)
def download_annotated_xlsx_analysis(n_clicks, validation_results):
    if not n_clicks or not validation_results or 'results' not in validation_results:
        raise PreventUpdate

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {}) or {}
    sample_types = validation_data.get('analysis_types_processed', []) or []

    excel_sheets = {}

    def _pick_lists_from_st_data(st_data: dict):
        valid_key = None
        invalid_key = None
        valid_list = []
        invalid_list = []

        for k, v in (st_data or {}).items():
            if isinstance(v, list) and k.startswith("valid_") and valid_key is None:
                valid_key = k
                valid_list = v
            if isinstance(v, list) and k.startswith("invalid_") and invalid_key is None:
                invalid_key = k
                invalid_list = v

        return valid_key, invalid_key, valid_list or [], invalid_list or []

    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = str(sample_type).replace(" ", "_")

        valid_key, invalid_key, valid_list, invalid_list = _pick_lists_from_st_data(st_data)

        invalid_rows_full = _flatten_data_rows(invalid_list, include_errors=True) or []
        rows_for_df_err = []
        for r in invalid_rows_full:
            rc = r.copy()
            rc.pop('errors', None)
            rc.pop('warnings', None)
            rows_for_df_err.append(rc)
        df_err = _df(rows_for_df_err)

        valid_rows_full = _flatten_data_rows(valid_list) or []
        warning_rows_full = [r for r in valid_rows_full if r.get('warnings')]
        rows_for_df_warn = []
        for r in warning_rows_full:
            rc = r.copy()
            rc.pop('warnings', None)
            rows_for_df_warn.append(rc)
        df_warn = _df(rows_for_df_warn)

        if not df_err.empty:
            excel_sheets[f"{st_key[:25]}"] = {"df": df_err, "rows": invalid_rows_full, "mode": "error"}

        if not df_warn.empty:
            excel_sheets[f"{st_key[:24]}"] = {"df": df_warn, "rows": warning_rows_full, "mode": "warn"}

        if df_err.empty and df_warn.empty:
            rows_for_df_all = []
            for r in valid_rows_full:
                rc = r.copy()
                rc.pop('errors', None)
                rc.pop('warnings', None)
                rows_for_df_all.append(rc)

            df_all = _df(rows_for_df_all)
            if not df_all.empty:
                excel_sheets[f"{st_key[:28]}"] = {"df": df_all, "rows": valid_rows_full, "mode": "all"}
            else:
                print(f"[{sample_type}] df_all is empty (no rows to export)")

    if not excel_sheets:
        raise PreventUpdate

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, payload in excel_sheets.items():
            payload["df"].to_excel(writer, sheet_name=sheet_name[:31], index=False)

        book = writer.book
        fmt_red = book.add_format({"bg_color": "#FFCCCC"})
        fmt_yellow = book.add_format({"bg_color": "#FFF4CC"})

        comment_opts = {"visible": False, "x_scale": 1.5, "y_scale": 1.8}

        def _short(kind, field, msgs):
            return (f"{kind}: {field} — " + " | ".join(str(m) for m in msgs)).replace(" | ", "\n• ")

        def _add_prompt(ws, r, c, title, full_text):
            ws.data_validation(
                r, c, r, c,
                {"validate": "any", "input_title": title[:32], "input_message": full_text[:3000], "show_input": True}
            )

        for sheet_name, payload in excel_sheets.items():
            mode = payload["mode"]
            if mode not in ("error", "warn"):
                continue

            ws = writer.sheets[sheet_name[:31]]
            df = payload["df"]
            cols = list(df.columns)
            rows_full = payload["rows"]

            for i, raw in enumerate(rows_full, start=1):
                if mode == "error":
                    row_err = raw.get("errors") or {}
                    if isinstance(row_err, dict) and "field_errors" in row_err:
                        row_err = row_err["field_errors"]

                    for field, msgs in (row_err or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)

                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        only_extra = all("extra inputs are not permitted" in str(m).lower() for m in msgs_list)

                        if only_extra:
                            ws.write(i, c, value, fmt_yellow)
                            kind = "Warning"
                        else:
                            ws.write(i, c, value, fmt_red)
                            kind = "Error"

                        text = _short(kind, field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        long_text = f"{kind}: {field} — " + " | ".join(str(m) for m in msgs_list)
                        if len(long_text) > 800:
                            _add_prompt(ws, i, c, field, long_text)

                else:
                    by_field = _warnings_by_field(raw.get("warnings", []))
                    for field, msgs in (by_field or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name or col_name not in cols:
                            continue

                        c = cols.index(col_name)
                        value = df.iat[i - 1, c] if (i - 1) < len(df) else ""
                        ws.write(i, c, value, fmt_yellow)

                        msgs_list = msgs if isinstance(msgs, list) else [msgs]
                        text = _short("Warning", field, msgs_list)
                        ws.write_comment(i, c, text, comment_opts)

                        if len(text) > 800:
                            _add_prompt(ws, i, c, field, text)

    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "annotated.xlsx")


@app.callback(
    Output('validation-results-container', 'children'),
    [Input('stored-json-validation-results', 'data')]
)
def populate_validation_results_tabs(validation_results):
    if not validation_results or 'results' not in validation_results:
        return []

    validation_data = validation_results['results']
    sample_types = validation_data.get('sample_types_processed', []) or []
    if not sample_types:
        return []

    sample_type_tabs = []
    sample_types_with_data = []
    
    for sample_type in sample_types:
        v, iv = _count_valid_invalid_for_type(validation_results, sample_type)
        # Only include tabs that have data (valid or invalid samples)
        if v > 0 or iv > 0:
            sample_types_with_data.append(sample_type)
            # Create label as string (dcc.Tab only accepts strings for label)
            # We'll use JavaScript to add colors after rendering
            label = f"{sample_type.capitalize()} ({v} valid / {iv} invalid)"
            sample_type_tabs.append(
                dcc.Tab(
                    label=label,
                    value=sample_type,
                    id={'type': 'sample-type-tab', 'sample_type': sample_type},
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
                    children=[html.Div(id={'type': 'sample-type-content', 'index': sample_type})]
                )
            )

    if not sample_type_tabs:
        return html.Div([
            html.P("No validation data available.", style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
        ])

    tabs = dcc.Tabs(
        id='sample-type-tabs',
        value=sample_types_with_data[0] if sample_types_with_data else None,
        children=sample_type_tabs,
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
    
    # Add a script component to style the tab labels with colors
    # This will be executed after the tabs are rendered
    style_script = html.Script("""
        (function() {
            function styleTabLabels() {
                // Find the tab container
                const tabContainer = document.getElementById('sample-type-tabs');
                if (!tabContainer) {
                    console.log('Tab container not found');
                    return;
                }
                
                // Dash tabs are typically rendered with role="tablist" and children with role="tab"
                // Try multiple selectors to find tab elements
                let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
                
                // If not found, try other common patterns
                if (tabLabels.length === 0) {
                    tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"]');
                }
                
                // Also try finding by data attributes or IDs
                if (tabLabels.length === 0) {
                    tabLabels = tabContainer.querySelectorAll('div, button, a');
                }
                
                console.log('Found', tabLabels.length, 'potential tab elements');
                
                tabLabels.forEach((tab, index) => {
                    // Try to find the actual text node or label element
                    let textElement = tab;
                    let originalText = tab.textContent || tab.innerText || '';
                    
                    // If the tab has children, try to find the label element
                    if (tab.children.length > 0) {
                        for (let child of tab.children) {
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
                        
                        // Create styled version with spans
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
                                console.log('Styled tab', index, ':', originalText.substring(0, 50));
                            } catch (e) {
                                console.error('Error styling tab', index, ':', e);
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
            
            // Run after short delays
            setTimeout(attemptStyle, 100);
            setTimeout(attemptStyle, 300);
            setTimeout(attemptStyle, 500);
            setTimeout(attemptStyle, 1000);
            
            // Use MutationObserver to watch for changes
            const observer = new MutationObserver(function(mutations) {
                setTimeout(attemptStyle, 50);
            });
            
            const container = document.getElementById('sheet-validation-tabs');
            if (container) {
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
            
            // Watch for Dash renderer updates
            if (window.dash_clientside) {
                window.dash_clientside.no_update = function() {
                    setTimeout(attemptStyle, 100);
                };
            }
        })();
    """
    )

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

    return html.Div([header_bar, tabs, style_script], style={"marginTop": "8px"})


# Callback to populate sample type content when tab is selected
@app.callback(
    Output({'type': 'sample-type-content', 'index': MATCH}, 'children'),
    [Input('sample-type-tabs', 'value')],
    [State('stored-json-validation-results', 'data'),
     State('stored-all-sheets-data', 'data')]
)
def populate_sample_type_content(selected_sample_type, validation_results, all_sheets_data):
    if validation_results is None or selected_sample_type is None:
        return []

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {})

    if selected_sample_type not in results_by_type:
        return html.Div("No data available for this sample type.")

    return make_sample_type_panel(selected_sample_type, results_by_type, all_sheets_data)


# ========== EXPERIMENTS VALIDATION RESULTS CALLBACKS ==========

@app.callback(
    Output('validation-results-container-experiments', 'children'),
    [Input('stored-json-validation-results-experiments', 'data')]
)
def populate_validation_results_tabs_experiments(validation_results):
    if not validation_results or 'results' not in validation_results:
        return []

    validation_data = validation_results['results']
    sample_types = validation_data.get('experiment_types_processed', []) or []
    if not sample_types:
        return []

    sample_type_tabs = []
    sample_types_with_data = []
    
    for sample_type in sample_types:
        v, iv = _count_valid_invalid_for_type(validation_results, sample_type)
        if v > 0 or iv > 0:
            sample_types_with_data.append(sample_type)
            label = f"{sample_type.capitalize()} ({v} valid / {iv} invalid)"
            sample_type_tabs.append(
                dcc.Tab(
                    label=label,
                    value=sample_type,
                    id={'type': 'sample-type-tab-experiments', 'sample_type': sample_type},
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
                    children=[html.Div(id={'type': 'sample-type-content-experiments', 'index': sample_type})]
                )
            )

    if not sample_type_tabs:
        return html.Div([
            html.P("No validation data available.", style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
        ])

    tabs = dcc.Tabs(
        id='sample-type-tabs-experiments',
        value=sample_types_with_data[0] if sample_types_with_data else None,
        children=sample_type_tabs,
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


@app.callback(
    Output({'type': 'sample-type-content-experiments', 'index': MATCH}, 'children'),
    [Input('sample-type-tabs-experiments', 'value')],
    [State('stored-json-validation-results-experiments', 'data'),
     State('stored-all-sheets-data-experiments', 'data')]
)
def populate_sample_type_content_experiments(selected_sample_type, validation_results, all_sheets_data):
    if validation_results is None or selected_sample_type is None:
        return []

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {})

    if selected_sample_type not in results_by_type:
        return html.Div("No data available for this sample type.")

    return make_sample_type_panel(selected_sample_type, results_by_type, all_sheets_data)


# ========== ANALYSIS VALIDATION RESULTS CALLBACKS ==========

@app.callback(
    Output('validation-results-container-analysis', 'children'),
    [Input('stored-json-validation-results-analysis', 'data')]
)
def populate_validation_results_tabs_analysis(validation_results):
    if not validation_results or 'results' not in validation_results:
        return []

    validation_data = validation_results['results']
    sample_types = validation_data.get('analysis_types_processed', []) or []
    if not sample_types:
        return []

    sample_type_tabs = []
    sample_types_with_data = []
    
    for sample_type in sample_types:
        v, iv = _count_valid_invalid_for_type(validation_results, sample_type)
        if v > 0 or iv > 0:
            sample_types_with_data.append(sample_type)
            label = f"{sample_type.capitalize()} ({v} valid / {iv} invalid)"
            sample_type_tabs.append(
                dcc.Tab(
                    label=label,
                    value=sample_type,
                    id={'type': 'sample-type-tab-analysis', 'sample_type': sample_type},
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
                    children=[html.Div(id={'type': 'sample-type-content-analysis', 'index': sample_type})]
                )
            )

    if not sample_type_tabs:
        return html.Div([
            html.P("No validation data available.", style={'textAlign': 'center', 'padding': '20px', 'color': '#666'})
        ])

    tabs = dcc.Tabs(
        id='sample-type-tabs-analysis',
        value=sample_types_with_data[0] if sample_types_with_data else None,
        children=sample_type_tabs,
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


@app.callback(
    Output({'type': 'sample-type-content-analysis', 'index': MATCH}, 'children'),
    [Input('sample-type-tabs-analysis', 'value')],
    [State('stored-json-validation-results-analysis', 'data'),
     State('stored-all-sheets-data-analysis', 'data')]
)
def populate_sample_type_content_analysis(selected_sample_type, validation_results, all_sheets_data):
    if validation_results is None or selected_sample_type is None:
        return []

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {})

    if selected_sample_type not in results_by_type:
        return html.Div("No data available for this sample type.")

    return make_sample_type_panel(selected_sample_type, results_by_type, all_sheets_data)


@app.callback(
    Output('sheet-tabs-container', 'children'),
    [Input('stored-sheet-names', 'data')],
    [State('active-sheet', 'data'),
     State('stored-all-sheets-data', 'data')]
)
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
        html.H4("Samples", style={'textAlign': 'center', 'marginTop': '30px', 'marginBottom': '15px'}),
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


def make_sample_type_panel(sample_type: str, results_by_type: dict, all_sheets_data: dict = None):
    import uuid
    panel_id = str(uuid.uuid4())

    st_data = results_by_type.get(sample_type, {}) or {}
    st_key = sample_type.replace(' ', '_')
    valid_key = f"valid_{st_key}s"
    invalid_key = f"invalid_{st_key}s"
    if invalid_key.endswith('ss'):
        invalid_key = invalid_key[:-1]

    # Handle both list format (from backend) and dict format (from old format)
    if isinstance(st_data, list):
        # Backend returns list directly - split into valid/invalid
        valid_list = []
        invalid_list = []
        for item in st_data:
            errors = item.get('errors')
            has_errors = False
            if errors:
                if isinstance(errors, dict):
                    has_errors = bool(errors.get('field_errors') or errors.get('errors'))
                elif isinstance(errors, list):
                    has_errors = len(errors) > 0
                else:
                    has_errors = bool(errors)
            if has_errors:
                invalid_list.append(item)
            else:
                valid_list.append(item)
        invalid_rows = _flatten_data_rows(invalid_list, include_errors=True)
        valid_rows = _flatten_data_rows(valid_list)
    else:
        # Old format with valid_* and invalid_* keys
        invalid_rows = _flatten_data_rows(st_data.get(invalid_key) or [], include_errors=True)
        valid_rows = _flatten_data_rows(st_data.get(valid_key) or [])
    
    all_rows = invalid_rows + valid_rows

    # Create a mapping of sample_name to error/warning info
    error_map = {}  # {sample_name: {field: [errors]}}
    warning_map = {}  # {sample_name: {field: [warnings]}}
    
    def _as_list(msgs):
        if isinstance(msgs, list):
            return [str(m) for m in msgs]
        return [str(msgs)]

    # Map errors from invalid rows
    for row in invalid_rows:
        sample_name = row.get("Sample Name", "")
        if not sample_name:
            continue
        row_err = row.get("errors") or {}
        if isinstance(row_err, dict) and "field_errors" in row_err:
            row_err = row_err["field_errors"]
        error_map[sample_name] = row_err

    # Map warnings from valid rows
    for row in valid_rows:
        sample_name = row.get("Sample Name", "")
        if not sample_name:
            continue
        warnings = row.get("warnings", [])
        if warnings:
            warning_map[sample_name] = _warnings_by_field(warnings)

    # Get original sheet data - try to find matching sheet
    original_data = []
    if all_sheets_data:
        # Find the first sheet that has data matching our sample names
        sample_names_set = {row.get("Sample Name", "") for row in all_rows if row.get("Sample Name")}
        for sheet_name, sheet_records in all_sheets_data.items():
            if sheet_records:
                # Check if this sheet has matching sample names
                sheet_sample_names = {str(record.get("Sample Name", "")) for record in sheet_records}
                if sample_names_set.intersection(sheet_sample_names):
                    original_data = sheet_records
                    break

    # Use validation data (which includes all fields from backend)
    # This ensures all fields like "Experiment Target" are included
    rows_for_df = []
    for row in all_rows:
        rc = row.copy()
        rc.pop('errors', None)
        rc.pop('warnings', None)
        rows_for_df.append(rc)
    df_all = _df(rows_for_df)
    
    # If original_data is available, merge any additional columns that might be missing
    if original_data and not df_all.empty:
        # Get columns from original data that might not be in validation data
        orig_df = pd.DataFrame(original_data)
        for col in orig_df.columns:
            if col not in df_all.columns:
                # Add missing column from original data, matching by Sample Name
                df_all[col] = None
                for idx, row in df_all.iterrows():
                    sample_name = str(row.get("Sample Name", ""))
                    matching_orig = [r for r in original_data if str(r.get("Sample Name", "")) == sample_name]
                    if matching_orig:
                        df_all.at[idx, col] = matching_orig[0].get(col)

    if df_all.empty:
        return html.Div([html.H4("No data available", style={'textAlign': 'center', 'margin': '10px 0'})])

    # Build cell styles and tooltips for all rows
    cell_styles = []
    tooltip_data = []
    cols_with_real_errors = set()

    for i, row in df_all.iterrows():
        sample_name = str(row.get("Sample Name", ""))
        tips = {}
        row_styles = []

        # Check for errors
        if sample_name in error_map:
            field_errors = error_map[sample_name]
            for field, msgs in (field_errors or {}).items():
                col = _resolve_col(field, df_all.columns)
                if not col:
                    continue
                msgs_list = _as_list(msgs)
                is_extra = any("extra inputs are not permitted" in m.lower() for m in msgs_list)

                if is_extra:
                    row_styles.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#fff4cc'})
                    tips[col] = {'value': f"**Warning**: {field} — " + " | ".join(msgs_list), 'type': 'markdown'}
                else:
                    row_styles.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#ffcccc'})
                    tips[col] = {'value': f"**Error**: {field} — " + " | ".join(msgs_list), 'type': 'markdown'}
                    cols_with_real_errors.add(col)

        # Check for warnings
        if sample_name in warning_map:
            field_warnings = warning_map[sample_name]
            for field, msgs in (field_warnings or {}).items():
                col = _resolve_col(field, df_all.columns)
                if not col:
                    continue
                # Only add warning style if not already styled by error
                if not any(s.get('column_id') == col for s in row_styles):
                    row_styles.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#fff4cc'})
                if col not in tips:
                    tips[col] = {'value': f"**Warning**: {field} — " + " | ".join(map(str, msgs)), 'type': 'markdown'}

        cell_styles.extend(row_styles)
        tooltip_data.append(tips)

    tint_whole_columns = [
        {'if': {'column_id': c}, 'backgroundColor': '#ffd6d6'}
        for c in sorted(cols_with_real_errors)
    ]

    base_cell = {"textAlign": "left", "padding": "6px", "minWidth": 120, "whiteSpace": "normal", "height": "auto"}
    zebra = [{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]

    # Create columns with cleaned display names (remove part after period)
    def clean_header_name(header):
        """Remove part after period from header name for display"""
        if '.' in header:
            return header.split('.')[0]
        return header
    
    columns = [{"name": clean_header_name(c), "id": c} for c in df_all.columns]
    
    blocks = [
        html.H4("Validation Results", style={'textAlign': 'center', 'margin': '10px 0'}),
        html.Div([
            DataTable(
                id={"type": "result-table-all", "sample_type": sample_type, "panel_id": panel_id},
                data=df_all.to_dict("records"),
                columns=columns,
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left", "padding": "6px"},
                style_header={"fontWeight": "bold", "backgroundColor": "rgb(230, 230, 230)"},
                # style_cell=base_cell,
                # style_header={"fontWeight": "bold"},
                style_data_conditional=zebra + cell_styles + tint_whole_columns,
                tooltip_data=tooltip_data,
                tooltip_duration=None
            )
        ], id={"type": "table-container-all", "sample_type": sample_type, "panel_id": panel_id},
            style={'display': 'block'}),
    ]

    return html.Div(blocks)

@app.callback(
    Output("biosamples-form-mount", "children"),
    Input("stored-json-validation-results", "data"),
    prevent_initial_call=True,
)
def _mount_biosamples_form(v):
    if not v or "results" not in v:
        raise PreventUpdate

    results = v.get("results", {})
    if not results.get("results_by_type"):
        raise PreventUpdate

    return biosamples_form()


# Callback to populate sheet content when tab is selected
@app.callback(
    Output({'type': 'sheet-validation-content', 'index': MATCH}, 'children'),
    [Input('sheet-validation-tabs', 'value')],
    [State('stored-json-validation-results', 'data'),
     State('stored-all-sheets-data', 'data')]
)
def populate_sheet_validation_content(selected_sheet_name, validation_results, all_sheets_data):
    if validation_results is None or selected_sheet_name is None:
        return []

    if not all_sheets_data or selected_sheet_name not in all_sheets_data:
        return html.Div("No data available for this sheet.")
    
    return make_sheet_validation_panel(selected_sheet_name, validation_results, all_sheets_data)


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
    results_by_type = validation_data.get('results_by_type', {}) or {}
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
        invalid_key = f"invalid_{st_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
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
                    col = field # Use field name if no column found
                
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
                    col = field # Use field name if no column found

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
                ], style={'padding': '10px', 'backgroundColor': '#ffebee', 'borderRadius': '6px', 'listStylePosition': 'inside'})
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
                ], style={'padding': '10px', 'backgroundColor': '#fff3e0', 'borderRadius': '6px', 'listStylePosition': 'inside'})
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
        html.H4(f"Validation Results - {sheet_name}", style={'textAlign': 'center', 'margin': '10px 0'})
        ,
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
    [
        Output("biosamples-submit-msg", "children"),
        Output("biosamples-results-table", "children"),
    ],
    Input("biosamples-submit-btn", "n_clicks"),
    State("biosamples-username", "value"),
    State("biosamples-password", "value"),
    State("biosamples-env", "value"),
    State("biosamples-action", "value"),
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
    data_label = _get_data_type_label(v)
    submit_label = "samples" if "sample" in data_label else ("experiments" if "experiment" in data_label else "analyses")
    if valid_cnt == 0:
        msg = html.Span(
            f"No valid {submit_label} to submit. Please fix errors and re-validate.",
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

            table = DataTable(
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
    Output("biosamples-form-mount", "children", allow_duplicate=True),
    Input("upload-data", "contents"),
    prevent_initial_call=True,
)
def _clear_biosamples_form_on_new_upload(_):
    return []


# ========== EXPERIMENTS SUBMISSION CALLBACKS ==========

@app.callback(
    Output("experiments-form-mount", "children"),
    Input("stored-json-validation-results-experiments", "data"),
    prevent_initial_call=True,
)
def _mount_experiments_form(v):
    if not v or "results" not in v:
        raise PreventUpdate

    results = v.get("results", {})
    if not results.get("results_by_type"):
        raise PreventUpdate

    return experiments_form()


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


@app.callback(
    [
        Output("experiments-form", "style"),
        Output("experiments-status-banner", "children"),
        Output("experiments-status-banner", "style"),
    ],
    Input("stored-json-validation-results-experiments", "data"),
)
def _toggle_experiments_form(v):
    base_style = {"display": "block", "marginTop": "16px"}

    if not v or "results" not in v:
        return ({"display": "none"}, "", {"display": "none"})

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
    data_label = _get_data_type_label(v)
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
                f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid {data_label}."
            ),
            html.Br(),
        ]
        return base_style, msg_children, style_ok
    else:
        submit_label = "samples" if "sample" in data_label else ("experiments" if "experiment" in data_label else "analyses")
        return (
            base_style,
            f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid {data_label}. No valid {submit_label} to submit.",
            style_warn,
        )


@app.callback(
    Output("experiments-submit-btn", "disabled"),
    [
        Input("stored-json-validation-results-experiments", "data"),
    ],
)
def _disable_experiments_submit(v):
    if not v or "results" not in v:
        return True
    valid_cnt, _ = _valid_invalid_counts(v)
    if valid_cnt == 0:
        return True
    return False


@app.callback(
    [
        Output("experiments-submit-msg", "children"),
        Output("experiments-results-table", "children"),
    ],
    Input("experiments-submit-btn", "n_clicks"),
    State("experiments-env", "value"),
    State("experiments-action", "value"),
    State("stored-json-validation-results-experiments", "data"),
    prevent_initial_call=True,
)
def _submit_to_experiments(n, env, action, v):
    if not n:
        raise PreventUpdate

    if not v or "results" not in v:
        msg = html.Span(
            "No validation results available. Please validate your file first.",
            style={"color": "#c62828", "fontWeight": 500},
        )
        return msg, dash.no_update

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
    data_label = _get_data_type_label(v)
    submit_label = "samples" if "sample" in data_label else ("experiments" if "experiment" in data_label else "analyses")
    if valid_cnt == 0:
        msg = html.Span(
            f"No valid {submit_label} to submit. Please fix errors and re-validate.",
            style={"color": "#c62828", "fontWeight": 500},
        )
        return msg, dash.no_update

    validation_results = v["results"]

    body = {
        "validation_results": validation_results,
        "mode": env,
        "update_existing": action == "update",
    }

    try:
        url = f"{BACKEND_API_URL}/submit-to-experiments"
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
        experiment_ids = data.get("experiment_ids") or {}

        color = "#388e3c" if success else "#c62828"

        msg_children = [html.Span(message, style={"fontWeight": 500})]
        if submitted_count is not None:
            msg_children += [
                html.Br(),
                html.Span(f"Submitted experiments: {submitted_count}"),
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

        if experiment_ids:
            table_data = [
                {"Experiment Name": name, "Experiment ID": exp_id}
                for name, exp_id in experiment_ids.items()
            ]

            table = DataTable(
                data=table_data,
                columns=[
                    {"name": "Experiment Name", "id": "Experiment Name"},
                    {"name": "Experiment ID", "id": "Experiment ID"},
                ],
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left"},
            )
        else:
            table = html.Div(
                "No experiment IDs returned.",
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
    Output("experiments-form-mount", "children", allow_duplicate=True),
    Input("upload-data-experiments", "contents"),
    prevent_initial_call=True,
)
def _clear_experiments_form_on_new_upload(_):
    return []


# ========== ANALYSIS SUBMISSION CALLBACKS ==========

@app.callback(
    Output("analysis-form-mount", "children"),
    Input("stored-json-validation-results-analysis", "data"),
    prevent_initial_call=True,
)
def _mount_analysis_form(v):
def _calculate_sheet_statistics(validation_results, all_sheets_data):
    """Calculate errors and warnings count for each Excel sheet."""
    sheet_stats = {}
    
    if not validation_results or 'results' not in validation_results:
        return sheet_stats
    
    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {}) or {}
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
        
        invalid_key = f"invalid_{st_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
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
    if not results.get("results_by_type"):
        raise PreventUpdate

    return analysis_form()


@app.callback(
    [
        Output("analysis-form", "style"),
        Output("analysis-status-banner", "children"),
        Output("analysis-status-banner", "style"),
    ],
    Input("stored-json-validation-results-analysis", "data"),
)
def _toggle_analysis_form(v):
    base_style = {"display": "block", "marginTop": "16px"}

    if not v or "results" not in v:
        return {"display": "none"}, "", {"display": "none"}

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
    data_label = _get_data_type_label(v)
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
                f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid {data_label}."
            ),
            html.Br(),
        ]
        return base_style, msg_children, style_ok
    else:
        submit_label = "samples" if "sample" in data_label else ("experiments" if "experiment" in data_label else "analyses")
        return (
            base_style,
            f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid {data_label}. No valid {submit_label} to submit.",
            style_warn,
        )


@app.callback(
    Output("analysis-submit-btn", "disabled"),
    [
        Input("stored-json-validation-results-analysis", "data"),
    ],
)
def _disable_analysis_submit(v):
    if not v or "results" not in v:
        return True
    valid_cnt, _ = _valid_invalid_counts(v)
    if valid_cnt == 0:
        return True
    return False


@app.callback(
    [
        Output("analysis-submit-msg", "children"),
        Output("analysis-results-table", "children"),
    ],
    Input("analysis-submit-btn", "n_clicks"),
    State("analysis-env", "value"),
    State("analysis-action", "value"),
    State("stored-json-validation-results-analysis", "data"),
    prevent_initial_call=True,
)
def _submit_to_analysis(n, env, action, v):
    if not n:
        raise PreventUpdate

    if not v or "results" not in v:
        msg = html.Span(
            "No validation results available. Please validate your file first.",
            style={"color": "#c62828", "fontWeight": 500},
        )
        return msg, dash.no_update

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
    data_label = _get_data_type_label(v)
    submit_label = "samples" if "sample" in data_label else ("experiments" if "experiment" in data_label else "analyses")
    if valid_cnt == 0:
        msg = html.Span(
            f"No valid {submit_label} to submit. Please fix errors and re-validate.",
            style={"color": "#c62828", "fontWeight": 500},
        )
        return msg, dash.no_update

    validation_results = v["results"]

    body = {
        "validation_results": validation_results,
        "mode": env,
        "update_existing": action == "update",
    }

    try:
        url = f"{BACKEND_API_URL}/submit-to-analysis"
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
        analysis_ids = data.get("analysis_ids") or {}

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

        if analysis_ids:
            table_data = [
                {"Analysis Name": name, "Analysis ID": anal_id}
                for name, anal_id in analysis_ids.items()
            ]

            table = DataTable(
                data=table_data,
                columns=[
                    {"name": "Analysis Name", "id": "Analysis Name"},
                    {"name": "Analysis ID", "id": "Analysis ID"},
                ],
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left"},
            )
        else:
            table = html.Div(
                "No analysis IDs returned.",
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
    Output("analysis-form-mount", "children", allow_duplicate=True),
    Input("upload-data-analysis", "contents"),
    prevent_initial_call=True,
)
def _clear_analysis_form_on_new_upload(_):
    return []


app.clientside_callback(
    """
    function(n1, n2, n3) {
        if ((n1 && n1 > 0) || (n2 && n2 > 0) || (n3 && n3 > 0)) {
            window.location.reload();
        }
        return '';
    }
    """,
    Output("dummy-output-for-reset", "children"),
    [Input("reset-button", "n_clicks"),
     Input("reset-button-experiments", "n_clicks"),
     Input("reset-button-analysis", "n_clicks")],
    prevent_initial_call=True,
)

# Clientside callback to style tab labels when validation results are updated
app.clientside_callback(
    """
    function(validation_results) {
        if (!validation_results) {
            return window.dash_clientside.no_update;
        }
        

        // Function to style tab labels
        function styleTabLabels() {
            const tabContainer = document.getElementById('sample-type-tabs');
            if (!tabContainer) {
                return;
            }
            

            // Try multiple selectors to find tab elements
            let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"], div[class*="tab"]');
            }
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('div, button, a, span');
            }

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
                    if (textElement.querySelector && textElement.querySelector('span[style*="color"]')) {
                        return;
                    }
                    

                    // Create styled version
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
                            console.error('Error styling tab:', e);
                        }
                    }
                }
            });
        }
        

        // Run after delays to ensure DOM is ready
        setTimeout(styleTabLabels, 100);
        setTimeout(styleTabLabels, 500);
        setTimeout(styleTabLabels, 1000);

        return window.dash_clientside.no_update;
    }
    """,
    Output('validation-results-container', 'children', allow_duplicate=True),
    [Input('stored-json-validation-results', 'data')],
    prevent_initial_call='initial_duplicate'
)

app.clientside_callback(
    """
    function(validation_results) {
        if (!validation_results) {
            return window.dash_clientside.no_update;
        }
        
        function styleTabLabels() {
            const tabContainer = document.getElementById('sample-type-tabs-experiments');
            if (!tabContainer) {
                return;
            }
            
            let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"], div[class*="tab"]');
            }
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('div, button, a, span');
            }
            
            tabLabels.forEach((tab) => {
                let textElement = tab;
                let originalText = tab.textContent || tab.innerText || '';
                
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
                    if (textElement.querySelector && textElement.querySelector('span[style*="color"]')) {
                        return;
                    }
                    
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
                            console.error('Error styling tab:', e);
                        }
                    }
                }
            });
        }
        
        setTimeout(styleTabLabels, 100);
        setTimeout(styleTabLabels, 500);
        setTimeout(styleTabLabels, 1000);
        
        return window.dash_clientside.no_update;
    }
    """,
    Output('validation-results-container-experiments', 'children', allow_duplicate=True),
    [Input('stored-json-validation-results-experiments', 'data')],
    prevent_initial_call='initial_duplicate'
)

app.clientside_callback(
    """
    function(validation_results) {
        if (!validation_results) {
            return window.dash_clientside.no_update;
        }
        
        function styleTabLabels() {
            const tabContainer = document.getElementById('sample-type-tabs-analysis');
            if (!tabContainer) {
                return;
            }
            
            let tabLabels = tabContainer.querySelectorAll('[role="tab"]');
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('.tab, [class*="tab"], [class*="Tab"], div[class*="tab"]');
            }
            if (tabLabels.length === 0) {
                tabLabels = tabContainer.querySelectorAll('div, button, a, span');
            }
            
            tabLabels.forEach((tab) => {
                let textElement = tab;
                let originalText = tab.textContent || tab.innerText || '';
                
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
                    if (textElement.querySelector && textElement.querySelector('span[style*="color"]')) {
                        return;
                    }
                    
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
                            console.error('Error styling tab:', e);
                        }
                    }
                }
            });
        }
        
        setTimeout(styleTabLabels, 100);
        setTimeout(styleTabLabels, 500);
        setTimeout(styleTabLabels, 1000);
        
        return window.dash_clientside.no_update;
    }
    """,
    Output('validation-results-container-analysis', 'children', allow_duplicate=True),
    [Input('stored-json-validation-results-analysis', 'data')],
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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
