import os
import base64
import io
import dash
import requests
from dash import dcc, html
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH, ALL
import pandas as pd
from dash.exceptions import PreventUpdate
from typing import List, Dict, Any

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
'''


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
    """
    grouped_data = []
    has_health_status = any(h.startswith("Health Status") for h in headers)
    has_cell_type = any(h.startswith("Cell Type") for h in headers)
    has_child_of = any(h == "Child Of" for h in headers)
    has_specimen_picture_url = any(h == "Specimen Picture URL" for h in headers)
    has_derived_from = any(h == "Derived From" for h in headers)

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
            col = headers[i]
            val = row[i] if i < len(row) else ""

            # Convert val to string, handling NaN and None
            if pd.isna(val) or val is None:
                val = ""
            else:
                val = str(val).strip()

            if has_health_status and col.startswith("Health Status"):
                # Skip if this header also contains "Term Source ID" (it's a processed header from process_headers)
                # In that case, it was already handled when we processed the previous "Health Status" column
                if "Term Source ID" in col:
                    i += 1
                    continue
                
                # Check next column for Term Source ID
                if i + 1 < len(headers) and "Term Source ID" in headers[i + 1]:
                    term_val = row[i + 1] if i + 1 < len(row) else ""
                    if pd.isna(term_val) or term_val is None:
                        term_val = ""
                    else:
                        term_val = str(term_val).strip()

                    # Ensure we're using the correct values: val should be text, term_val should be term
                    record["Health Status"].append({
                        "text": val,  # Current column value is the health status text
                        "term": term_val  # Next column value is the term source ID
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

            if has_cell_type and col.startswith("Cell Type"):
                # Check next column for Term Source ID
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

            elif has_child_of and col.startswith("Child Of"):
                if val:  # Only append non-empty values
                    record["Child Of"].append(val)
                i += 1
                continue

            elif has_specimen_picture_url and col.startswith("Specimen Picture URL"):
                if val:  # Only append non-empty values
                    record["Specimen Picture URL"].append(val)
                i += 1
                continue

            elif has_derived_from and col.startswith("Derived From"):
                if val:  # Only append non-empty values
                    record["Derived From"].append(val)
                i += 1
                continue

            if col in record:
                if not isinstance(record[col], list):
                    record[col] = [record[col]]
                record[col].append(val)
            else:
                record[col] = val
            i += 1

        grouped_data.append(record)

    return grouped_data


def _warnings_by_field(warnings_list):
    import re
    by_field = {}
    for w in warnings_list or []:
        m = re.search(r"Field '([^']+)'", str(w))
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
            errors = r.get("errors", {})
            if isinstance(errors, dict) and "field_errors" in errors:
                base['errors'] = errors["field_errors"]

        warnings = r.get("warnings", [])
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
        dcc.Store(id='stored-all-sheets-data'),
        dcc.Store(id='stored-sheet-names'),
        dcc.Store(id='stored-parsed-json'),  # Store parsed JSON from Excel for backend
        dcc.Store(id='error-popup-data', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet', data=None),
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
                    html.Div(id="biosamples-form-mount"),
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
            parsed_json_records = build_json_data(processed_headers, rows)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data  # Update to only include sheets with data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
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
            html.P("Conversion Status", style={'fontWeight': 'bold'}),
            html.P("Success", style={'color': 'green', 'fontWeight': 'bold'}),
            html.P("Validation Status", style={'fontWeight': 'bold'}),
            html.P("Finished", style={'color': 'green', 'fontWeight': 'bold'}),
        ], style={'margin': '10px 0'}),

        html.Div(id='error-table-container', style={'display': 'none'}),
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
                            elif (isinstance(c_props.get('children'), str) and
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
            parsed_json_records = build_json_data(processed_headers, rows)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data  # Update to only include sheets with data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading-experiments'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
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


# ========== ANALYSIS TAB CALLBACKS ==========

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
            parsed_json_records = build_json_data(processed_headers, rows)
            parsed_json_data[sheet] = parsed_json_records
            sheets_with_data.append(sheet)

        active_sheet = sheets_with_data[0] if sheets_with_data else None
        sheet_names = sheets_with_data  # Update to only include sheets with data

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading-analysis'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
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
                    id='uploaded-sheets-tabs-analysis',
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
            
            const container = document.getElementById('sample-type-tabs');
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


@app.callback(
    [
        Output("biosamples-form", "style"),
        Output("biosamples-status-banner", "children"),
        Output("biosamples-status-banner", "style"),
    ],
    Input("stored-json-validation-results", "data"),
)
def _toggle_biosamples_form(v):
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
    Output("biosamples-submit-btn", "disabled"),
    [
        Input("biosamples-username", "value"),
        Input("biosamples-password", "value"),
        Input("stored-json-validation-results", "data"),
    ],
)
def _disable_submit(u, p, v):
    if not v or "results" not in v:
        return True
    valid_cnt, _ = _valid_invalid_counts(v)
    if valid_cnt == 0:
        return True
    return not (u and p)


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
        
        function styleTabLabels() {
            const tabContainer = document.getElementById('sample-type-tabs');
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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
