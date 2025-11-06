import os
import base64
import io
import dash
import requests
from dash import dcc, html, dash_table
from dash.dash_table import DataTable
from dash.dependencies import Input, Output, State, MATCH, ALL
import pandas as pd
from dash.exceptions import PreventUpdate

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL',
                                 'https://faang-validator-backend-service-964531885708.europe-west2.run.app/api')

# Initialize the Dash app
app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server  # Expose server variable for gunicorn

# --- App Layout ---
app.layout = html.Div([
    html.Div([
        html.H1("FAANG Validation"),
        html.Div(id='dummy-output-for-reset'),
        # Store for uploaded file data
        dcc.Store(id='stored-file-data'),
        dcc.Store(id='stored-filename'),
        dcc.Store(id='stored-all-sheets-data'),
        dcc.Store(id='stored-sheet-names'),
        dcc.Store(id='error-popup-data', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet', data=None),
        dcc.Store(id='stored-json-validation-results', data=None),

        dcc.Download(id='download-table-csv'),

        # Error popup
        html.Div(
            id='error-popup-container',
            style={'display': 'none'},
            children=[
                html.Div(
                    className='error-popup-overlay',
                    id='error-popup-overlay'
                ),
                html.Div(
                    className='error-popup',
                    children=[
                        html.Div(
                            className='error-popup-close',
                            id='error-popup-close',
                            children='×'
                        ),
                        html.H3(
                            className='error-popup-title',
                            id='error-popup-title',
                            children='Error Details'
                        ),
                        html.Div(
                            className='error-popup-content',
                            id='error-popup-content',
                            children=[]
                        )
                    ]
                )
            ]
        ),

        # Tabs
        dcc.Tabs([
            # Samples Tab
            dcc.Tab(label='Samples', style={'border': 'none'},
                    selected_style={'border': 'none', 'borderBottom': '2px solid blue'}, children=[
                    # File Upload
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
                                                }
                                                ),
                                    html.Div('No file chosen', id='file-chosen-text')
                                ], style={'display': 'flex', 'alignItems': 'center', 'gap': '10px'}),
                                style={
                                    'width': 'auto',
                                    'margin': '10px 0',
                                },
                                className='upload-area',
                                multiple=False
                            ),
                            # Validate button container - initially hidden
                            html.Div(
                                html.Button(
                                    'Validate',
                                    id='validate-button',
                                    className='validate-button',
                                    disabled=True,  # Initially disabled until a file is uploaded
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
                                style={'display': 'none', 'marginLeft': '10px'}  # Initially hidden
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
                        html.Div(id='selected-file-display', style={'display': 'none'}),
                    ], style={'margin': '20px 0'}),

                    dcc.Loading(
                        id="loading-validation",
                        type="circle",
                        children=html.Div(id='output-data-upload')
                    ),
                    html.Div(id="biosamples-form-mount")
                ]),

            # Experiments Tab (empty for now)
            dcc.Tab(label='Experiments', style={'border': 'none'},
                    selected_style={'border': 'none', 'borderBottom': '2px solid blue'}, children=[
                    html.Div([], style={'margin': '20px 0'})
                ]),

            # Analysis Tab (empty for now)
            dcc.Tab(label='Analysis', style={'border': 'none'},
                    selected_style={'border': 'none', 'borderBottom': '2px solid blue'}, children=[
                    html.Div([], style={'margin': '20px 0'})
                ])
        ], style={'margin': '20px 0', 'border': 'none'},
            colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"})
    ], className='container')
])


@app.callback(
    Output("biosamples-form-mount", "children"),
    Input("stored-json-validation-results", "data"),
    prevent_initial_call=True
)
def _mount_biosamples_form(v):
    if not v or "results" not in v:
        raise PreventUpdate

    results = v.get("results", {})
    if not results.get("results_by_type"):
        raise PreventUpdate

    return biosamples_form()


# Callback to store uploaded file data and display filename
@app.callback(
    [Output('stored-file-data', 'data'),
     Output('stored-filename', 'data'),
     Output('file-chosen-text', 'children'),
     Output('selected-file-display', 'children'),
     Output('selected-file-display', 'style'),
     Output('output-data-upload', 'children'),
     Output('stored-all-sheets-data', 'data'),
     Output('stored-sheet-names', 'data'),
     Output('active-sheet', 'data')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def store_file_data(contents, filename):
    if contents is None:
        return None, None, "No file chosen", [], {'display': 'none'}, [], None, None, None

    try:
        content_type, content_string = contents.split(',')

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
            html.P("Click 'Validate' to process the file and see results."),
        ])

        output_data_upload_children = html.Div(id='sheet-tabs-container', style={'margin': '20px 0', 'display': 'none'})

        all_sheets_data = {}
        sheet_names = []
        active_sheet = None

        return contents, filename, filename, file_selected_display, {'display': 'block', 'margin': '20px 0'}, [
            output_data_upload_children], all_sheets_data, sheet_names, active_sheet

    except Exception as e:
        error_display = html.Div([
            html.H5(filename),
            html.P(f"Error processing file: {str(e)}", style={'color': 'red'})
        ])
        return contents, filename, filename, error_display, {'display': 'block',
                                                             'margin': '20px 0'}, [], None, None, None


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
     State('stored-sheet-names', 'data')],
    prevent_initial_call=True
)
def validate_data(n_clicks, contents, filename, current_children, all_sheets_data, sheet_names):
    if n_clicks is None or contents is None:
        return current_children if current_children else html.Div([]), None

    error_data = []
    records = []
    valid_count = 0
    invalid_count = 0
    all_sheets_validation_data = {}
    json_validation_results = None

    try:
        # with open('validation_results.json', 'r') as f:
        #     response_json = json.load(f)

        content_type, content_string = contents.split(',')
        decoded = io.BytesIO(base64.b64decode(content_string))

        # Send the file to the API
        files = {'file': (filename, decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post(
            'https://faang-validator-backend-service-964531885708.europe-west2.run.app/validate-file',
            files=files,
            headers={'accept': 'application/json'}
        )
        # Handle response and parse as JSON
        if response.status_code == 200:
            response_json = response.json()
        else:
            raise Exception(f"Error {response.status_code}: {response.text}")

        print("Using validation_results.json file for validation results")

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
        dcc.Store(id='stored-validation-results', data={
            'valid_count': valid_count,
            'invalid_count': invalid_count,
            'all_sheets_data': all_sheets_validation_data
        }),

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


# Callback to show/hide error table when "Invalid organisms" button is clicked
@app.callback(
    [Output('error-table-container', 'children'),
     Output('error-table-container', 'style'),
     Output('sheet-tabs-container', 'style'),
     Output('original-file-heading', 'style')],
    [Input('issues-validation-button', 'n_clicks')],
    [State('error-table-container', 'style'),
     State('stored-error-data', 'data'),
     State('sheet-tabs-container', 'style'),
     State('original-file-heading', 'style')]
)
def toggle_error_table(n_clicks, current_style, error_data, sheet_tabs_style, heading_style):
    if n_clicks is None or n_clicks == 0 or not error_data:
        return [], {'display': 'none'}, sheet_tabs_style, heading_style

    is_visible = current_style and current_style.get('display') == 'block'

    if is_visible:
        return [], {'display': 'none'}, sheet_tabs_style, heading_style
    else:
        error_table = [
            html.H3("2. Conversion and Validation results"),
            dash_table.DataTable(
                id='error-table',
                data=error_data,
                columns=[
                    {'name': 'Sheet', 'id': 'Sheet'},
                    {'name': 'Sample Name', 'id': 'Sample Name'},
                    {'name': 'Column Name', 'id': 'Column Name'},
                    {'name': 'Error', 'id': 'Error'}
                ],
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left'},
                style_header={
                    'backgroundColor': 'rgb(230, 230, 230)',
                    'fontWeight': 'bold'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': 'rgb(248, 248, 248)'
                    },
                    {
                        'if': {'column_id': 'Column Name'},
                        'color': '#ff0000',
                        'fontWeight': 'bold',
                        'cursor': 'pointer',
                        'textDecoration': 'underline'
                    }
                ],
                tooltip_data=[
                    {
                        'Column Name': {'value': 'Click to see error details', 'type': 'markdown'}
                    } for row in error_data
                ],
                tooltip_duration=None,
                cell_selectable=True
            )
        ]

        updated_sheet_tabs_style = {'margin': '20px 0'}
        updated_heading_style = {'display': 'none'}

        return error_table, {'display': 'block'}, updated_sheet_tabs_style, {}


# Callback to show error popup when a cell in the "Column Name" column is clicked
@app.callback(
    [Output('error-popup-container', 'style'),
     Output('error-popup-title', 'children'),
     Output('error-popup-content', 'children')],
    [Input('error-table', 'active_cell')],
    [State('error-table', 'data')]
)
def show_error_popup(active_cell, data):
    if active_cell is None or active_cell['column_id'] != 'Column Name':
        return {'display': 'none'}, 'Error Details', []

    row_idx = active_cell['row']
    column_name = data[row_idx]['Column Name']
    error_message = 'ERROR : ' + data[row_idx]['Error']

    error_parts = error_message.split('; ')
    error_elements = [html.P(error, style={'color': '#ff0000'}) for error in error_parts]

    return {'display': 'block'}, f"Error in column: {column_name}", [
        html.P(f"Sample: {data[row_idx]['Sample Name']}"),
        html.P(f"Sheet: {data[row_idx]['Sheet']}"),
        html.P("Error details:"),
        html.Div(
            error_elements,
            style={'marginLeft': '20px'}
        )
    ]


# Callback to close error popup when close button or overlay is clicked
@app.callback(
    Output('error-popup-container', 'style', allow_duplicate=True),
    [Input('error-popup-close', 'n_clicks'),
     Input('error-popup-overlay', 'n_clicks')],
    prevent_initial_call=True
)
def close_error_popup(close_clicks, overlay_clicks):
    return {'display': 'none'}


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

    for sample_type in sample_types:
        st_data = results_by_type.get(sample_type, {}) or {}
        st_key = sample_type.replace(' ', '_')

        invalid_key = f"invalid_{st_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
        valid_key = f"valid_{st_key}s"

        invalid_rows_full = _flatten_data_rows(st_data.get(invalid_key), include_errors=True) or []
        rows_for_df_err = []
        for r in invalid_rows_full:
            rc = r.copy()
            rc.pop('errors', None)
            rc.pop('warnings', None)
            rows_for_df_err.append(rc)
        df_err = _df(rows_for_df_err)

        valid_rows_full = _flatten_data_rows(st_data.get(valid_key)) or []
        warning_rows_full = [r for r in valid_rows_full if r.get('warnings')]
        rows_for_df_warn = []
        for r in warning_rows_full:
            rc = r.copy()
            rc.pop('warnings', None)
            rows_for_df_warn.append(rc)
        df_warn = _df(rows_for_df_warn)

        if not df_err.empty:
            excel_sheets[f"{st_key[:25]}_errors"] = {"df": df_err, "rows": invalid_rows_full, "mode": "error"}
        if not df_warn.empty:
            excel_sheets[f"{st_key[:24]}_warnings"] = {"df": df_warn, "rows": warning_rows_full, "mode": "warn"}

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
            ws.data_validation(r, c, r, c, {
                "validate": "any",
                "input_title": title[:32],
                "input_message": full_text[:3000],
                "show_input": True
            })

        for sheet_name, payload in excel_sheets.items():
            ws = writer.sheets[sheet_name[:31]]
            df = payload["df"]
            cols = list(df.columns)
            mode = payload["mode"]
            rows_full = payload["rows"]

            for i, raw in enumerate(rows_full, start=1):
                if mode == "error":
                    row_err = raw.get("errors") or {}
                    if isinstance(row_err, dict) and "field_errors" in row_err:
                        row_err = row_err["field_errors"]

                    for field, msgs in (row_err or {}).items():
                        col_name = _resolve_col(field, cols)
                        if not col_name:
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
                        if not col_name:
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


# Callback to populate validation results tabs
@app.callback(
    Output('validation-results-container', 'children'),
    [Input('stored-json-validation-results', 'data')]
)
def populate_validation_results_tabs(validation_results):
    if validation_results is None or 'results' not in validation_results:
        return []

    validation_data = validation_results['results']
    sample_types = validation_data.get('sample_types_processed', [])
    if not sample_types:
        return []

    sample_type_tabs = []
    for sample_type in sample_types:
        results_by_type = validation_data.get('results_by_type', {})
        st_data = results_by_type.get(sample_type, {})
        invalid_key = f"invalid_{sample_type.replace(' ', '_')}s"
        if invalid_key.endswith('ss'):  # fix for pool of specimens
            invalid_key = invalid_key[:-1]

        invalid_records = st_data.get(invalid_key, [])

        if invalid_records:
            sample_type_tabs.append(
                dcc.Tab(
                    label=sample_type.capitalize(),
                    value=sample_type,
                    style={'border': 'none'},
                    selected_style={'border': 'none', 'borderBottom': '2px solid blue'},
                    children=[
                        html.Div(id={'type': 'sample-type-content', 'index': sample_type})
                    ]
                )
            )

    if not sample_type_tabs:
        return html.Div("No invalid records found.")

    tabs = dcc.Tabs(
        id='sample-type-tabs',
        value=sample_type_tabs[0].value if sample_type_tabs else None,
        children=sample_type_tabs,
        style={'border': 'none'},
        colors={"border": "transparent", "primary": "#4CAF50", "background": "#f5f5f5"}
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

    return html.Div([header_bar, tabs], style={"marginTop": "8px"})


# Callback to populate sample type content when tab is selected
@app.callback(
    Output({'type': 'sample-type-content', 'index': MATCH}, 'children'),
    [Input('sample-type-tabs', 'value')],
    [State('stored-json-validation-results', 'data')]
)
def populate_sample_type_content(selected_sample_type, validation_results):
    if validation_results is None or selected_sample_type is None:
        return []

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {})

    if selected_sample_type not in results_by_type:
        return html.Div("No data available for this sample type.")

    return make_sample_type_panel(selected_sample_type, results_by_type)


def create_samples_table(samples, is_valid=True):
    if not samples:
        return html.Div("No samples available.")

    table_data = []
    for sample in samples:
        sample_data = sample.get('data', {})
        row = {
            'Index': sample.get('index', ''),
            'Sample Name': sample.get('sample_name', '')
        }

        for key, value in sample_data.items():
            if not isinstance(value, (str, int, float, bool)) and value is not None:
                continue
            row[key] = value

        if is_valid:
            warnings = sample.get('warnings', [])
            row['Warnings'] = ', '.join(warnings) if warnings else 'None'
        else:
            errors = sample.get('errors', {}).get('errors', [])
            row['Errors'] = ', '.join(errors) if errors else 'None'

        table_data.append(row)

    all_columns = set()
    for row in table_data:
        all_columns.update(row.keys())

    ordered_columns = ['Index', 'Sample Name']
    for col in sorted(all_columns):
        if col not in ordered_columns:
            ordered_columns.append(col)

    table = dash_table.DataTable(
        id='samples-table',
        data=table_data,
        columns=[{'name': col, 'id': col} for col in ordered_columns if any(col in row for row in table_data)],
        page_size=10,
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left'},
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': 'rgb(248, 248, 248)'
            }
        ]
    )

    return html.Div([
        html.H4(f"{'Valid' if is_valid else 'Invalid'} Samples"),
        table
    ])


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
                    style={
                        'padding': '10px 20px',
                        'borderRadius': '4px 4px 0 0',
                        'border': 'none',
                    },
                    selected_style={
                        'backgroundColor': '#4CAF50',
                        'color': 'white',
                        'padding': '10px 20px',
                        'borderRadius': '4px 4px 0 0',
                        'fontWeight': 'bold',
                        'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
                        'border': 'none',
                        'borderBottom': '2px solid blue',
                    }
                ) for i, sheet_name in enumerate(filtered_sheet_names)
            ],
            style={
                'width': '100%',
                'marginBottom': '20px',
                'border': 'none',
            },
            colors={
                "border": "transparent",
                "primary": "#4CAF50",
                "background": "#f5f5f5"
            }
        )
    ], style={'marginTop': '30px', 'borderTop': '1px solid #ddd', 'paddingTop': '20px'})

    return tabs


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


def make_sample_type_panel(sample_type: str, results_by_type: dict):
    import uuid
    panel_id = str(uuid.uuid4())

    st_data = results_by_type.get(sample_type, {}) or {}
    st_key = sample_type.replace(' ', '_')
    valid_key = f"valid_{st_key}s"
    invalid_key = f"invalid_{st_key}s"
    if invalid_key.endswith('ss'):
        invalid_key = invalid_key[:-1]

    invalid_rows = _flatten_data_rows(st_data.get(invalid_key), include_errors=True)
    valid_rows = _flatten_data_rows(st_data.get(valid_key))

    rows_for_df_err = []
    for row in invalid_rows:
        rc = row.copy()
        rc.pop('errors', None)
        rc.pop('warnings', None)
        rows_for_df_err.append(rc)
    df_err = _df(rows_for_df_err)

    cell_styles_err = []
    tooltip_err = []
    cols_with_real_errors = set()

    def _as_list(msgs):
        if isinstance(msgs, list):
            return [str(m) for m in msgs]
        return [str(msgs)]

    for i, row in enumerate(invalid_rows):
        tips = {}
        row_err = row.get("errors") or {}
        if isinstance(row_err, dict) and "field_errors" in row_err:
            row_err = row_err["field_errors"]

        for field, msgs in (row_err or {}).items():
            if df_err.empty:
                continue
            col = _resolve_col(field, df_err.columns)
            if not col:
                continue

            msgs_list = _as_list(msgs)
            is_extra = any("extra inputs are not permitted" in m.lower() for m in msgs_list)

            if is_extra:
                cell_styles_err.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#fff4cc'})
                tips[col] = {'value': f"**Warning**: {field} — " + " | ".join(msgs_list), 'type': 'markdown'}
            else:
                cell_styles_err.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#ffcccc'})
                tips[col] = {'value': f"**Error**: {field} — " + " | ".join(msgs_list), 'type': 'markdown'}
                cols_with_real_errors.add(col)

        tooltip_err.append(tips)

    tint_whole_columns = [
        {'if': {'column_id': c}, 'backgroundColor': '#ffd6d6'}
        for c in sorted(cols_with_real_errors)
    ]

    base_cell = {"textAlign": "left", "padding": "6px", "minWidth": 120, "whiteSpace": "normal", "height": "auto"}
    zebra = [{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]

    blocks = [
        html.H4("Records With Error", style={'textAlign': 'center', 'margin': '10px 0'}),
        html.Div([
            DataTable(
                id={"type": "result-table-error", "sample_type": sample_type, "panel_id": panel_id},
                data=df_err.to_dict("records"),
                columns=[{"name": c, "id": c} for c in (df_err.columns if not df_err.empty else [])],
                page_size=10,
                style_table={"overflowX": "auto"},
                style_cell=base_cell,
                style_header={"fontWeight": "bold"},
                style_data_conditional=zebra + cell_styles_err + tint_whole_columns,
                tooltip_data=tooltip_err,
                tooltip_duration=None
            )
        ], id={"type": "table-container-error", "sample_type": sample_type, "panel_id": panel_id},
            style={'display': 'block'}),
    ]

    warning_rows = [row for row in valid_rows if row.get('warnings')]
    if warning_rows:
        rows_for_df_warn = []
        for row in warning_rows:
            rc = row.copy()
            rc.pop('warnings', None)
            rows_for_df_warn.append(rc)
        df_warn = _df(rows_for_df_warn)

        cell_styles_warn = []
        tooltip_warn = []
        for i, row in enumerate(warning_rows):
            by_field = _warnings_by_field(row.get('warnings', []))
            tips = {}
            for field, msgs in (by_field or {}).items():
                col = _resolve_col(field, df_warn.columns)
                if not col:
                    continue
                cell_styles_warn.append({'if': {'row_index': i, 'column_id': col}, 'backgroundColor': '#fff4cc'})
                tips[col] = {'value': f"**Warning**: {field} — " + " | ".join(map(str, msgs)), 'type': 'markdown'}
            tooltip_warn.append(tips)

        blocks += [
            html.H4("Records With Warnings", style={'textAlign': 'center', 'margin': '20px 0 10px'}),
            html.Div([
                DataTable(
                    id={"type": "result-table-warning", "sample_type": sample_type, "panel_id": panel_id},
                    data=df_warn.to_dict("records"),
                    columns=[{"name": c, "id": c} for c in df_warn.columns],
                    page_size=10,
                    style_table={"overflowX": "auto"},
                    style_cell=base_cell,
                    style_header={"fontWeight": "bold"},
                    style_data_conditional=zebra + cell_styles_warn,
                    tooltip_data=tooltip_warn,
                    tooltip_duration=None
                )
            ], id={"type": "table-container-warning", "sample_type": sample_type, "panel_id": panel_id},
                style={'display': 'block'})
        ]

    return html.Div(blocks)


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
        s = v.get("results", {}).get("total_summary", {}) or {}
        return int(s.get("valid_samples", 0)), int(s.get("invalid_samples", 0))
    except Exception:
        return 0, 0


def _valid_sample_names(v, limit=20):
    names = []
    for rec in _collect_valid_records(v):
        name = rec.get("sample_name") or rec.get("Sample Name") or ""
        if name:
            names.append(str(name))
    names = [n for n in names if n]
    if len(names) > limit:
        head = ", ".join(names[:limit])
        return f"{head} (+{len(names) - limit} more)"
    return ", ".join(names)


def biosamples_form():
    return html.Div(
        [
            html.H3("3. Submit to BioSamples", style={"marginBottom": "10px"}),
            html.Div(
                id="biosamples-status-banner",
                style={
                    "backgroundColor": "#e8f0fe",
                    "border": "1px solid #c6dafc",
                    "color": "#1a73e8",
                    "padding": "10px 12px",
                    "borderRadius": "8px",
                    "marginBottom": "12px",
                    "fontWeight": 500,
                    "display": "none",
                }
            ),
            html.Label("Username", style={"fontWeight": 600}),
            dcc.Input(
                id="biosamples-username",
                type="text",
                placeholder="Username",
                style={"width": "100%", "padding": "10px", "borderRadius": "8px",
                       "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF", "margin": "6px 0 12px"}
            ),
            html.Label("Password", style={"fontWeight": 600}),
            dcc.Input(
                id="biosamples-password",
                type="password",
                placeholder="Password",
                style={"width": "100%", "padding": "10px", "borderRadius": "8px",
                       "border": "1px solid #cbd5e1", "backgroundColor": "#ECF2FF", "margin": "6px 0 12px"}
            ),
            dcc.RadioItems(
                id="biosamples-env",
                options=[
                    {"label": " Test server", "value": "test"},
                    {"label": " Production server", "value": "prod"},
                ],
                value="test",
                labelStyle={"marginRight": "18px"},
                style={"marginBottom": "12px"},
            ),
            html.Button(
                "Submit to BioSamples",
                id="biosamples-submit-btn",
                n_clicks=0,
                style={
                    "backgroundColor": "#673ab7",
                    "color": "white",
                    "padding": "10px 18px",
                    "border": "none",
                    "borderRadius": "8px",
                    "cursor": "pointer",
                    "fontSize": "16px",
                    "width": "220px",
                },
            ),
            html.Div(id="biosamples-submit-msg", style={"marginTop": "10px"}),
        ],
        id="biosamples-form",
        style={"display": "none", "marginTop": "16px"},
    )


@app.callback(
    [Output("biosamples-form", "style"),
     Output("biosamples-status-banner", "children"),
     Output("biosamples-status-banner", "style")],
    Input("stored-json-validation-results", "data")
)
def _toggle_biosamples_form(v):
    base_style = {"display": "block", "marginTop": "16px"}

    if not v or "results" not in v:
        return (
            {"display": "none"},
            "",
            {"display": "none"}
        )

    valid_cnt, invalid_cnt = _valid_invalid_counts(v)
    style_ok = {"display": "block", "backgroundColor": "#e6f4ea", "border": "1px solid #b7e1c5",
                "color": "#137333", "padding": "10px 12px", "borderRadius": "8px",
                "marginBottom": "12px", "fontWeight": 500}
    style_warn = {"display": "block", "backgroundColor": "#fff7e6", "border": "1px solid #ffd699",
                  "color": "#8a6d3b", "padding": "10px 12px", "borderRadius": "8px",
                  "marginBottom": "12px", "fontWeight": 500}

    if valid_cnt > 0:
        names_str = _valid_sample_names(v, limit=20)
        msg_children = [
            html.Span(f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid sample(s)."),
            html.Br(),
            html.Span("Submission will include: "),
            html.Span(names_str, style={"fontWeight": 600})
        ]
        return base_style, msg_children, style_ok
    else:
        return (
            base_style,
            f"Validation result: {valid_cnt} valid / {invalid_cnt} invalid sample(s). No valid samples to submit.",
            style_warn
        )


@app.callback(
    Output("biosamples-submit-btn", "disabled"),
    [Input("biosamples-username", "value"),
     Input("biosamples-password", "value"),
     Input("stored-json-validation-results", "data")]
)
def _disable_submit(u, p, v):
    if not u or not p or not v or "results" not in v:
        return True
    valid_cnt, _ = _valid_invalid_counts(v)
    return valid_cnt == 0


@app.callback(
    Output("biosamples-form-mount", "children", allow_duplicate=True),
    Input("upload-data", "contents"),
    prevent_initial_call=True,
)
def _clear_biosamples_form_on_new_upload(_):
    return []


@app.callback(
    Output("biosamples-submit-msg", "children"),
    Input("biosamples-submit-btn", "n_clicks"),
    [State("biosamples-username", "value"),
     State("biosamples-password", "value"),
     State("biosamples-env", "value"),
     State("stored-json-validation-results", "data")],
    prevent_initial_call=True
)
def _submit_to_biosamples(n, username, password, env, validation_results):
    if not n:
        raise PreventUpdate

    try:
        valid_records = _collect_valid_records(validation_results)
        if not valid_records:
            return html.Span("No valid samples to submit.", style={"color": "#c62828", "fontWeight": 500})

        payload = {
            "credentials": {"username": username, "password": password},
            "environment": env,
            "mode": "valid_only",
            "valid_samples": valid_records,
            "summary": dict(zip(("valid", "invalid"), _valid_invalid_counts(validation_results))),
            "validation_results": validation_results
        }
        url = f"{BACKEND_API_URL}/biosamples/submit"
        r = requests.post(url, json=payload, timeout=60)
        if r.status_code == 200:
            msg = r.json() if r.headers.get("content-type", "").startswith("application/json") else {"detail": r.text}
            return html.Span(f"Submitted successfully: {msg}", style={"color": "#388e3c", "fontWeight": 500})
        else:
            return html.Span(f"Submission failed [{r.status_code}]: {r.text}",
                             style={"color": "#c62828", "fontWeight": 500})
    except Exception as e:
        return html.Span(f"Submission error: {e}", style={"color": "#c62828", "fontWeight": 500})


app.clientside_callback(
    """
    function(n_clicks) {
        if (n_clicks > 0) {
            window.location.reload();
        }
        return '';
    }
    """,
    Output('dummy-output-for-reset', 'children'),
    [Input('reset-button', 'n_clicks')],
    prevent_initial_call=True
)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
