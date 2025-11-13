import json
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
                    html.Div(id="table-container"),
                    dcc.Loading(
                        id="loading-validation",
                        type="circle",
                        children=html.Div(id='output-data-upload')
                    )
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
    Output("table-container", "children"),
    Input("upload-excel", "contents"),
    State("upload-excel", "filename"),
    prevent_initial_call=True,
)

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
        return (
            None,  # stored-file-data
            None,  # stored-filename
            "No file chosen",
            [],
            {'display': 'none'},
            [],
            None,
            None,
            None,
        )

    try:
        content_type, content_string = contents.split(',')

        # Decode Excel and read all sheets
        decoded = io.BytesIO(base64.b64decode(content_string))
        xls = pd.ExcelFile(decoded)
        sheet_names = xls.sheet_names

        all_sheets_data = {}
        for sheet in sheet_names:
            df_sheet = pd.read_excel(xls, sheet_name=sheet)
            # store as list-of-dicts (JSON serializable)
            all_sheets_data[sheet] = df_sheet.to_dict("records")

        active_sheet = sheet_names[0] if sheet_names else None

        file_selected_display = html.Div([
            html.H3("File Selected", id='original-file-heading'),
            html.P(f"File: {filename}", style={'fontWeight': 'bold'}),
            html.P("Click 'Validate' to process the file and see results."),
        ])

        output_data_upload_children = html.Div(
            id='sheet-tabs-container',
            style={'margin': '20px 0', 'display': 'none'}
        )

        return (
            contents,                                # stored-file-data
            filename,                                # stored-filename
            filename,                                # file-chosen-text
            file_selected_display,                   # selected-file-display
            {'display': 'block', 'margin': '20px 0'},
            [output_data_upload_children],           # output-data-upload children
            all_sheets_data,                         # stored-all-sheets-data
            sheet_names,                             # stored-sheet-names
            active_sheet,                            # active-sheet
        )

    except Exception as e:
        error_display = html.Div([
            html.H5(filename),
            html.P(f"Error processing file: {str(e)}", style={'color': 'red'})
        ])
        return (
            contents,
            filename,
            filename,
            error_display,
            {'display': 'block', 'margin': '20px 0'},
            [],
            None,
            None,
            None,
        )



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
        df = pd.read_excel(decoded)
        print(json.dumps(df.columns.tolist()))
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
# Popup for cells clicked in the per-sample tables:
# - result-table-error
# - result-table-warning
@app.callback(
    [
        Output('error-popup-container', 'style'),
        Output('error-popup-title', 'children'),
        Output('error-popup-content', 'children'),
    ],
    [
        Input({'type': 'result-table-error', 'sample_type': ALL, 'panel_id': ALL}, 'active_cell'),
        Input({'type': 'result-table-warning', 'sample_type': ALL, 'panel_id': ALL}, 'active_cell'),
    ],
    [
        State({'type': 'result-table-error', 'sample_type': ALL, 'panel_id': ALL}, 'data'),
        State({'type': 'result-table-error', 'sample_type': ALL, 'panel_id': ALL}, 'id'),
        State({'type': 'result-table-warning', 'sample_type': ALL, 'panel_id': ALL}, 'data'),
        State({'type': 'result-table-warning', 'sample_type': ALL, 'panel_id': ALL}, 'id'),
        State('stored-json-validation-results', 'data'),
    ]
)
def show_error_popup(
    error_tables_active_cells,
    warning_tables_active_cells,
    error_tables_data,
    error_tables_ids,
    warning_tables_data,
    warning_tables_ids,
    validation_results,
):
    hidden = ({'display': 'none'}, 'Error Details', [])

    ctx = dash.callback_context
    if not ctx.triggered:
        return hidden

    # Which input fired?
    trigger_id = ctx.triggered[0]['prop_id'].rsplit('.', 1)[0]

    if not validation_results or 'results' not in validation_results:
        return hidden

    results = validation_results['results']
    results_by_type = results.get('results_by_type', {})

    def _st_data(sample_type):
        return results_by_type.get(sample_type, {}) or {}

    # Try to parse pattern-matching id
    try:
        trigger_obj = json.loads(trigger_id)
    except Exception:
        trigger_obj = None

    # ------------------------------------------------------------------
    # 1) Click from an ERROR table (result-table-error)
    # ------------------------------------------------------------------
    if isinstance(trigger_obj, dict) and trigger_obj.get('type') == 'result-table-error':
        # find which specific table index
        idx = None
        for i, tid in enumerate(error_tables_ids or []):
            if tid == trigger_obj:
                idx = i
                break
        if idx is None:
            return hidden

        active_cell = (error_tables_active_cells or [None])[idx]
        data = (error_tables_data or [None])[idx]
        table_id = (error_tables_ids or [None])[idx]

        if active_cell is None or data is None:
            return hidden

        row_idx = active_cell.get('row')
        col_id = active_cell.get('column_id')
        if row_idx is None or row_idx >= len(data) or col_id is None:
            return hidden

        row = data[row_idx]
        sample_name = row.get('Sample Name', f"row {row_idx + 1}")
        sample_type = table_id.get('sample_type', 'Unknown')

        st_data = _st_data(sample_type)
        st_key = sample_type.replace(' ', '_')
        invalid_key = f"invalid_{st_key}s"
        if invalid_key.endswith('ss'):
            invalid_key = invalid_key[:-1]
        invalid_records = st_data.get(invalid_key, []) or []

        # find JSON record for this sample
        rec = next((r for r in invalid_records if r.get('sample_name') == sample_name), None)
        if not rec:
            return hidden

        err_obj = rec.get('errors', {}) or {}
        field_errors = err_obj.get('field_errors', {}) or {}
        relationship_errors = err_obj.get('relationship_errors', []) or []

        msgs = []

        # field-specific errors for clicked column (supports Unit / Term Source ID mapping)
        field_key, col_errs_list = _find_msgs_for_col(field_errors, col_id)

        if col_errs_list:
            only_extra = all("extra inputs are not permitted" in str(m).lower()
                             for m in col_errs_list)
            prefix = "Warning" if only_extra else "Error"
            label_field = field_key or col_id
            msgs.append(html.P(f"{prefix}s for field '{label_field}':"))
            msgs.append(html.Ul([html.Li(str(m)) for m in col_errs_list]))

        # relationship errors associated with that sample — show when Sample Name clicked
        if relationship_errors and col_id == "Sample Name":
            msgs.append(html.P("Relationship errors:"))
            msgs.append(html.Ul([html.Li(str(m)) for m in relationship_errors]))

        if not msgs:
            return hidden

        content = [
            html.P(f"Sample: {sample_name}"),
            html.P(f"Sample type / sheet: {sample_type}"),
            html.Hr(),
        ] + msgs

        title = f"Issues for {sample_name} – {col_id}"
        return {'display': 'block'}, title, content

    # ------------------------------------------------------------------
    # 2) Click from a WARNING table (result-table-warning)
    # ------------------------------------------------------------------
    if isinstance(trigger_obj, dict) and trigger_obj.get('type') == 'result-table-warning':
        idx = None
        for i, tid in enumerate(warning_tables_ids or []):
            if tid == trigger_obj:
                idx = i
                break
        if idx is None:
            return hidden

        active_cell = (warning_tables_active_cells or [None])[idx]
        data = (warning_tables_data or [None])[idx]
        table_id = (warning_tables_ids or [None])[idx]

        if active_cell is None or data is None:
            return hidden

        row_idx = active_cell.get('row')
        col_id = active_cell.get('column_id')
        if row_idx is None or row_idx >= len(data) or col_id is None:
            return hidden

        row = data[row_idx]
        sample_name = row.get('Sample Name', f"row {row_idx + 1}")
        sample_type = table_id.get('sample_type', 'Unknown')

        st_data = _st_data(sample_type)
        st_key = sample_type.replace(' ', '_')
        valid_key = f"valid_{st_key}s"
        valid_records = st_data.get(valid_key, []) or []

        rec = next((r for r in valid_records if r.get('sample_name') == sample_name), None)
        if not rec:
            return hidden

        warnings_list = rec.get('warnings', []) or []
        by_field = _warnings_by_field(warnings_list)

        field_msgs = []
        for field, msgs in (by_field or {}).items():
            if field == col_id or _normalize_header(field) == _normalize_header(col_id):
                if isinstance(msgs, list):
                    field_msgs.extend(map(str, msgs))
                else:
                    field_msgs.append(str(msgs))

        # if still nothing, but there are some generic warnings, show them all
        if not field_msgs and warnings_list:
            field_msgs = [str(w) for w in warnings_list]

        if not field_msgs:
            return hidden

        content = [
            html.P(f"Sample: {sample_name}"),
            html.P(f"Sample type / sheet: {sample_type}"),
            html.Hr(),
            html.P(f"Warnings for field '{col_id}':"),
            html.Ul([html.Li(str(m)) for m in field_msgs]),
        ]

        title = f"Warnings for {sample_name} – {col_id}"
        return {'display': 'block'}, title, content

    # fallback
    return hidden

def build_invalid_table_from_json(invalid_records):
    """
    invalid_records: list from e.g. results_by_type["organism"]["invalid_organisms"]
    Returns:
      table_data: list[dict]  -> rows for DataTable
      tooltip_data: list[dict] -> per-row tooltips
      style_data_conditional: list[dict] -> colouring per cell
    """
    table_data = []
    tooltip_data = []
    style_data_conditional = []

    for row_idx, rec in enumerate(invalid_records):
        row_dict = {}

        # 1) basic flat data
        data = rec.get("data", {}) or {}
        row_dict["Sample Name"] = rec.get("sample_name", "")
        for key, value in data.items():
            # Health Status is a list of dicts like {"text":..., "term":...}
            if key == "Health Status" and isinstance(value, list):
                texts = []
                terms = []
                for item in value:
                    if isinstance(item, dict):
                        if item.get("text"):
                            texts.append(item["text"])
                        if item.get("term"):
                            terms.append(item["term"])
                if texts:
                    row_dict["Health Status"] = ", ".join(texts)
                if terms:
                    row_dict["Health Status Term Source ID"] = ", ".join(terms)
            else:
                # Child Of, Derived From may be lists of strings
                if isinstance(value, list):
                    row_dict[key] = ", ".join(map(str, value))
                else:
                    row_dict[key] = value

        # 2) error tooltips for each field with errors
        errs = rec.get("errors") or {}
        field_errors = errs.get("field_errors") or {}
        relationship_errors = errs.get("relationship_errors") or []

        tips_for_row = {}

        # field_errors → colour + tooltip per column
        cols_in_row = list(row_dict.keys())

        for field, msgs in field_errors.items():
            # map backend field -> real sheet column header
            col_id = _match_field_to_col(field, cols_in_row)
            if not col_id:
                # nothing in this row matches this field
                continue

            msg_list = [str(m) for m in (msgs if isinstance(msgs, list) else [msgs])]

            only_extra = all("extra inputs are not permitted" in m.lower()
                             for m in msg_list)

            style_data_conditional.append({
                "if": {"row_index": row_idx, "column_id": col_id},
                "backgroundColor": "#fff4cc" if only_extra else "#ffcccc",
            })

            tips_for_row[col_id] = {
                "value": ("**Warning**: " if only_extra else "**Error**: ")
                         + field + " — "
                         + " | ".join(msg_list),
                "type": "markdown",
            }

        # relationship errors → attach on Sample Name
        if relationship_errors:
            text = "**Relationship errors:** " + " | ".join(
                map(str, relationship_errors)
            )
            # if Sample Name already has tooltip, append
            existing = tips_for_row.get("Sample Name")
            if existing:
                existing["value"] += "\n\n" + text
            else:
                tips_for_row["Sample Name"] = {
                    "value": text,
                    "type": "markdown",
                }

            # highlight Sample Name cell if relationships fail
            style_data_conditional.append({
                "if": {"row_index": row_idx, "column_id": "Sample Name"},
                "backgroundColor": "#ffcccc",
            })

        tooltip_data.append(tips_for_row)
        table_data.append(row_dict)

    return table_data, tooltip_data, style_data_conditional
#
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


# Callback to show error/warning popup from:
# - main error table ("error-table")
# - per-sample tables ("result-table-error" and "result-table-warning")
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
    if validation_results is None:
        return []

    if 'results' not in validation_results:
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

    return html.Div([header_bar, tabs])


# Callback to populate sample type content when tab is selected
@app.callback(
    Output({'type': 'sample-type-content', 'index': MATCH}, 'children'),
    [Input('sample-type-tabs', 'value')],
    [
        State('stored-json-validation-results', 'data'),
        State('stored-all-sheets-data', 'data'),
        State('stored-sheet-names', 'data'),
    ]
)
def populate_sample_type_content(selected_sample_type, validation_results, all_sheets_data, sheet_names):
    if validation_results is None or selected_sample_type is None:
        return []

    validation_data = validation_results['results']
    results_by_type = validation_data.get('results_by_type', {})

    if selected_sample_type not in results_by_type:
        return html.Div("No data available for this sample type.")

    return make_sample_type_panel(
        selected_sample_type,
        results_by_type,
        all_sheets_data or {},
        sheet_names or [],
    )


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


def make_sample_type_panel(sample_type: str,
                           results_by_type: dict,
                           all_sheets_data=None,
                           sheet_names=None):
    import uuid
    panel_id = str(uuid.uuid4())

    st_data = results_by_type.get(sample_type, {}) or {}
    st_key = sample_type.replace(' ', '_')
    valid_key = f"valid_{st_key}s"
    invalid_key = f"invalid_{st_key}s"
    if invalid_key.endswith('ss'):  # fix for "pool of specimens"
        invalid_key = invalid_key[:-1]

    # --- Raw JSON records from backend ---
    invalid_records = st_data.get(invalid_key, []) or []
    valid_records = st_data.get(valid_key, []) or []

    # --- Build error table (with colours + tooltips) from JSON ---
    table_data, tooltip_err, style_data_err = build_invalid_table_from_json(invalid_records)

    base_cell = {
        "textAlign": "left",
        "padding": "6px",
        "minWidth": 120,
        "whiteSpace": "normal",
        "height": "auto",
    }
    zebra = [
        {"if": {"row_index": "odd"}, "backgroundColor": "rgb(248, 248, 248)"}
    ]

    if table_data:
        # Determine columns from keys present in the data
        all_cols = []
        seen = set()
        for row in table_data:
            for k in row.keys():
                if k not in seen:
                    seen.add(k)
                    all_cols.append(k)

        columns = [{"name": _normalize_header(c), "id": c} for c in all_cols]
    else:
        columns = []

    blocks = [
        html.H4("Records With Error", style={'textAlign': 'center', 'margin': '10px 0'}),
        html.Div(
            [
                DataTable(
                    id={"type": "result-table-error", "sample_type": sample_type, "panel_id": panel_id},
                    data=table_data,
                    columns=columns,
                    page_size=10,
                    style_table={"overflowX": "auto"},
                    style_cell=base_cell,
                    style_header={"fontWeight": "bold"},
                    style_data_conditional=zebra + style_data_err,
                    tooltip_data=tooltip_err,
                    tooltip_duration=None,
                )
            ],
            id={"type": "table-container-error", "sample_type": sample_type, "panel_id": panel_id},
            style={'display': 'block'},
        ),
    ]

    # --- Warnings section (keep existing logic, but based on valid_records) ---
    warning_rows = [row for row in _flatten_data_rows(valid_records) if row.get('warnings')]
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
                cell_styles_warn.append({
                    'if': {'row_index': i, 'column_id': col},
                    'backgroundColor': '#fff4cc'
                })
                tips[col] = {
                    'value': f"**Warning**: {field} — " + " | ".join(map(str, msgs)),
                    'type': 'markdown'
                }
            tooltip_warn.append(tips)

        blocks += [
            html.H4("Records With Warnings", style={'textAlign': 'center', 'margin': '20px 0 10px'}),
            html.Div(
                [
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
                        tooltip_duration=None,
                    )
                ],
                id={"type": "table-container-warning", "sample_type": sample_type, "panel_id": panel_id},
                style={'display': 'block'},
            ),
        ]

    return html.Div(blocks)



from collections import defaultdict
def _flatten_data_rows(rows, include_errors=False):
    flat = []

    for r in rows or []:
        base = {"Sample Name": r.get("sample_name")}
        data_fields = r.get("data", {}) or {}

        # ✅ list of objects, allows duplicates
        processed_fields = []

        for key, value in data_fields.items():
            display_title = _normalize_header(key)  # default title same as key

            # ---- Special handling for Health Status ----
            if key == "Health Status" and isinstance(value, list) and value:
                flattened = []
                for item in value:
                    if isinstance(item, list):
                        flattened.extend(item)
                    else:
                        flattened.append(item)

                for status in flattened:
                    if isinstance(status, dict):
                        text = status.get("text")
                        term = status.get("term")
                        if text:
                            processed_fields.append({
                                "key": key,
                                "title": "Health Status",
                                "value": text
                            })
                        if term:
                            processed_fields.append({
                                "key": key,
                                "title": "Term Source ID",
                                "value": term
                            })

            # ---- Handle Term Source ID fields ----
            elif "Term Source ID" in key:
                processed_fields.append({
                    "key": key,
                    "title": key,
                    "value": value
                })

            # ---- Handle Unit fields ----
            elif "Unit" in key:
                processed_fields.append({
                    "key": key,
                    "title": key,
                    "value": value
                })

            # ---- Handle Child Of fields ----
            elif key == "Child Of" and isinstance(value, list):
                processed_fields.append({
                    "key": key,
                    "title": "Child Of",
                    "value": ", ".join(str(item) for item in value if item)
                })

            # ---- Handle complex objects ----
            elif not isinstance(value, (str, int, float, bool, type(None))):
                processed_fields.append({
                    "key": key,
                    "title": key,
                    "value": str(value) if value else ""
                })

            # ---- Default case ----
            else:
                processed_fields.append({
                    "key": key,
                    "title": display_title,
                    "value": value
                })

        # attach processed fields list
        base["fields"] = processed_fields

        # include optional errors/warnings
        if include_errors:
            errors = r.get("errors", {})
            if isinstance(errors, dict) and "field_errors" in errors:
                base["errors"] = errors["field_errors"]

        warnings = r.get("warnings", [])
        if warnings:
            base["warnings"] = warnings

        flat.append(base)
    return flat




def _df(records):
    df = pd.DataFrame(records)
    if df.empty:
        return df
    lead = [c for c in ["Sample Name"] if c in df.columns]
    other = [c for c in df.columns if c not in lead]
    return df[lead + other]


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


def _normalize_header(header: str) -> str:
    """
    Normalize column / field names into a generic form.
    Examples:
      'Birth Location Latitude Unit' -> 'unit'
      'Body Weight Term Source ID'   -> 'term source id'
      'Unit'                         -> 'unit'
      'Term Source ID'               -> 'term source id'
    """
    if not isinstance(header, str):
        return header

    if "Term Source ID" in header:
        return "Term Source ID"
    if "Unit" in header:
        return "Unit"

    return header


def _match_field_to_col(field: str, available_cols) -> str | None:
    """
    Map a backend field name (e.g. 'Unit', 'Term Source ID') to
    the real sheet header (e.g. 'Birth Location Latitude Unit').
    """
    if not available_cols:
        return None

    field_raw = (field or "").strip().lower()
    field_norm = _normalize_header(field)

    # 1) exact match
    for c in available_cols:
        if str(c).strip().lower() == field_raw:
            return c

    # 2) normalized match (Unit, Term Source ID etc.)
    for c in available_cols:
        if _normalize_header(str(c)) == field_norm:
            return c

    # nothing found
    return None


def _find_msgs_for_col(field_errors: dict, col_id: str):
    """
    Given backend field_errors and a clicked column id (sheet header),
    find the matching field key and its messages.
    """
    if not field_errors:
        return None, []

    # 1) direct key match
    if col_id in field_errors:
        msgs = field_errors[col_id]
        return col_id, (msgs if isinstance(msgs, list) else [msgs])

    # 2) normalized match
    col_norm = _normalize_header(col_id)
    for f, msgs in field_errors.items():
        if _normalize_header(f) == col_norm:
            return f, (msgs if isinstance(msgs, list) else [msgs])

    return None, []


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
