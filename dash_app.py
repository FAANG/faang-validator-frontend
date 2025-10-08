import json
import sys
import os
import base64
import requests
import io

import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State, MATCH, ALL
import pandas as pd
from dash.exceptions import PreventUpdate
import uuid

# Backend API URL - can be configured via environment variable
BACKEND_API_URL = os.environ.get('BACKEND_API_URL', 'https://faang-validator-backend-service-964531885708.europe-west2.run.app/api')

# Initialize the Dash app
# Set suppress_callback_exceptions=True to allow callbacks for components that are created dynamically
# (like the issues-validation-button which is created after validation)
app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server  # Expose server variable for gunicorn

# --- App Layout ---
app.layout = html.Div([
    html.Div([
        html.H1("FAANG Validation"),
        # Store for uploaded file data
        dcc.Store(id='stored-file-data'),
        dcc.Store(id='stored-filename'),
        dcc.Store(id='stored-all-sheets-data'),
        dcc.Store(id='stored-sheet-names'),
        dcc.Store(id='error-popup-data', data={'visible': False, 'column': '', 'error': ''}),
        dcc.Store(id='active-sheet', data=None),

        # Error popup
        # html.Div(
        #     id='error-popup-container',
        #     style={'display': 'none'},
        #     children=[
        #         html.Div(
        #             className='error-popup-overlay',
        #             id='error-popup-overlay'
        #         ),
        #         html.Div(
        #             className='error-popup',
        #             children=[
        #                 html.Div(
        #                     className='error-popup-close',
        #                     id='error-popup-close',
        #                     children='×'
        #                 ),
        #                 html.H3(
        #                     className='error-popup-title',
        #                     id='error-popup-title',
        #                     children='Error Details'
        #                 ),
        #                 html.Div(
        #                     className='error-popup-content',
        #                     id='error-popup-content',
        #                     children=[]
        #                 )
        #             ]
        #         )
        #     ]
        # ),

        # Tabs
        dcc.Tabs([
            # Samples Tab
            dcc.Tab(label='Samples', children=[
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
                    ], style={'display': 'flex', 'alignItems': 'center'}),
                    html.Div(id='selected-file-display', style={'display': 'none'}),
                ], style={'margin': '20px 0'}),

                dcc.Loading(
                    id="loading-validation",
                    type="circle",
                    children=html.Div(id='output-data-upload')
                )
            ]),

            # Experiments Tab (empty for now)
            dcc.Tab(label='Experiments', children=[
                html.Div([], style={'margin': '20px 0'})
            ]),

            # Analysis Tab (empty for now)
            dcc.Tab(label='Analysis', children=[
                html.Div([], style={'margin': '20px 0'})
            ])
        ], style={'margin': '20px 0'})
    ], className='container')
])


# Callback to store uploaded file data and display filename
@app.callback(
    [Output('stored-file-data', 'data'),
     Output('stored-filename', 'data'),
     Output('file-chosen-text', 'children'),
     Output('output-data-upload', 'children'),
     Output('stored-all-sheets-data', 'data'),
     Output('stored-sheet-names', 'data'),
     Output('active-sheet', 'data')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def store_file_data(contents, filename):
    if contents is None:
        return None, None, "No file chosen", [], None, None, None

    # Parse the uploaded file to display its content
    try:
        # Extract the base64 content from the data URL
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)

        # Initialize variables for all sheets
        all_sheets_data = {}
        sheet_names = []

        # Parse the file based on its type
        if filename.endswith('.csv'):
            # Parse CSV - only one sheet
            df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
            sheet_name = "Sheet 1"  # Default name for CSV
            all_sheets_data = {sheet_name: df.to_dict('records')}
            sheet_names = [sheet_name]
            active_sheet = sheet_name
        elif filename.endswith(('.xls', '.xlsx')):
            # Parse Excel - multiple sheets
            excel_file = pd.ExcelFile(io.BytesIO(decoded), engine="openpyxl")
            sheet_names = excel_file.sheet_names

            for sheet_name in sheet_names:
                df = excel_file.parse(sheet_name)
                # Skip empty sheets
                if df.empty:
                    continue
                all_sheets_data[sheet_name] = df.to_dict('records')

            # Set active sheet to the third sheet (index 2) if available, otherwise use the first sheet
            start_index = 3  # Start from sheet number 3 (index 2)
            active_sheet = sheet_names[start_index] if len(sheet_names) > start_index else (sheet_names[0] if sheet_names else None)
        else:
            # Unsupported file type
            return contents, filename, filename, html.Div([
                html.H5(filename),
                html.P("Unsupported file type. Please upload a CSV or Excel file.", style={'color': 'red'})
            ]), None, None, None

        # If no valid sheets were found
        if not all_sheets_data:
            return contents, filename, filename, html.Div([
                html.H5(filename),
                html.P("No valid data found in the file.", style={'color': 'red'})
            ]), None, None, None

        # Get the active sheet's data for initial display
        # Use the third sheet (index 2) if available, otherwise use the first sheet
        start_index = 2  # Start from sheet number 3 (index 2)
        display_sheet_name = sheet_names[start_index] if len(sheet_names) > start_index else sheet_names[0]
        display_sheet_df = pd.DataFrame(all_sheets_data[display_sheet_name])

        # Display the active sheet's data but hide it initially
        file_display = html.Div([
            html.H3("Original File Data", id='original-file-heading', style={'display': 'none'}),
            dash_table.DataTable(
                id='file-data-table',
                data=display_sheet_df.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in display_sheet_df.columns],
                page_size=10,
                style_table={'overflowX': 'auto', 'display': 'none'},
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
            ),
            # Placeholder for validation results will be populated by validate_data callback
            # This will be before the sheet tabs
            # Add a placeholder for sheet tabs that will be populated later
            html.Div(id='sheet-tabs-container', style={'margin': '20px 0', 'display': 'none'})
        ], style={'margin': '20px 0'})

        # Store the file data, filename, all sheets data, and update the display
        return contents, filename, filename, file_display, all_sheets_data, sheet_names, active_sheet

    except Exception as e:
        # Handle any errors during file parsing
        return contents, filename, filename, html.Div([
            html.H5(filename),
            html.P(f"Error parsing file: {str(e)}", style={'color': 'red'})
        ]), None, None, None

# Callback to show and enable validate button when a file is uploaded
@app.callback(
    [Output('validate-button', 'disabled'),
     Output('validate-button-container', 'style')],
    [Input('stored-file-data', 'data')]
)
def show_and_enable_validate_button(file_data):
    if file_data is None:
        # No file selected: button is disabled and hidden
        return True, {'display': 'none', 'marginLeft': '10px'}
    else:
        # File selected: button is enabled and visible
        return False, {'display': 'block', 'marginLeft': '10px'}

# Callback to validate data when button is clicked
@app.callback(
    Output('output-data-upload', 'children', allow_duplicate=True),
    [Input('validate-button', 'n_clicks')],
    [State('stored-file-data', 'data'),
     State('stored-filename', 'data'),
     State('output-data-upload', 'children'),
     State('stored-all-sheets-data', 'data'),
     State('stored-sheet-names', 'data')]
    ,prevent_initial_call=True
)
def validate_data(n_clicks, contents, filename, current_children, all_sheets_data, sheet_names):
    if n_clicks is None or contents is None:
        return current_children if current_children else []

    error_data = []
    records = []
    valid_count = 0
    invalid_count = 0
    all_sheets_validation_data = {}

    try:
        # Extract the base64 content from the data URL
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)

        # Create a file-like object to send to the API
        files = {'file': (filename, decoded, content_type)}

        # Send the file to the backend API for validation
        response = requests.post(f"{BACKEND_API_URL}/validate", files=files)

        # Check if the request was successful
        if response.status_code != 200:
            error_message = response.json().get('detail', 'Unknown error occurred')
            error_div = html.Div([
                html.H5(filename),
                html.P(f"Error: {error_message}", style={'color': 'red'})
            ])
            # Return the original file display with the error message
            return current_children + [error_div] if isinstance(current_children, list) else [current_children, error_div]

        # Parse the response
        validation_data = response.json()
        records = validation_data.get('records', [])
        valid_count = validation_data.get('valid_count', 0)
        invalid_count = validation_data.get('invalid_count', 0)
        error_data = validation_data.get('errors', [])

        # Get all sheets data if available
        all_sheets_validation_data = validation_data.get('all_sheets_data', {})

        # If all_sheets_data is not in the response, use the first sheet's data
        if not all_sheets_validation_data and sheet_names:
            first_sheet = sheet_names[0]
            all_sheets_validation_data = {first_sheet: records}

    except Exception as e:
        error_div = html.Div([
            html.H5(filename),
            html.P(f"Error connecting to backend API: {str(e)}", style={'color': 'red'})
        ])
        # Return the original file display with the error message
        return current_children + [error_div] if isinstance(current_children, list) else [current_children, error_div]

    # Create a validation results display
    validation_components = [
        # Store for error data
        dcc.Store(id='stored-error-data', data=error_data),
        # Store for validation results by sheet
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

        html.Div([
            html.Button(
                f"Valid organisms: {valid_count}",
                id='passed-validation-button',
                className='summary-button success',
            ),
            html.Button(
                f"Invalid organisms: {invalid_count}",
                id='issues-validation-button',
                className='summary-button warning',
                n_clicks=0  # Initialize click counter
            ),
        ], style={'display': 'flex', 'justifyContent': 'center', 'gap': '20px', 'margin': '20px 0'}),

        # Container for error table that will be populated by callback
        html.Div(id='error-table-container', style={'display': 'none'})
    ]

    # Combine the original file display with the validation results
    # Insert validation results before the sheet tabs container
    if current_children is None:
        return validation_components
    elif isinstance(current_children, list):
        # Make a copy of current_children to modify
        modified_children = []

        # Process each child in the current_children list
        for child in current_children:
            if isinstance(child, dict) and child.get('props'):
                props = child.get('props', {})
                children = props.get('children', [])

                # Check if this is the main file_display div
                if isinstance(children, list) and any(
                    isinstance(c, dict) and c.get('props', {}).get('id') == 'file-data-table' 
                    for c in children
                ):
                    # This is the main file_display div, update its children
                    updated_child = child.copy()
                    updated_children = []

                    for c in children:
                        if isinstance(c, dict) and c.get('props'):
                            c_props = c.get('props', {})

                            # Update the heading style
                            if isinstance(c_props.get('children'), str) and c_props.get('children') == "Original File Data":
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                updated_c['props']['style'] = {}  # Remove display: none
                                updated_children.append(updated_c)

                            # Update the data table style
                            elif c_props.get('id') == 'file-data-table':
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                if 'style_table' in c_props:
                                    updated_c['props']['style_table'] = {'overflowX': 'auto'}  # Remove display: none
                                updated_children.append(updated_c)

                            # Update the sheet tabs container style
                            elif c_props.get('id') == 'sheet-tabs-container':
                                updated_c = c.copy()
                                updated_c['props'] = c_props.copy()
                                updated_c['props']['style'] = {'margin': '20px 0'}  # Remove display: none
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

        # Return validation components first, then the modified children with visible file data and tabs
        return validation_components + modified_children
    else:
        # If current_children is not a list, convert it to a list and append validation components
        return validation_components + [current_children]

# Callback to show/hide error table when "Invalid organisms" button is clicked
@app.callback(
    [Output('error-table-container', 'children'),
     Output('error-table-container', 'style'),
     Output('file-data-table', 'style_table'),
     Output('sheet-tabs-container', 'style'),
     Output('original-file-heading', 'style')],
    [Input('issues-validation-button', 'n_clicks')],
    [State('error-table-container', 'style'),
     State('stored-error-data', 'data'),
     State('file-data-table', 'style_table'),
     State('sheet-tabs-container', 'style'),
     State('original-file-heading', 'style')]
)
def toggle_error_table(n_clicks, current_style, error_data, file_table_style, sheet_tabs_style, heading_style):
    # If button hasn't been clicked or there's no error data, don't show table
    if n_clicks is None or n_clicks == 0 or not error_data:
        return [], {'display': 'none'}, file_table_style, sheet_tabs_style, heading_style

    # Toggle visibility based on current state
    is_visible = current_style and current_style.get('display') == 'block'

    if is_visible:
        # If currently visible, hide it
        return [], {'display': 'none'}, file_table_style, sheet_tabs_style, heading_style
    else:
        # If currently hidden, show it
        # Create the error table
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

        # Make file data table, sheet tabs, and heading visible
        updated_file_table_style = {'overflowX': 'auto'}  # Remove display: none
        updated_sheet_tabs_style = {'margin': '20px 0'}  # Remove display: none
        updated_heading_style = {}  # Remove display: none

        # Return the table and set display style to block
        return error_table, {'display': 'block'}, updated_file_table_style, updated_sheet_tabs_style, updated_heading_style

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
    error_message = data[row_idx]['Error']

    # Split error message by semicolons and create a list of paragraph elements
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

# Callback to create sheet tabs
@app.callback(
    Output('sheet-tabs-container', 'children'),
    [Input('stored-sheet-names', 'data')],
    [State('active-sheet', 'data'),
     State('stored-all-sheets-data', 'data')]
)
def create_sheet_tabs(sheet_names, active_sheet, all_sheets_data):
    # Use the helper function to create the tabs
    return create_sheet_tabs_ui(sheet_names, active_sheet, all_sheets_data)

# Callback to handle sheet tab selection
@app.callback(
    [Output('active-sheet', 'data', allow_duplicate=True),
     Output('file-data-table', 'data'),
     Output('file-data-table', 'columns'),
     Output('sheet-tabs-container', 'children', allow_duplicate=True)],
    [Input('sheet-tabs', 'value')],
    [State('stored-sheet-names', 'data'),
     State('stored-all-sheets-data', 'data'),
     State('active-sheet', 'data')],
    prevent_initial_call=True
)
def handle_sheet_tab_click(selected_tab_value, sheet_names, all_sheets_data, current_active_sheet):
    # If no tab is selected or the selected tab is already active, do nothing
    if selected_tab_value is None or selected_tab_value == current_active_sheet:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    # The selected_tab_value is the sheet name
    clicked_sheet = selected_tab_value

    # Get the data for the clicked sheet
    sheet_data = all_sheets_data.get(clicked_sheet, [])

    # If the sheet has no data, return empty data
    if not sheet_data:
        # Create updated tabs with the new active sheet
        updated_tabs = create_sheet_tabs_ui(sheet_names, clicked_sheet, all_sheets_data)
        return clicked_sheet, [], [], updated_tabs

    # Create columns for the data table
    columns = [{'name': key, 'id': key} for key in sheet_data[0].keys()]

    # Create updated tabs with the new active sheet
    updated_tabs = create_sheet_tabs_ui(sheet_names, clicked_sheet, all_sheets_data)

    # Return the new active sheet, data, and updated tabs
    return clicked_sheet, sheet_data, columns, updated_tabs

# Helper function to create sheet tabs UI
def create_sheet_tabs_ui(sheet_names, active_sheet, all_sheets_data=None):
    if not sheet_names or len(sheet_names) <= 1:
        # If there's only one sheet or no sheets, don't show tabs
        return []

    # Start from sheet number 3 (index 2 if zero-indexed)
    start_index = 2
    if len(sheet_names) <= start_index:
        # If there are fewer than 3 sheets, don't show tabs
        return []

    # Filter sheet names to start from index 2 (sheet 3)
    filtered_sheet_names = sheet_names[start_index:]

    # If all_sheets_data is provided, filter out sheets that don't have data
    if all_sheets_data:
        filtered_sheet_names = [sheet_name for sheet_name in filtered_sheet_names 
                               if all_sheets_data.get(sheet_name, [])]

    # Find the index of the active sheet in the filtered list
    active_tab_index = None
    for i, sheet_name in enumerate(filtered_sheet_names):
        if sheet_name == active_sheet:
            active_tab_index = i
            break

    # Create a tab for each sheet
    tabs = html.Div([
        html.H4("Samples", style={'textAlign': 'center', 'marginTop': '30px', 'marginBottom': '15px'}),
        dcc.Tabs(
            id='sheet-tabs',
            value=active_sheet if active_tab_index is not None else (filtered_sheet_names[0] if filtered_sheet_names else None),
            children=[
                dcc.Tab(
                    label=sheet_name,
                    value=sheet_name,
                    id={'type': 'sheet-tab', 'index': i + start_index},  # Maintain the same ID pattern for compatibility
                    style={
                        'padding': '10px 20px',
                        'borderRadius': '4px 4px 0 0',
                    },
                    selected_style={
                        'backgroundColor': '#4CAF50',
                        'color': 'white',
                        'padding': '10px 20px',
                        'borderRadius': '4px 4px 0 0',
                        'fontWeight': 'bold',
                        'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
                    }
                ) for i, sheet_name in enumerate(filtered_sheet_names)
            ],
            style={
                'width': '100%',
                'marginBottom': '20px',
            },
            colors={
                "border": "#ddd",
                "primary": "#4CAF50",
                "background": "#f5f5f5"
            }
        )
    ], style={'marginTop': '30px', 'borderTop': '1px solid #ddd', 'paddingTop': '20px'})

    return tabs

# --- Run the app ---
if __name__ == '__main__':
    # Get port from environment variable or use 8050 as default
    port = int(os.environ.get('PORT', 8050))

    # Determine if we're in a production environment
    debug = os.environ.get('ENVIRONMENT', 'development') != 'production'

    app.run(host='0.0.0.0', port=port, debug=debug)
