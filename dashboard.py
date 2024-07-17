import pandas as pd
import plotly.express as px
import dash_bootstrap_components as dbc
from dash import Dash, dcc, html, dash_table, no_update
from dash.dependencies import Input, Output, State
from dash import callback_context

def load_excel(file_name, sheet_name):
    def round_to_millions(x):
        if isinstance(x, (int, float)) and x is not None:
            return round(x / 1_000_000, 1)
        return x
    def extract_tables(sheet_data):
        tables = {}
        num_columns = sheet_data.shape[1]
        table_count = 1
        for start_col in range(0, num_columns, 3):
            end_col = start_col + 2
            if end_col < num_columns:
                current_table = sheet_data.iloc[:, start_col:end_col]
                current_table.columns = sheet_data.columns[start_col:end_col]
                tables[f"table_{table_count}"] = current_table
                table_count += 1
        return tables
    data = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
    data = data.applymap(round_to_millions)
    tables = extract_tables(data)
    return tables

cpdf = load_excel('data/PowerBI Dashboard data - Copy.xlsx', 'CP-DF Summary')
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

def create_plot(data, x_column, y_column, title, color='royalblue'):
    fig = px.bar(data, x=x_column, y=y_column, title=title, color_discrete_sequence=[color])
    fig.update_layout(
        barmode='group',
        xaxis_tickangle=-45,
        xaxis_title=None,
        yaxis_title=None,
        title={'font': {'size': 14}, 'x': 0.5, 'xanchor': 'center'},
        margin={'t': 50}
    )

    fig.update_traces(texttemplate='%{y}', textposition='outside')
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')

    return fig

figures = {
    'fig1': create_plot(cpdf['table_12'], cpdf['table_12'].columns[0], cpdf['table_12'].columns[1], 'Total Secondary Purchases and Sales by Counterparty from 6/30/2023 to 6/30/2024 (MM)'),
    'fig2': create_plot(cpdf['table_4'], cpdf['table_4'].columns[0], cpdf['table_4'].columns[1], 'Primary Buys by Counterparty From 6/30/2023 to 6/30/2024 (MM)'),
    'fig3': create_plot(cpdf['table_1'], cpdf['table_1'].columns[0], cpdf['table_1'].columns[1], 'Trading Volume by Counterparty From 6/30/2023 to 6/30/2024 (MM)'),
    'fig4': create_plot(cpdf['table_2'], cpdf['table_2'].columns[0], cpdf['table_2'].columns[1], 'Primary and Secondary Buys by Counterparty from 6/30/2023 to 6/30/2024 (MM)'),
    'fig5': create_plot(cpdf['table_3'], cpdf['table_3'].columns[0], cpdf['table_3'].columns[1], 'Sells by Counterparty From 6/30/2023 to 6/30/2024 (MM)')
}

app.layout = html.Div([
    html.H1('Dashboard Summary 6/30/24'),
    html.Div([
        html.Div([
            dcc.Graph(id='fig1', figure=figures['fig1'], style={'display': 'inline-block', 'width': '100%', 'height': '300px'}),
            html.Button('Enlarge Graph 1', id='btn-fig1', className='small-button')
        ], style={'display': 'inline-block', 'width': '48%', 'vertical-align': 'top'}),
        html.Div([
            dcc.Graph(id='fig2', figure=figures['fig2'], style={'display': 'inline-block', 'width': '100%', 'height': '300px'}),
            html.Button('Enlarge Graph 2', id='btn-fig2', className='small-button')
        ], style={'display': 'inline-block', 'width': '48%', 'vertical-align': 'top'}),
        html.Div([
            dcc.Graph(id='fig3', figure=figures['fig3'], style={'display': 'inline-block', 'width': '100%', 'height': '300px'}),
            html.Button('Enlarge Graph 3', id='btn-fig3', className='small-button')
        ], style={'display': 'inline-block', 'width': '48%', 'vertical-align': 'top'}),
        html.Div([
            dcc.Graph(id='fig4', figure=figures['fig4'], style={'display': 'inline-block', 'width': '100%', 'height': '300px'}),
            html.Button('Enlarge Graph 4', id='btn-fig4', className='small-button')
        ], style={'display': 'inline-block', 'width': '48%', 'vertical-align': 'top'}),
        html.Div([
            dcc.Graph(id='fig5', figure=figures['fig5'], style={'display': 'inline-block', 'width': '100%', 'height': '300px'}),
            html.Button('Enlarge Graph 5', id='btn-fig5', className='small-button')
        ], style={'display': 'inline-block', 'width': '48%', 'vertical-align': 'top'}),
    ]),
    dbc.Modal(
        [
            dbc.ModalHeader(dbc.ModalTitle("Enlarged Graph")),
            dbc.ModalBody(dcc.Graph(id='enlarged-plot', style={'width': '100%', 'height': '600px'})),
            dbc.ModalBody(html.Div(id='data-table')),
            dbc.ModalFooter(
                dbc.Button("Close", id="close", className="ms-auto", n_clicks=0)
            ),
        ],
        id="modal",
        size="xl",
        is_open=False,
    ),
], style={'font-size': '12px', 'height': '100vh'})

app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            .small-button {
                font-size: 12px;
                padding: 5px 10px;
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

@app.callback(
    [Output('enlarged-plot', 'figure'), Output('data-table', 'children'), Output('modal', 'is_open')],
    [Input('btn-fig1', 'n_clicks'), Input('btn-fig2', 'n_clicks'), Input('btn-fig3', 'n_clicks'), Input('btn-fig4', 'n_clicks'), Input('btn-fig5', 'n_clicks'), Input('close', 'n_clicks')],
    [State('modal', 'is_open')]
)

def update_graph(btn1, btn2, btn3, btn4, btn5, close, is_open):
    ctx = callback_context
    if not ctx.triggered:
        return no_update, no_update, is_open
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    if button_id == 'close':
        return no_update, no_update, not is_open
    button_to_fig = {
        'btn-fig1': 'fig1',
        'btn-fig2': 'fig2',
        'btn-fig3': 'fig3',
        'btn-fig4': 'fig4',
        'btn-fig5': 'fig5'
    }
    fig_id = button_to_fig[button_id]
    fig = figures[fig_id]
    data = cpdf[fig_id.replace('fig', 'table_')]
    columns = data.columns
    data_table = dash_table.DataTable(
        columns=[{"name": col, "id": col} for col in columns],
        data=data.to_dict('records'),
        page_size=10
    )
    return fig, data_table, True

if __name__ == '__main__':
    app.run_server(debug=True)

