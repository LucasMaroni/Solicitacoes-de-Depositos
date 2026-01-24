# app.py - Dashboard Completo de Análise de Custos Diários
import dash
from dash import dcc, html, Input, Output, dash_table, State
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import base64
import io
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURAÇÃO DA APLICAÇÃO
# ============================================================================
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        dbc.icons.FONT_AWESOME,
        'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap'
    ],
    suppress_callback_exceptions=True,
    meta_tags=[
        {'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}
    ]
)

app.title = "📊 Dashboard de Custos Diários - Solicitações de Depósitos"
server = app.server

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================
def parse_contents(contents, filename):
    """Processa arquivo Excel enviado"""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    try:
        if 'xlsx' in filename or 'xls' in filename:
            # Ler o arquivo Excel
            df = pd.read_excel(io.BytesIO(decoded))
            return df
    except Exception as e:
        print(f"Erro ao processar arquivo: {e}")
        return None

def process_dataframe(df):
    """Processa o dataframe para análise"""
    if df is None or df.empty:
        return None
    
    df_processed = df.copy()
    
    # Converter tipos de dados
    # ID
    if 'ID' in df_processed.columns:
        df_processed['ID'] = pd.to_numeric(df_processed['ID'], errors='coerce')
    
    # Valor - tratamento robusto
    if 'Valor' in df_processed.columns:
        def clean_value(x):
            if pd.isna(x):
                return np.nan
            try:
                # Converter para string
                x_str = str(x)
                # Remover caracteres não numéricos exceto ponto e vírgula
                x_str = ''.join(c for c in x_str if c.isdigit() or c in ',.-')
                # Substituir vírgula por ponto se for separador decimal
                if ',' in x_str and '.' not in x_str:
                    # Se tem uma vírgula e é seguida por menos de 3 dígitos, provavelmente é decimal
                    parts = x_str.split(',')
                    if len(parts) == 2 and len(parts[1]) <= 2:
                        x_str = x_str.replace(',', '.')
                    else:
                        # Remover todas as vírgulas
                        x_str = x_str.replace(',', '')
                elif ',' in x_str and '.' in x_str:
                    # Se tem ambos, remover vírgulas
                    x_str = x_str.replace(',', '')
                
                # Remover pontos extras (milhares)
                if '.' in x_str:
                    parts = x_str.split('.')
                    if len(parts) > 2:  # Tem separador de milhares
                        x_str = parts[0] + ''.join(parts[1:])
                
                # Converter para float
                return float(x_str)
            except:
                return np.nan
        
        df_processed['Valor'] = df_processed['Valor'].apply(clean_value)
    
    # Datas
    date_columns = ['Criado', 'Modificado']
    for col in date_columns:
        if col in df_processed.columns:
            df_processed[col] = pd.to_datetime(df_processed[col], errors='coerce')
    
    # Extrair features de data
    if 'Criado' in df_processed.columns:
        df_processed['Data_Criacao'] = df_processed['Criado'].dt.date
        df_processed['Ano_Mes'] = df_processed['Criado'].dt.to_period('M').astype(str)
        df_processed['Mes'] = df_processed['Criado'].dt.month
        df_processed['Dia'] = df_processed['Criado'].dt.day
        df_processed['Dia_Semana'] = df_processed['Criado'].dt.day_name()
        df_processed['Hora_Criacao'] = df_processed['Criado'].dt.hour
        df_processed['Minuto_Criacao'] = df_processed['Criado'].dt.minute
        
        # Traduzir dias da semana
        dias_portugues = {
            'Monday': 'Segunda',
            'Tuesday': 'Terça',
            'Wednesday': 'Quarta',
            'Thursday': 'Quinta',
            'Friday': 'Sexta',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }
        df_processed['Dia_Semana_PT'] = df_processed['Dia_Semana'].map(dias_portugues)
    
    # Criar faixas de valor
    if 'Valor' in df_processed.columns:
        bins = [0, 100, 500, 1000, 5000, float('inf')]
        labels = ['0-100', '101-500', '501-1000', '1001-5000', '5000+']
        df_processed['Faixa_Valor'] = pd.cut(df_processed['Valor'], bins=bins, labels=labels)
    
    return df_processed

def create_kpi_card(title, value, icon, color, subtitle=None, change=None):
    """Cria um card de KPI"""
    card_content = []
    
    # Ícone e título
    card_content.append(
        html.Div([
            html.I(className=f"{icon} me-2", style={'color': color}),
            html.Span(title, className="text-muted small")
        ], className="d-flex align-items-center mb-2")
    )
    
    # Valor principal
    card_content.append(
        html.H4(value, className="mb-1", style={'color': color})
    )
    
    # Subtitle e change se fornecidos
    if subtitle or change:
        sub_content = []
        if subtitle:
            sub_content.append(html.Span(subtitle, className="text-muted small me-2"))
        if change:
            change_color = "text-success" if change >= 0 else "text-danger"
            change_icon = "fa-arrow-up" if change >= 0 else "fa-arrow-down"
            sub_content.append(
                html.Span([
                    html.I(className=f"fas {change_icon} me-1"),
                    f"{abs(change):.1f}%"
                ], className=f"small {change_color}")
            )
        card_content.append(html.Div(sub_content))
    
    return dbc.Card(
        dbc.CardBody(card_content),
        className="h-100 shadow-sm border-0"
    )

def create_filter_card(title, children):
    """Cria um card de filtro"""
    return dbc.Card([
        dbc.CardHeader(title, className="bg-light"),
        dbc.CardBody(children)
    ], className="shadow-sm mb-3")

# ============================================================================
# LAYOUT PRINCIPAL
# ============================================================================
app.layout = dbc.Container([
    # Armazenamento de dados
    dcc.Store(id='stored-data'),
    dcc.Store(id='filtered-data'),
    
    # Upload de arquivo
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            html.I(className="fas fa-cloud-upload-alt me-2"),
            'Arraste ou clique para fazer upload do Excel'
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px 0',
            'cursor': 'pointer'
        },
        multiple=False
    ),
    
    # Loading overlay
    dcc.Loading(
        id="loading-overlay",
        type="circle",
        children=[
            # Cabeçalho
            dbc.Row([
                dbc.Col([
                    html.Div([
                        html.H1("📊 Dashboard de Custos Diários", 
                               className="text-primary mb-2"),
                        html.P("Análise inteligente de solicitações de depósitos e despesas operacionais", 
                              className="text-muted lead mb-4"),
                        html.Hr()
                    ])
                ], width=12)
            ], className="mb-4"),
            
            # Cards de KPIs
            dbc.Row(id="kpi-cards", className="mb-4"),
            
            # Filtros
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H5("🔍 Filtros Avançados", className="mb-0"),
                            dbc.Button("Aplicar Filtros", id="btn-apply-filters", 
                                      color="primary", size="sm", className="float-end")
                        ]),
                        dbc.CardBody([
                            dbc.Row([
                                dbc.Col([
                                    html.Label("Classificação:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-classificacao',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione classificação..."
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Finalidade:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-finalidade',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione finalidade..."
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Período:", className="form-label"),
                                    dcc.DatePickerRange(
                                        id='filtro-data',
                                        start_date=None,
                                        end_date=None,
                                        display_format='DD/MM/YYYY',
                                        className="w-100"
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Faixa de Valor (R$):", className="form-label"),
                                    dcc.RangeSlider(
                                        id='filtro-valor',
                                        min=0,
                                        max=10000,
                                        step=100,
                                        value=[0, 10000],
                                        marks={0: '0', 5000: '5.000', 10000: '10.000+'},
                                        tooltip={"placement": "bottom", "always_visible": True}
                                    )
                                ], md=6, lg=3)
                            ], className="mb-3"),
                            
                            dbc.Row([
                                dbc.Col([
                                    html.Label("Status:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-status',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione status..."
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Solicitante:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-solicitante',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione solicitante..."
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Motorista:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-motorista',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione motorista..."
                                    )
                                ], md=6, lg=3),
                                
                                dbc.Col([
                                    html.Label("Gestor:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-gestor',
                                        options=[],
                                        multi=True,
                                        placeholder="Selecione gestor..."
                                    )
                                ], md=6, lg=3)
                            ])
                        ])
                    ], className="shadow-sm mb-4")
                ], width=12)
            ]),
            
            # Gráficos Principais
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📈 Evolução de Gastos Diários"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-evolucao', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=8),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🏷️ Distribuição por Classificação"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-classificacao', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=4)
            ], className="mb-4"),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🎯 Top Finalidades"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-finalidades', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("⏰ Distribuição por Hora do Dia"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-horas', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6)
            ], className="mb-4"),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📊 Análise Detalhada por Dia da Semana"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-dias-semana', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("💰 Distribuição de Valores (Box Plot)"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-boxplot', config={'displayModeBar': True})
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6)
            ], className="mb-4"),
            
            # Tabela de Dados
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H5("📋 Detalhamento das Solicitações", className="mb-0 d-inline"),
                            html.Div([
                                dbc.Button(
                                    [html.I(className="fas fa-file-csv me-2"), "CSV"],
                                    id="btn-export-csv",
                                    color="success",
                                    size="sm",
                                    className="me-2"
                                ),
                                dbc.Button(
                                    [html.I(className="fas fa-file-excel me-2"), "Excel"],
                                    id="btn-export-excel",
                                    color="primary",
                                    size="sm",
                                    className="me-2"
                                ),
                                dbc.Button(
                                    [html.I(className="fas fa-print me-2"), "Imprimir"],
                                    id="btn-print",
                                    color="secondary",
                                    size="sm"
                                )
                            ], className="float-end")
                        ]),
                        dbc.CardBody([
                            dash_table.DataTable(
                                id='tabela-dados',
                                columns=[],
                                data=[],
                                page_size=15,
                                style_table={'overflowX': 'auto'},
                                style_cell={
                                    'textAlign': 'left',
                                    'padding': '10px',
                                    'fontSize': '12px',
                                    'fontFamily': 'Inter, sans-serif',
                                    'whiteSpace': 'normal',
                                    'height': 'auto'
                                },
                                style_header={
                                    'backgroundColor': '#f8f9fa',
                                    'fontWeight': '600',
                                    'color': '#495057',
                                    'border': 'none'
                                },
                                style_data={
                                    'border': '1px solid #e9ecef'
                                },
                                style_data_conditional=[
                                    {
                                        'if': {'row_index': 'odd'},
                                        'backgroundColor': '#f8f9fa'
                                    }
                                ],
                                filter_action="native",
                                sort_action="native",
                                sort_mode="multi",
                                page_action="native",
                                style_cell_conditional=[
                                    {'if': {'column_id': 'ID'}, 'width': '80px'},
                                    {'if': {'column_id': 'Valor'}, 'width': '120px'},
                                    {'if': {'column_id': 'Criado'}, 'width': '180px'},
                                    {'if': {'column_id': 'Status'}, 'width': '100px'},
                                ],
                                export_format="csv"
                            )
                        ])
                    ], className="shadow-sm")
                ], width=12)
            ], className="mb-4"),
            
            # Downloads
            dcc.Download(id="download-csv"),
            dcc.Download(id="download-excel"),
            
            # Footer
            dbc.Row([
                dbc.Col([
                    html.Hr(),
                    html.Div([
                        html.P([
                            "Dashboard desenvolvido para análise de custos diários | ",
                            html.Span(id="data-info", className="text-muted"),
                            html.Br(),
                            html.Small(
                                f"Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                                className="text-muted"
                            )
                        ], className="text-center mt-3")
                    ])
                ], width=12)
            ])
        ]
    )
], fluid=True, className="p-3", style={'fontFamily': 'Inter, sans-serif'})

# ============================================================================
# CALLBACKS
# ============================================================================

# Callback para upload e processamento de dados
@app.callback(
    [Output('stored-data', 'data'),
     Output('filtro-classificacao', 'options'),
     Output('filtro-finalidade', 'options'),
     Output('filtro-status', 'options'),
     Output('filtro-solicitante', 'options'),
     Output('filtro-motorista', 'options'),
     Output('filtro-gestor', 'options'),
     Output('filtro-data', 'start_date'),
     Output('filtro-data', 'end_date'),
     Output('filtro-data', 'min_date_allowed'),
     Output('filtro-data', 'max_date_allowed'),
     Output('data-info', 'children')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_data(contents, filename):
    """Processa upload de arquivo e atualiza filtros"""
    if contents is None:
        return [dash.no_update] * 12
    
    # Processar arquivo
    df = parse_contents(contents, filename)
    if df is None:
        return [dash.no_update] * 12
    
    # Processar dados
    df_processed = process_dataframe(df)
    
    if df_processed is None or df_processed.empty:
        return [dash.no_update] * 12
    
    # Preparar opções para filtros
    classificacao_opts = [{'label': cat, 'value': cat} 
                         for cat in sorted(df_processed['Classificação'].dropna().unique())]
    finalidade_opts = [{'label': fin, 'value': fin} 
                      for fin in sorted(df_processed['Finalidade'].dropna().unique())]
    status_opts = [{'label': stat, 'value': stat} 
                  for stat in sorted(df_processed['Status'].dropna().unique())]
    solicitante_opts = [{'label': sol, 'value': sol} 
                       for sol in sorted(df_processed['Solicitante'].dropna().unique())]
    motorista_opts = [{'label': mot, 'value': mot} 
                     for mot in sorted(df_processed['Nome Motorista'].dropna().unique())]
    
    # Gestor - converter para string
    gestor_opts = []
    if 'Gestor' in df_processed.columns:
        gestores = df_processed['Gestor'].dropna().unique()
        for gestor in sorted(gestores, key=lambda x: str(x)):
            gestor_opts.append({'label': str(gestor), 'value': str(gestor)})
    
    # Datas
    if 'Criado' in df_processed.columns:
        min_date = df_processed['Criado'].min().date()
        max_date = df_processed['Criado'].max().date()
        start_date = min_date
        end_date = max_date
    else:
        min_date = max_date = start_date = end_date = None
    
    # Info
    total_gasto = df_processed['Valor'].sum()
    total_registros = len(df_processed)
    data_info = f"Total: R$ {total_gasto:,.2f} | {total_registros:,} registros"
    
    # Converter para JSON para storage
    df_json = df_processed.to_json(date_format='iso', orient='split')
    
    return [
        df_json,
        classificacao_opts,
        finalidade_opts,
        status_opts,
        solicitante_opts,
        motorista_opts,
        gestor_opts,
        start_date,
        end_date,
        min_date,
        max_date,
        data_info
    ]

# Callback para aplicar filtros
@app.callback(
    [Output('filtered-data', 'data'),
     Output('kpi-cards', 'children')],
    [Input('btn-apply-filters', 'n_clicks')],
    [State('stored-data', 'data'),
     State('filtro-classificacao', 'value'),
     State('filtro-finalidade', 'value'),
     State('filtro-data', 'start_date'),
     State('filtro-data', 'end_date'),
     State('filtro-valor', 'value'),
     State('filtro-status', 'value'),
     State('filtro-solicitante', 'value'),
     State('filtro-motorista', 'value'),
     State('filtro-gestor', 'value')]
)
def apply_filters(n_clicks, stored_data, classificacao, finalidade, start_date, 
                  end_date, valor_range, status, solicitante, motorista, gestor):
    """Aplica filtros aos dados"""
    if stored_data is None:
        return None, []
    
    # Carregar dados
    df = pd.read_json(stored_data, orient='split')
    
    # Aplicar filtros
    df_filtered = df.copy()
    
    # Filtro por classificação
    if classificacao and len(classificacao) > 0:
        df_filtered = df_filtered[df_filtered['Classificação'].isin(classificacao)]
    
    # Filtro por finalidade
    if finalidade and len(finalidade) > 0:
        df_filtered = df_filtered[df_filtered['Finalidade'].isin(finalidade)]
    
    # Filtro por data
    if start_date and end_date:
        df_filtered = df_filtered[
            (df_filtered['Data_Criacao'] >= pd.to_datetime(start_date).date()) &
            (df_filtered['Data_Criacao'] <= pd.to_datetime(end_date).date())
        ]
    
    # Filtro por valor
    if valor_range:
        df_filtered = df_filtered[
            (df_filtered['Valor'] >= valor_range[0]) &
            (df_filtered['Valor'] <= valor_range[1])
        ]
    
    # Filtro por status
    if status and len(status) > 0:
        df_filtered = df_filtered[df_filtered['Status'].isin(status)]
    
    # Filtro por solicitante
    if solicitante and len(solicitante) > 0:
        df_filtered = df_filtered[df_filtered['Solicitante'].isin(solicitante)]
    
    # Filtro por motorista
    if motorista and len(motorista) > 0:
        df_filtered = df_filtered[df_filtered['Nome Motorista'].isin(motorista)]
    
    # Filtro por gestor
    if gestor and len(gestor) > 0:
        df_filtered = df_filtered[df_filtered['Gestor'].astype(str).isin(gestor)]
    
    # Criar cards de KPI
    if df_filtered.empty:
        kpi_cards = [
            dbc.Col([
                dbc.Alert("Nenhum dado encontrado com os filtros aplicados", 
                         color="warning")
            ], width=12)
        ]
    else:
        # Calcular métricas
        total_gasto = df_filtered['Valor'].sum()
        total_registros = len(df_filtered)
        valor_medio = df_filtered['Valor'].mean()
        valor_mediano = df_filtered['Valor'].median()
        
        # Categoria mais frequente
        if 'Classificação' in df_filtered.columns and not df_filtered['Classificação'].empty:
            categoria_top = df_filtered['Classificação'].mode().iloc[0] if not df_filtered['Classificação'].mode().empty else "N/A"
        else:
            categoria_top = "N/A"
        
        # Finalidade mais frequente
        if 'Finalidade' in df_filtered.columns and not df_filtered['Finalidade'].empty:
            finalidade_top = df_filtered['Finalidade'].mode().iloc[0] if not df_filtered['Finalidade'].mode().empty else "N/A"
        else:
            finalidade_top = "N/A"
        
        # Valor máximo
        valor_max = df_filtered['Valor'].max()
        
        # Criar cards
        kpi_cards = [
            dbc.Col([
                create_kpi_card(
                    title="Total Gasto",
                    value=f"R$ {total_gasto:,.2f}",
                    icon="fas fa-money-bill-wave",
                    color="#4361ee",
                    subtitle=f"{total_registros:,} registros"
                )
            ], md=6, lg=3, className="mb-3"),
            
            dbc.Col([
                create_kpi_card(
                    title="Valor Médio",
                    value=f"R$ {valor_medio:,.2f}",
                    icon="fas fa-calculator",
                    color="#4cc9f0",
                    subtitle=f"Mediana: R$ {valor_mediano:,.2f}"
                )
            ], md=6, lg=3, className="mb-3"),
            
            dbc.Col([
                create_kpi_card(
                    title="Categoria Top",
                    value=categoria_top,
                    icon="fas fa-tag",
                    color="#7209b7",
                    subtitle=f"Finalidade: {finalidade_top}"
                )
            ], md=6, lg=3, className="mb-3"),
            
            dbc.Col([
                create_kpi_card(
                    title="Valor Máximo",
                    value=f"R$ {valor_max:,.2f}",
                    icon="fas fa-chart-line",
                    color="#f72585",
                    subtitle="Maior solicitação"
                )
            ], md=6, lg=3, className="mb-3")
        ]
    
    # Converter dados filtrados para JSON
    filtered_json = df_filtered.to_json(date_format='iso', orient='split') if not df_filtered.empty else None
    
    return filtered_json, kpi_cards

# Callback para atualizar gráficos
@app.callback(
    [Output('grafico-evolucao', 'figure'),
     Output('grafico-classificacao', 'figure'),
     Output('grafico-finalidades', 'figure'),
     Output('grafico-horas', 'figure'),
     Output('grafico-dias-semana', 'figure'),
     Output('grafico-boxplot', 'figure'),
     Output('tabela-dados', 'data'),
     Output('tabela-dados', 'columns')],
    [Input('filtered-data', 'data')]
)
def update_charts(filtered_data):
    """Atualiza todos os gráficos com dados filtrados"""
    if filtered_data is None:
        return [go.Figure()] * 7
    
    # Carregar dados filtrados
    df = pd.read_json(filtered_data, orient='split')
    
    if df.empty:
        empty_fig = go.Figure()
        empty_fig.update_layout(
            title="Sem dados para exibir",
            xaxis_visible=False,
            yaxis_visible=False,
            annotations=[dict(
                text="Nenhum dado encontrado",
                xref="paper", yref="paper",
                showarrow=False,
                font=dict(size=20)
            )]
        )
        return [empty_fig] * 6 + [[], []]
    
    # 1. Gráfico de Evolução Diária
    fig_evolucao = create_evolucao_chart(df)
    
    # 2. Gráfico de Classificação
    fig_classificacao = create_classificacao_chart(df)
    
    # 3. Gráfico de Finalidades
    fig_finalidades = create_finalidades_chart(df)
    
    # 4. Gráfico de Horas
    fig_horas = create_horas_chart(df)
    
    # 5. Gráfico de Dias da Semana
    fig_dias_semana = create_dias_semana_chart(df)
    
    # 6. Box Plot
    fig_boxplot = create_boxplot_chart(df)
    
    # 7. Tabela de Dados
    tabela_data, tabela_columns = create_table_data(df)
    
    return [
        fig_evolucao,
        fig_classificacao,
        fig_finalidades,
        fig_horas,
        fig_dias_semana,
        fig_boxplot,
        tabela_data,
        tabela_columns
    ]

# Funções para criação de gráficos
def create_evolucao_chart(df):
    """Cria gráfico de evolução diária"""
    if df.empty:
        return go.Figure()
    
    # Agrupar por data
    daily_data = df.groupby('Data_Criacao').agg({
        'Valor': ['sum', 'count'],
        'ID': 'nunique'
    }).reset_index()
    
    daily_data.columns = ['Data', 'Valor_Total', 'Quantidade', 'Registros_Unicos']
    
    # Calcular média móvel
    daily_data['Media_Movel'] = daily_data['Valor_Total'].rolling(window=3, min_periods=1).mean()
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Adicionar linha de valor total
    fig.add_trace(
        go.Scatter(
            x=daily_data['Data'],
            y=daily_data['Valor_Total'],
            mode='lines+markers',
            name='Gasto Diário',
            line=dict(color='#4361ee', width=3),
            marker=dict(size=8, color='#4361ee'),
            hovertemplate='<b>%{x|%d/%m}</b><br>R$ %{y:,.2f}<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Adicionar linha de média móvel
    fig.add_trace(
        go.Scatter(
            x=daily_data['Data'],
            y=daily_data['Media_Movel'],
            mode='lines',
            name='Média Móvel (3 dias)',
            line=dict(color='#f72585', width=2, dash='dash'),
            hovertemplate='<b>%{x|%d/%m}</b><br>R$ %{y:,.2f}<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Adicionar barras de quantidade
    fig.add_trace(
        go.Bar(
            x=daily_data['Data'],
            y=daily_data['Quantidade'],
            name='Quantidade',
            marker_color='rgba(76, 201, 240, 0.5)',
            hovertemplate='<b>%{x|%d/%m}</b><br>%{y} solicitações<extra></extra>'
        ),
        secondary_y=True
    )
    
    # Configurar layout
    fig.update_layout(
        title='Evolução de Gastos Diários',
        template='plotly_white',
        hovermode='x unified',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        margin=dict(l=50, r=50, t=80, b=50),
        height=400
    )
    
    fig.update_xaxes(title_text="Data")
    fig.update_yaxes(title_text="Valor (R$)", secondary_y=False)
    fig.update_yaxes(title_text="Quantidade", secondary_y=True)
    
    return fig

def create_classificacao_chart(df):
    """Cria gráfico de pizza por classificação"""
    if df.empty:
        return go.Figure()
    
    # Agrupar por classificação
    class_data = df.groupby('Classificação')['Valor'].sum().reset_index()
    class_data = class_data.sort_values('Valor', ascending=False)
    
    fig = px.pie(
        class_data,
        values='Valor',
        names='Classificação',
        title='Distribuição por Classificação',
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>',
        marker=dict(line=dict(color='#fff', width=2))
    )
    
    fig.update_layout(
        height=400,
        showlegend=True,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.05
        )
    )
    
    return fig

def create_finalidades_chart(df):
    """Cria gráfico de barras para top finalidades"""
    if df.empty:
        return go.Figure()
    
    # Agrupar por finalidade
    fin_data = df.groupby('Finalidade').agg({
        'Valor': 'sum',
        'ID': 'count'
    }).reset_index()
    
    fin_data.columns = ['Finalidade', 'Valor_Total', 'Quantidade']
    fin_data = fin_data.sort_values('Valor_Total', ascending=True).tail(15)
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        y=fin_data['Finalidade'],
        x=fin_data['Valor_Total'],
        orientation='h',
        name='Valor Total',
        marker_color='#4cc9f0',
        hovertemplate='<b>%{y}</b><br>R$ %{x:,.2f}<br>%{customdata} solicitações<extra></extra>',
        customdata=fin_data['Quantidade']
    ))
    
    fig.update_layout(
        title='Top 15 Finalidades por Valor',
        template='plotly_white',
        height=400,
        margin=dict(l=20, r=20, t=60, b=20),
        xaxis_title='Valor Total (R$)',
        yaxis_title='Finalidade',
        yaxis={'categoryorder': 'total ascending'}
    )
    
    return fig

def create_horas_chart(df):
    """Cria gráfico de distribuição por hora"""
    if df.empty:
        return go.Figure()
    
    # Agrupar por hora
    hour_data = df.groupby('Hora_Criacao').agg({
        'Valor': 'sum',
        'ID': 'count'
    }).reset_index()
    
    hour_data.columns = ['Hora', 'Valor_Total', 'Quantidade']
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Adicionar linha de valor total
    fig.add_trace(
        go.Scatter(
            x=hour_data['Hora'],
            y=hour_data['Valor_Total'],
            mode='lines+markers',
            name='Valor Total',
            line=dict(color='#7209b7', width=3),
            marker=dict(size=8, color='#7209b7'),
            hovertemplate='<b>%{x}:00h</b><br>R$ %{y:,.2f}<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Adicionar barras de quantidade
    fig.add_trace(
        go.Bar(
            x=hour_data['Hora'],
            y=hour_data['Quantidade'],
            name='Quantidade',
            marker_color='rgba(114, 9, 183, 0.3)',
            hovertemplate='<b>%{x}:00h</b><br>%{y} solicitações<extra></extra>'
        ),
        secondary_y=True
    )
    
    fig.update_layout(
        title='Distribuição por Hora do Dia',
        template='plotly_white',
        hovermode='x unified',
        height=400,
        margin=dict(l=50, r=50, t=80, b=50),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    fig.update_xaxes(
        title_text="Hora do Dia",
        tickmode='linear',
        tick0=0,
        dtick=1
    )
    
    fig.update_yaxes(title_text="Valor Total (R$)", secondary_y=False)
    fig.update_yaxes(title_text="Quantidade", secondary_y=True)
    
    return fig

def create_dias_semana_chart(df):
    """Cria gráfico de análise por dia da semana"""
    if df.empty:
        return go.Figure()
    
    # Ordem dos dias
    dias_ordem = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    
    # Agrupar por dia da semana
    day_data = df.groupby('Dia_Semana_PT').agg({
        'Valor': ['sum', 'mean', 'count'],
        'ID': 'nunique'
    }).reset_index()
    
    day_data.columns = ['Dia_Semana', 'Valor_Total', 'Valor_Medio', 'Quantidade', 'Registros_Unicos']
    
    # Ordenar
    day_data['Dia_Semana'] = pd.Categorical(day_data['Dia_Semana'], categories=dias_ordem, ordered=True)
    day_data = day_data.sort_values('Dia_Semana')
    
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Valor Total por Dia', 'Quantidade por Dia', 
                       'Valor Médio por Dia', 'Eficiência por Dia'),
        vertical_spacing=0.15,
        horizontal_spacing=0.15
    )
    
    # 1. Valor Total
    fig.add_trace(
        go.Bar(
            x=day_data['Dia_Semana'],
            y=day_data['Valor_Total'],
            name='Valor Total',
            marker_color='#4361ee',
            hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>'
        ),
        row=1, col=1
    )
    
    # 2. Quantidade
    fig.add_trace(
        go.Bar(
            x=day_data['Dia_Semana'],
            y=day_data['Quantidade'],
            name='Quantidade',
            marker_color='#4cc9f0',
            hovertemplate='<b>%{x}</b><br>%{y} solicitações<extra></extra>'
        ),
        row=1, col=2
    )
    
    # 3. Valor Médio
    fig.add_trace(
        go.Bar(
            x=day_data['Dia_Semana'],
            y=day_data['Valor_Medio'],
            name='Valor Médio',
            marker_color='#7209b7',
            hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>'
        ),
        row=2, col=1
    )
    
    # 4. Eficiência (Valor/Quantidade)
    day_data['Eficiencia'] = day_data['Valor_Total'] / day_data['Quantidade']
    fig.add_trace(
        go.Bar(
            x=day_data['Dia_Semana'],
            y=day_data['Eficiencia'],
            name='Eficiência',
            marker_color='#f72585',
            hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}/solicitação<extra></extra>'
        ),
        row=2, col=2
    )
    
    fig.update_layout(
        title='Análise por Dia da Semana',
        template='plotly_white',
        height=600,
        showlegend=False,
        margin=dict(l=50, r=50, t=100, b=50)
    )
    
    # Atualizar eixos
    for i in range(1, 3):
        for j in range(1, 3):
            fig.update_xaxes(title_text="Dia da Semana", row=i, col=j)
    
    fig.update_yaxes(title_text="Valor (R$)", row=1, col=1)
    fig.update_yaxes(title_text="Quantidade", row=1, col=2)
    fig.update_yaxes(title_text="Valor Médio (R$)", row=2, col=1)
    fig.update_yaxes(title_text="Eficiência (R$/solicitação)", row=2, col=2)
    
    return fig

def create_boxplot_chart(df):
    """Cria box plot de valores por categoria"""
    if df.empty:
        return go.Figure()
    
    # Ordenar por mediana
    medians = df.groupby('Classificação')['Valor'].median().sort_values(ascending=False)
    df['Classificação'] = pd.Categorical(df['Classificação'], categories=medians.index, ordered=True)
    
    fig = px.box(
        df,
        x='Classificação',
        y='Valor',
        color='Classificação',
        points='all',
        title='Distribuição de Valores por Classificação',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    
    fig.update_layout(
        height=400,
        template='plotly_white',
        margin=dict(l=50, r=50, t=80, b=100),
        xaxis_tickangle=-45,
        showlegend=False
    )
    
    fig.update_yaxes(title_text="Valor (R$)")
    fig.update_xaxes(title_text="Classificação")
    
    return fig

def create_table_data(df):
    """Prepara dados para tabela"""
    if df.empty:
        return [], []
    
    # Selecionar colunas para exibir
    display_columns = ['ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
                      'Valor', 'Nome Motorista', 'Solicitante', 'Criado', 'Gestor']
    
    # Filtrar colunas disponíveis
    available_columns = [col for col in display_columns if col in df.columns]
    
    # Preparar dados
    table_data = df[available_columns].copy()
    
    # Formatar valores
    if 'Valor' in table_data.columns:
        table_data['Valor'] = table_data['Valor'].apply(lambda x: f"R$ {x:,.2f}")
    
    if 'Criado' in table_data.columns:
        table_data['Criado'] = table_data['Criado'].dt.strftime('%d/%m/%Y %H:%M')
    
    # Criar definições de colunas
    columns = []
    for col in available_columns:
        col_def = {
            'name': col,
            'id': col
        }
        
        # Formatação especial para algumas colunas
        if col == 'Valor':
            col_def['type'] = 'numeric'
        elif col == 'ID':
            col_def['type'] = 'numeric'
        
        columns.append(col_def)
    
    return table_data.to_dict('records'), columns

# Callbacks para exportação
@app.callback(
    Output('download-csv', 'data'),
    Input('btn-export-csv', 'n_clicks'),
    State('filtered-data', 'data'),
    prevent_initial_call=True
)
def export_csv(n_clicks, filtered_data):
    """Exporta dados como CSV"""
    if n_clicks and filtered_data:
        df = pd.read_json(filtered_data, orient='split')
        return dcc.send_data_frame(df.to_csv, "dados_filtrados.csv", index=False)
    return None

@app.callback(
    Output('download-excel', 'data'),
    Input('btn-export-excel', 'n_clicks'),
    State('filtered-data', 'data'),
    prevent_initial_call=True
)
def export_excel(n_clicks, filtered_data):
    """Exporta dados como Excel"""
    if n_clicks and filtered_data:
        df = pd.read_json(filtered_data, orient='split')
        return dcc.send_data_frame(df.to_excel, "dados_filtrados.xlsx", index=False)
    return None

# Callback para impressão
@app.callback(
    Output('btn-print', 'children'),
    Input('btn-print', 'n_clicks'),
    prevent_initial_call=True
)
def print_dashboard(n_clicks):
    """Simula impressão do dashboard"""
    if n_clicks:
        import time
        return [html.I(className="fas fa-spinner fa-spin me-2"), "Imprimindo..."]
    return [html.I(className="fas fa-print me-2"), "Imprimir"]

# ============================================================================
# INICIALIZAÇÃO
# ============================================================================
if __name__ == '__main__':
    print("=" * 60)
    print("🚀 DASHBOARD DE CUSTOS DIÁRIOS - INICIANDO")
    print("=" * 60)
    print("\n📊 Acesse o dashboard em: http://localhost:8050")
    print("📁 Faça upload do arquivo Excel para começar")
    print("🔄 Pressione Ctrl+C para parar o servidor\n")
    
    app.run_server(
        debug=True,
        port=8050,
        host='0.0.0.0',
        dev_tools_ui=True,
        dev_tools_hot_reload=True
    )
