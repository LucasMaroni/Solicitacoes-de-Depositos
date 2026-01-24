import dash
from dash import dcc, html, Input, Output, dash_table, State
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime, timedelta
import base64
import io
import warnings
warnings.filterwarnings('ignore')

# Inicializar aplicação Dash
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        dbc.icons.FONT_AWESOME,
        'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css'
    ],
    suppress_callback_exceptions=True,
    meta_tags=[
        {"name": "viewport", "content": "width=device-width, initial-scale=1"}
    ]
)

app.title = "Dashboard de Custos Diários - Solicitações de Depósitos"

# ============================================================================
# FUNÇÕES DE CARREGAMENTO E PROCESSAMENTO DE DADOS
# ============================================================================

def load_and_process_data(uploaded_file=None, file_path=None):
    """Carrega e processa os dados do Excel"""
    try:
        if uploaded_file:
            # Se foi feito upload
            content_type, content_string = uploaded_file.split(',')
            decoded = base64.b64decode(content_string)
            df = pd.read_excel(io.BytesIO(decoded), dtype=str)
        elif file_path:
            # Se está usando arquivo local
            df = pd.read_excel(file_path, dtype=str)
        else:
            # Dados de exemplo (para demonstração)
            df = create_sample_data()
        
        print(f"✅ Dados carregados: {len(df)} registros")
        
        # Padronizar nomes das colunas
        df.columns = df.columns.str.strip()
        
        # Verificar colunas obrigatórias
        required_columns = ['ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
                          'Valor', 'Nome Motorista', 'Solicitante', 'Criado']
        
        # Renomear colunas comuns
        column_mapping = {
            'Placa Cavalo/Carreta': 'Placa',
            'Ordem de Serviço': 'Ordem_Servico',
            'Conta corrente / poupança': 'Conta',
            'Nome Favorecido': 'Favorecido'
        }
        
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Converter tipos de dados
        # ID
        if 'ID' in df.columns:
            df['ID'] = pd.to_numeric(df['ID'], errors='coerce')
        
        # Valor - processamento robusto
        if 'Valor' in df.columns:
            df['Valor'] = df['Valor'].astype(str)
            # Remover caracteres não numéricos, exceto ponto, vírgula e hífen
            df['Valor'] = df['Valor'].str.replace(r'[^\d.,-]', '', regex=True)
            # Substituir vírgula por ponto para decimal
            df['Valor'] = df['Valor'].str.replace(',', '.', regex=False)
            # Remover múltiplos pontos
            df['Valor'] = df['Valor'].apply(lambda x: self._fix_multiple_dots(x) if isinstance(x, str) else x)
            df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
        
        # Datas
        date_columns = ['Criado', 'Modificado']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
        
        # Extrair features de data
        if 'Criado' in df.columns:
            df['Data_Criacao'] = df['Criado'].dt.date
            df['Mes_Criacao'] = df['Criado'].dt.to_period('M').astype(str)
            df['Ano_Criacao'] = df['Criado'].dt.year
            df['Dia_Semana'] = df['Criado'].dt.day_name()
            df['Hora_Criacao'] = df['Criado'].dt.hour
            df['Dia_Mes'] = df['Criado'].dt.day
        
        # Processar empresa
        if 'Empresa' in df.columns:
            df['Empresa'] = df['Empresa'].fillna('Não Informada').astype(str).str.strip()
        else:
            df['Empresa'] = 'Não Informada'
        
        # Processar ordem de serviço
        if 'Ordem_Servico' in df.columns:
            df['Ordem_Servico'] = df['Ordem_Servico'].fillna('Não Informada').astype(str).str.strip()
        else:
            df['Ordem_Servico'] = 'Não Informada'
        
        # Classificar por data
        if 'Criado' in df.columns:
            df = df.sort_values('Criado', ascending=False)
        
        print(f"✅ Processamento concluído: {len(df)} registros válidos")
        return df
        
    except Exception as e:
        print(f"❌ Erro ao carregar dados: {str(e)}")
        return None

def _fix_multiple_dots(x):
    """Corrige múltiplos pontos em números"""
    if not isinstance(x, str):
        return x
    parts = x.split('.')
    if len(parts) > 2:
        # Se tiver mais de um ponto, mantém apenas o último como decimal
        return parts[0] + '.' + ''.join(parts[1:])
    return x

def create_sample_data():
    """Cria dados de exemplo para demonstração"""
    dates = pd.date_range(start='2026-01-01', end='2026-01-07', freq='H')
    n_samples = min(100, len(dates))
    
    data = {
        'ID': range(1, n_samples + 1),
        'Title': [f'Solicitação {i}' for i in range(1, n_samples + 1)],
        'Status': ['Pago'] * n_samples,
        'Classificação': np.random.choice(['Despesa de Veiculo', 'Despesa de Viagem', 'Despesa Motorista'], n_samples),
        'Finalidade': np.random.choice(['Estacionamento', 'Uber', 'Passagem', 'Adto', 'Manutenção'], n_samples),
        'Valor': np.random.uniform(50, 1000, n_samples),
        'Nome Motorista': [f'Motorista {i}' for i in range(1, n_samples + 1)],
        'Solicitante': np.random.choice(['Solicitante A', 'Solicitante B', 'Solicitante C'], n_samples),
        'Criado': dates[:n_samples],
        'Empresa': np.random.choice(['Transmaroni', 'TKS', 'Outra'], n_samples),
        'Ordem_Servico': np.random.choice(['OS-001', 'OS-002', 'OS-003', 'Não Informada'], n_samples),
        'Descrição': [f'Descrição da solicitação {i}' for i in range(1, n_samples + 1)],
        'Placa': np.random.choice(['ABC1D23', 'XYZ4E56', 'DEF7G89'], n_samples)
    }
    
    return pd.DataFrame(data)

# ============================================================================
# LAYOUT DO DASHBOARD
# ============================================================================

# Inicializar com dados de exemplo
df = create_sample_data()

# Calcular estatísticas iniciais
total_gasto = df['Valor'].sum() if 'Valor' in df.columns else 0
total_solicitacoes = len(df)
media_valor = df['Valor'].mean() if 'Valor' in df.columns else 0
top_categoria = df['Classificação'].value_counts().index[0] if not df['Classificação'].empty else "N/A"
empresas_unicas = df['Empresa'].nunique()
ordens_servico = df['Ordem_Servico'].nunique()

# Layout principal
app.layout = dbc.Container([
    # Armazenamento de dados
    dcc.Store(id='data-store', data=df.to_dict('records')),
    dcc.Store(id='filtered-data-store'),
    
    # Upload de arquivo
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            '📤 Arraste ou ',
            html.A('Selecione um arquivo Excel')
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
    
    # Cabeçalho
    dbc.Row([
        dbc.Col([
            html.H1("📊 Dashboard de Custos Diários", 
                   className="text-primary mb-3"),
            html.P("Análise completa de solicitações de depósitos e despesas operacionais", 
                  className="text-muted lead")
        ], width=12)
    ], className="mb-4"),
    
    # Cards de resumo
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H4(f"R$ {total_gasto:,.2f}", 
                           className="card-title text-success"),
                    html.P("Total Gasto", className="card-text"),
                    html.Small(f"Média: R$ {media_valor:,.2f}", 
                             className="text-muted")
                ])
            ], className="shadow-sm h-100")
        ], md=3, sm=6),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H4(f"{total_solicitacoes:,}", 
                           className="card-title text-primary"),
                    html.P("Total de Solicitações", className="card-text"),
                    html.Small(f"Última: {df['Criado'].max().strftime('%d/%m/%Y') if 'Criado' in df.columns else 'N/A'}", 
                             className="text-muted")
                ])
            ], className="shadow-sm h-100")
        ], md=3, sm=6),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H4(f"{empresas_unicas}", 
                           className="card-title text-warning"),
                    html.P("Empresas", className="card-text"),
                    html.Small(f"Principal: {df['Empresa'].mode()[0] if not df['Empresa'].empty else 'N/A'}", 
                             className="text-muted")
                ])
            ], className="shadow-sm h-100")
        ], md=3, sm=6),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H4(f"{ordens_servico}", 
                           className="card-title text-info"),
                    html.P("Ordens de Serviço", className="card-text"),
                    html.Small(f"Mais comum: {df['Ordem_Servico'].mode()[0] if not df['Ordem_Servico'].empty and df['Ordem_Servico'].mode().any() else 'N/A'}", 
                             className="text-muted")
                ])
            ], className="shadow-sm h-100")
        ], md=3, sm=6)
    ], className="mb-4"),
    
    # Navegação por tabs
    dbc.Tabs([
        # Tab 1: Visão Geral
        dbc.Tab(label="🏠 Visão Geral", tab_id="tab-overview", children=[
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🔍 Filtros Avançados"),
                        dbc.CardBody([
                            dbc.Row([
                                dbc.Col([
                                    html.Label("Classificação:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-classificacao',
                                        options=[{'label': 'Todas', 'value': 'Todas'}] + 
                                                [{'label': cat, 'value': cat} for cat in sorted(df['Classificação'].unique())],
                                        value='Todas',
                                        multi=False,
                                        clearable=False,
                                        placeholder="Selecione a classificação..."
                                    )
                                ], md=3, sm=6),
                                
                                dbc.Col([
                                    html.Label("Empresa:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-empresa',
                                        options=[{'label': 'Todas', 'value': 'Todas'}] + 
                                                [{'label': emp, 'value': emp} for emp in sorted(df['Empresa'].unique())],
                                        value='Todas',
                                        multi=False,
                                        clearable=False,
                                        placeholder="Selecione a empresa..."
                                    )
                                ], md=3, sm=6),
                                
                                dbc.Col([
                                    html.Label("Finalidade:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-finalidade',
                                        options=[{'label': 'Todas', 'value': 'Todas'}] + 
                                                [{'label': fin, 'value': fin} for fin in sorted(df['Finalidade'].unique())],
                                        value='Todas',
                                        multi=False,
                                        clearable=False,
                                        placeholder="Selecione a finalidade..."
                                    )
                                ], md=3, sm=6),
                                
                                dbc.Col([
                                    html.Label("Status:", className="form-label"),
                                    dcc.Dropdown(
                                        id='filtro-status',
                                        options=[{'label': 'Todos', 'value': 'Todos'}] + 
                                                [{'label': status, 'value': status} for status in sorted(df['Status'].unique())],
                                        value='Todos',
                                        multi=False,
                                        clearable=False,
                                        placeholder="Selecione o status..."
                                    )
                                ], md=3, sm=6)
                            ], className="mb-3"),
                            
                            dbc.Row([
                                dbc.Col([
                                    html.Label("Período:", className="form-label"),
                                    dcc.DatePickerRange(
                                        id='filtro-data',
                                        min_date_allowed=df['Criado'].min().date() if 'Criado' in df.columns else datetime.now().date(),
                                        max_date_allowed=df['Criado'].max().date() if 'Criado' in df.columns else datetime.now().date(),
                                        start_date=df['Criado'].min().date() if 'Criado' in df.columns else datetime.now().date(),
                                        end_date=df['Criado'].max().date() if 'Criado' in df.columns else datetime.now().date(),
                                        display_format='DD/MM/YYYY',
                                        className="w-100"
                                    )
                                ], md=4),
                                
                                dbc.Col([
                                    html.Label("Valor Mínimo:", className="form-label"),
                                    dcc.Input(
                                        id='filtro-valor-min',
                                        type='number',
                                        placeholder='0',
                                        min=0,
                                        value=0,
                                        className="form-control"
                                    )
                                ], md=2),
                                
                                dbc.Col([
                                    html.Label("Valor Máximo:", className="form-label"),
                                    dcc.Input(
                                        id='filtro-valor-max',
                                        type='number',
                                        placeholder='10000',
                                        min=0,
                                        value=10000,
                                        className="form-control"
                                    )
                                ], md=2),
                                
                                dbc.Col([
                                    html.Label("Ordenar por:", className="form-label"),
                                    dcc.Dropdown(
                                        id='ordenar-por',
                                        options=[
                                            {'label': 'Data (Mais Recente)', 'value': 'data_desc'},
                                            {'label': 'Data (Mais Antiga)', 'value': 'data_asc'},
                                            {'label': 'Valor (Maior)', 'value': 'valor_desc'},
                                            {'label': 'Valor (Menor)', 'value': 'valor_asc'}
                                        ],
                                        value='data_desc',
                                        clearable=False,
                                        className="w-100"
                                    )
                                ], md=4)
                            ])
                        ])
                    ], className="mb-4")
                ], width=12)
            ]),
            
            # Gráficos principais
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📈 Evolução de Gastos Diários"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-evolucao-gastos')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🏢 Distribuição por Empresa"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-empresas')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12)
            ], className="mb-4"),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🏷️ Distribuição por Classificação"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-classificacao')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=4, md=6),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🎯 Top 10 Finalidades"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-finalidades')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=4, md=6),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("⏰ Distribuição por Hora"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-horas')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=4, md=12)
            ], className="mb-4")
        ]),
        
        # Tab 2: Análise Detalhada
        dbc.Tab(label="🔍 Análise Detalhada", tab_id="tab-detalhada", children=[
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            html.H5("📋 Tabela Detalhada de Solicitações"),
                            dbc.ButtonGroup([
                                dbc.Button("📥 Exportar CSV", id="btn-export-csv", color="success", size="sm"),
                                dbc.Button("📊 Exportar Excel", id="btn-export-excel", color="primary", size="sm"),
                                dbc.Button("🔄 Atualizar", id="btn-refresh", color="secondary", size="sm")
                            ], className="float-end")
                        ]),
                        dbc.CardBody([
                            dash_table.DataTable(
                                id='tabela-detalhada',
                                columns=[
                                    {"name": "ID", "id": "ID", "type": "numeric"},
                                    {"name": "Title", "id": "Title"},
                                    {"name": "Status", "id": "Status"},
                                    {"name": "Classificação", "id": "Classificação"},
                                    {"name": "Finalidade", "id": "Finalidade"},
                                    {"name": "Valor", "id": "Valor", "type": "numeric", 
                                     "format": {"specifier": "R$ ,.2f"}},
                                    {"name": "Empresa", "id": "Empresa"},
                                    {"name": "Ordem Serviço", "id": "Ordem_Servico"},
                                    {"name": "Motorista", "id": "Nome Motorista"},
                                    {"name": "Solicitante", "id": "Solicitante"},
                                    {"name": "Criado", "id": "Criado", "type": "datetime"},
                                    {"name": "Placa", "id": "Placa"}
                                ],
                                page_size=15,
                                page_current=0,
                                page_action='native',
                                style_table={
                                    'overflowX': 'auto',
                                    'maxHeight': '600px',
                                    'overflowY': 'auto'
                                },
                                style_cell={
                                    'textAlign': 'left',
                                    'padding': '10px',
                                    'fontSize': '12px',
                                    'fontFamily': 'Arial, sans-serif',
                                    'minWidth': '80px',
                                    'maxWidth': '200px',
                                    'whiteSpace': 'normal',
                                    'textOverflow': 'ellipsis'
                                },
                                style_header={
                                    'backgroundColor': 'rgb(230, 230, 230)',
                                    'fontWeight': 'bold',
                                    'textAlign': 'center'
                                },
                                style_data_conditional=[
                                    {
                                        'if': {'row_index': 'odd'},
                                        'backgroundColor': 'rgb(248, 248, 248)'
                                    },
                                    {
                                        'if': {'column_id': 'Valor'},
                                        'fontWeight': 'bold',
                                        'color': 'green'
                                    },
                                    {
                                        'if': {'filter_query': '{Status} = "Pago"'},
                                        'backgroundColor': 'rgba(144, 238, 144, 0.3)'
                                    }
                                ],
                                filter_action="native",
                                sort_action="native",
                                sort_mode="multi",
                                column_selectable="single",
                                row_selectable='multi',
                                selected_columns=[],
                                selected_rows=[],
                                tooltip_data=[],
                                tooltip_duration=None
                            )
                        ]),
                        dbc.CardFooter([
                            html.Div(id='table-info', className="text-muted small")
                        ])
                    ], className="shadow-sm")
                ], width=12)
            ], className="mb-4"),
            
            # Análises específicas
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🏢 Análise por Empresa"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-analise-empresa')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🔧 Análise por Ordem de Serviço"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-analise-os')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12)
            ], className="mb-4")
        ]),
        
        # Tab 3: Reunião Manutenção Corporativa
        dbc.Tab(label="🏢 Reunião Manutenção", tab_id="tab-manutencao", children=[
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📊 Relatório Corporativo - Foco em Manutenção"),
                        dbc.CardBody([
                            dbc.Row([
                                dbc.Col([
                                    html.H5("Filtros para Relatório de Manutenção", className="mb-3"),
                                    dbc.Row([
                                        dbc.Col([
                                            html.Label("Tipo de Manutenção:", className="form-label"),
                                            dcc.Dropdown(
                                                id='filtro-manutencao-tipo',
                                                options=[
                                                    {'label': 'Todas', 'value': 'Todas'},
                                                    {'label': 'Corretiva', 'value': 'Manutenção Corretiva'},
                                                    {'label': 'Preventiva', 'value': 'Manutenção Preventiva'},
                                                    {'label': 'Borracharia', 'value': 'Borracharia'},
                                                    {'label': 'Reformas', 'value': 'Reformas'},
                                                    {'label': 'Lavagem', 'value': 'Lavagem'}
                                                ],
                                                value='Todas',
                                                clearable=False
                                            )
                                        ], md=4),
                                        
                                        dbc.Col([
                                            html.Label("Empresa Responsável:", className="form-label"),
                                            dcc.Dropdown(
                                                id='filtro-manutencao-empresa',
                                                options=[{'label': 'Todas', 'value': 'Todas'}] + 
                                                        [{'label': emp, 'value': emp} for emp in sorted(df['Empresa'].unique())],
                                                value='Todas',
                                                clearable=False
                                            )
                                        ], md=4),
                                        
                                        dbc.Col([
                                            html.Label("Período Específico:", className="form-label"),
                                            dcc.DatePickerRange(
                                                id='filtro-manutencao-data',
                                                min_date_allowed=df['Criado'].min().date() if 'Criado' in df.columns else datetime.now().date(),
                                                max_date_allowed=df['Criado'].max().date() if 'Criado' in df.columns else datetime.now().date(),
                                                start_date=df['Criado'].min().date() if 'Criado' in df.columns else datetime.now().date(),
                                                end_date=df['Criado'].max().date() if 'Criado' in df.columns else datetime.now().date(),
                                                display_format='DD/MM/YYYY'
                                            )
                                        ], md=4)
                                    ])
                                ], width=12)
                            ])
                        ])
                    ], className="mb-4")
                ], width=12)
            ]),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("💰 Custos de Manutenção por Empresa"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-manutencao-empresa')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📈 Evolução dos Custos de Manutenção"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-manutencao-evolucao')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12)
            ], className="mb-4"),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader([
                            "📋 Tabela de Manutenções Corporativas",
                            dbc.Button("📥 Exportar Relatório", id="btn-export-manutencao", 
                                      color="primary", size="sm", className="float-end")
                        ]),
                        dbc.CardBody([
                            dash_table.DataTable(
                                id='tabela-manutencao',
                                columns=[
                                    {"name": "ID", "id": "ID"},
                                    {"name": "Title", "id": "Title"},
                                    {"name": "Classificação", "id": "Classificação"},
                                    {"name": "Finalidade", "id": "Finalidade"},
                                    {"name": "Valor", "id": "Valor", "type": "numeric", 
                                     "format": {"specifier": "R$ ,.2f"}},
                                    {"name": "Empresa", "id": "Empresa"},
                                    {"name": "Ordem Serviço", "id": "Ordem_Servico"},
                                    {"name": "Descrição", "id": "Descrição"},
                                    {"name": "Data", "id": "Criado"},
                                    {"name": "Placa", "id": "Placa"}
                                ],
                                page_size=10,
                                style_table={'overflowX': 'auto'},
                                style_cell={
                                    'textAlign': 'left',
                                    'padding': '8px',
                                    'fontSize': '11px',
                                    'minWidth': '100px'
                                },
                                style_header={
                                    'backgroundColor': 'rgb(240, 240, 240)',
                                    'fontWeight': 'bold'
                                },
                                filter_action="native",
                                sort_action="native",
                                sort_mode="multi"
                            )
                        ])
                    ], className="shadow-sm")
                ], width=12)
            ]),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📊 KPIs de Manutenção Corporativa"),
                        dbc.CardBody([
                            dbc.Row([
                                dbc.Col([
                                    html.H6("Custo Total de Manutenção", className="text-center"),
                                    html.H3(id='kpi-custo-total', className="text-center text-primary"),
                                    html.Small("Acumulado no período", className="text-muted text-center d-block")
                                ], md=3),
                                
                                dbc.Col([
                                    html.H6("Média por Ordem de Serviço", className="text-center"),
                                    html.H3(id='kpi-media-os', className="text-center text-success"),
                                    html.Small("Valor médio por OS", className="text-muted text-center d-block")
                                ], md=3),
                                
                                dbc.Col([
                                    html.H6("Empresa com Mais Custos", className="text-center"),
                                    html.H4(id='kpi-empresa-top', className="text-center text-warning"),
                                    html.Small("Maior custo de manutenção", className="text-muted text-center d-block")
                                ], md=3),
                                
                                dbc.Col([
                                    html.H6("OS Mais Frequente", className="text-center"),
                                    html.H4(id='kpi-os-top', className="text-center text-info"),
                                    html.Small("Ordem de serviço mais comum", className="text-muted text-center d-block")
                                ], md=3)
                            ])
                        ])
                    ], className="shadow-sm")
                ], width=12)
            ], className="mb-4")
        ]),
        
        # Tab 4: Análise por Motorista
        dbc.Tab(label="👤 Análise por Motorista", tab_id="tab-motorista", children=[
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📊 Top 10 Motoristas por Valor"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-top-motoristas')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12),
                
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("🏢 Distribuição por Empresa (Motoristas)"),
                        dbc.CardBody([
                            dcc.Graph(id='grafico-motorista-empresa')
                        ])
                    ], className="shadow-sm h-100")
                ], lg=6, md=12)
            ], className="mb-4"),
            
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardHeader("📋 Detalhamento por Motorista"),
                        dbc.CardBody([
                            dcc.Dropdown(
                                id='select-motorista',
                                options=[{'label': mot, 'value': mot} for mot in sorted(df['Nome Motorista'].unique())],
                                placeholder="Selecione um motorista...",
                                className="mb-3"
                            ),
                            html.Div(id='info-motorista')
                        ])
                    ], className="shadow-sm")
                ], width=12)
            ])
        ])
    ], id="tabs", active_tab="tab-overview"),
    
    # Componentes para download
    dcc.Download(id="download-csv"),
    dcc.Download(id="download-excel"),
    dcc.Download(id="download-manutencao"),
    
    # Footer
    dbc.Row([
        dbc.Col([
            html.Hr(),
            html.P([
                html.I(className="fas fa-info-circle me-2"),
                f"Dashboard de Custos Diários | ",
                html.Small(f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | "),
                html.Small(f"Registros: {len(df):,} | "),
                html.Small(f"Período: {df['Criado'].min().strftime('%d/%m/%Y') if 'Criado' in df.columns else 'N/A'} a {df['Criado'].max().strftime('%d/%m/%Y') if 'Criado' in df.columns else 'N/A'}")
            ], className="text-center text-muted mt-3")
        ], width=12)
    ])
], fluid=True, className="p-3")

# ============================================================================
# CALLBACKS PRINCIPAIS
# ============================================================================

# Callback para carregar dados do upload
@app.callback(
    [Output('data-store', 'data'),
     Output('filtered-data-store', 'data')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_data(contents, filename):
    """Atualiza os dados quando um novo arquivo é carregado"""
    if contents is not None:
        try:
            df = load_and_process_data(uploaded_file=contents)
            if df is not None:
                return df.to_dict('records'), df.to_dict('records')
        except Exception as e:
            print(f"Erro no upload: {str(e)}")
    
    # Retornar dados atuais se não houver upload
    return dash.no_update, dash.no_update

# Callback principal para atualizar todos os gráficos
@app.callback(
    [Output('grafico-evolucao-gastos', 'figure'),
     Output('grafico-empresas', 'figure'),
     Output('grafico-classificacao', 'figure'),
     Output('grafico-finalidades', 'figure'),
     Output('grafico-horas', 'figure'),
     Output('tabela-detalhada', 'data'),
     Output('table-info', 'children'),
     Output('grafico-analise-empresa', 'figure'),
     Output('grafico-analise-os', 'figure'),
     Output('filtered-data-store', 'data', allow_duplicate=True)],
    [Input('filtro-classificacao', 'value'),
     Input('filtro-empresa', 'value'),
     Input('filtro-finalidade', 'value'),
     Input('filtro-status', 'value'),
     Input('filtro-data', 'start_date'),
     Input('filtro-data', 'end_date'),
     Input('filtro-valor-min', 'value'),
     Input('filtro-valor-max', 'value'),
     Input('ordenar-por', 'value'),
     Input('data-store', 'data')],
    prevent_initial_call=True
)
def update_all_components(classificacao, empresa, finalidade, status, start_date, end_date, 
                         valor_min, valor_max, ordenar_por, data_store):
    """Atualiza todos os componentes com base nos filtros"""
    
    # Converter dados do store para DataFrame
    df = pd.DataFrame(data_store)
    
    if df.empty:
        empty_fig = go.Figure()
        empty_fig.add_annotation(text="Nenhum dado disponível", showarrow=False)
        empty_data = []
        table_info = "Nenhum registro encontrado"
        return [empty_fig] * 5 + [empty_data, table_info, empty_fig, empty_fig, dash.no_update]
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    # Converter datas
    for col in ['Criado', 'Modificado']:
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_datetime(df_filtrado[col])
    
    # Filtro de classificação
    if classificacao != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Classificação'] == classificacao]
    
    # Filtro de empresa
    if empresa != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Empresa'] == empresa]
    
    # Filtro de finalidade
    if finalidade != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Finalidade'] == finalidade]
    
    # Filtro de status
    if status != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Status'] == status]
    
    # Filtro de data
    if start_date and end_date:
        start_date = pd.to_datetime(start_date).date()
        end_date = pd.to_datetime(end_date).date()
        if 'Criado' in df_filtrado.columns:
            df_filtrado['Data_Criacao'] = pd.to_datetime(df_filtrado['Criado']).dt.date
            df_filtrado = df_filtrado[
                (df_filtrado['Data_Criacao'] >= start_date) & 
                (df_filtrado['Data_Criacao'] <= end_date)
            ]
    
    # Filtro de valor
    if valor_min is not None:
        df_filtrado = df_filtrado[df_filtrado['Valor'] >= float(valor_min)]
    
    if valor_max is not None:
        df_filtrado = df_filtrado[df_filtrado['Valor'] <= float(valor_max)]
    
    # Ordenação
    if ordenar_por == 'data_desc':
        df_filtrado = df_filtrado.sort_values('Criado', ascending=False)
    elif ordenar_por == 'data_asc':
        df_filtrado = df_filtrado.sort_values('Criado', ascending=True)
    elif ordenar_por == 'valor_desc':
        df_filtrado = df_filtrado.sort_values('Valor', ascending=False)
    elif ordenar_por == 'valor_asc':
        df_filtrado = df_filtrado.sort_values('Valor', ascending=True)
    
    # ========================================================================
    # 1. Gráfico de Evolução de Gastos
    # ========================================================================
    if not df_filtrado.empty and 'Criado' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        df_filtrado['Data'] = pd.to_datetime(df_filtrado['Criado']).dt.date
        daily_data = df_filtrado.groupby('Data')['Valor'].sum().reset_index()
        
        fig_evolucao = go.Figure()
        fig_evolucao.add_trace(go.Scatter(
            x=daily_data['Data'],
            y=daily_data['Valor'],
            mode='lines+markers',
            name='Gasto Diário',
            line=dict(color='#4361ee', width=3),
            marker=dict(size=8, color='#4361ee'),
            fill='tozeroy',
            fillcolor='rgba(67, 97, 238, 0.1)'
        ))
        
        # Adicionar média móvel de 7 dias
        if len(daily_data) >= 7:
            daily_data['Media_Movel'] = daily_data['Valor'].rolling(window=7, min_periods=1).mean()
            fig_evolucao.add_trace(go.Scatter(
                x=daily_data['Data'],
                y=daily_data['Media_Movel'],
                mode='lines',
                name='Média Móvel (7 dias)',
                line=dict(color='#f72585', width=2, dash='dash')
            ))
        
        fig_evolucao.update_layout(
            title='Evolução de Gastos Diários',
            xaxis_title='Data',
            yaxis_title='Valor (R$)',
            hovermode='x unified',
            template='plotly_white',
            height=400,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
    else:
        fig_evolucao = create_empty_figure("Sem dados para o período selecionado")
    
    # ========================================================================
    # 2. Gráfico de Distribuição por Empresa
    # ========================================================================
    if not df_filtrado.empty and 'Empresa' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        empresa_data = df_filtrado.groupby('Empresa')['Valor'].agg(['sum', 'count']).reset_index()
        empresa_data.columns = ['Empresa', 'Valor_Total', 'Quantidade']
        empresa_data = empresa_data.sort_values('Valor_Total', ascending=False)
        
        fig_empresas = go.Figure()
        fig_empresas.add_trace(go.Bar(
            x=empresa_data['Empresa'],
            y=empresa_data['Valor_Total'],
            name='Valor Total',
            marker_color='#4cc9f0',
            text=empresa_data['Valor_Total'].apply(lambda x: f'R$ {x:,.0f}'),
            textposition='auto'
        ))
        
        fig_empresas.update_layout(
            title='Distribuição por Empresa (Valor Total)',
            xaxis_title='Empresa',
            yaxis_title='Valor Total (R$)',
            template='plotly_white',
            height=400,
            xaxis_tickangle=-45
        )
    else:
        fig_empresas = create_empty_figure("Sem dados de empresa")
    
    # ========================================================================
    # 3. Gráfico de Classificação (Pizza)
    # ========================================================================
    if not df_filtrado.empty and 'Classificação' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        classificacao_data = df_filtrado.groupby('Classificação')['Valor'].sum().reset_index()
        
        fig_classificacao = px.pie(
            classificacao_data,
            values='Valor',
            names='Classificação',
            title='Distribuição por Classificação',
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig_classificacao.update_traces(
            textposition='inside',
            textinfo='percent+label',
            hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}",
            pull=[0.1 if i == 0 else 0 for i in range(len(classificacao_data))]
        )
        
        fig_classificacao.update_layout(
            height=400,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5
            )
        )
    else:
        fig_classificacao = create_empty_figure("Sem dados de classificação")
    
    # ========================================================================
    # 4. Gráfico de Top Finalidades
    # ========================================================================
    if not df_filtrado.empty and 'Finalidade' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        finalidade_data = df_filtrado.groupby('Finalidade')['Valor'].sum().reset_index()
        finalidade_data = finalidade_data.sort_values('Valor', ascending=True).tail(10)
        
        fig_finalidades = px.bar(
            finalidade_data,
            x='Valor',
            y='Finalidade',
            orientation='h',
            title='Top 10 Finalidades por Valor',
            color='Valor',
            color_continuous_scale='viridis',
            text='Valor'
        )
        
        fig_finalidades.update_traces(
            texttemplate='R$ %{x:,.0f}',
            textposition='outside',
            hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>"
        )
        
        fig_finalidades.update_layout(
            height=400,
            xaxis_title='Valor Total (R$)',
            yaxis_title='Finalidade',
            coloraxis_showscale=False,
            yaxis={'categoryorder': 'total ascending'}
        )
    else:
        fig_finalidades = create_empty_figure("Sem dados de finalidade")
    
    # ========================================================================
    # 5. Gráfico de Distribuição por Hora
    # ========================================================================
    if not df_filtrado.empty and 'Criado' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        df_filtrado['Hora'] = pd.to_datetime(df_filtrado['Criado']).dt.hour
        hora_data = df_filtrado.groupby('Hora')['Valor'].agg(['sum', 'count']).reset_index()
        hora_data.columns = ['Hora', 'Valor_Total', 'Quantidade']
        
        fig_horas = go.Figure()
        
        fig_horas.add_trace(go.Bar(
            x=hora_data['Hora'],
            y=hora_data['Valor_Total'],
            name='Valor Total',
            marker_color='#7209b7',
            yaxis='y'
        ))
        
        fig_horas.add_trace(go.Scatter(
            x=hora_data['Hora'],
            y=hora_data['Quantidade'],
            name='Quantidade',
            line=dict(color='#f72585', width=3),
            yaxis='y2',
            mode='lines+markers'
        ))
        
        fig_horas.update_layout(
            title='Distribuição por Hora do Dia',
            xaxis=dict(
                title='Hora do Dia',
                tickmode='linear',
                tick0=0,
                dtick=1,
                range=[-0.5, 23.5]
            ),
            yaxis=dict(
                title='Valor Total (R$)',
                side='left'
            ),
            yaxis2=dict(
                title='Quantidade de Solicitações',
                overlaying='y',
                side='right'
            ),
            height=400,
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
    else:
        fig_horas = create_empty_figure("Sem dados de hora")
    
    # ========================================================================
    # 6. Tabela Detalhada
    # ========================================================================
    if not df_filtrado.empty:
        # Preparar dados para tabela
        tabela_data = df_filtrado.copy()
        
        # Selecionar colunas para exibir
        display_columns = ['ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
                          'Valor', 'Empresa', 'Ordem_Servico', 'Nome Motorista', 
                          'Solicitante', 'Criado', 'Placa']
        
        # Manter apenas colunas que existem
        existing_columns = [col for col in display_columns if col in tabela_data.columns]
        tabela_data = tabela_data[existing_columns]
        
        # Formatar data
        if 'Criado' in tabela_data.columns:
            tabela_data['Criado'] = tabela_data['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Converter para dict
        tabela_dict = tabela_data.to_dict('records')
        
        # Info da tabela
        table_info = f"Mostrando {len(tabela_data)} de {len(df)} registros | Total filtrado: R$ {df_filtrado['Valor'].sum():,.2f}"
    else:
        tabela_dict = []
        table_info = "Nenhum registro encontrado com os filtros aplicados"
    
    # ========================================================================
    # 7. Gráfico de Análise por Empresa (Detalhada)
    # ========================================================================
    if not df_filtrado.empty and 'Empresa' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        empresa_detailed = df_filtrado.groupby('Empresa').agg({
            'Valor': ['sum', 'mean', 'count'],
            'ID': 'nunique'
        }).round(2).reset_index()
        
        empresa_detailed.columns = ['Empresa', 'Total', 'Média', 'Quantidade', 'Registros_Únicos']
        empresa_detailed = empresa_detailed.sort_values('Total', ascending=False).head(10)
        
        fig_analise_empresa = go.Figure()
        
        fig_analise_empresa.add_trace(go.Bar(
            x=empresa_detailed['Empresa'],
            y=empresa_detailed['Total'],
            name='Total (R$)',
            marker_color='#4361ee',
            yaxis='y'
        ))
        
        fig_analise_empresa.add_trace(go.Scatter(
            x=empresa_detailed['Empresa'],
            y=empresa_detailed['Média'],
            name='Média (R$)',
            line=dict(color='#f72585', width=3),
            yaxis='y2',
            mode='lines+markers'
        ))
        
        fig_analise_empresa.update_layout(
            title='Análise Detalhada por Empresa',
            xaxis_title='Empresa',
            yaxis=dict(
                title='Valor Total (R$)',
                side='left'
            ),
            yaxis2=dict(
                title='Valor Médio (R$)',
                overlaying='y',
                side='right'
            ),
            height=400,
            hovermode='x unified',
            xaxis_tickangle=-45
        )
    else:
        fig_analise_empresa = create_empty_figure("Sem dados para análise de empresa")
    
    # ========================================================================
    # 8. Gráfico de Análise por Ordem de Serviço
    # ========================================================================
    if not df_filtrado.empty and 'Ordem_Servico' in df_filtrado.columns and 'Valor' in df_filtrado.columns:
        # Filtrar ordens de serviço válidas (não vazias ou "Não Informada")
        os_data = df_filtrado[~df_filtrado['Ordem_Servico'].isin(['', 'Não Informada', 'nan', 'NaN'])]
        
        if not os_data.empty:
            os_analysis = os_data.groupby('Ordem_Servico').agg({
                'Valor': ['sum', 'count'],
                'Empresa': lambda x: x.mode().iloc[0] if not x.mode().empty else 'N/A'
            }).round(2).reset_index()
            
            os_analysis.columns = ['Ordem_Servico', 'Total', 'Quantidade', 'Empresa_Principal']
            os_analysis = os_analysis.sort_values('Total', ascending=False).head(15)
            
            fig_analise_os = go.Figure()
            
            colors = ['#4361ee' if total > os_analysis['Total'].median() else '#4cc9f0' 
                     for total in os_analysis['Total']]
            
            fig_analise_os.add_trace(go.Bar(
                x=os_analysis['Ordem_Servico'],
                y=os_analysis['Total'],
                name='Valor Total',
                marker_color=colors,
                text=os_analysis['Total'].apply(lambda x: f'R$ {x:,.0f}'),
                textposition='auto'
            ))
            
            fig_analise_os.update_layout(
                title='Top 15 Ordens de Serviço por Valor',
                xaxis_title='Ordem de Serviço',
                yaxis_title='Valor Total (R$)',
                height=400,
                xaxis_tickangle=-45
            )
        else:
            fig_analise_os = create_empty_figure("Nenhuma ordem de serviço válida encontrada")
    else:
        fig_analise_os = create_empty_figure("Sem dados de ordem de serviço")
    
    return [
        fig_evolucao, fig_empresas, fig_classificacao, fig_finalidades, fig_horas,
        tabela_dict, table_info, fig_analise_empresa, fig_analise_os,
        df_filtrado.to_dict('records')
    ]

# ============================================================================
# CALLBACKS PARA REUNIÃO DE MANUTENÇÃO
# ============================================================================

@app.callback(
    [Output('grafico-manutencao-empresa', 'figure'),
     Output('grafico-manutencao-evolucao', 'figure'),
     Output('tabela-manutencao', 'data'),
     Output('kpi-custo-total', 'children'),
     Output('kpi-media-os', 'children'),
     Output('kpi-empresa-top', 'children'),
     Output('kpi-os-top', 'children')],
    [Input('filtro-manutencao-tipo', 'value'),
     Input('filtro-manutencao-empresa', 'value'),
     Input('filtro-manutencao-data', 'start_date'),
     Input('filtro-manutencao-data', 'end_date'),
     Input('filtered-data-store', 'data')]
)
def update_manutencao_analysis(tipo_manutencao, empresa, start_date, end_date, filtered_data):
    """Atualiza a análise de manutenção corporativa"""
    
    if filtered_data is None or len(filtered_data) == 0:
        empty_fig = create_empty_figure("Sem dados disponíveis")
        empty_data = []
        kpis = ["R$ 0,00", "R$ 0,00", "N/A", "N/A"]
        return [empty_fig, empty_fig, empty_data] + kpis
    
    df = pd.DataFrame(filtered_data)
    
    # Filtrar por tipo de manutenção
    if tipo_manutencao != 'Todas':
        manutencao_keywords = []
        if tipo_manutencao == 'Manutenção Corretiva':
            manutencao_keywords = ['Manutenção Corretiva', 'corretiva', 'reparo', 'conserto']
        elif tipo_manutencao == 'Manutenção Preventiva':
            manutencao_keywords = ['Manutenção Preventiva', 'preventiva', 'insulfime']
        elif tipo_manutencao == 'Borracharia':
            manutencao_keywords = ['Borracharia', 'pneu', 'borracha']
        elif tipo_manutencao == 'Reformas':
            manutencao_keywords = ['Reformas', 'reforma', 'insulfilm']
        elif tipo_manutencao == 'Lavagem':
            manutencao_keywords = ['Lavagem', 'lavador', 'lavar']
        
        # Filtrar por título, classificação ou finalidade
        mask = False
        for keyword in manutencao_keywords:
            mask = mask | (
                df['Title'].str.contains(keyword, case=False, na=False) |
                df['Classificação'].str.contains(keyword, case=False, na=False) |
                df['Finalidade'].str.contains(keyword, case=False, na=False)
            )
        df = df[mask]
    
    # Filtrar por empresa
    if empresa != 'Todas':
        df = df[df['Empresa'] == empresa]
    
    # Filtrar por data
    if start_date and end_date:
        if 'Criado' in df.columns:
            start_date = pd.to_datetime(start_date).date()
            end_date = pd.to_datetime(end_date).date()
            df['Data_Criacao'] = pd.to_datetime(df['Criado']).dt.date
            df = df[(df['Data_Criacao'] >= start_date) & (df['Data_Criacao'] <= end_date)]
    
    # ========================================================================
    # 1. Gráfico de Custos de Manutenção por Empresa
    # ========================================================================
    if not df.empty and 'Empresa' in df.columns and 'Valor' in df.columns:
        empresa_manutencao = df.groupby('Empresa')['Valor'].agg(['sum', 'count']).reset_index()
        empresa_manutencao.columns = ['Empresa', 'Total', 'Quantidade']
        empresa_manutencao = empresa_manutencao.sort_values('Total', ascending=True)
        
        fig_manutencao_empresa = px.bar(
            empresa_manutencao,
            x='Total',
            y='Empresa',
            orientation='h',
            title='Custos de Manutenção por Empresa',
            color='Total',
            color_continuous_scale='viridis',
            text='Total'
        )
        
        fig_manutencao_empresa.update_traces(
            texttemplate='R$ %{x:,.0f}',
            textposition='outside'
        )
        
        fig_manutencao_empresa.update_layout(
            height=400,
            xaxis_title='Custo Total (R$)',
            yaxis_title='Empresa',
            coloraxis_showscale=False
        )
    else:
        fig_manutencao_empresa = create_empty_figure("Sem dados de manutenção por empresa")
    
    # ========================================================================
    # 2. Gráfico de Evolução dos Custos de Manutenção
    # ========================================================================
    if not df.empty and 'Criado' in df.columns and 'Valor' in df.columns:
        df['Data'] = pd.to_datetime(df['Criado']).dt.date
        evolucao_manutencao = df.groupby('Data')['Valor'].sum().reset_index()
        
        fig_manutencao_evolucao = go.Figure()
        
        fig_manutencao_evolucao.add_trace(go.Scatter(
            x=evolucao_manutencao['Data'],
            y=evolucao_manutencao['Valor'],
            mode='lines+markers',
            name='Custo Diário',
            line=dict(color='#f72585', width=3),
            marker=dict(size=8),
            fill='tozeroy',
            fillcolor='rgba(247, 37, 133, 0.1)'
        ))
        
        fig_manutencao_evolucao.update_layout(
            title='Evolução dos Custos de Manutenção',
            xaxis_title='Data',
            yaxis_title='Custo (R$)',
            height=400,
            template='plotly_white'
        )
    else:
        fig_manutencao_evolucao = create_empty_figure("Sem dados para evolução")
    
    # ========================================================================
    # 3. Tabela de Manutenções
    # ========================================================================
    if not df.empty:
        # Preparar dados para tabela
        tabela_manutencao = df.copy()
        
        # Selecionar e renomear colunas
        tabela_colunas = {
            'ID': 'ID',
            'Title': 'Title',
            'Classificação': 'Classificação',
            'Finalidade': 'Finalidade',
            'Valor': 'Valor',
            'Empresa': 'Empresa',
            'Ordem_Servico': 'Ordem_Servico',
            'Descrição': 'Descrição',
            'Criado': 'Criado',
            'Placa': 'Placa'
        }
        
        # Manter apenas colunas existentes
        existing_cols = {k: v for k, v in tabela_colunas.items() if k in tabela_manutencao.columns}
        tabela_manutencao = tabela_manutencao[list(existing_cols.keys())]
        tabela_manutencao = tabela_manutencao.rename(columns=existing_cols)
        
        # Formatar data
        if 'Criado' in tabela_manutencao.columns:
            tabela_manutencao['Criado'] = pd.to_datetime(tabela_manutencao['Criado']).dt.strftime('%d/%m/%Y %H:%M')
        
        tabela_manutencao_data = tabela_manutencao.to_dict('records')
    else:
        tabela_manutencao_data = []
    
    # ========================================================================
    # 4. KPIs de Manutenção
    # ========================================================================
    if not df.empty:
        # Custo Total
        custo_total = df['Valor'].sum()
        kpi_custo_total = f"R$ {custo_total:,.2f}"
        
        # Média por Ordem de Serviço
        os_validas = df[~df['Ordem_Servico'].isin(['', 'Não Informada', 'nan', 'NaN'])]
        if not os_validas.empty:
            media_os = os_validas.groupby('Ordem_Servico')['Valor'].sum().mean()
            kpi_media_os = f"R$ {media_os:,.2f}"
        else:
            kpi_media_os = "R$ 0,00"
        
        # Empresa com Mais Custos
        if 'Empresa' in df.columns:
            empresa_top = df.groupby('Empresa')['Valor'].sum().idxmax()
            kpi_empresa_top = empresa_top[:15] + "..." if len(empresa_top) > 15 else empresa_top
        else:
            kpi_empresa_top = "N/A"
        
        # OS Mais Frequente
        if 'Ordem_Servico' in df.columns and not os_validas.empty:
            os_top = os_validas['Ordem_Servico'].mode()
            kpi_os_top = os_top.iloc[0][:15] + "..." if len(os_top.iloc[0]) > 15 else os_top.iloc[0]
        else:
            kpi_os_top = "N/A"
    else:
        kpi_custo_total = "R$ 0,00"
        kpi_media_os = "R$ 0,00"
        kpi_empresa_top = "N/A"
        kpi_os_top = "N/A"
    
    return [
        fig_manutencao_empresa,
        fig_manutencao_evolucao,
        tabela_manutencao_data,
        kpi_custo_total,
        kpi_media_os,
        kpi_empresa_top,
        kpi_os_top
    ]

# ============================================================================
# CALLBACKS PARA EXPORTAÇÃO
# ============================================================================

@app.callback(
    Output("download-excel", "data"),
    [Input("btn-export-excel", "n_clicks")],
    [State("filtered-data-store", "data")],
    prevent_initial_call=True
)
def export_to_excel(n_clicks, filtered_data):
    """Exporta dados filtrados para Excel"""
    if n_clicks is None or filtered_data is None:
        return dash.no_update
    
    df = pd.DataFrame(filtered_data)
    
    # Criar múltiplas planilhas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Planilha principal
        df.to_excel(writer, sheet_name='Dados_Filtrados', index=False)
        
        # Planilha de resumo por classificação
        if 'Classificação' in df.columns and 'Valor' in df.columns:
            resumo_class = df.groupby('Classificação').agg({
                'Valor': ['sum', 'mean', 'count'],
                'ID': 'nunique'
            }).round(2)
            resumo_class.columns = ['Total', 'Média', 'Quantidade', 'Registros_Únicos']
            resumo_class.to_excel(writer, sheet_name='Resumo_Classificação')
        
        # Planilha de resumo por empresa
        if 'Empresa' in df.columns and 'Valor' in df.columns:
            resumo_empresa = df.groupby('Empresa').agg({
                'Valor': ['sum', 'mean', 'count'],
                'Ordem_Servico': lambda x: x.mode().iloc[0] if not x.mode().empty else 'N/A'
            }).round(2)
            resumo_empresa.columns = ['Total', 'Média', 'Quantidade', 'OS_Mais_Comum']
            resumo_empresa.to_excel(writer, sheet_name='Resumo_Empresa')
        
        # Planilha de ordens de serviço
        if 'Ordem_Servico' in df.columns:
            os_data = df[~df['Ordem_Servico'].isin(['', 'Não Informada', 'nan', 'NaN'])]
            if not os_data.empty:
                resumo_os = os_data.groupby('Ordem_Servico').agg({
                    'Valor': ['sum', 'count'],
                    'Empresa': lambda x: x.mode().iloc[0] if not x.mode().empty else 'N/A',
                    'Classificação': lambda x: x.mode().iloc[0] if not x.mode().empty else 'N/A'
                }).round(2)
                resumo_os.columns = ['Total', 'Quantidade', 'Empresa_Principal', 'Classificação_Principal']
                resumo_os.to_excel(writer, sheet_name='Resumo_Ordens_Serviço')
    
    output.seek(0)
    
    return dcc.send_bytes(
        output.read(),
        filename=f"relatorio_custos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

@app.callback(
    Output("download-csv", "data"),
    [Input("btn-export-csv", "n_clicks")],
    [State("filtered-data-store", "data")],
    prevent_initial_call=True
)
def export_to_csv(n_clicks, filtered_data):
    """Exporta dados filtrados para CSV"""
    if n_clicks is None or filtered_data is None:
        return dash.no_update
    
    df = pd.DataFrame(filtered_data)
    
    # Converter para CSV
    csv_string = df.to_csv(index=False, encoding='utf-8-sig')
    
    return dcc.send_string(
        csv_string,
        filename=f"dados_custos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    )

@app.callback(
    Output("download-manutencao", "data"),
    [Input("btn-export-manutencao", "n_clicks")],
    [State("tabela-manutencao", "data")],
    prevent_initial_call=True
)
def export_manutencao_report(n_clicks, tabela_data):
    """Exporta relatório de manutenção para Excel"""
    if n_clicks is None or not tabela_data:
        return dash.no_update
    
    df = pd.DataFrame(tabela_data)
    
    # Criar Excel com formatação
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Planilha de dados
        df.to_excel(writer, sheet_name='Manutenções', index=False)
        
        # Adicionar planilha de resumo
        resumo = pd.DataFrame({
            'Métrica': ['Total de Registros', 'Custo Total', 'Custo Médio', 'Data Início', 'Data Fim'],
            'Valor': [
                len(df),
                f"R$ {df['Valor'].sum():,.2f}" if 'Valor' in df.columns else 'N/A',
                f"R$ {df['Valor'].mean():,.2f}" if 'Valor' in df.columns else 'N/A',
                df['Criado'].min() if 'Criado' in df.columns else 'N/A',
                df['Criado'].max() if 'Criado' in df.columns else 'N/A'
            ]
        })
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
    
    output.seek(0)
    
    return dcc.send_bytes(
        output.read(),
        filename=f"relatorio_manutencao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

# ============================================================================
# CALLBACKS ADICIONAIS
# ============================================================================

@app.callback(
    [Output('grafico-top-motoristas', 'figure'),
     Output('grafico-motorista-empresa', 'figure'),
     Output('info-motorista', 'children')],
    [Input('select-motorista', 'value'),
     Input('filtered-data-store', 'data')]
)
def update_motorista_analysis(motorista_selecionado, filtered_data):
    """Atualiza análises relacionadas a motoristas"""
    
    if filtered_data is None or len(filtered_data) == 0:
        empty_fig = create_empty_figure("Sem dados disponíveis")
        return empty_fig, empty_fig, "Selecione um motorista para ver detalhes"
    
    df = pd.DataFrame(filtered_data)
    
    # ========================================================================
    # 1. Gráfico de Top 10 Motoristas
    # ========================================================================
    if not df.empty and 'Nome Motorista' in df.columns and 'Valor' in df.columns:
        top_motoristas = df.groupby('Nome Motorista')['Valor'].agg(['sum', 'count']).reset_index()
        top_motoristas.columns = ['Motorista', 'Total', 'Quantidade']
        top_motoristas = top_motoristas.sort_values('Total', ascending=False).head(10)
        
        fig_top_motoristas = go.Figure()
        
        fig_top_motoristas.add_trace(go.Bar(
            x=top_motoristas['Motorista'],
            y=top_motoristas['Total'],
            name='Valor Total',
            marker_color='#4361ee',
            text=top_motoristas['Total'].apply(lambda x: f'R$ {x:,.0f}'),
            textposition='auto'
        ))
        
        fig_top_motoristas.update_layout(
            title='Top 10 Motoristas por Valor',
            xaxis_title='Motorista',
            yaxis_title='Valor Total (R$)',
            height=400,
            xaxis_tickangle=-45
        )
    else:
        fig_top_motoristas = create_empty_figure("Sem dados de motoristas")
    
    # ========================================================================
    # 2. Gráfico de Distribuição por Empresa (Motoristas)
    # ========================================================================
    if not df.empty and 'Empresa' in df.columns and 'Nome Motorista' in df.columns:
        motorista_empresa = df.groupby(['Empresa', 'Nome Motorista']).size().reset_index(name='Quantidade')
        empresa_dist = motorista_empresa.groupby('Empresa')['Quantidade'].sum().reset_index()
        
        fig_motorista_empresa = px.pie(
            empresa_dist,
            values='Quantidade',
            names='Empresa',
            title='Distribuição de Motoristas por Empresa',
            hole=0.3,
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        
        fig_motorista_empresa.update_traces(
            textposition='inside',
            textinfo='percent+label'
        )
        
        fig_motorista_empresa.update_layout(height=400)
    else:
        fig_motorista_empresa = create_empty_figure("Sem dados para distribuição")
    
    # ========================================================================
    # 3. Informações Detalhadas do Motorista Selecionado
    # ========================================================================
    if motorista_selecionado:
        motorista_data = df[df['Nome Motorista'] == motorista_selecionado]
        
        if not motorista_data.empty:
            total_motorista = motorista_data['Valor'].sum()
            qtd_solicitacoes = len(motorista_data)
            primeira_solicitacao = motorista_data['Criado'].min()
            ultima_solicitacao = motorista_data['Criado'].max()
            empresa_principal = motorista_data['Empresa'].mode().iloc[0] if not motorista_data['Empresa'].mode().empty else 'N/A'
            
            # Top finalidades
            top_finalidades = motorista_data['Finalidade'].value_counts().head(3)
            finalidades_str = ', '.join([f"{k} ({v})" for k, v in top_finalidades.items()])
            
            info_card = dbc.Card([
                dbc.CardHeader(f"📋 Informações do Motorista: {motorista_selecionado}"),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            html.H6("Total Gasto:", className="text-muted"),
                            html.H4(f"R$ {total_motorista:,.2f}", className="text-success")
                        ], md=3),
                        
                        dbc.Col([
                            html.H6("Solicitações:", className="text-muted"),
                            html.H4(f"{qtd_solicitacoes}", className="text-primary")
                        ], md=3),
                        
                        dbc.Col([
                            html.H6("Empresa Principal:", className="text-muted"),
                            html.H4(empresa_principal, className="text-warning")
                        ], md=3),
                        
                        dbc.Col([
                            html.H6("Média por Solicitação:", className="text-muted"),
                            html.H4(f"R$ {total_motorista/qtd_solicitacoes:,.2f}", className="text-info")
                        ], md=3)
                    ]),
                    
                    html.Hr(),
                    
                    dbc.Row([
                        dbc.Col([
                            html.H6("Período de Atuação:", className="text-muted"),
                            html.P(f"{primeira_solicitacao.strftime('%d/%m/%Y')} a {ultima_solicitacao.strftime('%d/%m/%Y')}")
                        ], md=6),
                        
                        dbc.Col([
                            html.H6("Finalidades Mais Comuns:", className="text-muted"),
                            html.P(finalidades_str)
                        ], md=6)
                    ])
                ])
            ])
        else:
            info_card = dbc.Alert(
                f"Nenhum dado encontrado para o motorista {motorista_selecionado}",
                color="warning"
            )
    else:
        info_card = html.P("Selecione um motorista na lista acima para ver informações detalhadas.", 
                          className="text-muted")
    
    return fig_top_motoristas, fig_motorista_empresa, info_card

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def create_empty_figure(message="Sem dados disponíveis"):
    """Cria uma figura vazia com uma mensagem"""
    fig = go.Figure()
    fig.add_annotation(
        text=message,
        xref="paper", yref="paper",
        x=0.5, y=0.5,
        showarrow=False,
        font=dict(size=16, color="gray")
    )
    fig.update_layout(
        plot_bgcolor='white',
        height=400
    )
    return fig

# ============================================================================
# EXECUÇÃO DA APLICAÇÃO
# ============================================================================

if __name__ == '__main__':
    print("=" * 60)
    print("🚀 DASHBOARD DE CUSTOS DIÁRIOS")
    print("=" * 60)
    print("\n📊 Iniciando aplicação...")
    print(f"📁 Dados carregados: {len(df)} registros")
    print(f"💰 Total gasto: R$ {total_gasto:,.2f}")
    print(f"🏢 Empresas: {empresas_unicas}")
    print(f"🔧 Ordens de Serviço: {ordens_servico}")
    print("\n🌐 Acesse: http://localhost:8050")
    print("=" * 60)
    
    app.run_server(
        debug=True,
        port=8050,
        host='0.0.0.0',
        dev_tools_ui=True,
        dev_tools_hot_reload=True
    )
