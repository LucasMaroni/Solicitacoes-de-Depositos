import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import locale
import io
import numpy as np
import re
import os

# Configurar locale para português brasileiro
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

st.set_page_config(
    page_title="Dashboard de Custos", 
    layout="wide",
    page_icon="💰"
)

# Função para formatar números no formato brasileiro (simplificada)
def formatar_brasileiro(valor, decimais=2):
    """
    Formata um número no formato brasileiro: 1.234,56
    """
    if pd.isna(valor) or valor is None:
        return "0,00"
    
    try:
        # Converter para float
        valor_float = float(valor)
        
        # Formatar usando locale se disponível
        try:
            return locale.format_string(f"%.{decimais}f", valor_float, grouping=True)
        except:
            # Fallback manual
            valor_str = f"{valor_float:,.{decimais}f}"
            valor_str = valor_str.replace(",", "X").replace(".", ",").replace("X", ".")
            return valor_str
    except:
        return str(valor)

@st.cache_data(ttl=3600)  # Cache por 1 hora
def load_data(caminho_arquivo):
    """Carrega e processa os dados do arquivo Excel"""
    
    # Verificar se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        st.error(f"❌ Arquivo não encontrado: {caminho_arquivo}")
        return pd.DataFrame()
    
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_arquivo, engine='openpyxl')
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo Excel: {str(e)}")
        return pd.DataFrame()
    
    if df.empty:
        st.warning("⚠️ O arquivo Excel está vazio.")
        return df
    
    # Criar DataFrame padronizado
    novo_df = pd.DataFrame()
    
    # Mapear colunas
    colunas = df.columns.tolist()
    
    # ID (coluna A)
    if len(colunas) > 0:
        novo_df['ID'] = df.iloc[:, 0].fillna('').astype(str)
    
    # Title (coluna B)
    if len(colunas) > 1:
        novo_df['Title'] = df.iloc[:, 1].fillna('Sem título')
    
    # Status (coluna C) - com valor padrão
    if len(colunas) > 2:
        novo_df['Status'] = df.iloc[:, 2].fillna('Pago')
    else:
        novo_df['Status'] = 'Pago'
    
    # Classificação (coluna D)
    if len(colunas) > 3:
        novo_df['Classificação'] = df.iloc[:, 3].fillna('Despesa de veículo')
    else:
        novo_df['Classificação'] = 'Despesa de veículo'
    
    # Finalidade (coluna E)
    if len(colunas) > 4:
        novo_df['Finalidade'] = df.iloc[:, 4].fillna('Outros')
    else:
        novo_df['Finalidade'] = novo_df.get('Title', 'Outros')
    
    # Descrição (coluna F)
    if len(colunas) > 5:
        novo_df['Descrição'] = df.iloc[:, 5].fillna('')
    
    # Solicitante (coluna G)
    if len(colunas) > 6:
        novo_df['Solicitante'] = df.iloc[:, 6].fillna('Não informado')
    else:
        novo_df['Solicitante'] = 'Não informado'
    
    # Nome Motorista (coluna H)
    if len(colunas) > 7:
        novo_df['Nome Motorista'] = df.iloc[:, 7].fillna('')
    
    # VALOR (coluna I) - Conversão robusta
    if len(colunas) > 8:
        valor_col = df.iloc[:, 8]
        
        def converter_valor(valor):
            if pd.isna(valor):
                return 0.0
            
            # Se já for numérico
            if isinstance(valor, (int, float, np.integer, np.floating)):
                return float(valor)
            
            # Converter string
            valor_str = str(valor).strip()
            if not valor_str:
                return 0.0
            
            # Remover espaços
            valor_str = valor_str.replace(' ', '')
            
            # Padrões brasileiros
            # Formato: 1.234,56 ou 1234,56
            if ',' in valor_str:
                # Remover pontos de milhar
                if '.' in valor_str:
                    partes = valor_str.split(',')
                    parte_inteira = partes[0].replace('.', '')
                    valor_str = parte_inteira + '.' + (partes[1] if len(partes) > 1 else '00')
                else:
                    valor_str = valor_str.replace(',', '.')
            
            # Tentar converter
            try:
                return float(valor_str)
            except:
                # Tentar remover caracteres não numéricos
                valor_str = re.sub(r'[^\d\.\-]', '', valor_str)
                try:
                    return float(valor_str) if valor_str else 0.0
                except:
                    return 0.0
        
        novo_df['Valor'] = valor_col.apply(converter_valor)
    else:
        novo_df['Valor'] = 0.0
    
    # Data Criado (coluna O)
    if len(colunas) > 14:
        data_col = df.iloc[:, 14]
        
        # Converter datas
        def converter_data(data):
            if pd.isna(data):
                return pd.NaT
            
            try:
                # Tentar vários formatos
                for fmt in ['%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']:
                    try:
                        return pd.to_datetime(data, format=fmt)
                    except:
                        continue
                
                # Última tentativa
                return pd.to_datetime(data, errors='coerce')
            except:
                return pd.NaT
        
        datas_convertidas = data_col.apply(converter_data)
        
        # Verificar se temos datas válidas
        if datas_convertidas.notna().sum() > 0:
            novo_df['Criado'] = datas_convertidas
        else:
            # Criar datas baseadas no índice (fallback)
            novo_df['Criado'] = pd.date_range(
                start='2024-01-01', 
                periods=len(novo_df), 
                freq='D'
            )
    else:
        # Criar datas fictícias
        novo_df['Criado'] = pd.date_range(
            start='2024-01-01', 
            periods=len(novo_df), 
            freq='D'
        )
    
    # CPF Motorista (coluna J)
    if len(colunas) > 9:
        novo_df['CPF Motorista'] = df.iloc[:, 9].fillna('')
    
    # Conta Bancaria (coluna K)
    if len(colunas) > 10:
        novo_df['Conta Bancaria'] = df.iloc[:, 10].fillna('')
    
    # Gestor (coluna L)
    if len(colunas) > 11:
        novo_df['Gestor'] = df.iloc[:, 11].fillna('Gestor não especificado')
    else:
        novo_df['Gestor'] = novo_df.get('Solicitante', 'Gestor não especificado')
    
    # Placa (coluna W)
    if len(colunas) > 22:
        novo_df['Placa'] = df.iloc[:, 22].fillna('')
    
    # Criar colunas de data auxiliares
    novo_df['Ano'] = novo_df['Criado'].dt.year
    novo_df['Mes'] = novo_df['Criado'].dt.month
    novo_df['Dia'] = novo_df['Criado'].dt.day
    novo_df['Mes_Nome'] = novo_df['Criado'].dt.strftime('%b/%Y')
    
    # Remover linhas completamente vazias
    colunas_essenciais = ['Valor', 'Criado']
    col_existentes = [col for col in colunas_essenciais if col in novo_df.columns]
    
    if col_existentes:
        novo_df = novo_df.dropna(subset=col_existentes)
    
    # Garantir tipos de dados
    if 'Valor' in novo_df.columns:
        novo_df['Valor'] = pd.to_numeric(novo_df['Valor'], errors='coerce').fillna(0)
    
    return novo_df

def convert_df(df):
    """Converte DataFrame para Excel em bytes"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def gerar_projecao_mes_atual(df):
    """Gera projeção de custos para o mês atual"""
    hoje = datetime.now().date()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia_mes = (primeiro_dia + pd.offsets.MonthEnd(0)).date()
    
    # Dados do mês atual até hoje
    df_mes = df[
        (df['Criado'].dt.date >= primeiro_dia) & 
        (df['Criado'].dt.date <= hoje)
    ].copy()
    
    if df_mes.empty or len(df_mes) < 2:
        return pd.DataFrame(), 0, 0
    
    # Realizado por dia
    realizado = df_mes.groupby(df_mes['Criado'].dt.date)['Valor'].sum().reset_index()
    realizado.columns = ['Data', 'Valor']
    realizado['Tipo'] = 'Realizado'
    
    # Calcular médias para projeção
    if not realizado.empty:
        # Separar dias úteis e fins de semana
        dias_uteis = realizado[realizado['Data'].apply(lambda d: d.weekday() < 5)]
        dias_fds = realizado[realizado['Data'].apply(lambda d: d.weekday() >= 5)]
        
        media_uteis = dias_uteis['Valor'].mean() if not dias_uteis.empty else realizado['Valor'].mean()
        media_fds = dias_fds['Valor'].mean() if not dias_fds.empty else (media_uteis * 0.3)
    else:
        media_uteis = 0
        media_fds = 0
    
    # Gerar projeção para dias futuros
    dias_futuros = []
    data_atual = hoje + timedelta(days=1)
    
    while data_atual <= ultimo_dia_mes:
        if data_atual.weekday() < 5:  # Dia útil
            valor_proj = media_uteis
        else:  # Fim de semana
            valor_proj = media_fds
        
        # Adicionar pequena variação
        if valor_proj > 0:
            valor_proj *= np.random.uniform(0.8, 1.2)
        
        dias_futuros.append({
            'Data': data_atual,
            'Valor': max(0, valor_proj),
            'Tipo': 'Projetado'
        })
        data_atual += timedelta(days=1)
    
    # Combinar realizado e projetado
    if dias_futuros:
        df_projetado = pd.DataFrame(dias_futuros)
        df_completo = pd.concat([realizado, df_projetado], ignore_index=True)
    else:
        df_completo = realizado
    
    total_esperado = df_completo['Valor'].sum()
    
    return df_completo, total_esperado, media_uteis

def obter_data_inicio_padrao():
    """Retorna data de início padrão baseada no dia da semana"""
    hoje = datetime.now().date()
    dia_semana = hoje.weekday()  # 0=segunda, 6=domingo
    
    if dia_semana == 0:  # Segunda-feira
        return hoje - timedelta(days=3)  # Sexta-feira anterior
    else:
        return hoje - timedelta(days=1)  # Dia anterior

def obter_data_fim_padrao():
    """Retorna data de fim padrão baseada no dia da semana"""
    hoje = datetime.now().date()
    return hoje - timedelta(days=1)

# ======================== CARREGAR DADOS ========================
st.sidebar.title("⚙️ Configurações")

# Upload de arquivo alternativo
arquivo_padrao = "Projeto-custo-diário-solicitações-de-depósitos.xlsx"
upload_arquivo = st.sidebar.file_uploader(
    "📁 Carregar arquivo Excel", 
    type=['xlsx', 'xls'],
    help="Carregue seu arquivo de dados ou use o padrão"
)

if upload_arquivo is not None:
    # Salvar arquivo temporariamente
    with open("temp_upload.xlsx", "wb") as f:
        f.write(upload_arquivo.getbuffer())
    arquivo_carregar = "temp_upload.xlsx"
else:
    arquivo_carregar = arquivo_padrao

try:
    with st.spinner("📊 Carregando dados..."):
        df = load_data(arquivo_carregar)
    
    if df.empty:
        st.error("""
        ❌ Não foi possível carregar dados válidos do arquivo.
        
        **Possíveis causas:**
        1. O arquivo não está no formato correto
        2. O arquivo está vazio
        3. As colunas esperadas não foram encontradas
        
        **Solução:** Carregue um arquivo Excel com a estrutura esperada.
        """)
        st.stop()
        
except Exception as e:
    st.error(f"❌ Erro crítico ao carregar dados: {str(e)}")
    st.stop()

# ======================== MENU LATERAL ========================
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "📌 Navegação",
    ["📊 Dashboard Geral", "👤 Análise Detalhada", "🏗️ Reunião Manutenção"],
    index=0
)

# ======================== DASHBOARD GERAL ========================
if menu == "📊 Dashboard Geral":
    st.title("📊 Dashboard de Custos - Solicitações de Depósitos")
    
    st.sidebar.header("🔍 Filtros")
    
    # Filtros de data
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        data_inicio = st.date_input(
            "📅 Data início", 
            value=min_date,
            min_value=min_date,
            max_value=max_date
        )
    with col2:
        data_fim = st.date_input(
            "📅 Data fim", 
            value=max_date,
            min_value=min_date,
            max_value=max_date
        )
    
    # Outros filtros
    st.sidebar.subheader("Filtros Adicionais")
    
    # Obter opções únicas
    solicitantes = sorted(df['Solicitante'].dropna().unique())
    status_opcoes = sorted(df['Status'].dropna().unique())
    gestores = sorted(df['Gestor'].dropna().unique())
    classificacoes = sorted(df['Classificação'].dropna().unique())
    finalidades = sorted(df['Finalidade'].dropna().unique())
    
    # Gestores padrão
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores]
    
    # Widgets de filtro
    solicitante_filtro = st.sidebar.multiselect(
        "🙋‍♂️ Solicitante",
        options=solicitantes,
        help="Selecione um ou mais solicitantes"
    )
    
    status_filtro = st.sidebar.multiselect(
        "📌 Status",
        options=status_opcoes,
        default=status_opcoes[:3] if len(status_opcoes) > 0 else []
    )
    
    gestor_filtro = st.sidebar.multiselect(
        "👔 Gestor",
        options=gestores,
        default=gestores_disponiveis
    )
    
    classificacao_filtro = st.sidebar.multiselect(
        "🏷️ Classificação",
        options=classificacoes
    )
    
    finalidade_filtro = st.sidebar.multiselect(
        "🎯 Finalidade",
        options=finalidades
    )
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    # Filtro de data
    df_filtrado = df_filtrado[
        (df_filtrado['Criado'].dt.date >= data_inicio) & 
        (df_filtrado['Criado'].dt.date <= data_fim)
    ]
    
    # Aplicar outros filtros
    if solicitante_filtro:
        df_filtrado = df_filtrado[df_filtrado['Solicitante'].isin(solicitante_filtro)]
    
    if status_filtro:
        df_filtrado = df_filtrado[df_filtrado['Status'].isin(status_filtro)]
    
    if gestor_filtro:
        df_filtrado = df_filtrado[df_filtrado['Gestor'].isin(gestor_filtro)]
    
    if classificacao_filtro:
        df_filtrado = df_filtrado[df_filtrado['Classificação'].isin(classificacao_filtro)]
    
    if finalidade_filtro:
        df_filtrado = df_filtrado[df_filtrado['Finalidade'].isin(finalidade_filtro)]
    
    # Resumo dos filtros
    with st.expander("📋 Resumo dos Filtros Aplicados", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Período", f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
            st.metric("Solicitantes", len(solicitante_filtro) if solicitante_filtro else "Todos")
            st.metric("Status", len(status_filtro) if status_filtro else "Todos")
        with col2:
            st.metric("Gestores", len(gestor_filtro) if gestor_filtro else "Todos")
            st.metric("Classificações", len(classificacao_filtro) if classificacao_filtro else "Todos")
            st.metric("Finalidades", len(finalidade_filtro) if finalidade_filtro else "Todos")
    
    # Métricas principais
    st.markdown("### 📈 Métricas Principais")
    
    if not df_filtrado.empty:
        custo_total = df_filtrado['Valor'].sum()
        qtd_registros = len(df_filtrado)
        dias_cobertos = df_filtrado['Criado'].dt.date.nunique()
        custo_medio_diario = custo_total / dias_cobertos if dias_cobertos > 0 else 0
        custo_medio_registro = custo_total / qtd_registros if qtd_registros > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "💰 Custo Total", 
                f"R$ {formatar_brasileiro(custo_total)}",
                help="Soma de todos os valores no período"
            )
        
        with col2:
            st.metric(
                "📅 Custo Médio Diário", 
                f"R$ {formatar_brasileiro(custo_medio_diario)}",
                help=f"Baseado em {dias_cobertos} dias"
            )
        
        with col3:
            st.metric(
                "📊 Custo Médio por Registro", 
                f"R$ {formatar_brasileiro(custo_medio_registro)}",
                help=f"Baseado em {qtd_registros} registros"
            )
        
        with col4:
            st.metric(
                "📝 Total de Registros", 
                f"{qtd_registros:,}",
                delta=None
            )
        
        # Gráfico 1: Evolução Mensal
        st.markdown("### 📅 Evolução Mensal de Custos")
        
        df_mensal = df_filtrado.groupby('Mes_Nome').agg({
            'Valor': 'sum',
            'ID': 'count'
        }).reset_index()
        
        df_mensal = df_mensal.sort_values('Mes_Nome')
        
        if not df_mensal.empty:
            fig_mensal = px.line(
                df_mensal, 
                x='Mes_Nome', 
                y='Valor',
                markers=True,
                title="Custo Total por Mês",
                labels={'Valor': 'Custo (R$)', 'Mes_Nome': 'Mês'}
            )
            
            fig_mensal.update_traces(
                hovertemplate="<b>%{x}</b><br>Custo: R$ %{y:,.2f}<br>Registros: %{customdata}",
                customdata=df_mensal['ID']
            )
            
            fig_mensal.update_layout(
                hoverlabel=dict(font_size=14),
                xaxis_title="Mês",
                yaxis_title="Custo Total (R$)"
            )
            
            st.plotly_chart(fig_mensal, use_container_width=True)
        
        # Gráfico 2: Top Finalidades
        st.markdown("### 🎯 Top Finalidades por Custo")
        
        df_finalidade = df_filtrado.groupby('Finalidade').agg({
            'Valor': 'sum',
            'ID': 'count'
        }).reset_index()
        
        df_finalidade = df_finalidade.sort_values('Valor', ascending=False).head(10)
        
        if not df_finalidade.empty:
            df_finalidade['Valor_Formatado'] = df_finalidade['Valor'].apply(formatar_brasileiro)
            
            fig_finalidade = px.bar(
                df_finalidade,
                x='Valor',
                y='Finalidade',
                orientation='h',
                title="Top 10 Finalidades por Custo Total",
                labels={'Valor': 'Custo (R$)', 'Finalidade': ''},
                text=df_finalidade.apply(
                    lambda row: f"R$ {row['Valor_Formatado']} ({row['ID']})", 
                    axis=1
                )
            )
            
            fig_finalidade.update_traces(
                textposition='outside',
                marker_color='steelblue'
            )
            
            fig_finalidade.update_layout(
                yaxis={'categoryorder': 'total ascending'},
                height=500
            )
            
            st.plotly_chart(fig_finalidade, use_container_width=True)
        
        # Gráfico 3: Distribuição por Classificação
        st.markdown("### 🏷️ Distribuição por Classificação")
        
        df_classificacao = df_filtrado.groupby('Classificação')['Valor'].sum().reset_index()
        
        if not df_classificacao.empty:
            fig_classificacao = px.pie(
                df_classificacao,
                names='Classificação',
                values='Valor',
                hole=0.4,
                title="Distribuição de Custos por Classificação"
            )
            
            fig_classificacao.update_traces(
                textposition='inside',
                textinfo='percent+label',
                hovertemplate="<b>%{label}</b><br>Custo: R$ %{value:,.2f}<br>Percentual: %{percent}"
            )
            
            st.plotly_chart(fig_classificacao, use_container_width=True)
        
        # Projeção do mês atual
        st.markdown("### 🔮 Projeção para o Mês Atual")
        
        df_projecao, total_projetado, media_diaria = gerar_projecao_mes_atual(df_filtrado)
        
        if not df_projecao.empty:
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric(
                    "📈 Custo Total Projetado",
                    f"R$ {formatar_brasileiro(total_projetado)}"
                )
            
            with col2:
                st.metric(
                    "📊 Média Diária Projetada",
                    f"R$ {formatar_brasileiro(media_diaria)}"
                )
            
            fig_projecao = px.line(
                df_projecao,
                x='Data',
                y='Valor',
                color='Tipo',
                title="Projeção de Custos Diários",
                labels={'Valor': 'Custo (R$)', 'Data': 'Data', 'Tipo': ''},
                color_discrete_map={'Realizado': 'blue', 'Projetado': 'green'}
            )
            
            fig_projecao.update_layout(
                hovermode='x unified',
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
            )
            
            st.plotly_chart(fig_projecao, use_container_width=True)
        
        # Tabela de dados
        st.markdown("### 📋 Visualização dos Dados")
        
        colunas_visiveis = ['ID', 'Title', 'Valor', 'Finalidade', 'Solicitante', 
                           'Status', 'Classificação', 'Criado']
        colunas_disponiveis = [col for col in colunas_visiveis if col in df_filtrado.columns]
        
        df_display = df_filtrado[colunas_disponiveis].copy()
        df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        df_display['Valor'] = df_display['Valor'].apply(lambda x: f"R$ {formatar_brasileiro(x)}")
        
        st.dataframe(
            df_display,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # Botão de download
        st.download_button(
            label="📥 Baixar Dados Filtrados (Excel)",
            data=convert_df(df_filtrado),
            file_name=f"dados_filtrados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    else:
        st.warning("⚠️ Nenhum dado encontrado com os filtros aplicados.")

# ======================== ANÁLISE DETALHADA ========================
elif menu == "👤 Análise Detalhada":
    st.title("👤 Análise Detalhada por Solicitante")
    
    st.sidebar.header("🔍 Filtros da Análise")
    
    # Filtros de data
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    data_inicio = st.sidebar.date_input(
        "📅 Data início análise", 
        value=data_inicio_padrao,
        min_value=min_date,
        max_value=max_date
    )
    
    data_fim = st.sidebar.date_input(
        "📅 Data fim análise", 
        value=data_fim_padrao,
        min_value=min_date,
        max_value=max_date
    )
    
    # Filtros adicionais
    status_opcoes = sorted(df['Status'].dropna().unique())
    classificacoes = sorted(df['Classificação'].dropna().unique())
    gestores = sorted(df['Gestor'].dropna().unique())
    
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores]
    
    status_filtro = st.sidebar.multiselect(
        "📌 Status",
        options=status_opcoes,
        default=status_opcoes[:3] if status_opcoes else []
    )
    
    classificacao_filtro = st.sidebar.multiselect(
        "🏷️ Classificação",
        options=classificacoes
    )
    
    gestor_filtro = st.sidebar.multiselect(
        "👔 Gestor",
        options=gestores,
        default=gestores_disponiveis
    )
    
    # Aplicar filtros iniciais
    df_filtrado = df.copy()
    df_filtrado = df_filtrado[
        (df_filtrado['Criado'].dt.date >= data_inicio) & 
        (df_filtrado['Criado'].dt.date <= data_fim)
    ]
    
    if status_filtro:
        df_filtrado = df_filtrado[df_filtrado['Status'].isin(status_filtro)]
    
    if classificacao_filtro:
        df_filtrado = df_filtrado[df_filtrado['Classificação'].isin(classificacao_filtro)]
    
    if gestor_filtro:
        df_filtrado = df_filtrado[df_filtrado['Gestor'].isin(gestor_filtro)]
    
    # Seleção de solicitante
    solicitantes_disponiveis = sorted(df_filtrado['Solicitante'].dropna().unique())
    
    st.subheader("Seleção de Solicitante")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        analise_tipo = st.radio(
            "Tipo de análise",
            ["Solicitante específico", "Visão geral"],
            horizontal=True
        )
    
    with col2:
        if analise_tipo == "Solicitante específico":
            solicitante_selecionado = st.selectbox(
                "Selecione o solicitante",
                options=solicitantes_disponiveis,
                index=0 if solicitantes_disponiveis else None
            )
            
            df_analise = df_filtrado[df_filtrado['Solicitante'] == solicitante_selecionado]
            
            # Métricas do solicitante
            if not df_analise.empty:
                custo_total = df_analise['Valor'].sum()
                qtd_solicitacoes = len(df_analise)
                custo_medio = custo_total / qtd_solicitacoes if qtd_solicitacoes > 0 else 0
                dias_ativos = df_analise['Criado'].dt.date.nunique()
                custo_diario_medio = custo_total / dias_ativos if dias_ativos > 0 else 0
                
                st.markdown(f"### 📊 Estatísticas de **{solicitante_selecionado}**")
                
                col_met1, col_met2, col_met3, col_met4 = st.columns(4)
                
                with col_met1:
                    st.metric("💰 Custo Total", f"R$ {formatar_brasileiro(custo_total)}")
                
                with col_met2:
                    st.metric("📝 Solicitações", qtd_solicitacoes)
                
                with col_met3:
                    st.metric("⚖️ Custo Médio", f"R$ {formatar_brasileiro(custo_medio)}")
                
                with col_met4:
                    st.metric("📅 Dias Ativos", dias_ativos)
                
                # Gráfico de evolução
                st.markdown("### 📈 Evolução Temporal")
                
                df_evolucao = df_analise.groupby(df_analise['Criado'].dt.date).agg({
                    'Valor': 'sum',
                    'ID': 'count'
                }).reset_index()
                
                if not df_evolucao.empty:
                    fig_evolucao = px.line(
                        df_evolucao,
                        x='Criado',
                        y='Valor',
                        markers=True,
                        title=f"Evolução de Custos - {solicitante_selecionado}",
                        labels={'Valor': 'Custo Diário (R$)', 'Criado': 'Data'}
                    )
                    
                    fig_evolucao.update_traces(
                        hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Custo: R$ %{y:,.2f}<br>Solicitações: %{customdata}",
                        customdata=df_evolucao['ID']
                    )
                    
                    st.plotly_chart(fig_evolucao, use_container_width=True)
        else:
            # Visão geral
            st.markdown("### 📊 Visão Geral dos Solicitantes")
            
            df_agrupado = df_filtrado.groupby('Solicitante').agg({
                'Valor': 'sum',
                'ID': 'count',
                'Criado': lambda x: x.nunique()  # Dias distintos
            }).reset_index()
            
            df_agrupado.columns = ['Solicitante', 'Custo_Total', 'Qtd_Solicitacoes', 'Dias_Ativos']
            df_agrupado['Custo_Medio_Diario'] = df_agrupado['Custo_Total'] / df_agrupado['Dias_Ativos']
            df_agrupado = df_agrupado.sort_values('Custo_Total', ascending=False)
            
            # Métricas gerais
            total_geral = df_agrupado['Custo_Total'].sum()
            total_solicitacoes = df_agrupado['Qtd_Solicitacoes'].sum()
            
            col_ger1, col_ger2 = st.columns(2)
            
            with col_ger1:
                st.metric("💰 Custo Total Geral", f"R$ {formatar_brasileiro(total_geral)}")
            
            with col_ger2:
                st.metric("📝 Total de Solicitações", total_solicitacoes)
            
            # Gráfico de top solicitantes
            st.markdown("### 🏆 Top Solicitantes por Custo")
            
            top_n = st.slider("Número de solicitantes para mostrar", 5, 20, 10)
            df_top = df_agrupado.head(top_n)
            
            if not df_top.empty:
                fig_top = px.bar(
                    df_top,
                    x='Custo_Total',
                    y='Solicitante',
                    orientation='h',
                    title=f"Top {top_n} Solicitantes por Custo Total",
                    labels={'Custo_Total': 'Custo Total (R$)', 'Solicitante': ''},
                    text=df_top.apply(
                        lambda row: f"R$ {formatar_brasileiro(row['Custo_Total'])} ({row['Qtd_Solicitacoes']} solic.)", 
                        axis=1
                    )
                )
                
                fig_top.update_traces(
                    textposition='outside',
                    marker_color='teal'
                )
                
                fig_top.update_layout(
                    yaxis={'categoryorder': 'total ascending'},
                    height=500
                )
                
                st.plotly_chart(fig_top, use_container_width=True)
            
            df_analise = df_filtrado
    
    # Tabela de dados
    st.markdown("### 📋 Detalhamento das Solicitações")
    
    if not df_analise.empty:
        colunas_detalhe = ['ID', 'Title', 'Valor', 'Finalidade', 'Classificação', 
                          'Status', 'Gestor', 'Criado']
        colunas_disponiveis = [col for col in colunas_detalhe if col in df_analise.columns]
        
        df_detalhe = df_analise[colunas_disponiveis].copy()
        df_detalhe['Criado'] = df_detalhe['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        df_detalhe['Valor'] = df_detalhe['Valor'].apply(lambda x: f"R$ {formatar_brasileiro(x)}")
        
        st.dataframe(
            df_detalhe,
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # Botão de download
        nome_arquivo = "analise_detalhada"
        if analise_tipo == "Solicitante específico" and 'solicitante_selecionado' in locals():
            nome_arquivo = f"analise_{solicitante_selecionado.replace(' ', '_')}"
        
        st.download_button(
            label="📥 Baixar Dados da Análise (Excel)",
            data=convert_df(df_analise),
            file_name=f"{nome_arquivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Nenhum dado para exibir com os filtros aplicados.")

# ======================== REUNIÃO MANUTENÇÃO CORPORATIVA ========================
elif menu == "🏗️ Reunião Manutenção":
    st.title("🏗️ Relatório para Reunião de Manutenção Corporativa")
    
    st.info("""
    Este relatório é otimizado para apresentações em reuniões de manutenção corporativa, 
    focando em métricas estratégicas e visualizações claras.
    """)
    
    st.sidebar.header("⚙️ Configurações do Relatório")
    
    # Filtros de data
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    periodo = st.sidebar.selectbox(
        "📅 Período de Análise",
        ["Personalizado", "Última semana", "Último mês", "Último trimestre"],
        index=1
    )
    
    if periodo == "Personalizado":
        data_inicio = st.sidebar.date_input(
            "Data início", 
            value=data_inicio_padrao,
            min_value=min_date,
            max_value=max_date
        )
        data_fim = st.sidebar.date_input(
            "Data fim", 
            value=data_fim_padrao,
            min_value=min_date,
            max_value=max_date
        )
    elif periodo == "Última semana":
        data_fim = datetime.now().date()
        data_inicio = data_fim - timedelta(days=7)
    elif periodo == "Último mês":
        data_fim = datetime.now().date()
        data_inicio = data_fim - timedelta(days=30)
    else:  # Último trimestre
        data_fim = datetime.now().date()
        data_inicio = data_fim - timedelta(days=90)
    
    # Filtros principais
    gestores = sorted(df['Gestor'].dropna().unique())
    status_opcoes = sorted(df['Status'].dropna().unique())
    classificacoes = sorted(df['Classificação'].dropna().unique())
    
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores]
    
    gestor_filtro = st.sidebar.multiselect(
        "👔 Gestores para Análise",
        options=gestores,
        default=gestores_disponiveis
    )
    
    status_filtro = st.sidebar.multiselect(
        "📌 Status das Solicitações",
        options=status_opcoes,
        default=status_opcoes[:3] if status_opcoes else []
    )
    
    classificacao_filtro = st.sidebar.multiselect(
        "🏷️ Classificações",
        options=classificacoes
    )
    
    # Aplicar filtros
    df_relatorio = df.copy()
    df_relatorio = df_relatorio[
        (df_relatorio['Criado'].dt.date >= data_inicio) & 
        (df_relatorio['Criado'].dt.date <= data_fim)
    ]
    
    if gestor_filtro:
        df_relatorio = df_relatorio[df_relatorio['Gestor'].isin(gestor_filtro)]
    
    if status_filtro:
        df_relatorio = df_relatorio[df_relatorio['Status'].isin(status_filtro)]
    
    if classificacao_filtro:
        df_relatorio = df_relatorio[df_relatorio['Classificação'].isin(classificacao_filtro)]
    
    # Cabeçalho do relatório
    st.markdown(f"""
    ### Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}
    
    **Gestores analisados:** {', '.join(gestor_filtro) if gestor_filtro else 'Todos'}
    """)
    
    # Métricas de alto nível
    st.markdown("### 📊 Métricas de Performance")
    
    if not df_relatorio.empty:
        custo_total = df_relatorio['Valor'].sum()
        qtd_solicitacoes = len(df_relatorio)
        dias_periodo = (data_fim - data_inicio).days + 1
        custo_diario_medio = custo_total / dias_periodo
        custo_medio_solicitacao = custo_total / qtd_solicitacoes if qtd_solicitacoes > 0 else 0
        
        # Maior solicitação
        maior_idx = df_relatorio['Valor'].idxmax()
        maior_solicitacao = df_relatorio.loc[maior_idx]
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "💰 Investimento Total",
                f"R$ {formatar_brasileiro(custo_total)}",
                help="Soma de todas as solicitações no período"
            )
        
        with col2:
            st.metric(
                "📈 Custo Médio Diário",
                f"R$ {formatar_brasileiro(custo_diario_medio)}",
                help=f"Média por dia em {dias_periodo} dias"
            )
        
        with col3:
            st.metric(
                "📊 Custo por Solicitação",
                f"R$ {formatar_brasileiro(custo_medio_solicitacao)}",
                help=f"Média por solicitação ({qtd_solicitacoes} total)"
            )
        
        with col4:
            st.metric(
                "🏆 Maior Solicitação",
                f"R$ {formatar_brasileiro(maior_solicitacao['Valor'])}",
                help=f"ID: {maior_solicitacao.get('ID', 'N/A')}"
            )
        
        # Análise por gestor
        st.markdown("### 👥 Performance por Gestor")
        
        df_gestor = df_relatorio.groupby('Gestor').agg({
            'Valor': ['sum', 'mean', 'count'],
            'ID': 'count'
        }).reset_index()
        
        df_gestor.columns = ['Gestor', 'Custo_Total', 'Custo_Medio', 'Contagem', 'Solicitacoes']
        df_gestor = df_gestor.sort_values('Custo_Total', ascending=False)
        
        if not df_gestor.empty:
            fig_gestor = px.bar(
                df_gestor,
                x='Gestor',
                y='Custo_Total',
                color='Gestor',
                title="Custo Total por Gestor",
                labels={'Custo_Total': 'Custo Total (R$)', 'Gestor': 'Gestor'},
                text=df_gestor['Custo_Total'].apply(lambda x: f"R$ {formatar_brasileiro(x)}")
            )
            
            fig_gestor.update_traces(
                textposition='outside',
                textfont=dict(size=12)
            )
            
            fig_gestor.update_layout(
                xaxis_tickangle=-45,
                showlegend=False
            )
            
            st.plotly_chart(fig_gestor, use_container_width=True)
        
        # Tendências temporais
        st.markdown("### 📈 Tendência de Custos")
        
        df_tendencia = df_relatorio.groupby(df_relatorio['Criado'].dt.date).agg({
            'Valor': 'sum',
            'ID': 'count'
        }).reset_index()
        
        df_tendencia.columns = ['Data', 'Custo_Diario', 'Solicitacoes_Diarias']
        
        if not df_tendencia.empty:
            fig_tendencia = px.line(
                df_tendencia,
                x='Data',
                y='Custo_Diario',
                markers=True,
                title="Evolução Diária de Custos",
                labels={'Custo_Diario': 'Custo Diário (R$)', 'Data': 'Data'}
            )
            
            fig_tendencia.update_traces(
                hovertemplate="<b>%{x|%d/%m/%Y}</b><br>Custo: R$ %{y:,.2f}<br>Solicitações: %{customdata}",
                customdata=df_tendencia['Solicitacoes_Diarias']
            )
            
            # Adicionar média móvel
            if len(df_tendencia) > 5:
                df_tendencia['Media_Movel'] = df_tendencia['Custo_Diario'].rolling(window=3).mean()
                
                fig_tendencia.add_scatter(
                    x=df_tendencia['Data'],
                    y=df_tendencia['Media_Movel'],
                    mode='lines',
                    name='Média Móvel (3 dias)',
                    line=dict(dash='dash', color='orange')
                )
            
            st.plotly_chart(fig_tendencia, use_container_width=True)
        
        # Análise de categorias
        st.markdown("### 🏷️ Distribuição por Categoria")
        
        col_cat1, col_cat2 = st.columns(2)
        
        with col_cat1:
            # Por classificação
            df_class = df_relatorio.groupby('Classificação')['Valor'].sum().reset_index()
            
            if not df_class.empty:
                fig_class = px.pie(
                    df_class,
                    names='Classificação',
                    values='Valor',
                    title="Distribuição por Classificação",
                    hole=0.4
                )
                
                st.plotly_chart(fig_class, use_container_width=True)
        
        with col_cat2:
            # Por finalidade (top 10)
            df_final = df_relatorio.groupby('Finalidade')['Valor'].sum().reset_index()
            df_final = df_final.sort_values('Valor', ascending=False).head(10)
            
            if not df_final.empty:
                fig_final = px.bar(
                    df_final,
                    x='Finalidade',
                    y='Valor',
                    title="Top 10 Finalidades",
                    labels={'Valor': 'Custo (R$)', 'Finalidade': ''}
                )
                
                fig_final.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_final, use_container_width=True)
        
        # Projeções e insights
        st.markdown("### 🔮 Insights e Projeções")
        
        col_ins1, col_ins2 = st.columns(2)
        
        with col_ins1:
            st.markdown("#### 📊 Projeção para Próxima Semana")
            
            if not df_tendencia.empty and len(df_tendencia) >= 5:
                media_semanal = df_tendencia['Custo_Diario'].tail(5).mean()
                projecao_semanal = media_semanal * 5
                
                st.success(f"""
                **Estimativa baseada nos últimos 5 dias:**
                
                • Média diária: **R$ {formatar_brasileiro(media_semanal)}**
                • Projeção semanal (5 dias): **R$ {formatar_brasileiro(projecao_semanal)}**
                """)
            else:
                st.info("ℹ️ Dados insuficientes para projeção semanal.")
        
        with col_ins2:
            st.markdown("#### ⚠️ Alertas e Observações")
            
            # Verificar anomalias
            if not df_tendencia.empty:
                ultimo_custo = df_tendencia['Custo_Diario'].iloc[-1]
                media_historica = df_tendencia['Custo_Diario'].mean()
                
                if ultimo_custo > media_historica * 1.5:
                    st.warning(f"""
                    **Alerta:** Custo do último dia está **{((ultimo_custo/media_historica)-1)*100:.0f}% acima** da média.
                    
                    • Último dia: R$ {formatar_brasileiro(ultimo_custo)}
                    • Média histórica: R$ {formatar_brasileiro(media_historica)}
                    """)
                else:
                    st.success("✅ Nenhuma anomalia significativa detectada.")
        
        # Tabela resumo
        st.markdown("### 📋 Resumo Executivo")
        
        resumo_colunas = ['ID', 'Title', 'Valor', 'Finalidade', 'Gestor', 
                         'Classificação', 'Status', 'Criado']
        colunas_disponiveis = [col for col in resumo_colunas if col in df_relatorio.columns]
        
        df_resumo = df_relatorio[colunas_disponiveis].copy()
        df_resumo['Criado'] = df_resumo['Criado'].dt.strftime('%d/%m/%Y')
        df_resumo['Valor'] = df_resumo['Valor'].apply(lambda x: f"R$ {formatar_brasileiro(x)}")
        df_resumo = df_resumo.sort_values('Valor', ascending=False)
        
        st.dataframe(
            df_resumo.head(20),  # Mostrar apenas as 20 maiores
            use_container_width=True,
            height=400,
            hide_index=True
        )
        
        # Botões de ação
        st.markdown("---")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            st.download_button(
                label="📥 Baixar Relatório Completo",
                data=convert_df(df_relatorio),
                file_name=f"relatorio_reuniao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_btn2:
            if st.button("🖨️ Gerar PDF (Simulado)", use_container_width=True):
                st.success("✅ Relatório preparado para impressão (função PDF em desenvolvimento)")
        
        with col_btn3:
            if st.button("📧 Enviar por E-mail", use_container_width=True):
                st.info("📤 Função de envio por e-mail em desenvolvimento")
    
    else:
        st.warning("""
        ⚠️ Nenhum dado encontrado para os filtros aplicados.
        
        **Sugestões:**
        1. Amplie o período de análise
        2. Verifique os filtros de gestor e status
        3. Carregue um arquivo com dados do período desejado
        """)

# ======================== RODAPÉ ========================
st.sidebar.markdown("---")
st.sidebar.markdown("""
**ℹ️ Sobre este dashboard:**
- Desenvolvido para análise de custos
- Atualizado automaticamente
- Formato brasileiro de valores
""")

st.sidebar.markdown("""
**🔄 Atualização dos dados:**
- Carregamento automático do arquivo Excel
- Cache de 1 hora para performance
- Suporte a upload de novos arquivos
""")

# Informações de versão
st.sidebar.markdown(f"""
**📊 Estatísticas do dataset:**
- Período: {df['Criado'].min().strftime('%d/%m/%Y')} a {df['Criado'].max().strftime('%d/%m/%Y')}
- Registros: {len(df):,}
- Custo total: R$ {formatar_brasileiro(df['Valor'].sum())}
""")
