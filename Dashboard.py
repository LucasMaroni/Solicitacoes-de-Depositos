import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import locale
import io
import numpy as np
import re

# Configurar locale para português brasileiro
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')
    except:
        pass

st.set_page_config(page_title="Dashboard de Custos", layout="wide")

# Função para formatar números no formato brasileiro
def formatar_brasileiro(valor, decimais=2):
    """
    Formata um número no formato brasileiro: 1.234,56
    """
    if pd.isna(valor):
        return "0,00"
    
    try:
        # Converter para float se for string
        if isinstance(valor, str):
            # Limpar a string
            valor_str = valor.replace('.', '').replace(',', '.')
            valor_float = float(valor_str)
        else:
            valor_float = float(valor)
        
        # Formatar com separador de milhar e vírgula decimal
        parte_inteira = int(valor_float)
        parte_decimal = int(round((valor_float - parte_inteira) * 100))
        
        # Formatar parte inteira com separadores de milhar
        parte_inteira_str = f"{parte_inteira:,}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # Juntar com parte decimal
        return f"{parte_inteira_str},{parte_decimal:02d}"
    except:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

@st.cache_data
def load_data(caminho_arquivo):
    # Ler o arquivo Excel
    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        return pd.DataFrame()
    
    # Baseado na estrutura fornecida, criar um dataframe com colunas padrão
    novo_df = pd.DataFrame()
    colunas_disponiveis = df.columns.tolist()
    
    # Mapear colunas baseado na estrutura fornecida
    # A coluna A (índice 0) é ID
    if len(colunas_disponiveis) > 0:
        novo_df['ID'] = df.iloc[:, 0] if len(df.columns) > 0 else range(1, len(df) + 1)
    
    # A coluna B (índice 1) é Title
    if len(colunas_disponiveis) > 1:
        novo_df['Title'] = df.iloc[:, 1]
    
    # A coluna C (índice 2) é Status
    if len(colunas_disponiveis) > 2:
        novo_df['Status'] = df.iloc[:, 2]
    else:
        novo_df['Status'] = 'Pago'  # Valor padrão
    
    # A coluna D (índice 3) é Classificação
    if len(colunas_disponiveis) > 3:
        novo_df['Classificação'] = df.iloc[:, 3]
    else:
        novo_df['Classificação'] = 'Despesa de veículo'  # Valor padrão
    
    # A coluna E (índice 4) é Finalidade
    if len(colunas_disponiveis) > 4:
        novo_df['Finalidade'] = df.iloc[:, 4]
    elif 'Title' in novo_df.columns:
        novo_df['Finalidade'] = novo_df['Title']
    else:
        novo_df['Finalidade'] = 'Outros'
    
    # A coluna F (índice 5) é Descrição
    if len(colunas_disponiveis) > 5:
        novo_df['Descrição'] = df.iloc[:, 5]
    
    # A coluna G (índice 6) é Solicitante
    if len(colunas_disponiveis) > 6:
        novo_df['Solicitante'] = df.iloc[:, 6]
    else:
        novo_df['Solicitante'] = 'Não informado'
    
    # A coluna H (índice 7) é Nome Motorista
    if len(colunas_disponiveis) > 7:
        novo_df['Nome Motorista'] = df.iloc[:, 7]
    
    # A coluna I (índice 8) é VALOR - Esta é a coluna crítica!
    if len(colunas_disponiveis) > 8:
        valor_col = df.iloc[:, 8]
        
        # Função para converter valor brasileiro para float
        def converter_valor_brasileiro(valor):
            if pd.isna(valor):
                return 0.0
            
            # Converter para string
            valor_str = str(valor).strip()
            
            # Se for vazio, retornar 0
            if not valor_str:
                return 0.0
            
            # Remover espaços em branco
            valor_str = valor_str.replace(' ', '')
            
            # Caso 1: Já é um número (float ou int)
            try:
                if isinstance(valor, (int, float, np.integer, np.floating)):
                    return float(valor)
            except:
                pass
            
            # Caso 2: Tem vírgula como separador decimal
            # Exemplos: "439,75", "1160,53", "1834,73"
            if ',' in valor_str:
                # Verificar se tem ponto como separador de milhar
                if '.' in valor_str:
                    # Formato: "1.160,53" ou "1.834,73"
                    # Remover pontos (separadores de milhar)
                    valor_str = valor_str.replace('.', '')
                
                # Substituir vírgula por ponto para decimal
                valor_str = valor_str.replace(',', '.')
            
            # Caso 3: Tem ponto como separador decimal (formato internacional)
            # Neste caso, manter como está
            
            # Remover qualquer caractere não numérico, exceto ponto e sinal negativo
            # Primeiro, extrair o sinal negativo se existir
            negativo = False
            if valor_str.startswith('-'):
                negativo = True
                valor_str = valor_str[1:]
            
            # Remover caracteres não numéricos, exceto ponto
            valor_str = re.sub(r'[^\d\.]', '', valor_str)
            
            # Restaurar sinal negativo
            if negativo:
                valor_str = '-' + valor_str
            
            # Se estiver vazio após limpeza, retornar 0
            if not valor_str:
                return 0.0
            
            # Se terminar com ponto, remover
            if valor_str.endswith('.'):
                valor_str = valor_str[:-1]
            
            # Converter para float
            try:
                return float(valor_str)
            except:
                # Última tentativa: remover todos os pontos exceto o último
                if valor_str.count('.') > 1:
                    partes = valor_str.split('.')
                    valor_str = ''.join(partes[:-1]) + '.' + partes[-1]
                    try:
                        return float(valor_str)
                    except:
                        return 0.0
                return 0.0
        
        # Aplicar conversão a todos os valores
        valores_convertidos = []
        
        for i, valor in enumerate(valor_col):
            convertido = converter_valor_brasileiro(valor)
            valores_convertidos.append(convertido)
        
        novo_df['Valor'] = valores_convertidos
        
    else:
        novo_df['Valor'] = 0
    
    # A coluna O (índice 14) é Criado (data)
    if len(colunas_disponiveis) > 14:
        data_col = df.iloc[:, 14]
        
        # Tentar convertir a data
        try:
            # Função para converter data brasileira
            def converter_data_brasileira(data_str):
                if pd.isna(data_str):
                    return pd.NaT
                
                data_str = str(data_str).strip()
                
                # Tentar formato "dd/mm/yyyy HH:MM:SS"
                try:
                    return pd.to_datetime(data_str, format='%d/%m/%Y %H:%M:%S', errors='coerce')
                except:
                    pass
                
                # Tentar formato "dd/mm/yyyy"
                try:
                    return pd.to_datetime(data_str, format='%d/%m/%Y', errors='coerce')
                except:
                    pass
                
                # Tentar qualquer formato que o pandas reconheça
                try:
                    return pd.to_datetime(data_str, errors='coerce')
                except:
                    return pd.NaT
            
            datas_convertidas = data_col.apply(converter_data_brasileira)
            
            # Verificar se alguma data foi convertida
            datas_validas = datas_convertidas.notna().sum()
            if datas_validas > 0:
                novo_df['Criado'] = datas_convertidas
            else:
                # Criar datas baseadas no índice
                datas_convertidas = pd.date_range(start='2026-01-01', periods=len(novo_df), freq='D')
                novo_df['Criado'] = datas_convertidas
            
        except Exception as e:
            # Criar datas fictícias para evitar erros
            novo_df['Criado'] = pd.date_range(start='2026-01-01', periods=len(novo_df), freq='D')
    else:
        # Se não encontrar data, criar datas baseadas no índice
        novo_df['Criado'] = pd.date_range(start='2026-01-01', periods=len(novo_df), freq='D')
    
    # A coluna J (índice 9) é CPF Motorista
    if len(colunas_disponiveis) > 9:
        novo_df['CPF Motorista'] = df.iloc[:, 9]
    
    # A coluna K (índice 10) é Conta Bancaria
    if len(colunas_disponiveis) > 10:
        novo_df['Conta Bancaria'] = df.iloc[:, 10]
    
    # A coluna L (índice 11) é Gestor
    if len(colunas_disponiveis) > 11:
        novo_df['Gestor'] = df.iloc[:, 11]
    elif 'Solicitante' in novo_df.columns:
        novo_df['Gestor'] = novo_df['Solicitante']
    else:
        novo_df['Gestor'] = 'Gestor não especificado'
    
    # Coluna W (índice 22) é Placa Cavalo/Carreta
    if len(colunas_disponiveis) > 22:
        novo_df['Placa'] = df.iloc[:, 22]
    
    # Adicionar colunas para Ordem de Serviço e Empresa se existirem no arquivo
    # Ordem de Serviço - assumindo que pode estar na coluna M (índice 12) ou outra
    if len(colunas_disponiveis) > 12:
        novo_df['Ordem de Serviço'] = df.iloc[:, 12]
    else:
        novo_df['Ordem de Serviço'] = 'Não informada'
    
    # Empresa - assumindo que pode estar na coluna N (índice 13) ou outra
    if len(colunas_disponiveis) > 13:
        novo_df['Empresa'] = df.iloc[:, 13]
    else:
        novo_df['Empresa'] = 'Não informada'
    
    # Criar colunas Ano e Mês
    if 'Criado' in novo_df.columns:
        novo_df['Ano'] = novo_df['Criado'].dt.year
        novo_df['Mes'] = novo_df['Criado'].dt.month
        novo_df['Dia'] = novo_df['Criado'].dt.day
    
    # Remover linhas com valores NaN em colunas essenciais
    colunas_essenciais = []
    if 'Criado' in novo_df.columns:
        colunas_essenciais.append('Criado')
    if 'Valor' in novo_df.columns:
        colunas_essenciais.append('Valor')
    if 'Solicitante' in novo_df.columns:
        colunas_essenciais.append('Solicitante')
    
    if colunas_essenciais:
        novo_df = novo_df.dropna(subset=colunas_essenciais)
    
    return novo_df

def get_label_color():
    # Detectar tema escuro no Streamlit
    theme = st.get_option("theme.base")
    return "white" if theme == "dark" else "black"

def convert_df(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def get_default_options(available_options, default_list):
    return [opt for opt in default_list if opt in available_options]

def gerar_projecao_mes_atual(df):
    hoje = datetime.today().date()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia = (primeiro_dia + pd.offsets.MonthEnd(0)).date()

    # Filtrar apenas os dados do mês atual
    df_mes = df[(df['Criado'].dt.date >= primeiro_dia) & (df['Criado'].dt.date <= hoje)]

    if df_mes.empty or len(df_mes) < 2:
        # Se não há dados suficientes, retornar dataframe vazio
        return pd.DataFrame(columns=['Data', 'Valor', 'Tipo']), 0, 0

    # Realizado por dia
    realizado = df_mes.groupby(df_mes['Criado'].dt.date)['Valor'].sum().reset_index()
    realizado.columns = ['Data', 'Valor']
    realizado['Tipo'] = 'Realizado'

    # Estatísticas para projeção
    # Separar dias úteis e fins de semana
    dias_uteis = realizado[realizado['Data'].apply(lambda d: d.weekday() < 5)]
    dias_fds = realizado[realizado['Data'].apply(lambda d: d.weekday() >= 5)]
    
    media_uteis = dias_uteis['Valor'].mean() if not dias_uteis.empty else realizado['Valor'].mean()
    media_fds = dias_fds['Valor'].mean() if not dias_fds.empty else (media_uteis * 0.5 if media_uteis > 0 else 0)

    # Dias futuros
    datas_futuras = pd.date_range(hoje + timedelta(days=1), ultimo_dia).date

    previsao = []
    for data in datas_futuras:
        if data.weekday() >= 5:
            valor = media_fds
        else:
            valor = media_uteis
        
        # Adicionar pequena variação aleatória (10%)
        if valor > 0:
            valor = valor * np.random.uniform(0.9, 1.1)
        
        previsao.append({'Data': data, 'Valor': valor, 'Tipo': 'Projetado'})

    df_proj = pd.DataFrame(previsao)
    df_resultado = pd.concat([realizado, df_proj], ignore_index=True)

    total_esperado = df_resultado['Valor'].sum()
    return df_resultado, total_esperado, media_uteis

# Função para obter a data inicial padrão com base no dia da semana
def obter_data_inicio_padrao():
    hoje = datetime.today().date()
    dia_semana = hoje.weekday()  # 0 = segunda, 1 = terça, ..., 6 = domingo
    
    if dia_semana == 0:  # Segunda-feira
        # Na segunda, mostrar sexta, sábado e domingo anteriores
        return hoje - timedelta(days=3)  # Sexta-feira
    else:
        # Nos outros dias, mostrar apenas o dia anterior
        return hoje - timedelta(days=1)

# Função para obter a data final padrão com base no dia da semana
def obter_data_fim_padrao():
    hoje = datetime.today().date()
    dia_semana = hoje.weekday()  # 0 = segunda, 1 = terça, ..., 6 = domingo
    
    if dia_semana == 0:  # Segunda-feira
        # Na segunda, mostrar sexta, sábado e domingo
        return hoje - timedelta(days=1)  # Domingo
    else:
        # Nos outros dias, mostrar apenas o dia anterior
        return hoje - timedelta(days=1)

# ======================== CARREGAR DADOS ========================
try:
    df = load_data("Projeto-custo-diário-solicitações-de-depósitos.xlsx")
    
    # Verificar se os dados foram carregados corretamente
    if df.empty:
        st.error("❌ Nenhum dado foi carregado do arquivo.")
        st.stop()
        
except Exception as e:
    st.error(f"❌ Erro ao carregar o arquivo: {e}")
    st.stop()

# ======================== MENU LATERAL ========================
menu = st.sidebar.radio("📌 Menu", [
    "Dashboard Geral",
    "Análise Detalhada",
    "Reunião Manutenção Corporativa"
])

# ======================== DASHBOARD GERAL ========================
if menu == "Dashboard Geral":
    st.sidebar.header("🧮 Filtros")

    # Definir datas padrão baseadas nos dados
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    data_inicio = st.sidebar.date_input("📅 Data Início", value=min_date)
    data_fim = st.sidebar.date_input("📅 Data Fim", value=max_date)

    # Obter opções únicas para os filtros
    solicitantes_unicos = sorted(df['Solicitante'].dropna().unique())
    status_unicos = sorted(df['Status'].dropna().unique())
    gestores_unicos = sorted(df['Gestor'].dropna().unique())
    classificacoes_unicas = sorted(df['Classificação'].dropna().unique())
    finalidades_unicas = sorted(df['Finalidade'].dropna().unique())

    # Aplicar filtros - por padrão apenas os gestores especificados
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]
    
    solicitante_sel = st.sidebar.multiselect("🙋‍♂️ Solicitante", solicitantes_unicos)
    status_sel = st.sidebar.multiselect("📌 Status", status_unicos, default=status_unicos[:5] if len(status_unicos) > 0 else [])
    gestor_sel = st.sidebar.multiselect("👔 Gestor", gestores_unicos, default=gestores_disponiveis)
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", classificacoes_unicas)  # SEM VALOR PADRÃO - CONSIDERA TODOS
    finalidade_sel = st.sidebar.multiselect("🎯 Finalidade", finalidades_unicas)  # SEM VALOR PADRÃO - CONSIDERA TODOS

    # Aplicar filtros aos dados
    df_filtrado = df.copy()
    df_filtrado = df_filtrado[df_filtrado['Criado'].dt.date.between(data_inicio, data_fim)]
    
    if solicitante_sel:
        df_filtrado = df_filtrado[df_filtrado['Solicitante'].isin(solicitante_sel)]
    if status_sel:
        df_filtrado = df_filtrado[df_filtrado['Status'].isin(status_sel)]
    if gestor_sel:
        df_filtrado = df_filtrado[df_filtrado['Gestor'].isin(gestor_sel)]
    if classif_sel:
        df_filtrado = df_filtrado[df_filtrado['Classificação'].isin(classif_sel)]
    if finalidade_sel:
        df_filtrado = df_filtrado[df_filtrado['Finalidade'].isin(finalidade_sel)]

    st.title("📊 Relatório de custos gerais | Solicitações de Depósitos")
    
    with st.expander("📌 Filtros aplicados", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Período:** {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}")
            st.write(f"**Solicitante:** {solicitante_sel if solicitante_sel else 'Todos'}")
            st.write(f"**Status:** {status_sel if status_sel else 'Todos'}")
        with col2:
            st.write(f"**Gestor:** {gestor_sel if gestor_sel else 'Todos'}")
            st.write(f"**Classificação:** {classif_sel if classif_sel else 'Todos'}")
            st.write(f"**Finalidade:** {finalidade_sel if finalidade_sel else 'Todos'}")

    # Calcular métricas
    custo_total = df_filtrado['Valor'].sum()
    qtd_registros = df_filtrado.shape[0]
    dias_distintos = df_filtrado['Criado'].dt.date.nunique()
    custo_medio_diario = custo_total / dias_distintos if dias_distintos else 0
    custo_medio_por_registro = custo_total / qtd_registros if qtd_registros else 0

    # Mostrar métricas no formato brasileiro
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Custo Total", f"R$ {formatar_brasileiro(custo_total)}")
    col2.metric("📅 Custo Médio Diário", f"R$ {formatar_brasileiro(custo_medio_diario)}")
    col3.metric("📄 Custo Médio por Registro", f"R$ {formatar_brasileiro(custo_medio_por_registro)}")
    col4.metric("📝 Qtd. Registros", f"{qtd_registros:,}")

    st.markdown("### 📈 Projeção de Custos do Mês Atual")
    
    # Gerar projeção
    df_proj_mes, total_projetado, media_diaria_proj = gerar_projecao_mes_atual(df_filtrado)
    
    if not df_proj_mes.empty:
        # Formatar valores para exibição
        df_proj_mes_display = df_proj_mes.copy()
        df_proj_mes_display['Valor_Formatado'] = df_proj_mes_display['Valor'].apply(formatar_brasileiro)
        
        fig_proj_mes = px.bar(df_proj_mes, x='Data', y='Valor', color='Tipo',
                              color_discrete_map={'Realizado': 'blue', 'Projetado': 'green'},
                              labels={'Valor': 'R$', 'Data': 'Data'},
                              title="Projeção de Custos Diários no Mês Atual")

        fig_proj_mes.update_layout(
            barmode='group',
            xaxis_title="Data",
            yaxis_title="Valor (R$)",
            hoverlabel=dict(font_size=25),
            showlegend=True
        )

        # Atualizar textos com formato brasileiro
        fig_proj_mes.update_traces(
            text=df_proj_mes_display['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            textposition='outside',
            textangle=0,
            cliponaxis=False,
            textfont=dict(size=12, color='black')
        )

        st.plotly_chart(fig_proj_mes, use_container_width=True)
        st.success(f"📌 Estimativa de custo total até o fim do mês: **R$ {formatar_brasileiro(total_projetado)}**")
    else:
        st.info("ℹ️ Não há dados suficientes para gerar projeção deste mês.")

    # Gráfico temporal por mês
    st.markdown("### 📅 Custo por Mês")
    
    df_temporal = df_filtrado.groupby(df_filtrado['Criado'].dt.to_period('M')).agg({'Valor': 'sum'}).reset_index()
    df_temporal['Criado'] = df_temporal['Criado'].astype(str)

    if not df_temporal.empty:
        # Formatar valores para exibição
        df_temporal['Valor_Formatado'] = df_temporal['Valor'].apply(formatar_brasileiro)
        
        fig_temporal = px.line(
            df_temporal,
            x='Criado',
            y='Valor',
            markers=True,
            text=df_temporal['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            title="Custo por Mês",
            labels={'Valor': 'R$', 'Criado': 'Mês'}
        )

        fig_temporal.update_traces(
            textposition='top center',
            textfont=dict(size=14, color='black'),
            mode='lines+markers+text'
        )

        fig_temporal.update_layout(
            hoverlabel=dict(font_size=20),
            xaxis_title="Mês",
            yaxis_title="Valor (R$)"
        )

        st.plotly_chart(fig_temporal, use_container_width=True)
    else:
        st.info("ℹ️ Não há dados para mostrar o gráfico temporal.")

    # Gráfico por finalidade
    st.markdown("### 🎯 Custo por Finalidade")
    
    df_finalidade = df_filtrado.groupby('Finalidade')['Valor'].sum().reset_index()
    df_finalidade = df_finalidade.sort_values('Valor', ascending=False).head(20)  # Limitar a top 20

    if not df_finalidade.empty:
        # Formatar valores para exibição
        df_finalidade['Valor_Formatado'] = df_finalidade['Valor'].apply(formatar_brasileiro)
        
        fig_finalidade = px.bar(
            df_finalidade,
            x='Valor',
            y='Finalidade',
            color='Finalidade',
            orientation='h',
            text=df_finalidade['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            labels={'Valor': 'R$', 'Finalidade': 'Finalidade'}
        )

        fig_finalidade.update_traces(
            textposition='outside',
            textfont=dict(size=12, color="black"),
            cliponaxis=False
        )

        fig_finalidade.update_layout(
            xaxis_title="Valor (R$)",
            yaxis_title="Finalidade",
            yaxis=dict(categoryorder='total ascending'),
            hoverlabel=dict(font_size=15),
            margin=dict(l=120, r=40, t=40, b=40),
            height=600
        )

        st.plotly_chart(fig_finalidade, use_container_width=True)
    else:
        st.info("ℹ️ Não há dados para mostrar o gráfico por finalidade.")

    # Gráfico por classificação
    st.markdown("### 🏷️ Custo por Classificação")
    
    df_classificacao = df_filtrado.groupby('Classificação')['Valor'].sum().reset_index()

    if not df_classificacao.empty:
        # Formatar valores para exibição no hover
        df_classificacao['Valor_Formatado'] = df_classificacao['Valor'].apply(formatar_brasileiro)
        
        fig_classificacao = px.pie(df_classificacao, names='Classificação', values='Valor', hole=0.3)
        fig_classificacao.update_traces(
            hovertemplate='%{label}<br>R$ %{customdata}<extra></extra>',
            customdata=df_classificacao['Valor_Formatado'],
            textinfo='label+percent'
        )
        fig_classificacao.update_layout(hoverlabel=dict(font_size=20))
        st.plotly_chart(fig_classificacao, use_container_width=True)
    else:
        st.info("ℹ️ Não há dados para mostrar o gráfico por classificação.")

    # Botão para baixar dados
    if not df_filtrado.empty:
        nome_arquivo = f"dados_filtrados_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        st.download_button(
            "📥 Baixar dados filtrados (Excel)",
            data=convert_df(df_filtrado),
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("⚠️ Não há dados para exportar com os filtros aplicados.")

# ======================== ANÁLISE DETALHADA ========================
elif menu == "Análise Detalhada":
    st.title("👤 Análise por Solicitante")

    # Definir datas padrão baseadas no dia da semana
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    # Obter min_date e max_date dos dados para os inputs
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    data_inicio = st.sidebar.date_input("📅 Data Início", value=data_inicio_padrao, min_value=min_date, max_value=max_date)
    data_fim = st.sidebar.date_input("📅 Data Fim", value=data_fim_padrao, min_value=min_date, max_value=max_date)

    # Obter opções únicas
    status_unicos = sorted(df['Status'].dropna().unique())
    classificacoes_unicas = sorted(df['Classificação'].dropna().unique())
    gestores_unicos = sorted(df['Gestor'].dropna().unique())
    
    # Por padrão apenas os gestores especificados
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]

    # Aplicar filtros básicos
    status_sel = st.sidebar.multiselect("📌 Status", status_unicos, default=status_unicos[:3] if status_unicos else [])
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", classificacoes_unicas)  # SEM VALOR PADRÃO - CONSIDERA TODOS
    gestor_sel = st.sidebar.multiselect("👔 Gestor", gestores_unicos, default=gestores_disponiveis)

    df_filtrado = df.copy()
    df_filtrado = df_filtrado[df_filtrado['Criado'].dt.date.between(data_inicio, data_fim)]
    
    if status_sel:
        df_filtrado = df_filtrado[df_filtrado['Status'].isin(status_sel)]
    if classif_sel:
        df_filtrado = df_filtrado[df_filtrado['Classificação'].isin(classif_sel)]
    if gestor_sel:
        df_filtrado = df_filtrado[df_filtrado['Gestor'].isin(gestor_sel)]

    # Selecionar solicitante
    solicitantes_disponiveis = sorted(df_filtrado['Solicitante'].dropna().unique())
    solicitante_select = st.selectbox("🙋‍♂️ Selecione um Solicitante", options=["Todos"] + solicitantes_disponiveis)

    if solicitante_select != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Solicitante'] == solicitante_select]

        # Cálculos do solicitante
        custo_total = df_filtrado['Valor'].sum()
        qtd_registros = df_filtrado.shape[0]
        custo_medio_por_solicitacao = custo_total / qtd_registros if qtd_registros else 0
        dias_distintos = df_filtrado['Criado'].dt.date.nunique()
        custo_medio_diario = custo_total / dias_distintos if dias_distintos else 0

        col1, col2 = st.columns([1, 3])
        with col1:
            st.image("https://cdn-icons-png.flaticon.com/512/1144/1144760.png", width=120)
            st.subheader(f"{solicitante_select}")
        with col2:
            mcol1, mcol2, mcol3 = st.columns(3)
            mcol1.metric("💰 Custo Total", f"R$ {formatar_brasileiro(custo_total)}")
            mcol2.metric("⚖️ Custo Médio por Solicitação", f"R$ {formatar_brasileiro(custo_medio_por_solicitacao)}")
            mcol3.metric("📅 Custo Médio Diário", f"R$ {formatar_brasileiro(custo_medio_diario)}")

        # Gráfico de evolução temporal
        st.markdown("### 📈 Evolução Temporal")
        if not df_filtrado.empty:
            df_evolucao = df_filtrado.groupby(df_filtrado['Criado'].dt.date)['Valor'].sum().reset_index()
            df_evolucao['Valor_Formatado'] = df_evolucao['Valor'].apply(formatar_brasileiro)
            
            fig_evolucao = px.line(df_evolucao, x='Criado', y='Valor', markers=True,
                                   title=f"Evolução de Custos - {solicitante_select}")
            fig_evolucao.update_traces(
                hovertemplate='Data: %{x|%d/%m/%Y}<br>Valor: R$ %{customdata}<extra></extra>',
                customdata=df_evolucao['Valor_Formatado']
            )
            st.plotly_chart(fig_evolucao, use_container_width=True)
    else:
        st.markdown("### 📊 Resumo Geral")
        total_geral = df_filtrado['Valor'].sum()
        qtd_registros = df_filtrado.shape[0]

        col1, col2 = st.columns(2)
        col1.metric("💰 Custo Total no Período", f"R$ {formatar_brasileiro(total_geral)}")
        col2.metric("📝 Qtd. Registros", f"{qtd_registros:,}")

        st.markdown("### 👥 Custo por Solicitante")
        if not df_filtrado.empty:
            df_por_solicitante = df_filtrado.groupby('Solicitante').agg({
                'Valor': 'sum',
                'ID': 'count'
            }).reset_index()
            df_por_solicitante.rename(columns={'ID': 'QtdSolicitações'}, inplace=True)
            df_por_solicitante = df_por_solicitante.sort_values('Valor', ascending=False).head(20)
            
            # Formatar valores
            df_por_solicitante['Valor_Formatado'] = df_por_solicitante['Valor'].apply(formatar_brasileiro)
            
            fig_solicitante = px.bar(
                df_por_solicitante,
                x='Valor',
                y='Solicitante',
                color='Solicitante',
                orientation='h',
                text=df_por_solicitante.apply(lambda row: f"R$ {row['Valor_Formatado']} ({row['QtdSolicitações']})", axis=1),
                labels={'Valor': 'R$', 'Solicitante': 'Solicitante'}
            )

            fig_solicitante.update_traces(
                textposition='outside',
                textfont=dict(size=12, color="black"),
                cliponaxis=False
            )

            fig_solicitante.update_layout(
                xaxis_title="Valor (R$)",
                yaxis_title="Solicitante",
                yaxis=dict(categoryorder='total ascending'),
                hoverlabel=dict(font_size=15),
                margin=dict(l=120, r=40, t=40, b=40),
                height=600
            )

            st.plotly_chart(fig_solicitante, use_container_width=True)

    # Tabela de dados - AGORA COM AS NOVAS COLUNAS
    st.markdown("### 📋 Tabela de Solicitações")
    if not df_filtrado.empty:
        # Selecionar colunas relevantes para mostrar - INCLUINDO AS NOVAS COLUNAS
        colunas_para_mostrar = ['ID', 'Solicitante', 'Valor', 'Finalidade', 'Classificação', 
                               'Status', 'Criado', 'Ordem de Serviço', 'Empresa']  # Adicionadas as novas colunas
        colunas_existentes = [col for col in colunas_para_mostrar if col in df_filtrado.columns]
        
        # Formatar a tabela
        df_display = df_filtrado[colunas_existentes].copy()
        if 'Criado' in df_display.columns:
            df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Formatar a coluna Valor no formato brasileiro
        if 'Valor' in df_display.columns:
            df_display['Valor'] = df_display['Valor'].apply(formatar_brasileiro)
        
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Botão para download (mantém valores numéricos no Excel)
        nome_arquivo = f"analise_{solicitante_select if solicitante_select != 'Todos' else 'geral'}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        st.download_button(
            "📥 Baixar Dados (Excel)",
            data=convert_df(df_filtrado),
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("⚠️ Não há dados para mostrar com os filtros aplicados.")

# ======================== REUNIÃO MANUTENÇÃO CORPORATIVA ========================
elif menu == "Reunião Manutenção Corporativa":
    st.title("🏗️ Relatório | Reunião de Manutenção Corporativa")

    # Definir datas padrão baseadas no dia da semana
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    # Obter min_date e max_date dos dados para os inputs
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    data_inicio = st.sidebar.date_input("📅 Data Início", value=data_inicio_padrao, min_value=min_date, max_value=max_date)
    data_fim = st.sidebar.date_input("📅 Data Fim", value=data_fim_padrao, min_value=min_date, max_value=max_date)

    # Obter opções únicas
    gestores_unicos = sorted(df['Gestor'].dropna().unique())
    status_unicos = sorted(df['Status'].dropna().unique())
    classificacoes_unicas = sorted(df['Classificação'].dropna().unique())
    
    # Por padrão apenas os gestores especificados
    gestores_padrao = ["Wesley Duarte Assumpção", "José Marcos", "José Wítalo", "Alex de França Silva"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]

    # Aplicar filtros
    gestor_sel = st.sidebar.multiselect("Gestor", gestores_unicos, default=gestores_disponiveis)
    status_sel = st.sidebar.multiselect("Status", status_unicos, default=status_unicos[:3] if status_unicos else [])
    classif_sel = st.sidebar.multiselect("Classificação", classificacoes_unicas)  # SEM VALOR PADRÃO - CONSIDERA TODOS

    # Filtrar dados
    df_rm = df[df['Criado'].dt.date.between(data_inicio, data_fim)]
    
    if gestor_sel:
        df_rm = df_rm[df_rm['Gestor'].isin(gestor_sel)]
    if status_sel:
        df_rm = df_rm[df_rm['Status'].isin(status_sel)]
    if classif_sel:
        df_rm = df_rm[df_rm['Classificação'].isin(classif_sel)]

    # Calcular métricas
    if not df_rm.empty:
        custo_total = df_rm['Valor'].sum()
        qtd_registros = df_rm.shape[0]
        custo_medio = custo_total / qtd_registros if qtd_registros else 0
        
        # Encontrar maior solicitação
        if not df_rm.empty:
            maior_idx = df_rm['Valor'].idxmax()
            maior_solicitacao = df_rm.loc[maior_idx]
        else:
            maior_solicitacao = None
    else:
        custo_total = 0
        qtd_registros = 0
        custo_medio = 0
        maior_solicitacao = None

    # Mostrar métricas no formato brasileiro
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Custo Total", f"R$ {formatar_brasileiro(custo_total)}")
    col2.metric("📋 Custo Médio por Solicitação", f"R$ {formatar_brasileiro(custo_medio)}")
    col3.metric("📝 Qtd. Registros", f"{qtd_registros:,}")
    
    if maior_solicitacao is not None and not pd.isna(maior_solicitacao['Valor']):
        col4.markdown(f"""
            <div style='font-size: 24px; font-weight: bold;'>R$ {formatar_brasileiro(maior_solicitacao['Valor'])}</div>
            <div style='font-size: 14px;'>ID: {maior_solicitacao.get('ID', 'N/A')} - {maior_solicitacao.get('Solicitante', 'N/A')}</div>
        """, unsafe_allow_html=True)
    else:
        col4.metric("🔝 Maior Solicitação", "Sem dados")

    # Gráfico por finalidade
    st.markdown("### 📈 Custo por Finalidade no Período")
    if not df_rm.empty:
        df_final = df_rm.groupby('Finalidade')['Valor'].sum().reset_index()
        df_final = df_final.sort_values('Valor', ascending=False).head(15)
        
        # Formatar valores para exibição
        df_final['Valor_Formatado'] = df_final['Valor'].apply(formatar_brasileiro)

        fig_final = px.bar(
            df_final,
            x='Finalidade',
            y='Valor',
            color='Finalidade',
            text=df_final['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            labels={'Valor': 'R$'}
        )

        fig_final.update_traces(
            textposition='outside',
            textfont=dict(size=12, color="black"),
            cliponaxis=False
        )

        fig_final.update_layout(
            hoverlabel=dict(font_size=15),
            xaxis_title="Finalidade",
            yaxis_title="Valor (R$)",
            margin=dict(t=40, b=100, l=60, r=40),
            xaxis_tickangle=-45
        )

        st.plotly_chart(fig_final, use_container_width=True)
    else:
        st.info("ℹ️ Não há dados para mostrar o gráfico por finalidade.")

    # Projeção
    st.markdown("### 🔮 Projeção de Custo HOJE (base últimos 5 dias)")
    if not df.empty:
        # Usar status selecionado ou todos se nenhum selecionado
        status_filtro = status_sel if status_sel else df['Status'].unique()
        
        ultimos_5_dias = df[(df['Criado'].dt.date >= max_date - timedelta(days=5)) & 
                           (df['Status'].isin(status_filtro))]
        
        if not ultimos_5_dias.empty:
            custo_diario = ultimos_5_dias.groupby(ultimos_5_dias['Criado'].dt.date)['Valor'].sum().reset_index()
            if not custo_diario.empty:
                media_diaria = custo_diario['Valor'].mean()
                st.success(f"🔜 Projeção de custo para HOJE: **R$ {formatar_brasileiro(media_diaria)}**")
            else:
                st.warning("⚠️ Não foi possível calcular a média diária.")
        else:
            st.warning("⚠️ Não há dados suficientes nos últimos 5 dias para estimar uma projeção.")
    else:
        st.warning("⚠️ Não há dados disponíveis para fazer projeções.")

    # Tabela detalhada
    st.markdown("### 📋 Tabela Detalhada")
    if not df_rm.empty:
        # Definir colunas para mostrar - TAMBÉM ADICIONANDO AS NOVAS COLUNAS AQUI
        colunas_desejadas = ['ID', 'Title', 'Valor', 'Finalidade', 'Solicitante', 
                            'Descrição', 'Gestor', 'Classificação', 'Criado',
                            'Ordem de Serviço', 'Empresa']  # Adicionadas as novas colunas
        colunas_existentes = [col for col in colunas_desejadas if col in df_rm.columns]
        
        # Formatar a tabela
        df_display = df_rm[colunas_existentes].copy()
        if 'Criado' in df_display.columns:
            df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Formatar valores no formato brasileiro
        if 'Valor' in df_display.columns:
            df_display['Valor'] = df_display['Valor'].apply(formatar_brasileiro)
        
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Botão para download
        nome_arquivo = f"reuniao_manutencao_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        st.download_button(
            "📥 Baixar Relatório (Excel)",
            data=convert_df(df_rm),
            file_name=nome_arquivo,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("⚠️ Não há dados para mostrar com os filtros aplicados.")
