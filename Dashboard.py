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
        # Se já for string no formato brasileiro, converter para float primeiro
        if isinstance(valor, str):
            # Remover caracteres não numéricos exceto vírgula, ponto e sinal negativo
            valor_limpo = re.sub(r'[^\d,\-\.]', '', valor)
            
            # Se tem vírgula e ponto, assumir que vírgula é decimal
            if ',' in valor_limpo and '.' in valor_limpo:
                # Remover pontos de milhar
                valor_limpo = valor_limpo.replace('.', '')
                # Substituir vírgula por ponto
                valor_limpo = valor_limpo.replace(',', '.')
            elif ',' in valor_limpo:
                # Se só tem vírgula, verificar se é decimal ou milhar
                partes = valor_limpo.split(',')
                if len(partes[-1]) == 2 or len(partes[-1]) == 3:  # Provavelmente decimal
                    valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
                else:
                    # Pode ser milhar europeu
                    valor_limpo = valor_limpo.replace(',', '')
            
            valor_float = float(valor_limpo)
        else:
            valor_float = float(valor)
        
        # Formatar com separador de milhar e vírgula decimal
        format_str = f"{{:,.{decimais}f}}"
        formatted = format_str.format(valor_float)
        # Substituir ponto por vírgula e vírgula por ponto
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        # Fallback simples
        try:
            return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(valor)

@st.cache_data
def load_data(caminho_arquivo):
    try:
        # Ler o arquivo Excel mantendo todos os dados
        df = pd.read_excel(caminho_arquivo, dtype=str)  # Ler tudo como string inicialmente
        
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        return pd.DataFrame()
    
    # Criar novo DataFrame com mapeamento mais flexível
    novo_df = pd.DataFrame()
    
    # Primeiro, tentar identificar colunas por nome
    col_map = {}
    
    # Mapear nomes de colunas possíveis (case insensitive)
    for col in df.columns:
        col_lower = str(col).lower().strip()
        
        if 'id' in col_lower or 'código' in col_lower:
            col_map['ID'] = col
        elif 'title' in col_lower or 'título' in col_lower:
            col_map['Title'] = col
        elif 'status' in col_lower or 'situação' in col_lower:
            col_map['Status'] = col
        elif 'classificação' in col_lower or 'categoria' in col_lower or 'tipo' in col_lower:
            col_map['Classificação'] = col
        elif 'finalidade' in col_lower or 'descrição' in col_lower:
            col_map['Finalidade'] = col
        elif 'descrição' in col_lower and 'Finalidade' not in col_map:
            col_map['Descrição'] = col
        elif 'solicitante' in col_lower or 'requerente' in col_lower:
            col_map['Solicitante'] = col
        elif 'motorista' in col_lower and 'nome' in col_lower:
            col_map['Nome Motorista'] = col
        elif 'valor' in col_lower or 'custo' in col_lower or 'preço' in col_lower:
            col_map['Valor'] = col
        elif 'criado' in col_lower or 'data' in col_lower or 'criação' in col_lower:
            col_map['Criado'] = col
        elif 'cpf' in col_lower:
            col_map['CPF Motorista'] = col
        elif 'conta' in col_lower or 'banc' in col_lower:
            col_map['Conta Bancaria'] = col
        elif 'gestor' in col_lower or 'responsável' in col_lower:
            col_map['Gestor'] = col
        elif 'placa' in col_lower:
            col_map['Placa'] = col
    
    # Se não encontrou por nome, usar mapeamento por posição como fallback
    for i, col in enumerate(df.columns):
        if i == 0 and 'ID' not in col_map:
            col_map['ID'] = col
        elif i == 1 and 'Title' not in col_map:
            col_map['Title'] = col
        elif i == 2 and 'Status' not in col_map:
            col_map['Status'] = col
        elif i == 3 and 'Classificação' not in col_map:
            col_map['Classificação'] = col
        elif i == 4 and 'Finalidade' not in col_map:
            col_map['Finalidade'] = col
        elif i == 5 and 'Descrição' not in col_map:
            col_map['Descrição'] = col
        elif i == 6 and 'Solicitante' not in col_map:
            col_map['Solicitante'] = col
        elif i == 7 and 'Nome Motorista' not in col_map:
            col_map['Nome Motorista'] = col
        elif i == 8 and 'Valor' not in col_map:
            col_map['Valor'] = col
        elif i == 9 and 'CPF Motorista' not in col_map:
            col_map['CPF Motorista'] = col
        elif i == 10 and 'Conta Bancaria' not in col_map:
            col_map['Conta Bancaria'] = col
        elif i == 11 and 'Gestor' not in col_map:
            col_map['Gestor'] = col
        elif i == 14 and 'Criado' not in col_map:
            col_map['Criado'] = col
        elif i == 22 and 'Placa' not in col_map:
            col_map['Placa'] = col
    
    # Criar as colunas no novo dataframe
    for target_col, source_col in col_map.items():
        novo_df[target_col] = df[source_col]
    
    # Verificar colunas obrigatórias e criar se não existirem
    if 'ID' not in novo_df.columns:
        novo_df['ID'] = range(1, len(df) + 1)
    
    if 'Status' not in novo_df.columns:
        novo_df['Status'] = 'Pago'
    
    if 'Classificação' not in novo_df.columns:
        novo_df['Classificação'] = 'Despesa de veículo'
    
    # CORREÇÃO: Garantir que 'Finalidade' seja distinta de 'Title'
    if 'Finalidade' not in novo_df.columns:
        if 'Descrição' in novo_df.columns:
            novo_df['Finalidade'] = novo_df['Descrição']
        elif 'Title' in novo_df.columns:
            # Usar Title como fallback, mas tentar extrair finalidade
            novo_df['Finalidade'] = novo_df['Title'].apply(lambda x: str(x)[:100] if pd.notna(x) else 'Outros')
        else:
            novo_df['Finalidade'] = 'Outros'
    
    # Se Finalidade for igual a Title, tentar usar outra coluna
    if 'Finalidade' in novo_df.columns and 'Title' in novo_df.columns:
        # Verificar se as colunas são idênticas
        if novo_df['Finalidade'].equals(novo_df['Title']):
            # Se for igual, usar Descrição se disponível
            if 'Descrição' in novo_df.columns:
                novo_df['Finalidade'] = novo_df['Descrição']
            else:
                # Criar uma finalidade baseada na classificação
                if 'Classificação' in novo_df.columns:
                    novo_df['Finalidade'] = novo_df['Classificação']
    
    if 'Solicitante' not in novo_df.columns:
        novo_df['Solicitante'] = 'Não informado'
    
    if 'Gestor' not in novo_df.columns:
        # Lista de gestores padrão
        gestores_padrao = ["Wesley Duarte Assumpção", "Alex de França Silva", "José Wítalo", "José Marcos"]
        # Atribuir gestor com base no solicitante ou usar padrão
        if 'Solicitante' in novo_df.columns:
            # Mapear solicitantes para gestores
            def atribuir_gestor(solicitante):
                solicitante_str = str(solicitante).lower()
                if 'wesley' in solicitante_str:
                    return "Wesley Duarte Assumpção"
                elif 'alex' in solicitante_str:
                    return "Alex de França Silva"
                elif 'josé wítalo' in solicitante_str or 'jose witalo' in solicitante_str:
                    return "José Wítalo"
                elif 'josé marcos' in solicitante_str or 'jose marcos' in solicitante_str:
                    return "José Marcos"
                else:
                    # Distribuir aleatoriamente entre os gestores padrão
                    return np.random.choice(gestores_padrao)
            
            novo_df['Gestor'] = novo_df['Solicitante'].apply(atribuir_gestor)
        else:
            novo_df['Gestor'] = 'Gestor não especificado'
    
    # Converter coluna de VALOR para numérico
    if 'Valor' in novo_df.columns:
        def converter_valor(valor):
            if pd.isna(valor):
                return 0.0
            
            # Se já for numérico
            if isinstance(valor, (int, float, np.integer, np.floating)):
                return float(valor)
            
            # Converter string
            valor_str = str(valor).strip()
            
            # Se vazio
            if not valor_str:
                return 0.0
            
            # Remover espaços
            valor_str = valor_str.replace(' ', '')
            
            # Se começar com R$
            if valor_str.startswith('R$'):
                valor_str = valor_str[2:].strip()
            
            # Se tem vírgula e ponto
            if ',' in valor_str and '.' in valor_str:
                # Remover pontos (separadores de milhar)
                valor_str = valor_str.replace('.', '')
                # Substituir vírgula por ponto decimal
                valor_str = valor_str.replace(',', '.')
            elif ',' in valor_str:
                # Verificar se vírgula é decimal ou separador de milhar
                if valor_str.count(',') == 1:
                    partes = valor_str.split(',')
                    # Se a parte depois da vírgula tem 2 ou 3 dígitos, é decimal
                    if len(partes[1]) <= 3:
                        valor_str = valor_str.replace(',', '.')
                    else:
                        # Pode ser separador de milhar europeu
                        valor_str = valor_str.replace(',', '')
                else:
                    # Múltiplas vírgulas - provavelmente separadores de milhar
                    valor_str = valor_str.replace(',', '')
            
            # Remover caracteres não numéricos exceto ponto e sinal negativo
            valor_str = re.sub(r'[^\d\.\-]', '', valor_str)
            
            # Se vazio após limpeza
            if not valor_str:
                return 0.0
            
            # Converter para float
            try:
                return float(valor_str)
            except:
                return 0.0
        
        novo_df['Valor'] = novo_df['Valor'].apply(converter_valor)
    else:
        novo_df['Valor'] = 0.0
    
    # Converter coluna de DATA
    if 'Criado' in novo_df.columns:
        def converter_data(data_str):
            if pd.isna(data_str):
                return pd.NaT
            
            # Se já for datetime
            if isinstance(data_str, (datetime, pd.Timestamp)):
                return pd.to_datetime(data_str)
            
            data_str = str(data_str).strip()
            
            # Tentar formatos comuns
            formatos = [
                '%d/%m/%Y %H:%M:%S',
                '%d/%m/%Y %H:%M',
                '%d/%m/%Y',
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d',
                '%m/%d/%Y %H:%M:%S',
                '%m/%d/%Y'
            ]
            
            for formato in formatos:
                try:
                    return pd.to_datetime(data_str, format=formato)
                except:
                    continue
            
            # Tentar parsing automático
            try:
                return pd.to_datetime(data_str, errors='coerce')
            except:
                return pd.NaT
        
        novo_df['Criado'] = novo_df['Criado'].apply(converter_data)
        
        # Preencher datas faltantes com data mais recente
        if novo_df['Criado'].isna().any():
            # Usar datas sequenciais baseadas no índice
            datas_base = pd.date_range(start='2026-01-01', periods=len(novo_df), freq='D')
            for i in range(len(novo_df)):
                if pd.isna(novo_df.loc[i, 'Criado']):
                    novo_df.loc[i, 'Criado'] = datas_base[i]
    else:
        # Criar datas fictícias
        novo_df['Criado'] = pd.date_range(start='2026-01-01', periods=len(novo_df), freq='D')
    
    # Criar colunas Ano e Mês
    novo_df['Ano'] = novo_df['Criado'].dt.year
    novo_df['Mes'] = novo_df['Criado'].dt.month
    novo_df['Dia'] = novo_df['Criado'].dt.day
    
    # Remover linhas com valores essenciais faltantes
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

    if df_mes.empty:
        # Se não há dados, retornar dataframe vazio
        return pd.DataFrame(columns=['Data', 'Valor', 'Tipo']), 0, 0

    # Realizado por dia
    realizado = df_mes.groupby(df_mes['Criado'].dt.date)['Valor'].sum().reset_index()
    realizado.columns = ['Data', 'Valor']
    realizado['Tipo'] = 'Realizado'

    # Calcular média diária REAL dos dados existentes
    if len(realizado) >= 3:  # Precisa de dados suficientes para projeção
        # Estatísticas para projeção
        # Separar dias úteis e fins de semana
        dias_uteis = realizado[realizado['Data'].apply(lambda d: d.weekday() < 5)]
        dias_fds = realizado[realizado['Data'].apply(lambda d: d.weekday() >= 5)]
        
        media_uteis = dias_uteis['Valor'].mean() if not dias_uteis.empty else realizado['Valor'].mean()
        media_fds = dias_fds['Valor'].mean() if not dias_fds.empty else (media_uteis * 0.3 if media_uteis > 0 else 0)  # Reduzido para 30% em fins de semana
    else:
        # Se poucos dados, usar média simples
        media_uteis = realizado['Valor'].mean() if not realizado.empty else 0
        media_fds = media_uteis * 0.3  # 30% em fins de semana

    # Dias futuros
    datas_futuras = pd.date_range(hoje + timedelta(days=1), ultimo_dia).date

    previsao = []
    for data in datas_futuras:
        if data.weekday() >= 5:  # Fim de semana
            valor = media_fds
        else:  # Dia útil
            valor = media_uteis
        
        # Adicionar pequena variação aleatória (±5%) apenas para visualização
        if valor > 0:
            valor = valor * np.random.uniform(0.95, 1.05)
        
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

# Função para ajustar data padrão se estiver fora do intervalo
def ajustar_data_padrao(data_padrao, min_date, max_date):
    if data_padrao < min_date:
        return min_date
    elif data_padrao > max_date:
        return max_date
    else:
        return data_padrao

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

    # CORREÇÃO: Lista de gestores padrão exata
    gestores_padrao = ["Wesley Duarte Assumpção", "Alex de França Silva", "José Wítalo", "José Marcos"]
    
    # Filtrar apenas os gestores que existem nos dados
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]
    
    # Se não encontrar nenhum, usar os primeiros gestores disponíveis
    if not gestores_disponiveis and gestores_unicos:
        gestores_disponiveis = gestores_unicos[:min(4, len(gestores_unicos))]
    
    solicitante_sel = st.sidebar.multiselect("🙋‍♂️ Solicitante", solicitantes_unicos)
    status_sel = st.sidebar.multiselect("📌 Status", status_unicos, default=status_unicos[:5] if len(status_unicos) > 0 else [])
    gestor_sel = st.sidebar.multiselect("👔 Gestor", gestores_unicos, default=gestores_disponiveis)
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", classificacoes_unicas)
    finalidade_sel = st.sidebar.multiselect("🎯 Finalidade", finalidades_unicas)

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

    # Calcular métricas CORRETAMENTE
    if not df_filtrado.empty:
        custo_total = df_filtrado['Valor'].sum()
        qtd_registros = df_filtrado.shape[0]
        
        # Calcular número de dias úteis DISTINTOS no período
        dias_distintos = df_filtrado['Criado'].dt.date.nunique()
        
        # Se não há dias distintos (por exemplo, todos os registros são do mesmo dia)
        if dias_distintos == 0:
            dias_distintos = 1
            
        # Calcular custo médio diário CORRETO
        custo_medio_diario = custo_total / dias_distintos if dias_distintos > 0 else 0
        
        # Calcular custo médio por registro
        custo_medio_por_registro = custo_total / qtd_registros if qtd_registros > 0 else 0
        
        # Mostrar métricas no formato brasileiro
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💰 Custo Total", f"R$ {formatar_brasileiro(custo_total)}")
        col2.metric("📅 Custo Médio Diário", f"R$ {formatar_brasileiro(custo_medio_diario)}")
        col3.metric("📄 Custo Médio por Registro", f"R$ {formatar_brasileiro(custo_medio_por_registro)}")
        col4.metric("📝 Qtd. Registros", f"{qtd_registros:,}")
    else:
        st.warning("⚠️ Não há dados para calcular métricas com os filtros aplicados.")
        # Métricas vazias
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("💰 Custo Total", "R$ 0,00")
        col2.metric("📅 Custo Médio Diário", "R$ 0,00")
        col3.metric("📄 Custo Médio por Registro", "R$ 0,00")
        col4.metric("📝 Qtd. Registros", "0")

    st.markdown("### 📈 Projeção de Custos do Mês Atual")
    
    # Gerar projeção baseada nos dados FILTRADOS
    df_proj_mes, total_projetado, media_diaria_proj = gerar_projecao_mes_atual(df_filtrado)
    
    if not df_proj_mes.empty and total_projetado > 0:
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

        # Atualizar textos com formato brasileiro nos hovers
        fig_proj_mes.update_traces(
            hovertemplate='Data: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>',
            textposition='outside',
            textangle=0,
            cliponaxis=False,
            textfont=dict(size=12, color='black')
        )

        st.plotly_chart(fig_proj_mes, use_container_width=True)
        
        # Calcular custo realizado até agora
        hoje = datetime.today().date()
        primeiro_dia = hoje.replace(day=1)
        
        df_mes_atual = df_filtrado[(df_filtrado['Criado'].dt.date >= primeiro_dia) & 
                                  (df_filtrado['Criado'].dt.date <= hoje)]
        custo_realizado = df_mes_atual['Valor'].sum() if not df_mes_atual.empty else 0
        
        st.success(f"""
        📌 **Resumo da Projeção:**
        - Custo realizado até agora: **R$ {formatar_brasileiro(custo_realizado)}**
        - Projeção até o fim do mês: **R$ {formatar_brasileiro(total_projetado)}**
        - Média diária projetada: **R$ {formatar_brasileiro(media_diaria_proj)}**
        """)
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

    # Gráfico por finalidade (colunas VERTICAIS) - CORRIGIDO
    st.markdown("### 🎯 Custo por Finalidade")
    
    # Agrupar por Finalidade corretamente
    df_finalidade = df_filtrado.groupby('Finalidade')['Valor'].sum().reset_index()
    df_finalidade = df_finalidade.sort_values('Valor', ascending=False).head(15)

    if not df_finalidade.empty:
        # Formatar valores para exibição
        df_finalidade['Valor_Formatado'] = df_finalidade['Valor'].apply(formatar_brasileiro)
        
        # Limitar o tamanho dos labels para melhor visualização
        df_finalidade['Finalidade_Truncada'] = df_finalidade['Finalidade'].apply(
            lambda x: (x[:50] + '...') if len(str(x)) > 50 else str(x)
        )
        
        fig_finalidade = px.bar(
            df_finalidade,
            x='Finalidade_Truncada',
            y='Valor',
            color='Finalidade',
            text=df_finalidade['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            labels={'Valor': 'R$', 'Finalidade_Truncada': 'Finalidade'},
            title="Custo por Finalidade (Top 15)"
        )

        fig_finalidade.update_traces(
            textposition='outside',
            textfont=dict(size=12, color="black"),
            cliponaxis=False
        )

        fig_finalidade.update_layout(
            xaxis_title="Finalidade",
            yaxis_title="Valor (R$)",
            xaxis_tickangle=-45,
            hoverlabel=dict(font_size=15),
            margin=dict(t=40, b=120, l=60, r=40),
            height=500,
            showlegend=False
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

    # Tabela de dados com TODAS as colunas
    st.markdown("### 📋 Tabela de Dados Completa")
    if not df_filtrado.empty:
        # Criar uma cópia para exibição
        df_display = df_filtrado.copy()
        
        # Formatar a coluna Valor no formato brasileiro
        if 'Valor' in df_display.columns:
            df_display['Valor_Formatado'] = df_display['Valor'].apply(formatar_brasileiro)
        
        # Formatar a coluna Criado
        if 'Criado' in df_display.columns:
            df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Selecionar todas as colunas para exibição
        colunas_para_exibir = [
            'ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
            'Descrição', 'Solicitante', 'Nome Motorista', 'Valor_Formatado',
            'CPF Motorista', 'Conta Bancaria', 'Gestor', 'Placa', 'Criado',
            'Ano', 'Mes', 'Dia'
        ]
        
        # Filtrar apenas colunas que existem
        colunas_existentes = [col for col in colunas_para_exibir if col in df_display.columns]
        
        # Renomear colunas para melhor legibilidade
        df_display = df_display[colunas_existentes].copy()
        df_display = df_display.rename(columns={
            'Valor_Formatado': 'Valor (R$)',
            'CPF Motorista': 'CPF',
            'Conta Bancaria': 'Conta Bancária',
            'Nome Motorista': 'Motorista'
        })
        
        # Mostrar tabela com rolagem
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Botão para baixar dados
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
    st.sidebar.header("🧮 Filtros")

    # Definir datas padrão baseadas no dia da semana
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    # Obter min_date e max_date dos dados
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    # Ajustar datas padrão se necessário
    data_inicio_padrao_ajustada = ajustar_data_padrao(data_inicio_padrao, min_date, max_date)
    data_fim_padrao_ajustada = ajustar_data_padrao(data_fim_padrao, min_date, max_date)
    
    data_inicio = st.sidebar.date_input("📅 Data Início", 
                                       value=data_inicio_padrao_ajustada, 
                                       min_value=min_date, 
                                       max_value=max_date)
    data_fim = st.sidebar.date_input("📅 Data Fim", 
                                    value=data_fim_padrao_ajustada, 
                                    min_value=min_date, 
                                    max_value=max_date)

    # Obter opções únicas
    status_unicos = sorted(df['Status'].dropna().unique())
    classificacoes_unicas = sorted(df['Classificação'].dropna().unique())
    gestores_unicos = sorted(df['Gestor'].dropna().unique())
    
    # CORREÇÃO: Lista de gestores padrão exata
    gestores_padrao = ["Wesley Duarte Assumpção", "Alex de França Silva", "José Wítalo", "José Marcos"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]
    
    # Se não encontrar nenhum, usar os primeiros gestores disponíveis
    if not gestores_disponiveis and gestores_unicos:
        gestores_disponiveis = gestores_unicos[:min(4, len(gestores_unicos))]

    # Aplicar filtros básicos
    status_sel = st.sidebar.multiselect("📌 Status", status_unicos, default=status_unicos[:3] if status_unicos else [])
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", classificacoes_unicas)
    gestor_sel = st.sidebar.multiselect("👔 Gestor", gestores_unicos, default=gestores_disponiveis)

    # Aplicar filtros aos dados
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

    # Tabela de dados com TODAS as colunas
    st.markdown("### 📋 Tabela de Solicitações Completa")
    if not df_filtrado.empty:
        # Criar uma cópia para exibição
        df_display = df_filtrado.copy()
        
        # Formatar a coluna Criado
        if 'Criado' in df_display.columns:
            df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Formatar a coluna Valor no formato brasileiro
        if 'Valor' in df_display.columns:
            df_display['Valor_Formatado'] = df_display['Valor'].apply(formatar_brasileiro)
        
        # Selecionar todas as colunas para exibição
        colunas_para_exibir = [
            'ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
            'Descrição', 'Solicitante', 'Nome Motorista', 'Valor_Formatado',
            'CPF Motorista', 'Conta Bancaria', 'Gestor', 'Placa', 'Criado',
            'Ano', 'Mes', 'Dia'
        ]
        
        # Filtrar apenas colunas que existem
        colunas_existentes = [col for col in colunas_para_exibir if col in df_display.columns]
        
        # Renomear colunas para melhor legibilidade
        df_display = df_display[colunas_existentes].copy()
        df_display = df_display.rename(columns={
            'Valor_Formatado': 'Valor (R$)',
            'CPF Motorista': 'CPF',
            'Conta Bancaria': 'Conta Bancária',
            'Nome Motorista': 'Motorista'
        })
        
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
    st.sidebar.header("🧮 Filtros")

    # Definir datas padrão baseadas no dia da semana
    data_inicio_padrao = obter_data_inicio_padrao()
    data_fim_padrao = obter_data_fim_padrao()
    
    # Obter min_date e max_date dos dados
    min_date = df['Criado'].min().date()
    max_date = df['Criado'].max().date()
    
    # Ajustar datas padrão se necessário
    data_inicio_padrao_ajustada = ajustar_data_padrao(data_inicio_padrao, min_date, max_date)
    data_fim_padrao_ajustada = ajustar_data_padrao(data_fim_padrao, min_date, max_date)
    
    data_inicio = st.sidebar.date_input("📅 Data Início", 
                                       value=data_inicio_padrao_ajustada, 
                                       min_value=min_date, 
                                       max_value=max_date)
    data_fim = st.sidebar.date_input("📅 Data Fim", 
                                    value=data_fim_padrao_ajustada, 
                                    min_value=min_date, 
                                    max_value=max_date)

    # Obter opções únicas
    gestores_unicos = sorted(df['Gestor'].dropna().unique())
    status_unicos = sorted(df['Status'].dropna().unique())
    classificacoes_unicas = sorted(df['Classificação'].dropna().unique())
    
    # CORREÇÃO: Lista de gestores padrão exata
    gestores_padrao = ["Wesley Duarte Assumpção", "Alex de França Silva", "José Wítalo", "José Marcos"]
    gestores_disponiveis = [g for g in gestores_padrao if g in gestores_unicos]
    
    # Se não encontrar nenhum, usar os primeiros gestores disponíveis
    if not gestores_disponiveis and gestores_unicos:
        gestores_disponiveis = gestores_unicos[:min(4, len(gestores_unicos))]

    # Aplicar filtros
    gestor_sel = st.sidebar.multiselect("Gestor", gestores_unicos, default=gestores_disponiveis)
    status_sel = st.sidebar.multiselect("Status", status_unicos, default=status_unicos[:3] if status_unicos else [])
    classif_sel = st.sidebar.multiselect("Classificação", classificacoes_unicas)

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

    # Gráfico por finalidade (colunas HORIZONTAIS) - CORRIGIDO
    st.markdown("### 📈 Custo por Finalidade no Período")
    if not df_rm.empty:
        df_final = df_rm.groupby('Finalidade')['Valor'].sum().reset_index()
        df_final = df_final.sort_values('Valor', ascending=False).head(15)
        
        # Formatar valores para exibição
        df_final['Valor_Formatado'] = df_final['Valor'].apply(formatar_brasileiro)
        
        # Limitar o tamanho dos labels para melhor visualização
        df_final['Finalidade_Truncada'] = df_final['Finalidade'].apply(
            lambda x: (x[:60] + '...') if len(str(x)) > 60 else str(x)
        )

        fig_final = px.bar(
            df_final,
            x='Valor',
            y='Finalidade_Truncada',
            color='Finalidade',
            orientation='h',
            text=df_final['Valor_Formatado'].apply(lambda x: f"R$ {x}"),
            labels={'Valor': 'R$', 'Finalidade_Truncada': 'Finalidade'},
            title="Custo por Finalidade (Top 15)"
        )

        fig_final.update_traces(
            textposition='outside',
            textfont=dict(size=12, color="black"),
            cliponaxis=False
        )

        fig_final.update_layout(
            xaxis_title="Valor (R$)",
            yaxis_title="Finalidade",
            yaxis=dict(categoryorder='total ascending'),
            hoverlabel=dict(font_size=15),
            margin=dict(l=150, r=40, t=40, b=40),
            height=600,
            showlegend=False
        )

        st.plotly_chart(fig_final, use_container_width=True)
    else:
        st.info("ℹ️ Não há dados para mostrar o gráfico por finalidade.")

    # Projeção baseada nos últimos 5 dias do período selecionado
    st.markdown("### 🔮 Projeção de Custo HOJE (base últimos 5 dias do período)")
    if not df_rm.empty:
        # Calcular a média dos últimos 5 dias do período selecionado
        data_max_periodo = df_rm['Criado'].max().date()
        data_min_projecao = data_max_periodo - timedelta(days=4)
        
        ultimos_5_dias = df_rm[(df_rm['Criado'].dt.date >= data_min_projecao) & 
                              (df_rm['Criado'].dt.date <= data_max_periodo)]
        
        if not ultimos_5_dias.empty and len(ultimos_5_dias) > 0:
            custo_diario = ultimos_5_dias.groupby(ultimos_5_dias['Criado'].dt.date)['Valor'].sum().reset_index()
            if not custo_diario.empty:
                media_diaria = custo_diario['Valor'].mean()
                st.success(f"🔜 Projeção de custo para HOJE (base média dos últimos 5 dias): **R$ {formatar_brasileiro(media_diaria)}**")
            else:
                st.warning("⚠️ Não foi possível calcular a média diária.")
        else:
            st.warning("⚠️ Não há dados suficientes nos últimos 5 dias do período para estimar uma projeção.")
    else:
        st.warning("⚠️ Não há dados disponíveis para fazer projeções.")

    # Tabela detalhada com TODAS as colunas
    st.markdown("### 📋 Tabela Detalhada Completa")
    if not df_rm.empty:
        # Criar uma cópia para exibição
        df_display = df_rm.copy()
        
        # Formatar a coluna Criado
        if 'Criado' in df_display.columns:
            df_display['Criado'] = df_display['Criado'].dt.strftime('%d/%m/%Y %H:%M')
        
        # Formatar a coluna Valor no formato brasileiro
        if 'Valor' in df_display.columns:
            df_display['Valor_Formatado'] = df_display['Valor'].apply(formatar_brasileiro)
        
        # Selecionar todas as colunas para exibição
        colunas_para_exibir = [
            'ID', 'Title', 'Status', 'Classificação', 'Finalidade', 
            'Descrição', 'Solicitante', 'Nome Motorista', 'Valor_Formatado',
            'CPF Motorista', 'Conta Bancaria', 'Gestor', 'Placa', 'Criado',
            'Ano', 'Mes', 'Dia'
        ]
        
        # Filtrar apenas colunas que existem
        colunas_existentes = [col for col in colunas_para_exibir if col in df_display.columns]
        
        # Renomear colunas para melhor legibilidade
        df_display = df_display[colunas_existentes].copy()
        df_display = df_display.rename(columns={
            'Valor_Formatado': 'Valor (R$)',
            'CPF Motorista': 'CPF',
            'Conta Bancaria': 'Conta Bancária',
            'Nome Motorista': 'Motorista'
        })
        
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
