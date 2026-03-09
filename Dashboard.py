import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import locale
import io

st.set_page_config(page_title="Dashboard de Custos", layout="wide")

@st.cache_data
def load_data(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    df['Criado'] = pd.to_datetime(df['Criado'], errors='coerce')
    df['Ano'] = df['Criado'].dt.year
    df['Mes'] = df['Criado'].dt.month
    return df
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

# Carregar dados
try:
    df = load_data("Projeto-custo-diário-solicitações-de-depósitos.xlsx")
except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")
    st.stop()

# Menu lateral
menu = st.sidebar.radio("📌 Menu", [
    "Dashboard Geral",
    "Análise Detalhada",
    "Reunião Manutenção Corporativa"
])

import numpy as np
def gerar_projecao_mes_atual(df):
    hoje = datetime.today().date()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia = (primeiro_dia + pd.offsets.MonthEnd(0)).date()

    # Filtrar apenas os dados do mês atual
    df_mes = df[(df['Criado'].dt.date >= primeiro_dia) & (df['Criado'].dt.date <= hoje)]

    if df_mes.empty:
        # Retorna vazio para evitar erro ao agrupar
        return pd.DataFrame(columns=['Data', 'Valor', 'Tipo']), 0, 0

    # Realizado por dia
    realizado = df_mes.groupby(df_mes['Criado'].dt.date)['Valor'].sum().reset_index()
    realizado.columns = ['Data', 'Valor']
    realizado['Tipo'] = 'Realizado'

    # Estatísticas para projeção
    media_uteis = realizado[realizado['Data'].apply(lambda d: d.weekday() < 5)]['Valor'].mean()
    media_fds = realizado[realizado['Data'].apply(lambda d: d.weekday() >= 5)]['Valor'].mean()

    if np.isnan(media_fds):
        media_fds = media_uteis * 0.5  # fallback

    # Dias futuros
    datas_futuras = pd.date_range(hoje + timedelta(days=1), ultimo_dia).date

    previsao = []
    for data in datas_futuras:
        if data.weekday() >= 5:
            valor = media_fds * np.random.uniform(0.9, 1.1)
        else:
            valor = media_uteis * np.random.uniform(0.9, 1.1)
        previsao.append({'Data': data, 'Valor': valor, 'Tipo': 'Projetado'})

    df_proj = pd.DataFrame(previsao)
    df_resultado = pd.concat([realizado, df_proj], ignore_index=True)

    total_esperado = df_resultado['Valor'].sum()
    return df_resultado, total_esperado, media_uteis

# ----------------------- DASHBOARD GERAL -----------------------
if menu == "Dashboard Geral":
    st.sidebar.header("🧮 Filtros")

    data_inicio = st.sidebar.date_input("📅 Data Início", value=pd.to_datetime("2025-01-01").date())
    data_fim = st.sidebar.date_input("📅 Data Fim", value=datetime.today().date())

    gestores_padrao = [
        "José Marcos", "Alex de França Silva",
        "Wesley Duarte Assumpcao", "Renan Francisco Cunha"
    ]
    status_padrao = ["Pago"]
    classificacao_padrao = ["Despesa de veículo"]

    solicitante_sel = st.sidebar.multiselect("🙋‍♂️ Solicitante", df['Solicitante'].dropna().unique())
    status_sel = st.sidebar.multiselect("📌 Status", df['Status'].dropna().unique(), default=get_default_options(df['Status'].dropna().unique(), status_padrao))
    gestor_sel = st.sidebar.multiselect("👔 Gestor", df['Gestor'].dropna().unique(), default=get_default_options(df['Gestor'].dropna().unique(), gestores_padrao))
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", df['Classificação'].dropna().unique(), default=get_default_options(df['Classificação'].dropna().unique(), classificacao_padrao))
    finalidade_sel = st.sidebar.multiselect("🎯 Finalidade", df['Finalidade'].dropna().unique())

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
        st.write(f"**Período:** {data_inicio} até {data_fim}")
        st.write(f"**Solicitante:** {solicitante_sel if solicitante_sel else 'Todos'}")
        st.write(f"**Status:** {status_sel if status_sel else 'Todos'}")
        st.write(f"**Gestor:** {gestor_sel if gestor_sel else 'Todos'}")
        st.write(f"**Classificação:** {classif_sel if classif_sel else 'Todos'}")
        st.write(f"**Finalidade:** {finalidade_sel if finalidade_sel else 'Todos'}")

    custo_total = df_filtrado['Valor'].sum()
    qtd_registros = df_filtrado.shape[0]
    dias_distintos = df_filtrado['Criado'].dt.date.nunique()
    custo_medio_diario = custo_total / dias_distintos if dias_distintos else 0
    custo_medio_por_registro = custo_total / qtd_registros if qtd_registros else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Custo Total", f"R$ {custo_total:,.2f}")
    col2.metric("📅 Custo Médio Diário", f"R$ {custo_medio_diario:,.2f}")
    col3.metric("📄 Custo Médio por Registro", f"R$ {custo_medio_por_registro:,.2f}")
    col4.metric("📝 Qtd. Registros", qtd_registros)

    st.markdown("### 📈 Projeção de Custos do Mês Atual")

    df_proj_mes, total_projetado, media_diaria_proj = gerar_projecao_mes_atual(df_filtrado)

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

    fig_proj_mes.update_traces(
    texttemplate='R$ %{y:,.2f}',
    textposition='outside',
    textangle=0,
    cliponaxis=False,
    textfont=dict(size=40, color='black')  # 👈 AQUI
)


    st.plotly_chart(fig_proj_mes, use_container_width=True)

    st.success(f"📌 Estimativa de custo total até o fim do mês: **R$ {total_projetado:,.2f}**")

    df_temporal = df_filtrado.groupby(df_filtrado['Criado'].dt.to_period('M')).agg({'Valor': 'sum'}).reset_index()
    df_temporal['Criado'] = df_temporal['Criado'].astype(str)

    fig_temporal = px.line(
        df_temporal,
        x='Criado',
        y='Valor',
        markers=True,
        text=df_temporal['Valor'].apply(lambda x: f"R$ {x:,.2f}"),
        title="Custo por Mês",
        labels={'Valor': 'R$', 'Criado': 'Mês'}
    )

    fig_temporal.update_traces(
        textposition='top center',  # ✅ posição do valor acima do ponto
        textfont=dict(size=14, color='black'),
        mode='lines+markers+text'  # ✅ exibe linha + ponto + texto
    )

    fig_temporal.update_layout(
        hoverlabel=dict(font_size=20),
        xaxis_title="Mês",
        yaxis_title="Valor (R$)"
    )

    st.plotly_chart(fig_temporal, use_container_width=True)




    st.markdown("### 🎯 Custo por Finalidade")

    # Agrupar os dados por finalidade
    df_finalidade = df_filtrado.groupby('Finalidade')['Valor'].sum().reset_index()

    # Criar gráfico horizontal com os valores fixos ao final da barra
    fig_finalidade = px.bar(
    df_finalidade,
    x='Valor',
    y='Finalidade',
    color='Finalidade',
    orientation='h',
    text=df_finalidade['Valor'].apply(lambda x: f"R$ {x:,.2f}"),  # ✅ Correção AQUI
    labels={'Valor': 'R$', 'Finalidade': 'Finalidade'}
    )

    fig_finalidade.update_traces(
        textposition='outside',
        textfont=dict(size=16, color="black"),
        cliponaxis=False
    )

    # Ajustar layout
    fig_finalidade.update_layout(
        xaxis_title="Valor (R$)",
        yaxis_title="Finalidade",
        yaxis=dict(categoryorder='total ascending'),
        hoverlabel=dict(font_size=15),
        margin=dict(l=120, r=40, t=40, b=40)  # aumenta margem esquerda para texto
    )

    # Exibir gráfico
    st.plotly_chart(fig_finalidade, use_container_width=True)
    

    st.markdown("### 🏷️ Custo por Classificação")
    df_classificacao = df_filtrado.groupby('Classificação')['Valor'].sum().reset_index()
    fig_classificacao = px.pie(df_classificacao, names='Classificação', values='Valor', hole=0.3)
    fig_classificacao.update_traces(
        hovertemplate='%{label}<br>R$ %{value:,.2f}<extra></extra>',
        textinfo='label+percent')
    fig_classificacao.update_layout(hoverlabel=dict(font_size=20))

    st.plotly_chart(fig_classificacao, use_container_width=True)

    nome_arquivo = f"dados_filtrados_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    st.download_button("📥 Baixar dados filtrados (Excel)", data=convert_df(df_filtrado), file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ----------------------- ANÁLISE DETALHADA -----------------------
elif menu == "Análise Detalhada":
    st.title("👤 Análise por Solicitante")

    ontem = (datetime.today() - timedelta(days=1)).date()

    data_inicio = st.sidebar.date_input("📅 Data Início", value=ontem)
    data_fim = st.sidebar.date_input("📅 Data Fim", value=ontem)

    status_padrao = ["Pago"]
    classificacao_padrao = ["Despesa de Veiculo"]
    gestores_padrao = ["José Marcos", "Alex de França Silva", "Wesley Duarte Assumpcao", "Renan Francisco Cunha"]

    status_sel = st.sidebar.multiselect("📌 Status", df['Status'].dropna().unique(),
                                        default=get_default_options(df['Status'].dropna().unique(), status_padrao))
    classif_sel = st.sidebar.multiselect("🏷️ Classificação", df['Classificação'].dropna().unique(),
                                         default=get_default_options(df['Classificação'].dropna().unique(), classificacao_padrao))
    gestor_sel = st.sidebar.multiselect("👔 Gestor", df['Gestor'].dropna().unique(),
                                        default=get_default_options(df['Gestor'].dropna().unique(), gestores_padrao))

    df_filtrado = df.copy()
    df_filtrado = df_filtrado[df_filtrado['Criado'].dt.date.between(data_inicio, data_fim)]
    if status_sel:
        df_filtrado = df_filtrado[df_filtrado['Status'].isin(status_sel)]
    if classif_sel:
        df_filtrado = df_filtrado[df_filtrado['Classificação'].isin(classif_sel)]
    if gestor_sel:
        df_filtrado = df_filtrado[df_filtrado['Gestor'].isin(gestor_sel)]

    solicitantes_disponiveis = sorted(df_filtrado['Solicitante'].dropna().unique())
    solicitante_select = st.selectbox("🙋‍♂️ Selecione um Solicitante", options=["Todos"] + solicitantes_disponiveis)

    # Aplicar filtro extra se um solicitante for escolhido
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
            mcol1.metric("💰 Custo Total", f"R$ {custo_total:,.2f}")
            mcol2.metric("⚖️ Custo Médio por Solicitação", f"R$ {custo_medio_por_solicitacao:,.2f}")
            mcol3.metric("📅 Custo Médio Diário", f"R$ {custo_medio_diario:,.2f}")

    else:
        st.markdown("### 📊 Resumo Geral")
        total_geral = df_filtrado['Valor'].sum()
        qtd_registros = df_filtrado.shape[0]

        col1, col2 = st.columns(2)
        col1.metric("💰 Custo Total no Período", f"R$ {total_geral:,.2f}")
        col2.metric("📝 Qtd. Registros", qtd_registros)


        st.markdown("### 👥 Custo por Solicitante")
        df_por_solicitante = df_filtrado.groupby('Solicitante')['Valor'].sum().reset_index()

        df_por_solicitante = df_filtrado.groupby('Solicitante').agg({
            'Valor': 'sum',
            'ID': 'count'  # ou outro campo não nulo para contar registros
        }).reset_index()
        df_por_solicitante.rename(columns={'ID': 'QtdSolicitações'}, inplace=True)

        fig_solicitante = px.bar(
            df_por_solicitante,
            x='Valor',
            y='Solicitante',
            color='Solicitante',
            orientation='h',
            text=df_por_solicitante.apply(lambda row: f"R$ {row['Valor']:,.2f} ({row['QtdSolicitações']})", axis=1),
            labels={'Valor': 'R$', 'Solicitante': 'Solicitante'},
            custom_data=['QtdSolicitações']  # 👉 Aqui está a diferença
        )

        fig_solicitante.update_traces(
            textposition='outside',
            textfont=dict(size=14, color="black"),
            cliponaxis=False,
            hovertemplate='<b>%{y}</b><br>Valor Total: R$ %{x:,.2f}<br>Qtd Solicitações: %{customdata[0]}<extra></extra>'
        )



        fig_solicitante.update_traces(
            textposition='outside',
            textfont=dict(size=14, color="black"),
            cliponaxis=False
        )

        fig_solicitante.update_layout(
            xaxis_title="Valor (R$)",
            yaxis_title="Solicitante",
            yaxis=dict(categoryorder='total ascending'),
            hoverlabel=dict(font_size=20),
            margin=dict(l=120, r=40, t=40, b=40)
        )

        st.plotly_chart(fig_solicitante, use_container_width=True)

    # 🔄 Custo por finalidade (respeita o filtro do solicitante)
    st.markdown("### 🎯 Custo por Finalidade")
    df_finalidade = df_filtrado.groupby('Finalidade')['Valor'].sum().reset_index()

    fig_finalidade = px.bar(
        df_finalidade,
        x='Finalidade',
        y='Valor',
        color='Finalidade',
        text=df_finalidade['Valor'].apply(lambda x: f"R$ {x:,.2f}"),
        labels={'Valor': 'R$'}
    )
    fig_finalidade.update_traces(
        textposition='outside',
        textfont=dict(size=14, color="black"),
        cliponaxis=False
    )
    fig_finalidade.update_layout(
        xaxis={'categoryorder': 'total descending'},
        yaxis_title="Valor (R$)",
        hoverlabel=dict(font_size=20)
    )
    st.plotly_chart(fig_finalidade, use_container_width=True)

    # 🔄 Tabela respeita filtros
    st.markdown("### 📋 Tabela de Solicitações")
    st.dataframe(df_filtrado.style.format({'Valor': 'R$ {:,.2f}'}), use_container_width=True)

    nome_arquivo = f"analise_{solicitante_select}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    st.download_button("📥 Baixar Dados (Excel)", data=convert_df(df_filtrado), file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ----------------------- REUNIÃO MANUTENÇÃO CORPORATIVA -----------------------
elif menu == "Reunião Manutenção Corporativa":
    st.title("🏗️ Relatório | Reunião de Manutenção Corporativa")

    ontem = (datetime.today() - timedelta(days=1)).date()
    data_inicio = st.sidebar.date_input("📅 Data Início", value=ontem)
    data_fim = st.sidebar.date_input("📅 Data Fim", value=ontem)

    gestores_default = ["José Marcos", "Alex de França Silva", "Wesley Duarte Assumpcao", "Renan Francisco Cunha"]
    status_default = ["Pago"]
    classif_default = ["Despesa de Veiculo"]

    gestor_sel = st.sidebar.multiselect("Gestor", df['Gestor'].dropna().unique(),
                                        default=get_default_options(df['Gestor'].dropna().unique(), gestores_default))
    status_sel = st.sidebar.multiselect("Status", df['Status'].dropna().unique(),
                                        default=get_default_options(df['Status'].dropna().unique(), status_default))
    classif_sel = st.sidebar.multiselect("Classificação", df['Classificação'].dropna().unique(),
                                         default=get_default_options(df['Classificação'].dropna().unique(), classif_default))

    df_rm = df[df['Criado'].dt.date.between(data_inicio, data_fim)]
    df_rm = df_rm[df_rm['Gestor'].isin(gestor_sel)]
    df_rm = df_rm[df_rm['Status'].isin(status_sel)]
    df_rm = df_rm[df_rm['Classificação'].isin(classif_sel)]

    custo_total = df_rm['Valor'].sum()
    qtd_registros = df_rm.shape[0]
    custo_medio = custo_total / qtd_registros if qtd_registros else 0
    maior_solicitacao = df_rm.loc[df_rm['Valor'].idxmax()] if not df_rm.empty else None

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Custo Total", f"R$ {custo_total:,.2f}")
    col2.metric("📋 Custo Médio por Solicitação", f"R$ {custo_medio:,.2f}")
    col3.metric("📝 Qtd. Registros", qtd_registros)
    if maior_solicitacao is not None:
        col4.markdown(f"""
            <div style='font-size: 24px; font-weight: bold;'>R$ {maior_solicitacao['Valor']:,.2f}</div>
            <div style='font-size: 14px;'>ID: {maior_solicitacao['ID']} - {maior_solicitacao['Solicitante']}</div>
        """, unsafe_allow_html=True)
    else:
        col4.metric("🔝 Maior Solicitação", "Sem dados")


    st.markdown("### 📈 Custo por Finalidade no Período")
    df_final = df_rm.groupby('Finalidade')['Valor'].sum().reset_index()

    fig_final = px.bar(
        df_final,
        x='Finalidade',
        y='Valor',
        color='Finalidade',
        text=df_final['Valor'].apply(lambda x: f"R$ {x:,.2f}"),
        labels={'Valor': 'R$'}
    )

    fig_final.update_traces(
        textposition='outside',
        textfont=dict(size=14, color="black"),
        cliponaxis=False
    )

    fig_final.update_layout(
        hoverlabel=dict(font_size=20),
        xaxis_title="Finalidade",
        yaxis_title="Valor (R$)",
        margin=dict(t=40, b=40, l=60, r=40)
    )

    st.plotly_chart(fig_final, use_container_width=True)

    st.markdown("### 🔮 Projeção de Custo HOJE (base últimos 5 dias)")
    ultimos_5_dias = df[(df['Criado'] >= datetime.today() - timedelta(days=5)) & (df['Status'].isin(status_sel))]
    custo_diario = ultimos_5_dias.groupby(ultimos_5_dias['Criado'].dt.date)['Valor'].sum().reset_index()
    if not custo_diario.empty:
        media_diaria = custo_diario['Valor'].mean()
        st.success(f"🔜 Projeção de custo para HOJE: **R$ {media_diaria:,.2f}**")
    else:
        st.warning("⚠️ Não há dados suficientes nos últimos 5 dias para estimar uma projeção.")

    st.markdown("### 📋 Tabela Detalhada")
    colunas_desejadas = ['ID', 'Title', 'Valor', 'Finalidade', 'Solicitante', 'Descrição', 'Gestor', 'Nome do favorecido']
    colunas_existentes = [col for col in colunas_desejadas if col in df_rm.columns]
    st.dataframe(df_rm[colunas_existentes].style.format({'Valor': 'R$ {:,.2f}'}), use_container_width=True)

    import time
    import threading
    import webbrowser           

    def abrir_navegador():
        time.sleep(2)
        webbrowser.open("http://localhost:8501")
        
    threading.Thread(target=abrir_navegador).start()
