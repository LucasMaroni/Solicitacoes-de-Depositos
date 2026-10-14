import io
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Dashboard de Custos", layout="wide")

# ----------------------- CONFIGURAÇÕES -----------------------
ARQUIVO_DADOS = Path(__file__).parent / "Projeto-custo-diário-solicitações-de-depósitos.xlsx"

# Critério de custo de manutenção
INICIO_ANALISE = date(2026, 1, 1)
STATUS_PADRAO = ["Pago"]
CLASSIFICACAO_PADRAO = ["Despesa de Veiculo"]
GESTORES_PADRAO = [
    "Dilton Dantas", "José Marcos", "Lucas Martins", "Tiago Baldo Ribeiro", "José Witalo",
    "Alex de França Silva", "Wesley Duarte Assumpcao", "Flávio Marcante", "Carlos Silva Cipriano",
]

# Agrupamento das finalidades para a análise (o que não estiver aqui vira "Operacional / outros")
GRUPOS_FINALIDADE = {
    "Manutenção Corretiva": "Corretiva",
    "Manutenção Preventiva": "Preventiva",
    "Socorro na estrada cavalos": "Socorro e guincho",
    "Socorro na estrada carretas": "Socorro e guincho",
    "Guincho Acidentes": "Socorro e guincho",
    "Guincho Erros da operação": "Socorro e guincho",
    "Borracharia": "Borracharia e baterias",
    "Troca baterias": "Borracharia e baterias",
    "Reformas": "Outros de manutenção",
    "Manutenção Seminovos": "Outros de manutenção",
    "Troca de Lâmpada de Veículo": "Outros de manutenção",
    "lona": "Outros de manutenção",
    "Dedetização de veiculo": "Outros de manutenção",
}
GRUPO_OUTROS = "Operacional / outros"
ORDEM_GRUPOS = ["Corretiva", "Preventiva", "Socorro e guincho", "Borracharia e baterias",
                "Outros de manutenção", GRUPO_OUTROS]
CORES_GRUPOS = dict(zip(ORDEM_GRUPOS, ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#898781"]))
COR_PRINCIPAL = "#2a78d6"

# Placa antiga (ABC1234) ou Mercosul (ABC1D23), com ou sem hífen/espaço
PADRAO_PLACA = r"(?<![A-Z])([A-Z]{3})[\s-]?(\d[A-Z0-9]\d{2})(?!\d)"


# ----------------------- FUNÇÕES AUXILIARES -----------------------
def brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(valor):
    return f"{valor:.1f}%".replace(".", ",")


def extrair_placa(serie):
    m = serie.fillna("").astype(str).str.upper().str.extract(PADRAO_PLACA)
    return m[0] + m[1]


@st.cache_data(show_spinner="Carregando base de solicitações...")
def load_data(caminho_arquivo, modificado_em):
    # modificado_em só serve para invalidar o cache quando o Excel for substituído
    df = pd.read_excel(caminho_arquivo)
    df.columns = df.columns.str.strip()
    for col in ["Status", "Classificação", "Finalidade", "Gestor", "Solicitante"]:
        df[col] = df[col].str.strip().str.replace(r"\s+", " ", regex=True)
    df["Criado"] = pd.to_datetime(df["Criado"], errors="coerce")
    df = df.dropna(subset=["Criado"])
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    df["Ano"] = df["Criado"].dt.year
    df["Mês"] = df["Criado"].dt.to_period("M").dt.to_timestamp()
    df["Grupo"] = df["Finalidade"].map(GRUPOS_FINALIDADE).fillna(GRUPO_OUTROS)
    df["Placa"] = extrair_placa(df["Placa Cavalo/Carreta"]).fillna(extrair_placa(df["Title"]))
    df["Tem OS"] = df["Ordem de Serviço"].fillna("").astype(str).str.contains(r"\d")
    return df


def convert_df(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


def botao_download(df, prefixo):
    nome = f"{prefixo}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    st.download_button("📥 Baixar dados filtrados (Excel)", data=convert_df(df), file_name=nome,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def opcoes(df, col):
    return sorted(df[col].dropna().unique())


def padrao(disponiveis, lista):
    return [o for o in lista if o in disponiveis]


def periodos_rapidos():
    ini_mes = HOJE.replace(day=1)
    fim_mes_ant = ini_mes - timedelta(days=1)
    return {
        f"Desde {INICIO_ANALISE:%d/%m/%Y}": (INICIO_ANALISE, HOJE),
        "Mês atual": (ini_mes, HOJE),
        "Mês anterior": (fim_mes_ant.replace(day=1), fim_mes_ant),
        "Últimos 30 dias": (HOJE - timedelta(days=29), HOJE),
        "Últimos 90 dias": (HOJE - timedelta(days=89), HOJE),
        "Ontem": (ONTEM, ONTEM),
        "Hoje": (HOJE, HOJE),
        "Personalizado": None,
    }


def restaurar_filtros(pagina):
    for chave in [k for k in st.session_state if k.startswith(f"{pagina}_")]:
        del st.session_state[chave]


def filtros_sidebar(df, pagina, periodo_padrao, extras=()):
    """Monta os filtros da barra lateral e devolve um dict com as seleções."""
    sb = st.sidebar
    sb.markdown("#### Filtros")

    # Período
    presets = periodos_rapidos()
    nomes = list(presets)
    escolha = sb.selectbox("Período", nomes, index=nomes.index(periodo_padrao), key=f"{pagina}_periodo")
    if presets[escolha] is None:
        intervalo = sb.date_input("Intervalo", value=(INICIO_ANALISE, HOJE), format="DD/MM/YYYY",
                                  key=f"{pagina}_intervalo", label_visibility="collapsed")
        ini, fim = (intervalo[0], intervalo[-1]) if intervalo else (INICIO_ANALISE, HOJE)
    else:
        ini, fim = presets[escolha]
        sb.caption(f"{ini:%d/%m/%Y} a {fim:%d/%m/%Y}")

    selecao = {}

    # Gestores
    disp = opcoes(df, "Gestor")
    selecao["Gestor"] = sb.multiselect("Gestores", disp, default=padrao(disp, GESTORES_PADRAO),
                                       placeholder="Todos os gestores", key=f"{pagina}_Gestor")
    sb.caption(f"{len(selecao['Gestor']) or len(disp)} de {len(disp)} gestores")

    # Filtros adicionais da página
    rotulos = {"Finalidade": ("Finalidade", "Todas as finalidades"),
               "Solicitante": ("Solicitante", "Todos os solicitantes")}
    for col in extras:
        rotulo, vazio = rotulos[col]
        selecao[col] = sb.multiselect(rotulo, opcoes(df, col), placeholder=vazio, key=f"{pagina}_{col}")

    # Critério de custo (normalmente fixo)
    with sb.expander("Critério de custo", expanded=False):
        for col, lista_padrao in [("Status", STATUS_PADRAO), ("Classificação", CLASSIFICACAO_PADRAO)]:
            disp = opcoes(df, col)
            selecao[col] = st.multiselect(col, disp, default=padrao(disp, lista_padrao),
                                          placeholder="Todos", key=f"{pagina}_{col}")

    sb.button("Restaurar filtros padrão", on_click=restaurar_filtros, args=(pagina,),
              width="stretch", key=f"{pagina}_restaurar")
    return {"ini": ini, "fim": fim, "sel": selecao}


def aplicar_filtros(df, f, usar_datas=True, usar_status=True):
    mask = pd.Series(True, index=df.index)
    if usar_datas:
        mask &= df["Criado"].dt.date.between(f["ini"], f["fim"])
    for col, valores in f["sel"].items():
        if valores and (usar_status or col != "Status"):
            mask &= df[col].isin(valores)
    return df[mask]


def resumo_filtros(f):
    with st.expander("📌 Filtros aplicados", expanded=False):
        st.write(f"**Período:** {f['ini']:%d/%m/%Y} até {f['fim']:%d/%m/%Y}")
        for col, valores in f["sel"].items():
            st.write(f"**{col}:** {', '.join(valores) if valores else 'Todos'}")


def layout_padrao(fig, titulo=None, altura=None):
    fig.update_layout(
        title=titulo, separators=",.", hoverlabel=dict(font_size=14),
        legend_title_text="", margin=dict(l=10, r=10, t=60 if titulo else 20, b=10),
    )
    if altura:
        fig.update_layout(height=altura)
    return fig


def mostrar(fig):
    st.plotly_chart(fig, width="stretch")


def barras_horizontais(df_agg, categoria, titulo=None, altura=None, total=None):
    """Barra horizontal ordenada com rótulo 'R$ valor (x%)'."""
    df_agg = df_agg.sort_values("Valor")
    base = total if total else df_agg["Valor"].sum()
    rotulos = df_agg["Valor"].apply(lambda v: f"{brl(v)} ({pct(v / base * 100)})" if base else brl(v))
    fig = px.bar(df_agg, x="Valor", y=categoria, orientation="h", text=rotulos,
                 color_discrete_sequence=[COR_PRINCIPAL])
    fig.update_traces(textposition="outside", cliponaxis=False,
                      hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>")
    fig.update_xaxes(title="Valor (R$)", range=[0, df_agg["Valor"].max() * 1.35])
    fig.update_yaxes(title="")
    return layout_padrao(fig, titulo, altura or max(300, 32 * len(df_agg) + 80))


JANELA_PROJECAO = 30   # dias fechados usados como base (melhor resultado no backtest jan–set/2026)
MESES_CALIBRACAO = 12  # meses anteriores usados para medir o erro real da projeção (faixa provável)


def custo_diario(d):
    return d.groupby(d["Criado"].dt.date)["Valor"].sum()


def serie_diaria(diario, ini, fim):
    """Custo por dia entre ini e fim, com zero nos dias sem lançamento."""
    dias = pd.date_range(ini, fim).date
    return pd.Series(diario.reindex(dias, fill_value=0).values, index=dias, dtype=float)


def eh_dia_util(dia):
    return dia.weekday() < 5


def historico_ritmo(diario, hoje):
    """Custo diário dos últimos dias fechados, separado em dias úteis (True) e fins de semana (False)."""
    hist = serie_diaria(diario, hoje - timedelta(days=JANELA_PROJECAO), hoje - timedelta(days=1))
    util = pd.Series([eh_dia_util(x) for x in hist.index], index=hist.index)
    return {True: hist[util], False: hist[~util]}


def projetar_aberto(diario, hoje, fim_mes):
    """Valor esperado de hoje até o fim do mês, pela média de cada tipo de dia."""
    medias = {k: v.mean() for k, v in historico_ritmo(diario, hoje).items()}
    return sum(medias[eh_dia_util(x)] for x in pd.date_range(hoje, fim_mes).date)


def faixa_empirica(diario, hoje):
    """Refaz a projeção nos meses anteriores, no mesmo ponto do mês (±2 dias), e mede quanto o
    realizado ficou acima/abaixo do projetado. Devolve os percentis 10 e 90 dessa razão (faixa de 80%)."""
    if diario.empty:
        return None
    primeiro_lancamento = min(diario.index)
    razoes = []
    ini_atual = pd.Timestamp(hoje.replace(day=1))
    for k in range(1, MESES_CALIBRACAO + 1):
        ini = (ini_atual - pd.DateOffset(months=k)).date()
        fim = (pd.Timestamp(ini) + pd.offsets.MonthEnd(0)).date()
        for desloc in range(-2, 3):
            corte = ini.replace(day=min(max(hoje.day + desloc, 1), fim.day))
            if corte - timedelta(days=JANELA_PROJECAO) < primeiro_lancamento:
                continue  # sem histórico completo para projetar naquele ponto
            previsto = projetar_aberto(diario, corte, fim)
            if previsto > 0:
                razoes.append(serie_diaria(diario, corte, fim).sum() / previsto)
    if len(razoes) < 15:
        return None
    return pd.Series(razoes).quantile(0.10), pd.Series(razoes).quantile(0.90)


def projetar_mes(d, hoje):
    """Projeção do custo do mês corrente.

    - Dias fechados (dia 1 até ontem): valor realizado.
    - Hoje (dia em andamento): o maior entre o já lançado e a média esperada para o dia.
    - Dias restantes: média diária dos últimos 30 dias fechados, separada em dias úteis e fins de semana.
    - Faixa provável (80%): erro real que essa mesma projeção teve nos 12 meses anteriores,
      no mesmo ponto do mês, aplicado sobre a parte ainda não realizada.
    """
    diario = custo_diario(d)
    ini_mes = hoje.replace(day=1)
    fim_mes = (pd.Timestamp(ini_mes) + pd.offsets.MonthEnd(0)).date()
    medias = {k: v.mean() for k, v in historico_ritmo(diario, hoje).items()}

    fechado = serie_diaria(diario, ini_mes, hoje - timedelta(days=1)) if hoje > ini_mes else pd.Series(dtype=float)
    lancado_hoje = serie_diaria(diario, hoje, hoje).sum()
    esperado_hoje = max(lancado_hoje, medias[eh_dia_util(hoje)])
    futuros = pd.date_range(hoje + timedelta(days=1), fim_mes).date
    proj = pd.Series([medias[eh_dia_util(x)] for x in futuros], index=futuros, dtype=float)

    aberto = esperado_hoje + proj.sum()
    realizado = fechado.sum() + lancado_hoje
    total = fechado.sum() + aberto
    q = faixa_empirica(diario, hoje)
    faixa = (max(realizado, fechado.sum() + aberto * q[0]), fechado.sum() + aberto * q[1]) if q else None

    grafico = pd.concat([
        pd.DataFrame({"Data": list(fechado.index) + [hoje], "Valor": list(fechado.values) + [lancado_hoje],
                      "Tipo": "Realizado"}),
        pd.DataFrame({"Data": [hoje] + list(proj.index), "Valor": [esperado_hoje - lancado_hoje] + list(proj.values),
                      "Tipo": "Projetado"}),
    ], ignore_index=True)
    return {"grafico": grafico, "total": total, "realizado": realizado, "faixa": faixa,
            "media_util": medias[True], "media_fds": medias[False]}


def grafico_projecao(grafico):
    fig = px.bar(grafico, x="Data", y="Valor", color="Tipo",
                 color_discrete_map={"Realizado": COR_PRINCIPAL, "Projetado": "#86b6ef"})
    fig.update_traces(hovertemplate="%{x|%d/%m/%Y}<br>R$ %{y:,.2f}<extra>%{fullData.name}</extra>",
                      marker_line_width=0)
    fig.update_xaxes(title="", tickformat="%d/%m", dtick=86400000 * 2)
    fig.update_yaxes(title="Valor (R$)")
    fig.update_layout(barmode="stack")
    return layout_padrao(fig, "Custo diário no mês atual — realizado e projetado")


# ----------------------- CARREGAR DADOS -----------------------
try:
    df = load_data(str(ARQUIVO_DADOS), ARQUIVO_DADOS.stat().st_mtime)
except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")
    st.stop()

HOJE = date.today()
ONTEM = HOJE - timedelta(days=1)

st.sidebar.markdown("## Custos de Manutenção")
st.sidebar.caption(f"Base atualizada em {df['Criado'].max():%d/%m/%Y %H:%M} · "
                   f"{len(df):,} solicitações".replace(",", "."))
menu = st.sidebar.radio("Página", ["Custos da Manutenção", "Análise Detalhada", "Reunião Manutenção Corporativa"],
                        key="menu")
st.sidebar.divider()


# ----------------------- CUSTOS DA MANUTENÇÃO -----------------------
def pagina_manutencao():
    st.title("🔧 Análise de Custos da Manutenção")
    st.caption("Custo de manutenção = Status **Pago** · Classificação **Despesa de Veiculo** · gestores selecionados.")

    f = filtros_sidebar(df, "man", f"Desde {INICIO_ANALISE:%d/%m/%Y}", extras=("Finalidade",))
    dfm = aplicar_filtros(df, f)
    resumo_filtros(f)
    if dfm.empty:
        st.warning("Nenhum registro encontrado para os filtros selecionados.")
        return

    # --- KPIs ---
    total = dfm["Valor"].sum()
    qtd = len(dfm)
    meses = dfm["Mês"].nunique()
    corr_prev = dfm[dfm["Grupo"].isin(["Corretiva", "Preventiva"])]["Valor"].sum()
    perc_prev = dfm.loc[dfm["Grupo"] == "Preventiva", "Valor"].sum() / corr_prev * 100 if corr_prev else 0
    perc_emerg = dfm.loc[dfm["Grupo"] == "Socorro e guincho", "Valor"].sum() / total * 100 if total else 0

    # Mês atual x mesmo intervalo do mês anterior, comparando só dias fechados (hoje ainda está em andamento)
    sem_data = aplicar_filtros(df, f, usar_datas=False)
    ini_mes = HOJE.replace(day=1)
    ini_ant = (pd.Timestamp(ini_mes) - pd.DateOffset(months=1)).date()
    fim_ant = min(ini_ant + (ONTEM - ini_mes), ini_mes - timedelta(days=1))
    datas = sem_data["Criado"].dt.date
    mes_atual = sem_data.loc[datas.between(ini_mes, HOJE), "Valor"].sum()
    fechado_atual = sem_data.loc[datas.between(ini_mes, ONTEM), "Valor"].sum()
    mes_ant = sem_data.loc[datas.between(ini_ant, fim_ant), "Valor"].sum()
    variacao = (fechado_atual / mes_ant - 1) * 100 if mes_ant and HOJE.day > 1 else None
    proj = projetar_mes(sem_data, HOJE)
    faixa = f"{brl(proj['faixa'][0])} a {brl(proj['faixa'][1])}" if proj["faixa"] else "histórico insuficiente"

    c = st.columns(4)
    c[0].metric("💰 Custo total no período", brl(total))
    c[1].metric("📝 Solicitações pagas", f"{qtd:,}".replace(",", "."))
    c[2].metric("🧾 Ticket médio", brl(total / qtd))
    c[3].metric("📅 Média mensal", brl(total / meses), help="Custo total ÷ meses com lançamento no período")
    c = st.columns(4)
    c[0].metric(f"📆 Mês atual (até {HOJE:%d/%m})", brl(mes_atual),
                delta=f"{pct(variacao)} vs 01–{fim_ant:%d/%m}" if variacao is not None else None,
                delta_color="inverse",
                help=f"Variação dos dias fechados (01/{HOJE:%m} a {ONTEM:%d/%m}) contra o mesmo intervalo do mês anterior")
    c[1].metric("🔮 Projeção do mês", brl(proj["total"]),
                help=f"Faixa provável (80%): {faixa}. Realizado + média dos últimos 30 dias "
                     f"({brl(proj['media_util'])}/dia útil, {brl(proj['media_fds'])}/dia de fim de semana)")
    c[2].metric("🛠️ Preventiva / (Corretiva + Preventiva)", pct(perc_prev),
                help="Quanto maior, menos reativa é a manutenção")
    c[3].metric("🚨 Socorro e guincho", pct(perc_emerg), help="% do custo total gasto com socorro e guincho")

    # --- Evolução mensal por grupo ---
    st.markdown("### 📈 Evolução mensal por grupo de finalidade")
    mensal = dfm.groupby(["Mês", "Grupo"])["Valor"].sum().reset_index()
    totais = dfm.groupby("Mês")["Valor"].sum().reset_index()
    fig = px.bar(mensal, x="Mês", y="Valor", color="Grupo", category_orders={"Grupo": ORDEM_GRUPOS},
                 color_discrete_map=CORES_GRUPOS)
    fig.update_traces(hovertemplate="%{x|%m/%Y}<br>R$ %{y:,.2f}<extra>%{fullData.name}</extra>",
                      marker_line_width=1.5, marker_line_color="rgba(0,0,0,0)")
    fig.add_scatter(x=totais["Mês"], y=totais["Valor"], mode="text", text=totais["Valor"].apply(brl),
                    textposition="top center", showlegend=False, hoverinfo="skip")
    fig.update_xaxes(title="", dtick="M1", tickformat="%m/%Y")
    fig.update_yaxes(title="Valor (R$)", range=[0, totais["Valor"].max() * 1.15])
    mostrar(layout_padrao(fig, altura=450))

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🎯 Custo por finalidade")
        por_fin = dfm.groupby("Finalidade")["Valor"].sum().reset_index()
        mostrar(barras_horizontais(por_fin, "Finalidade"))
    with col2:
        st.markdown("### 🛠️ Preventiva × Corretiva")
        cp = dfm[dfm["Grupo"].isin(["Corretiva", "Preventiva"])].pivot_table(
            index="Mês", columns="Grupo", values="Valor", aggfunc="sum", fill_value=0)
        if {"Corretiva", "Preventiva"} <= set(cp.columns):
            cp["% Preventiva"] = cp["Preventiva"] / (cp["Corretiva"] + cp["Preventiva"]) * 100
            fig = px.line(cp.reset_index(), x="Mês", y="% Preventiva", markers=True,
                          text=cp["% Preventiva"].apply(pct).values, color_discrete_sequence=[CORES_GRUPOS["Preventiva"]])
            fig.update_traces(textposition="top center", line_width=2, marker_size=8,
                              hovertemplate="%{x|%m/%Y}<br>%{y:.1f}% preventiva<extra></extra>")
            fig.update_xaxes(title="", dtick="M1", tickformat="%m/%Y")
            fig.update_yaxes(title="% do custo de manutenção preventiva", range=[0, max(cp["% Preventiva"].max() * 1.3, 10)])
            mostrar(layout_padrao(fig, "Participação da preventiva no custo (corretiva + preventiva)"))
        else:
            st.info("Sem lançamentos de corretiva e preventiva no período.")

        st.markdown("### 🚨 Socorro e guincho")
        emerg = dfm[dfm["Grupo"] == "Socorro e guincho"].groupby("Finalidade")["Valor"].sum().reset_index()
        if emerg.empty:
            st.info("Sem lançamentos de socorro ou guincho no período.")
        else:
            mostrar(barras_horizontais(emerg, "Finalidade"))

    # --- Gestor x mês ---
    st.markdown("### 👔 Custo por gestor e mês")
    mapa = dfm.pivot_table(index="Gestor", columns="Mês", values="Valor", aggfunc="sum", fill_value=0)
    mapa = mapa.loc[mapa.sum(axis=1).sort_values(ascending=False).index]
    mapa.columns = [c.strftime("%m/%Y") for c in mapa.columns]
    fig = px.imshow(mapa, aspect="auto", color_continuous_scale=["#f0f6fd", "#86b6ef", "#2a78d6", "#0d366b"])
    fig.update_traces(texttemplate="%{z:,.0f}", hovertemplate="<b>%{y}</b><br>%{x}<br>R$ %{z:,.2f}<extra></extra>",
                      xgap=2, ygap=2)
    fig.update_xaxes(title="", side="top")
    fig.update_yaxes(title="")
    fig.update_coloraxes(colorbar_title="R$")
    mostrar(layout_padrao(fig, altura=max(300, 45 * len(mapa) + 100)))

    # --- Placas ---
    st.markdown("### 🚛 Veículos com maior custo")
    com_placa = dfm.dropna(subset=["Placa"])
    st.caption(f"Placa identificada em {pct(len(com_placa) / qtd * 100)} das solicitações "
               "(extraída do campo Placa ou do título; quando há várias, considera a primeira — normalmente o cavalo).")
    if not com_placa.empty:
        placas = com_placa.groupby("Placa").agg(
            Qtd=("Valor", "size"), Valor=("Valor", "sum"),
            Corretivas=("Grupo", lambda s: (s == "Corretiva").sum()),
            Socorros=("Grupo", lambda s: (s == "Socorro e guincho").sum()),
            Primeira=("Criado", "min"), Última=("Criado", "max"),
        ).reset_index()
        placas["Intervalo médio (dias)"] = ((placas["Última"] - placas["Primeira"]).dt.days /
                                           (placas["Qtd"] - 1).where(placas["Qtd"] > 1)).round(1)
        placas = placas.sort_values("Valor", ascending=False)

        col1, col2 = st.columns([2, 3])
        with col1:
            mostrar(barras_horizontais(placas.head(15)[["Placa", "Valor"]], "Placa", "Top 15 placas por custo",
                                       total=com_placa["Valor"].sum()))
        with col2:
            st.markdown("**Reincidência** — placas com 3 ou mais solicitações no período")
            reinc = placas[placas["Qtd"] >= 3].sort_values(["Qtd", "Valor"], ascending=False)
            st.dataframe(reinc, hide_index=True, width="stretch", height=480, column_config={
                "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
                "Primeira": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
                "Última": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
            })

    # --- Ticket médio por finalidade ---
    st.markdown("### 🧾 Ticket médio por finalidade")
    ultimo_mes = dfm["Mês"].max()
    tk = dfm.groupby("Finalidade").agg(Qtd=("Valor", "size"), Valor=("Valor", "sum"),
                                       Maior=("Valor", "max")).reset_index()
    tk["% do total"] = tk["Valor"] / total * 100
    tk["Ticket médio"] = tk["Valor"] / tk["Qtd"]
    ant = dfm[dfm["Mês"] < ultimo_mes].groupby("Finalidade")["Valor"].mean().rename("Ticket meses anteriores")
    atual = dfm[dfm["Mês"] == ultimo_mes].groupby("Finalidade")["Valor"].mean().rename(f"Ticket {ultimo_mes:%m/%Y}")
    tk = tk.join(ant, on="Finalidade").join(atual, on="Finalidade")
    tk["Variação %"] = (tk[f"Ticket {ultimo_mes:%m/%Y}"] / tk["Ticket meses anteriores"] - 1) * 100
    tk = tk.sort_values("Valor", ascending=False)
    moeda = st.column_config.NumberColumn(format="localized")
    st.dataframe(tk, hide_index=True, width="stretch", column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
        "Maior": st.column_config.NumberColumn("Maior valor (R$)", format="localized"),
        "% do total": st.column_config.NumberColumn(format="%.1f%%"),
        "Ticket médio": moeda, "Ticket meses anteriores": moeda, f"Ticket {ultimo_mes:%m/%Y}": moeda,
        "Variação %": st.column_config.NumberColumn(format="%+.1f%%"),
    })

    # --- Conformidade ---
    st.markdown("### ✅ Conformidade do processo por gestor")
    st.caption("Taxa de recusa = recusadas ÷ (pagas + recusadas), mesmos filtros exceto Status.")
    todos_status = aplicar_filtros(df, f, usar_status=False)
    conf = dfm.groupby("Gestor").agg(Pagas=("Valor", "size"), Valor=("Valor", "sum"),
                                     sem_os=("Tem OS", lambda s: (~s).mean() * 100)).reset_index()
    recusadas = todos_status[todos_status["Status"] == "Recusado"].groupby("Gestor").size().rename("Recusadas")
    conf = conf.join(recusadas, on="Gestor").fillna({"Recusadas": 0})
    conf["Taxa de recusa"] = conf["Recusadas"] / (conf["Pagas"] + conf["Recusadas"]) * 100
    conf = conf.rename(columns={"sem_os": "% sem OS"}).sort_values("Valor", ascending=False)
    st.dataframe(conf, hide_index=True, width="stretch", column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
        "% sem OS": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100),
        "Recusadas": st.column_config.NumberColumn(format="%d"),
        "Taxa de recusa": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100),
    })

    # --- Projeção ---
    st.markdown("### 🔮 Projeção do mês atual")
    mostrar(grafico_projecao(proj["grafico"]))
    st.success(f"📌 Estimativa de custo de manutenção até o fim do mês: **{brl(proj['total'])}** "
               f"(faixa provável: {faixa})")
    st.caption(f"Como é calculado: dias fechados entram pelo valor realizado; hoje entra pelo maior valor entre "
               f"o já lançado e a média esperada; os dias restantes recebem a média dos últimos {JANELA_PROJECAO} "
               f"dias fechados ({brl(proj['media_util'])} por dia útil e {brl(proj['media_fds'])} por dia de fim de semana). "
               f"A faixa provável (80%) vem do erro real que esta mesma projeção teve nos {MESES_CALIBRACAO} meses "
               "anteriores, no mesmo ponto do mês.")

    # --- Detalhe ---
    st.markdown("### 📋 Solicitações")
    colunas = ["ID", "Criado", "Gestor", "Solicitante", "Finalidade", "Grupo", "Placa", "Valor",
               "Title", "Descrição", "Ordem de Serviço", "Nome Favorecido"]
    detalhe = dfm[colunas].sort_values("Criado", ascending=False)
    st.dataframe(detalhe, hide_index=True, width="stretch", column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
        "Criado": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
        "ID": st.column_config.NumberColumn(format="%d"),
    })
    botao_download(detalhe, "custos_manutencao")


# ----------------------- ANÁLISE DETALHADA -----------------------
def pagina_detalhada():
    st.title("👤 Análise por Solicitante")
    f = filtros_sidebar(df, "det", "Ontem", extras=("Finalidade",))
    resumo_filtros(f)
    df_filtrado = aplicar_filtros(df, f)

    solicitante = st.selectbox("🙋 Selecione um Solicitante", options=["Todos"] + opcoes(df_filtrado, "Solicitante"))
    if solicitante != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Solicitante"] == solicitante]
        custo_total = df_filtrado["Valor"].sum()
        qtd = len(df_filtrado)
        dias = df_filtrado["Criado"].dt.date.nunique()
        col1, col2 = st.columns([1, 3])
        col1.subheader(f"🙋 {solicitante}")
        m1, m2, m3 = col2.columns(3)
        m1.metric("💰 Custo Total", brl(custo_total))
        m2.metric("⚖️ Custo Médio por Solicitação", brl(custo_total / qtd if qtd else 0))
        m3.metric("📅 Custo Médio Diário", brl(custo_total / dias if dias else 0))
    else:
        st.markdown("### 📊 Resumo Geral")
        col1, col2 = st.columns(2)
        col1.metric("💰 Custo Total no Período", brl(df_filtrado["Valor"].sum()))
        col2.metric("📝 Qtd. Registros", f"{len(df_filtrado):,}".replace(",", "."))

        if not df_filtrado.empty:
            st.markdown("### 👥 Custo por Solicitante")
            por_sol = df_filtrado.groupby("Solicitante").agg(Valor=("Valor", "sum"), Qtd=("ID", "count")).reset_index()
            por_sol = por_sol.sort_values("Valor")
            fig = px.bar(por_sol, x="Valor", y="Solicitante", orientation="h", custom_data=["Qtd"],
                         text=por_sol.apply(lambda r: f"{brl(r['Valor'])} ({r['Qtd']})", axis=1),
                         color_discrete_sequence=[COR_PRINCIPAL])
            fig.update_traces(textposition="outside", cliponaxis=False,
                              hovertemplate="<b>%{y}</b><br>Valor Total: R$ %{x:,.2f}<br>Qtd Solicitações: %{customdata[0]}<extra></extra>")
            fig.update_xaxes(title="Valor (R$)", range=[0, por_sol["Valor"].max() * 1.35])
            fig.update_yaxes(title="")
            mostrar(layout_padrao(fig, altura=max(300, 32 * len(por_sol) + 80)))

    if not df_filtrado.empty:
        st.markdown("### 🎯 Custo por Finalidade")
        mostrar(barras_horizontais(df_filtrado.groupby("Finalidade")["Valor"].sum().reset_index(), "Finalidade"))

    st.markdown("### 📋 Tabela de Solicitações")
    st.dataframe(df_filtrado, hide_index=True, width="stretch", column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
        "Criado": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
        "ID": st.column_config.NumberColumn(format="%d"),
    })
    botao_download(df_filtrado, f"analise_{solicitante}")


# ----------------------- REUNIÃO MANUTENÇÃO CORPORATIVA -----------------------
def pagina_reuniao():
    st.title("🏗️ Relatório | Reunião de Manutenção Corporativa")
    f = filtros_sidebar(df, "reu", "Ontem", extras=("Finalidade",))
    resumo_filtros(f)
    df_rm = aplicar_filtros(df, f)

    custo_total = df_rm["Valor"].sum()
    qtd = len(df_rm)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Custo Total", brl(custo_total))
    col2.metric("📋 Custo Médio por Solicitação", brl(custo_total / qtd if qtd else 0))
    col3.metric("📝 Qtd. Registros", qtd)
    if qtd:
        maior = df_rm.loc[df_rm["Valor"].idxmax()]
        col4.metric("🔝 Maior Solicitação", brl(maior["Valor"]))
        col4.caption(f"ID {maior['ID']} — {maior['Solicitante']}")
    else:
        col4.metric("🔝 Maior Solicitação", "Sem dados")

    if qtd:
        st.markdown("### 📈 Custo por Finalidade no Período")
        mostrar(barras_horizontais(df_rm.groupby("Finalidade")["Valor"].sum().reset_index(), "Finalidade"))

    st.markdown("### 🔮 Projeção de custo para hoje")
    sem_data = aplicar_filtros(df, f, usar_datas=False)
    eh_util = HOJE.weekday() < 5
    diario = custo_diario(sem_data)
    hist = historico_ritmo(diario, HOJE)[eh_util]
    lancado = serie_diaria(diario, HOJE, HOJE).sum()
    tipo = "dia útil" if eh_util else "dia de fim de semana"
    st.success(f"🔜 Custo esperado para hoje ({HOJE:%d/%m}, {tipo}): **{brl(max(hist.mean(), lancado))}** "
               f"· já lançado: {brl(lancado)}")
    st.caption(f"Média dos dias do mesmo tipo nos últimos {JANELA_PROJECAO} dias fechados. "
               f"Em 80% desses dias o custo ficou entre {brl(hist.quantile(0.10))} e {brl(hist.quantile(0.90))}.")

    st.markdown("### 📋 Tabela Detalhada")
    colunas = ["ID", "Title", "Valor", "Finalidade", "Solicitante", "Descrição", "Gestor", "Nome Favorecido"]
    st.dataframe(df_rm[colunas], hide_index=True, width="stretch", column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="localized"),
        "ID": st.column_config.NumberColumn(format="%d"),
    })


PAGINAS = {
    "Custos da Manutenção": pagina_manutencao,
    "Análise Detalhada": pagina_detalhada,
    "Reunião Manutenção Corporativa": pagina_reuniao,
}
PAGINAS[menu]()
