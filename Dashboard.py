import io
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Custos de Manutenção", layout="wide", initial_sidebar_state="collapsed")
pd.set_option("styler.render.max_elements", 10_000_000)

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
# Paleta neutra (cinza-azulado), do mais escuro ao mais claro
CORES_GRUPOS = dict(zip(ORDEM_GRUPOS, ["#2f4560", "#5b7590", "#8ea2b6", "#b9c6d3", "#d5dde5", "#e6eaef"]))
COR_PRINCIPAL = "#4a6785"
COR_SECUNDARIA = "#c3ccd6"

# Placa antiga (ABC1234) ou Mercosul (ABC1D23), com ou sem hífen/espaço
PADRAO_PLACA = r"(?<![A-Z])([A-Z]{3})[\s-]?(\d[A-Z0-9]\d{2})(?!\d)"

# Projeção
JANELA_PROJECAO = 30   # dias fechados usados como base (melhor resultado no backtest jan–set/2026)
JANELA_PERFIL = 90     # dias usados para o perfil de cada dia da semana (seg..dom)
MESES_CALIBRACAO = 12  # meses anteriores usados para medir o erro real da projeção (faixa provável)

ESTILO = """
<style>
.block-container {padding-top: 4rem; padding-bottom: 2rem; max-width: 1500px;}
.stAppDeployButton {display: none;}
h1 {font-size: 1.75rem !important; font-weight: 600 !important; padding: 0 !important;}
h3 {font-size: 1.1rem !important; font-weight: 600 !important; padding: 0.5rem 0 0.25rem 0 !important;}
[data-testid="stHeading"] {margin-bottom: 0.5rem;}
[data-testid="stMetricLabel"] p {opacity: 0.7;}
[data-testid="stMetricValue"] {font-weight: 600;}
</style>
"""


# ----------------------- FORMATAÇÃO -----------------------
def brl(valor):
    """Valor em Real brasileiro: R$1.234,56"""
    if pd.isna(valor):
        return "—"
    texto = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R${texto}" if valor < 0 else f"R${texto}"


def md(texto):
    """Escapa o '$' para textos exibidos como Markdown (senão 'R$...R$' vira fórmula matemática)."""
    return texto.replace("$", r"\$")


def pct(valor, sinal=False):
    if pd.isna(valor):
        return "—"
    return (f"{valor:+.1f}%" if sinal else f"{valor:.1f}%").replace(".", ",")


def inteiro(valor):
    return f"{int(valor):,}".replace(",", ".")


def tabela(dados, moeda=(), percentual=(), variacao=(), datas=(), data_hora=(), altura=None):
    """Tabela com formatação brasileira fixa (não depende do idioma do navegador)."""
    fmt = {c: brl for c in moeda}
    fmt |= {c: pct for c in percentual}
    fmt |= {c: (lambda v: pct(v, sinal=True)) for c in variacao}
    fmt |= {c: (lambda v: f"{v:%d/%m/%Y}") for c in datas}
    fmt |= {c: (lambda v: f"{v:%d/%m/%Y %H:%M}") for c in data_hora}
    dados = dados.copy()
    texto = dados.select_dtypes(include=["object", "string"]).columns
    dados[texto] = dados[texto].fillna("—")
    extra = {"height": altura} if altura else {}
    st.dataframe(dados.style.format(fmt, na_rep="—"), hide_index=True, width="stretch", **extra)


# ----------------------- DADOS -----------------------
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
    st.download_button("Baixar Excel", data=convert_df(df), file_name=nome, icon=":material/download:",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def opcoes(df, col):
    return sorted(df[col].dropna().unique())


def padrao(disponiveis, lista):
    return [o for o in lista if o in disponiveis]


# ----------------------- FILTROS (barra horizontal) -----------------------
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


def resumo_selecao(valores, total, vazio="Todos"):
    if not valores or len(valores) == total:
        return vazio
    return valores[0] if len(valores) == 1 else f"{len(valores)} selecionados"


def filtro_multiplo(container, df, col, rotulo, lista_padrao, chave, vazio="Todos"):
    disp = opcoes(df, col)
    atual = st.session_state.get(chave, padrao(disp, lista_padrao))
    with container.popover(f"{rotulo}: {resumo_selecao(atual, len(disp), vazio)}", width="stretch",
                           key=f"{chave}_pop"):
        return st.multiselect(rotulo, disp, default=padrao(disp, lista_padrao), key=chave,
                              placeholder=vazio, label_visibility="collapsed")


def barra_filtros(df, pagina, periodo_padrao):
    """Filtros em uma linha no topo da página. Devolve um dict com as seleções."""
    presets = periodos_rapidos()
    nomes = list(presets)
    k = lambda nome: f"{pagina}_{nome}"

    c = st.columns([1.3, 1.5, 1.3, 1.3, 1.1, 0.5], vertical_alignment="bottom")
    escolha = c[0].selectbox("Período", nomes, index=nomes.index(periodo_padrao), key=k("periodo"),
                              label_visibility="collapsed")
    if presets[escolha] is None:
        intervalo = c[1].date_input("Intervalo", value=(INICIO_ANALISE, HOJE), format="DD/MM/YYYY",
                                    key=k("intervalo"), label_visibility="collapsed")
        ini, fim = (intervalo[0], intervalo[-1]) if len(intervalo) else (INICIO_ANALISE, HOJE)
    else:
        ini, fim = presets[escolha]
        c[1].date_input("Intervalo", value=(ini, fim), format="DD/MM/YYYY", disabled=True, label_visibility="collapsed")

    selecao = {
        "Gestor": filtro_multiplo(c[2], df, "Gestor", "Gestores", GESTORES_PADRAO, k("Gestor")),
        "Finalidade": filtro_multiplo(c[3], df, "Finalidade", "Finalidades", [], k("Finalidade"), "Todas"),
    }
    with c[4].popover("Critério de custo", width="stretch", key=k("criterio_pop")):
        for col, lista_padrao in [("Status", STATUS_PADRAO), ("Classificação", CLASSIFICACAO_PADRAO)]:
            disp = opcoes(df, col)
            selecao[col] = st.multiselect(col, disp, default=padrao(disp, lista_padrao),
                                          placeholder="Todos", key=k(col))
    c[5].button("", icon=":material/restart_alt:", help="Restaurar filtros padrão", on_click=restaurar_filtros,
                args=(pagina,), width="stretch", key=k("restaurar"))
    return {"ini": ini, "fim": fim, "sel": selecao}


def aplicar_filtros(df, f, usar_datas=True, usar_status=True):
    mask = pd.Series(True, index=df.index)
    if usar_datas:
        mask &= df["Criado"].dt.date.between(f["ini"], f["fim"])
    for col, valores in f["sel"].items():
        if valores and (usar_status or col != "Status"):
            mask &= df[col].isin(valores)
    return df[mask]


# ----------------------- GRÁFICOS -----------------------
def layout_padrao(fig, titulo=None, altura=None, eixo_moeda="y"):
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=15)) if titulo else None,
        font=dict(size=12), separators=",.", hoverlabel=dict(font_size=13), legend_title_text="",
        legend=dict(orientation="h", yanchor="top", y=-0.1, x=0),
        margin=dict(l=8, r=8, t=48 if titulo else 12, b=8),
    )
    if eixo_moeda == "y":
        fig.update_yaxes(tickprefix="R$", tickformat=",.0f")
    elif eixo_moeda == "x":
        fig.update_xaxes(tickprefix="R$", tickformat=",.0f")
    if altura:
        fig.update_layout(height=altura)
    return fig


def mostrar(fig):
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def barras_horizontais(df_agg, categoria, titulo=None, altura=None, top=None):
    """Barra horizontal ordenada com o valor no fim da barra. Com top=N, o restante vira 'Outros'."""
    df_agg = df_agg.sort_values("Valor", ascending=False)
    if top and len(df_agg) > top:
        outros = pd.DataFrame({categoria: ["Outros"], "Valor": [df_agg["Valor"].iloc[top:].sum()]})
        df_agg = pd.concat([df_agg.head(top), outros], ignore_index=True)
    df_agg = df_agg.iloc[::-1]
    fig = px.bar(df_agg, x="Valor", y=categoria, orientation="h", text=df_agg["Valor"].apply(brl),
                 color_discrete_sequence=[COR_PRINCIPAL])
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0,
                      hovertemplate="<b>%{y}</b><br>R$%{x:,.2f}<extra></extra>")
    fig.update_xaxes(title="", range=[0, df_agg["Valor"].max() * 1.3], showgrid=False, showticklabels=False)
    fig.update_yaxes(title="", type="category")
    fig.update_layout(bargap=0.35)
    return layout_padrao(fig, titulo, altura or max(180, 32 * len(df_agg) + 60), eixo_moeda=None)


# ----------------------- PROJEÇÃO -----------------------
def custo_diario(d):
    return d.groupby(d["Criado"].dt.date)["Valor"].sum()


def serie_diaria(diario, ini, fim):
    """Custo por dia entre ini e fim, com zero nos dias sem lançamento."""
    dias = pd.date_range(ini, fim).date
    return pd.Series(diario.reindex(dias, fill_value=0).values, index=dias, dtype=float)


def taxas_dia_semana(diario, hoje):
    """Custo esperado para cada dia da semana (0 = segunda ... 6 = domingo).

    O nível vem da média dos últimos 30 dias fechados (dias úteis e fins de semana separados) e é
    distribuído pelo perfil de cada dia da semana nos últimos 90 dias (ex.: segunda costuma ser
    menor que terça; domingo quase não tem lançamento)."""
    ontem = hoje - timedelta(days=1)
    h30 = serie_diaria(diario, hoje - timedelta(days=JANELA_PROJECAO), ontem)
    h90 = serie_diaria(diario, hoje - timedelta(days=JANELA_PERFIL), ontem)
    dia30 = pd.Series([x.weekday() for x in h30.index], index=h30.index)
    dia90 = pd.Series([x.weekday() for x in h90.index], index=h90.index)
    taxas = {}
    for dias in (range(0, 5), range(5, 7)):
        nivel = h30[dia30.isin(dias)].mean()
        base = h90[dia90.isin(dias)].mean()
        for dia in dias:
            fator = h90[dia90 == dia].mean() / base if base else 1
            taxas[dia] = nivel * fator
    return taxas


def projetar_aberto(diario, hoje, fim_mes):
    """Valor esperado de hoje até o fim do mês."""
    taxas = taxas_dia_semana(diario, hoje)
    return sum(taxas[x.weekday()] for x in pd.date_range(hoje, fim_mes).date)


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
    - Hoje (dia em andamento): o maior entre o já lançado e o valor esperado para o dia.
    - Dias restantes: valor esperado de cada dia da semana (ver taxas_dia_semana).
    - Faixa provável (80%): erro real que essa mesma projeção teve nos 12 meses anteriores,
      no mesmo ponto do mês, aplicado sobre a parte ainda não realizada.
    """
    diario = custo_diario(d)
    ini_mes = hoje.replace(day=1)
    fim_mes = (pd.Timestamp(ini_mes) + pd.offsets.MonthEnd(0)).date()
    taxas = taxas_dia_semana(diario, hoje)

    fechado = serie_diaria(diario, ini_mes, hoje - timedelta(days=1)) if hoje > ini_mes else pd.Series(dtype=float)
    lancado_hoje = serie_diaria(diario, hoje, hoje).sum()
    esperado_hoje = max(lancado_hoje, taxas[hoje.weekday()])
    futuros = pd.date_range(hoje + timedelta(days=1), fim_mes).date
    proj = pd.Series([taxas[x.weekday()] for x in futuros], index=futuros, dtype=float)

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
    return {"grafico": grafico, "total": total, "realizado": realizado, "faixa": faixa}


def grafico_projecao(grafico):
    fig = px.bar(grafico, x="Data", y="Valor", color="Tipo",
                 color_discrete_map={"Realizado": COR_PRINCIPAL, "Projetado": COR_SECUNDARIA})
    fig.update_traces(hovertemplate="%{x|%d/%m/%Y}<br>R$%{y:,.2f}<extra>%{fullData.name}</extra>",
                      marker_line_width=0)
    fig.update_xaxes(title="", tickformat="%d/%m", dtick=86400000 * 2)
    fig.update_yaxes(title="")
    fig.update_layout(barmode="stack")
    return layout_padrao(fig, altura=380)


# ----------------------- CABEÇALHO -----------------------
def cabecalho(titulo, pagina, periodo_padrao):
    """Título, data da base e barra de filtros. Devolve os filtros."""
    st.title(titulo)
    st.caption(f"Base atualizada em {df['Criado'].max():%d/%m/%Y %H:%M}")
    return barra_filtros(df, pagina, periodo_padrao)


# ----------------------- CUSTOS DA MANUTENÇÃO -----------------------
def pagina_manutencao():
    f = cabecalho("Custos da Manutenção", "man", f"Desde {INICIO_ANALISE:%d/%m/%Y}")
    dfm = aplicar_filtros(df, f)
    if dfm.empty:
        st.info("Nenhum registro encontrado para os filtros selecionados.")
        return

    # --- Indicadores ---
    total = dfm["Valor"].sum()
    qtd = len(dfm)
    corr_prev = dfm[dfm["Grupo"].isin(["Corretiva", "Preventiva"])]["Valor"].sum()
    perc_prev = dfm.loc[dfm["Grupo"] == "Preventiva", "Valor"].sum() / corr_prev * 100 if corr_prev else 0

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
    c[0].metric("Custo no período", brl(total), border=True,
                help=md(f"{inteiro(qtd)} solicitações pagas · ticket médio {brl(total / qtd)}"))
    c[1].metric(f"Mês atual · até {HOJE:%d/%m}", brl(mes_atual), border=True,
                delta=f"{pct(variacao, sinal=True)} vs mês anterior" if variacao is not None else None,
                delta_color="inverse",
                help=f"Variação dos dias fechados (01/{HOJE:%m} a {ONTEM:%d/%m}) contra o mesmo intervalo do mês anterior")
    c[2].metric("Projeção do mês", brl(proj["total"]), border=True,
                help=md(f"Faixa provável (80%): {faixa}"))
    c[3].metric("Participação da preventiva", pct(perc_prev), border=True,
                help="Preventiva ÷ (corretiva + preventiva). Quanto maior, menos reativa é a manutenção")

    abas = st.tabs(["Visão geral", "Projeção", "Gestores", "Veículos", "Solicitações"])

    # --- Visão geral ---
    with abas[0]:
        mensal = dfm.groupby(["Mês", "Grupo"])["Valor"].sum().reset_index()
        totais = dfm.groupby("Mês")["Valor"].sum().reset_index()
        fig = px.bar(mensal, x="Mês", y="Valor", color="Grupo", category_orders={"Grupo": ORDEM_GRUPOS},
                     color_discrete_map=CORES_GRUPOS)
        fig.update_traces(hovertemplate="%{x|%m/%Y}<br>R$%{y:,.2f}<extra>%{fullData.name}</extra>",
                          marker_line_width=0)
        fig.add_scatter(x=totais["Mês"], y=totais["Valor"], mode="text", text=totais["Valor"].apply(brl),
                        textposition="top center", showlegend=False, hoverinfo="skip")
        fig.update_xaxes(title="", dtick="M1", tickformat="%m/%Y")
        fig.update_yaxes(title="", range=[0, totais["Valor"].max() * 1.15])
        mostrar(layout_padrao(fig, "Custo mensal", altura=420))

        col1, col2 = st.columns(2, gap="large")
        with col1:
            por_fin = dfm.groupby("Finalidade")["Valor"].sum().reset_index()
            mostrar(barras_horizontais(por_fin, "Finalidade", "Custo por finalidade", top=7))
        with col2:
            cp = dfm[dfm["Grupo"].isin(["Corretiva", "Preventiva"])].pivot_table(
                index="Mês", columns="Grupo", values="Valor", aggfunc="sum", fill_value=0)
            if {"Corretiva", "Preventiva"} <= set(cp.columns):
                cp["% Preventiva"] = cp["Preventiva"] / (cp["Corretiva"] + cp["Preventiva"]) * 100
                fig = px.line(cp.reset_index(), x="Mês", y="% Preventiva", markers=True,
                              text=cp["% Preventiva"].apply(pct).values, color_discrete_sequence=[COR_PRINCIPAL])
                fig.update_traces(textposition="top center", line_width=2, marker_size=7,
                                  hovertemplate="%{x|%m/%Y}<br>%{y:.1f}% preventiva<extra></extra>")
                fig.update_xaxes(title="", dtick="M1", tickformat="%m/%Y")
                fig.update_yaxes(title="", ticksuffix="%", range=[0, max(cp["% Preventiva"].max() * 1.3, 10)])
                mostrar(layout_padrao(fig, "Participação da preventiva", altura=320, eixo_moeda=None))

    # --- Projeção ---
    with abas[1]:
        mostrar(grafico_projecao(proj["grafico"]))
        st.caption(md(f"Realizado até agora: {brl(proj['realizado'])} · Projeção: {brl(proj['total'])} · "
                      f"Faixa provável (80%): {faixa}"))
        st.caption(f"Cada dia restante recebe o custo esperado para o seu dia da semana: o nível dos últimos "
                   f"{JANELA_PROJECAO} dias fechados, distribuído pelo perfil de segunda a domingo dos últimos "
                   f"{JANELA_PERFIL} dias. A faixa vem do erro real desta projeção nos {MESES_CALIBRACAO} "
                   "meses anteriores.")

    # --- Gestores ---
    with abas[2]:
        ges = dfm.groupby("Gestor").agg(Solicitações=("Valor", "size"), Valor=("Valor", "sum"),
                                        sem_os=("Tem OS", lambda s: (~s).mean() * 100)).reset_index()
        ges["Ticket médio"] = ges["Valor"] / ges["Solicitações"]
        ges = ges.rename(columns={"sem_os": "% sem OS"}).sort_values("Valor", ascending=False)
        col1, col2 = st.columns(2, gap="large")
        with col1:
            mostrar(barras_horizontais(ges[["Gestor", "Valor"]], "Gestor", "Custo por gestor"))
        with col2:
            st.markdown("### Resumo")
            tabela(ges[["Gestor", "Solicitações", "Ticket médio", "% sem OS"]],
                   moeda=["Ticket médio"], percentual=["% sem OS"])

    # --- Veículos ---
    with abas[3]:
        com_placa = dfm.dropna(subset=["Placa"])
        if com_placa.empty:
            st.info("Nenhuma placa identificada no período.")
        else:
            placas = com_placa.groupby("Placa").agg(Solicitações=("Valor", "size"), Valor=("Valor", "sum"),
                                                    Última=("Criado", "max")).reset_index()
            col1, col2 = st.columns(2, gap="large")
            with col1:
                mostrar(barras_horizontais(placas[["Placa", "Valor"]].nlargest(10, "Valor"), "Placa",
                                           "10 placas com maior custo"))
            with col2:
                st.markdown("### Reincidência")
                st.caption("Placas com 3 ou mais solicitações no período")
                reinc = placas[placas["Solicitações"] >= 3].sort_values(["Solicitações", "Valor"], ascending=False)
                tabela(reinc, moeda=["Valor"], datas=["Última"], altura=380)

    # --- Solicitações ---
    with abas[4]:
        detalhe = dfm[["Criado", "Gestor", "Finalidade", "Placa", "Valor", "Descrição"]] \
            .sort_values("Criado", ascending=False).rename(columns={"Criado": "Data"})
        tabela(detalhe, moeda=["Valor"], datas=["Data"], altura=460)
        botao_download(dfm.sort_values("Criado", ascending=False), "custos_manutencao")


# ----------------------- ANÁLISE POR SOLICITANTE -----------------------
def pagina_detalhada():
    f = cabecalho("Análise por Solicitante", "det", "Ontem")
    df_filtrado = aplicar_filtros(df, f)

    solicitante = st.selectbox("Solicitante", options=["Todos"] + opcoes(df_filtrado, "Solicitante"), key="det_sol")
    if solicitante != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Solicitante"] == solicitante]

    custo_total = df_filtrado["Valor"].sum()
    qtd = len(df_filtrado)
    c = st.columns(3)
    c[0].metric("Custo total", brl(custo_total), border=True)
    c[1].metric("Solicitações", inteiro(qtd), border=True)
    c[2].metric("Ticket médio", brl(custo_total / qtd if qtd else 0), border=True)

    if df_filtrado.empty:
        st.info("Nenhum registro encontrado para os filtros selecionados.")
        return

    col1, col2 = st.columns(2, gap="large")
    with col1:
        if solicitante == "Todos":
            por_sol = df_filtrado.groupby("Solicitante")["Valor"].sum().reset_index()
            mostrar(barras_horizontais(por_sol, "Solicitante", "Custo por solicitante", top=10))
    with col2:
        por_fin = df_filtrado.groupby("Finalidade")["Valor"].sum().reset_index()
        mostrar(barras_horizontais(por_fin, "Finalidade", "Custo por finalidade", top=7))

    st.markdown("### Solicitações")
    detalhe = df_filtrado[["Criado", "Solicitante", "Finalidade", "Placa", "Valor", "Descrição"]] \
        .sort_values("Criado", ascending=False).rename(columns={"Criado": "Data"})
    tabela(detalhe, moeda=["Valor"], data_hora=["Data"])
    botao_download(df_filtrado.sort_values("Criado", ascending=False), f"analise_{solicitante}")


# ----------------------- EXECUÇÃO -----------------------
st.markdown(ESTILO, unsafe_allow_html=True)
try:
    df = load_data(str(ARQUIVO_DADOS), ARQUIVO_DADOS.stat().st_mtime)
except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")
    st.stop()

HOJE = date.today()
ONTEM = HOJE - timedelta(days=1)

st.navigation([
    st.Page(pagina_manutencao, title="Custos da Manutenção", default=True),
    st.Page(pagina_detalhada, title="Análise por Solicitante", url_path="solicitantes"),
], position="top").run()


