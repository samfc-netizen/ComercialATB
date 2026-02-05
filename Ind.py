import re
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Indicadores - Vendas", layout="wide")

# =============================
# CONFIG
# =============================
ARQUIVO_EXCEL = "base.xlsx"
MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
MESES_LONG = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}

# Faturamento 2024 (fornecido por você) — usado quando o ano anterior não existir na base e for 2024
FAT_2024_MES = {
    1: 421_375.43,
    2: 478_839.00,
    3: 514_630.18,
    4: 491_583.50,
    5: 561_725.99,
    6: 440_306.20,
    7: 360_277.10,
    8: 339_108.52,
    9: 480_860.64,
    10: 557_455.19,
    11: 515_291.01,
    12: 629_538.77,
}

# =============================
# HELPERS
# =============================
def normalize_col(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def parse_brl_number(v):
    """Converte número BR (1.234,56) / textos / floats em float."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)

    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return 0.0

    s = s.replace("\u00a0", " ")
    s = s.replace("R$", "").replace(" ", "")

    # Padrão BR: 1.234,56
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def format_brl(v: float) -> str:
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def safe_to_datetime(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def pct_br(x: float) -> str:
    try:
        return f"{x*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def _to_ascii_lower(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("ç", "c").replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
    s = s.replace("é", "e").replace("ê", "e").replace("í", "i").replace("ó", "o").replace("ô", "o")
    s = s.replace("ú", "u")
    return s


def parse_mes_to_num(v):
    """
    Tenta extrair MES_NUM (1..12) de:
    - 1..12
    - 'jan', 'fev', ...
    - 'janeiro', ...
    - '01/2026', '2026-01', etc. (pega o mês)
    Retorna int ou None.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = int(v)
        return n if 1 <= n <= 12 else None

    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return None

    s_low = _to_ascii_lower(s)

    # abreviado (jan, fev...)
    for i, abv in enumerate(MESES_PT, start=1):
        if s_low == abv:
            return i

    # mês por extenso
    if s_low in MESES_LONG:
        return MESES_LONG[s_low]

    # tenta extrair algo tipo mm/aaaa, aaaa-mm, etc.
    m = re.search(r"(?<!\d)(0?[1-9]|1[0-2])(?!\d)", s_low)
    if m:
        n = int(m.group(1))
        return n if 1 <= n <= 12 else None

    return None


def abc_classification(df_in: pd.DataFrame, value_col: str, label_col: str = "Produto") -> pd.DataFrame:
    """
    Gera Curva ABC baseada em value_col (Quantidade ou Faturamento).
    Regras:
      A: até 80% acumulado
      B: 80% a 95%
      C: acima de 95%
    """
    d = df_in[[label_col, value_col]].copy()
    d[value_col] = d[value_col].fillna(0.0)
    d = d.groupby(label_col, as_index=False)[value_col].sum()
    d = d.sort_values(value_col, ascending=False)

    total = float(d[value_col].sum())
    if total <= 0:
        d["%"] = 0.0
        d["% Acum"] = 0.0
        d["Curva"] = "C"
        return d

    d["%"] = d[value_col] / total
    d["% Acum"] = d["%"].cumsum()

    def _curva(p):
        if p <= 0.80:
            return "A"
        if p <= 0.95:
            return "B"
        return "C"

    d["Curva"] = d["% Acum"].apply(_curva)
    return d


def sum_fat_2024_for_months(meses_nums):
    return float(sum(FAT_2024_MES.get(m, 0.0) for m in meses_nums))


# =============================
# LOAD EXCEL + SELEÇÃO DAS ABAS
# =============================
st.title("Indicadores de Vendas")

try:
    xls = pd.ExcelFile(ARQUIVO_EXCEL)
except Exception as e:
    st.error(f"Erro ao abrir o arquivo '{ARQUIVO_EXCEL}': {e}")
    st.stop()

abas = xls.sheet_names

# =============================
# LOAD EXCEL (abas fixas)
# =============================

ABA_VENDAS = "RELATÓRIO DE VENDAS"
ABA_PRODUTOS = "BASE DE PRODUTOS"
ABA_CLIENTES = "BASE DE CLIENTES"

abas = xls.sheet_names

faltando = [
    aba for aba in [ABA_VENDAS, ABA_PRODUTOS, ABA_CLIENTES]
    if aba not in abas
]

if faltando:
    st.error(
        "As seguintes abas não foram encontradas no Excel:\n\n"
        + "\n".join(f"- {a}" for a in faltando)
        + "\n\nVerifique os nomes das abas no arquivo."
    )
    st.stop()

try:
    df_v = pd.read_excel(xls, sheet_name=ABA_VENDAS)
    df_p = pd.read_excel(xls, sheet_name=ABA_PRODUTOS)
    df_c = pd.read_excel(xls, sheet_name=ABA_CLIENTES)

except Exception as e:
    st.error(f"Erro ao ler as abas selecionadas: {e}")
    st.stop()

df_v.columns = [normalize_col(c) for c in df_v.columns]
df_p.columns = [normalize_col(c) for c in df_p.columns]
df_c.columns = [normalize_col(c) for c in df_c.columns]

# =============================
# PREP VENDAS
# =============================
required_cols = ["DATA2", "Valor total", "Valor custo", "Cliente", "UF", "LOCALIZAÇÃO", "BAIRRO", "CLASSIFICAÇÃO"]
missing = [c for c in required_cols if c not in df_v.columns]
if missing:
    st.error(
        "A aba de vendas não contém as colunas esperadas. Faltando: "
        + ", ".join(missing)
        + "\n\nConfira nomes, espaços e acentos (ex.: DATA2, Valor total, Valor custo, LOCALIZAÇÃO...)."
    )
    st.stop()

df = df_v.copy()

df["DATA2"] = safe_to_datetime(df["DATA2"])
df = df[df["DATA2"].notna()].copy()

df["Valor total"] = df["Valor total"].apply(parse_brl_number)
df["Valor custo"] = df["Valor custo"].apply(parse_brl_number)

for col in ["Cliente", "UF", "LOCALIZAÇÃO", "BAIRRO", "CLASSIFICAÇÃO"]:
    df[col] = df[col].astype(str).fillna("").str.strip()

# Remove linhas duplicadas após normalização (evita somas duplicadas no faturamento)
df = df.drop_duplicates()

df["ANO"] = df["DATA2"].dt.year
df["MES_NUM"] = df["DATA2"].dt.month
df["MES"] = df["MES_NUM"].apply(lambda m: MESES_PT[m - 1])

df["MARGEM_BRUTA_R$"] = df["Valor total"] - df["Valor custo"]
df["MARGEM_BRUTA_%"] = df.apply(
    lambda r: (r["MARGEM_BRUTA_R$"] / r["Valor total"]) if r["Valor total"] else 0.0,
    axis=1
)

# =============================
# FILTROS (ANO + PERÍODO)
# =============================
with st.sidebar:
    st.header("Filtros")

    anos = sorted(df["ANO"].dropna().unique().tolist())
    if not anos:
        st.error("Não há dados com DATA2 válida para filtrar por ano.")
        st.stop()

    ano_sel = st.selectbox("Ano", anos, index=len(anos) - 1)

    df_ano = df[df["ANO"] == ano_sel].copy()
    if df_ano.empty:
        st.warning("Não há dados para o ano selecionado.")
        st.stop()

    min_d = df_ano["DATA2"].min().date()
    max_d = df_ano["DATA2"].max().date()

    periodo = st.date_input(
        "Período (calendário BR)",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
    )
    if isinstance(periodo, tuple) and len(periodo) == 2:
        d_ini, d_fim = periodo
    else:
        d_ini, d_fim = min_d, max_d

df_f = df_ano[(df_ano["DATA2"].dt.date >= d_ini) & (df_ano["DATA2"].dt.date <= d_fim)].copy()

# meses selecionados no período (para filtrar BASE DE PRODUTOS por MÊS)
meses_sel = sorted(df_f["MES_NUM"].dropna().unique().tolist())

# =============================
# KPIs
# =============================
fat_total = df_f["Valor total"].sum()
custo_total = df_f["Valor custo"].sum()
margem_rs = fat_total - custo_total
margem_pct = (margem_rs / fat_total) if fat_total else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Faturamento (R$)", f"R$ {format_brl(fat_total)}")
k2.metric("Valor Custo (R$)", f"R$ {format_brl(custo_total)}")
k3.metric("Margem Bruta (R$)", f"R$ {format_brl(margem_rs)}")
k4.metric("Margem Bruta (%)", pct_br(margem_pct))

# =============================
# INDICADOR: CRESCIMENTO ANO-1 (mesmo período)
# =============================
st.subheader("Crescimento Ano-1 (mesmo período)")

ano_ant = int(ano_sel) - 1

# Receita atual (já está filtrada por período)
fat_atual_periodo = float(df_f["Valor total"].sum())

# Define período do ano anterior (mesmo range de datas)
d_ini_ts = pd.Timestamp(d_ini)
d_fim_ts = pd.Timestamp(d_fim)
d_ini_ant = (d_ini_ts - pd.DateOffset(years=1)).date()
d_fim_ant = (d_fim_ts - pd.DateOffset(years=1)).date()

df_ant_ano = df[df["ANO"] == ano_ant].copy()
tem_ano_ant_na_base = not df_ant_ano.empty

if tem_ano_ant_na_base:
    df_ant_periodo = df_ant_ano[
        (df_ant_ano["DATA2"].dt.date >= d_ini_ant) &
        (df_ant_ano["DATA2"].dt.date <= d_fim_ant)
    ].copy()
    fat_ant_periodo = float(df_ant_periodo["Valor total"].sum())
    origem_ant = f"Base XLSX (ano {ano_ant})"
else:
    # fallback apenas para 2024
    if ano_ant == 2024:
        fat_ant_periodo = sum_fat_2024_for_months(meses_sel)
        origem_ant = "Tabela fixa 2024 (por mês)"
    else:
        fat_ant_periodo = 0.0
        origem_ant = f"Sem dados do ano {ano_ant} (base vazia)"

crescimento_rs = fat_atual_periodo - fat_ant_periodo
crescimento_pct = (crescimento_rs / fat_ant_periodo) if fat_ant_periodo else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric(f"Faturamento {ano_sel} (período)", f"R$ {format_brl(fat_atual_periodo)}")
c2.metric(f"Faturamento {ano_ant} (período)", f"R$ {format_brl(fat_ant_periodo)}")
c3.metric("Crescimento (R$)", f"R$ {format_brl(crescimento_rs)}")
c4.metric("Crescimento (%)", pct_br(crescimento_pct))

st.caption(f"Fonte do Ano-1: **{origem_ant}**. Período comparado: {d_ini}–{d_fim} vs {d_ini_ant}–{d_fim_ant}.")

# Tabela mês-a-mês: atual vs ano-1
fat_mes_atual = (
    df_f.groupby("MES_NUM", as_index=False)["Valor total"].sum()
    .rename(columns={"Valor total": "FAT_ATUAL"})
)
fat_mes_atual["MÊS"] = fat_mes_atual["MES_NUM"].apply(lambda m: MESES_PT[m - 1])

if tem_ano_ant_na_base:
    df_ant_periodo["MES_NUM"] = df_ant_periodo["DATA2"].dt.month
    fat_mes_ant = (
        df_ant_periodo.groupby("MES_NUM", as_index=False)["Valor total"].sum()
        .rename(columns={"Valor total": "FAT_ANO_1"})
    )
else:
    # 2024 fixo por meses (sem recorte de dia) — usa apenas meses selecionados
    fat_mes_ant = pd.DataFrame({
        "MES_NUM": meses_sel if meses_sel else list(range(1, 13)),
    })
    fat_mes_ant["FAT_ANO_1"] = fat_mes_ant["MES_NUM"].apply(lambda m: float(FAT_2024_MES.get(m, 0.0)) if ano_ant == 2024 else 0.0)

tbl_yoy = pd.merge(fat_mes_atual, fat_mes_ant, on="MES_NUM", how="left")
tbl_yoy["FAT_ANO_1"] = tbl_yoy["FAT_ANO_1"].fillna(0.0)
tbl_yoy["DIF_R$"] = tbl_yoy["FAT_ATUAL"] - tbl_yoy["FAT_ANO_1"]
tbl_yoy["DIF_%"] = tbl_yoy.apply(lambda r: (r["DIF_R$"] / r["FAT_ANO_1"]) if r["FAT_ANO_1"] else 0.0, axis=1)

tbl_yoy_show = tbl_yoy[["MÊS", "FAT_ATUAL", "FAT_ANO_1", "DIF_R$", "DIF_%"]].copy()
tbl_yoy_show["FAT_ATUAL"] = tbl_yoy_show["FAT_ATUAL"].apply(lambda x: f"R$ {format_brl(x)}")
tbl_yoy_show["FAT_ANO_1"] = tbl_yoy_show["FAT_ANO_1"].apply(lambda x: f"R$ {format_brl(x)}")
tbl_yoy_show["DIF_R$"] = tbl_yoy_show["DIF_R$"].apply(lambda x: f"R$ {format_brl(x)}")
tbl_yoy_show["DIF_%"] = tbl_yoy_show["DIF_%"].apply(pct_br)

st.dataframe(tbl_yoy_show, use_container_width=True, hide_index=True)

st.divider()

# =============================
# 1) FATURAMENTO POR MÊS (BARRAS)
# =============================
fat_mes = (
    df_f.groupby("MES_NUM", as_index=False)["Valor total"].sum()
    .sort_values("MES_NUM")
)
fat_mes["MES"] = fat_mes["MES_NUM"].apply(lambda m: MESES_PT[m - 1])

fig_fat_mes = px.bar(
    fat_mes,
    x="MES",
    y="Valor total",
    title="Faturamento Total por Mês (R$)",
    hover_data={"Valor total": ":,.2f"},
)
st.plotly_chart(fig_fat_mes, use_container_width=True)

# =============================
# 2) RELAÇÃO FINANCEIRA POR MÊS (TABELA)
# =============================
st.subheader("Relação Financeira por Mês (Tabela)")

rel_mes = df_f.groupby("MES_NUM", as_index=False).agg(
    FATURAMENTO=("Valor total", "sum"),
    VALOR_CUSTO=("Valor custo", "sum"),
).sort_values("MES_NUM")

rel_mes["MARGEM_BRUTA_R$"] = rel_mes["FATURAMENTO"] - rel_mes["VALOR_CUSTO"]
rel_mes["MARGEM_BRUTA_%"] = rel_mes.apply(
    lambda r: (rel_mes.loc[r.name, "MARGEM_BRUTA_R$"] / r["FATURAMENTO"]) if r["FATURAMENTO"] else 0.0,
    axis=1
)
rel_mes["MÊS"] = rel_mes["MES_NUM"].apply(lambda m: MESES_PT[m - 1])

rel_mes_show = rel_mes[["MÊS", "FATURAMENTO", "VALOR_CUSTO", "MARGEM_BRUTA_R$", "MARGEM_BRUTA_%"]].copy()
rel_mes_show["FATURAMENTO"] = rel_mes_show["FATURAMENTO"].apply(lambda x: f"R$ {format_brl(x)}")
rel_mes_show["VALOR_CUSTO"] = rel_mes_show["VALOR_CUSTO"].apply(lambda x: f"R$ {format_brl(x)}")
rel_mes_show["MARGEM_BRUTA_R$"] = rel_mes_show["MARGEM_BRUTA_R$"].apply(lambda x: f"R$ {format_brl(x)}")
rel_mes_show["MARGEM_BRUTA_%"] = rel_mes_show["MARGEM_BRUTA_%"].apply(pct_br)

st.dataframe(rel_mes_show, use_container_width=True, hide_index=True)

st.divider()

# =============================
# 3) MAPA (TREEMAP) CONDICIONAL EM 1 GRÁFICO
# =============================
st.subheader("Mapa de Vendas (UF → Localização → Bairro no DF | demais: UF → Localização)")

base_map = df_f.copy()
for c in ["UF", "LOCALIZAÇÃO", "BAIRRO"]:
    base_map[c] = base_map[c].fillna("").astype(str).str.strip()
    base_map.loc[base_map[c] == "", c] = "(vazio)"

base_map["UF_UP"] = base_map["UF"].str.upper()
base_map["BAIRRO_MAPA"] = base_map.apply(
    lambda r: r["BAIRRO"] if r["UF_UP"] == "DF" else "— (sem detalhamento)",
    axis=1
)

map_agg = (
    base_map.groupby(["UF", "LOCALIZAÇÃO", "BAIRRO_MAPA"], as_index=False)
    .agg(FATURAMENTO=("Valor total", "sum"))
)

fig_map = px.treemap(
    map_agg,
    path=["UF", "LOCALIZAÇÃO", "BAIRRO_MAPA"],
    values="FATURAMENTO",
    title="Interaja no hover: caminho (UF/Localização/Bairro), Faturamento e % do Total (todas as UFs)"
)

fig_map.update_traces(
    hovertemplate=(
        "<b>%{label}</b><br>"
        "Caminho: %{currentPath}<br>"
        "Faturamento: R$ %{value:,.2f}<br>"
        "Representatividade (Total): %{percentRoot:.2%}"
        "<extra></extra>"
    )
)

st.plotly_chart(fig_map, use_container_width=True)

st.divider()

# =============================
# 4) TABELA POR UF: FATURAMENTO, CUSTO, MARGEM R$, MARGEM %
# =============================
st.subheader("Tabela por UF: Faturamento × Custo × Margem")

uf_tbl = df_f.groupby("UF", as_index=False).agg(
    FATURAMENTO=("Valor total", "sum"),
    VALOR_CUSTO=("Valor custo", "sum"),
)

uf_tbl["MARGEM_BRUTA_R$"] = uf_tbl["FATURAMENTO"] - uf_tbl["VALOR_CUSTO"]
uf_tbl["MARGEM_BRUTA_%"] = uf_tbl.apply(
    lambda r: (r["MARGEM_BRUTA_R$"] / r["FATURAMENTO"]) if r["FATURAMENTO"] else 0.0,
    axis=1
)

uf_tbl = uf_tbl.sort_values("FATURAMENTO", ascending=False)

uf_tbl_show = uf_tbl.copy()
uf_tbl_show["FATURAMENTO"] = uf_tbl_show["FATURAMENTO"].apply(lambda x: f"R$ {format_brl(x)}")
uf_tbl_show["VALOR_CUSTO"] = uf_tbl_show["VALOR_CUSTO"].apply(lambda x: f"R$ {format_brl(x)}")
uf_tbl_show["MARGEM_BRUTA_R$"] = uf_tbl_show["MARGEM_BRUTA_R$"].apply(lambda x: f"R$ {format_brl(x)}")
uf_tbl_show["MARGEM_BRUTA_%"] = uf_tbl_show["MARGEM_BRUTA_%"].apply(pct_br)

st.dataframe(uf_tbl_show, use_container_width=True, hide_index=True)

st.divider()

# =============================
# 5) CLIENTES POR UF (com linha de totais dinâmica)
# =============================
st.subheader("Clientes por UF (Faturamento × Custo × Margem)")

ufs_disp = sorted([u for u in df_f["UF"].dropna().unique().tolist() if str(u).strip() != ""])
uf_sel = st.selectbox("Selecione a UF", ["(Selecione)"] + ufs_disp, index=0)

if uf_sel == "(Selecione)":
    st.info("Selecione uma UF para listar os clientes e seus indicadores no período filtrado.")
else:
    df_uf = df_f[df_f["UF"] == uf_sel].copy()

    tab_cli = df_uf.groupby("Cliente", as_index=False).agg(
        FATURAMENTO=("Valor total", "sum"),
        VALOR_CUSTO=("Valor custo", "sum"),
    )
    tab_cli["MARGEM_BRUTA_R$"] = tab_cli["FATURAMENTO"] - tab_cli["VALOR_CUSTO"]
    tab_cli["MARGEM_BRUTA_%"] = tab_cli.apply(
        lambda r: (r["MARGEM_BRUTA_R$"] / r["FATURAMENTO"]) if r["FATURAMENTO"] else 0.0,
        axis=1
    )

    total_uf = tab_cli["FATURAMENTO"].sum()
    tab_cli["% UF (Fat)"] = tab_cli["FATURAMENTO"].apply(lambda x: (x / total_uf) if total_uf else 0.0)

    tab_cli = tab_cli.sort_values("FATURAMENTO", ascending=False)

    tot_fat = tab_cli["FATURAMENTO"].sum()
    tot_custo = tab_cli["VALOR_CUSTO"].sum()
    tot_margem = tot_fat - tot_custo
    tot_margem_pct = (tot_margem / tot_fat) if tot_fat else 0.0

    total_row = pd.DataFrame([{
        "Cliente": "TOTAL",
        "FATURAMENTO": tot_fat,
        "VALOR_CUSTO": tot_custo,
        "MARGEM_BRUTA_R$": tot_margem,
        "MARGEM_BRUTA_%": tot_margem_pct,
        "% UF (Fat)": 1.0 if total_uf else 0.0
    }])

    tab_cli2 = pd.concat([tab_cli, total_row], ignore_index=True)

    tab_show = tab_cli2.copy()
    tab_show["FATURAMENTO"] = tab_show["FATURAMENTO"].apply(lambda x: f"R$ {format_brl(x)}")
    tab_show["VALOR_CUSTO"] = tab_show["VALOR_CUSTO"].apply(lambda x: f"R$ {format_brl(x)}")
    tab_show["MARGEM_BRUTA_R$"] = tab_show["MARGEM_BRUTA_R$"].apply(lambda x: f"R$ {format_brl(x)}")
    tab_show["MARGEM_BRUTA_%"] = tab_show["MARGEM_BRUTA_%"].apply(pct_br)
    tab_show["% UF (Fat)"] = tab_show["% UF (Fat)"].apply(pct_br)

    st.dataframe(
        tab_show[["Cliente", "FATURAMENTO", "VALOR_CUSTO", "MARGEM_BRUTA_R$", "MARGEM_BRUTA_%", "% UF (Fat)"]],
        use_container_width=True,
        hide_index=True
    )

st.divider()

# =============================
# 6) PIZZA: FATURAMENTO POR CLASSIFICAÇÃO (TIPO DE CLIENTE)
# =============================
st.subheader("Faturamento por Tipo de Cliente (Classificação)")

cls_tbl = df_f.copy()
cls_tbl["CLASSIFICAÇÃO"] = cls_tbl["CLASSIFICAÇÃO"].fillna("").astype(str).str.strip()
cls_tbl.loc[cls_tbl["CLASSIFICAÇÃO"] == "", "CLASSIFICAÇÃO"] = "(vazio)"

cls = cls_tbl.groupby("CLASSIFICAÇÃO", as_index=False)["Valor total"].sum()
cls = cls.sort_values("Valor total", ascending=False)

fig_pizza = px.pie(
    cls,
    names="CLASSIFICAÇÃO",
    values="Valor total",
    title="Faturamento por Classificação",
)
fig_pizza.update_traces(texttemplate="%{percent:.1%}<br>R$ %{value:,.2f}")
st.plotly_chart(fig_pizza, use_container_width=True)

st.divider()

# =============================
# 7) RANKING DE CLIENTES (faturamento + % geral)
# =============================
st.subheader("Ranking de Clientes (Faturamento e % do Total)")

rank = df_f.groupby("Cliente", as_index=False)["Valor total"].sum().sort_values("Valor total", ascending=False)
tot_geral = rank["Valor total"].sum()
rank["% Geral"] = rank["Valor total"].apply(lambda x: (x / tot_geral) if tot_geral else 0.0)

rank_show = rank.copy()
rank_show["Valor total"] = rank_show["Valor total"].apply(lambda x: f"R$ {format_brl(x)}")
rank_show["% Geral"] = rank_show["% Geral"].apply(pct_br)

st.dataframe(rank_show, use_container_width=True, hide_index=True)

st.divider()

# =============================
# 8) EVOLUÇÃO DE VENDAS | CLIENTES (jan..dez + Total Geral) com zeros em vermelho
# =============================
st.subheader("Evolução de Vendas | Clientes (jan..dez)")

top_n = st.slider("Quantos clientes mostrar (por faturamento no período)?", 10, 300, 50, step=10)

top_clientes = rank.head(top_n)["Cliente"].tolist()
df_ev = df_f[df_f["Cliente"].isin(top_clientes)].copy()

pivot = df_ev.pivot_table(
    index="Cliente",
    columns="MES_NUM",
    values="Valor total",
    aggfunc="sum",
    fill_value=0.0
)

for m in range(1, 13):
    if m not in pivot.columns:
        pivot[m] = 0.0
pivot = pivot[list(range(1, 13))]

pivot.columns = [MESES_PT[m - 1] for m in pivot.columns]
pivot["Total Geral"] = pivot.sum(axis=1)
pivot = pivot.sort_values("Total Geral", ascending=False)

def style_zeros_red(v):
    try:
        val = float(v)
    except Exception:
        return ""
    if val == 0.0:
        return "background-color: #ffdddd"
    return ""

pivot_fmt = pivot.copy()
for c in MESES_PT + ["Total Geral"]:
    pivot_fmt[c] = pivot_fmt[c].apply(lambda x: f"R$ {format_brl(x)}")

st.dataframe(
    pivot_fmt.style.applymap(style_zeros_red, subset=MESES_PT),
    use_container_width=True
)

st.caption("Meses sem faturamento ficam zerados e destacados em vermelho.")

st.divider()

# =============================
# 9) PRODUTOS (BASE DE PRODUTOS)
# =============================
st.header("Produtos (Base de Produtos)")

required_prod = ["Produto", "Quantidade", "MÊS", "ANO", "Valor total", "Custo total"]
missing_prod = [c for c in required_prod if c not in df_p.columns]
if missing_prod:
    st.warning(
        "Não foi possível montar os indicadores de produtos porque faltam colunas na BASE DE PRODUTOS: "
        + ", ".join(missing_prod)
        + "\n\nConfira se os nomes estão exatamente assim (incluindo acentos) e tente novamente."
    )
else:
    df_prod = df_p.copy()

    df_prod["Produto"] = df_prod["Produto"].astype(str).fillna("").str.strip()
    df_prod["Quantidade"] = df_prod["Quantidade"].apply(parse_brl_number)
    df_prod["Valor total"] = df_prod["Valor total"].apply(parse_brl_number)
    df_prod["Custo total"] = df_prod["Custo total"].apply(parse_brl_number)

    df_prod["MES_NUM"] = df_prod["MÊS"].apply(parse_mes_to_num)

    # Usa coluna ANO para diferenciar meses repetidos (Jan..Dez e depois Jan..)
    df_prod["ANO"] = pd.to_numeric(df_prod["ANO"], errors="coerce")
    df_prod = df_prod[df_prod["ANO"].notna()].copy()

    # Filtra produtos pelo mesmo ano selecionado (Ano do filtro)
    df_prod = df_prod[df_prod["ANO"].astype(int) == int(ano_sel)].copy()

    # Filtro por meses do período selecionado (vendas)
    if meses_sel:
        df_prod_f = df_prod[df_prod["MES_NUM"].isin(meses_sel)].copy()
    else:
        df_prod_f = df_prod.copy()

    if df_prod_f.empty:
        st.info("Sem dados de produtos para os meses do período filtrado.")
    else:
        st.subheader("Tabela Mensal de Produtos (Quantidade)")

        tab_qtd = df_prod_f.pivot_table(
            index="Produto",
            columns="MES_NUM",
            values="Quantidade",
            aggfunc="sum",
            fill_value=0.0
        )

        for m in range(1, 13):
            if m not in tab_qtd.columns:
                tab_qtd[m] = 0.0
        tab_qtd = tab_qtd[list(range(1, 13))]

        tab_qtd.columns = [MESES_PT[m - 1] for m in tab_qtd.columns]
        tab_qtd["Total (Qtd)"] = tab_qtd.sum(axis=1)
        tab_qtd = tab_qtd.sort_values("Total (Qtd)", ascending=False)

        st.dataframe(tab_qtd, use_container_width=True)

        st.subheader("Curva ABC por Quantidade (Produtos)")
        abc_qtd = abc_classification(df_prod_f, value_col="Quantidade", label_col="Produto")
        abc_qtd_show = abc_qtd.copy()
        abc_qtd_show["Quantidade"] = abc_qtd_show["Quantidade"].apply(
            lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        abc_qtd_show["%"] = abc_qtd_show["%"].apply(pct_br)
        abc_qtd_show["% Acum"] = abc_qtd_show["% Acum"].apply(pct_br)
        st.dataframe(abc_qtd_show[["Produto", "Quantidade", "%", "% Acum", "Curva"]], use_container_width=True, hide_index=True)

        st.subheader("Curva ABC por Faturamento (Produtos)")
        abc_fat = abc_classification(df_prod_f, value_col="Valor total", label_col="Produto")
        abc_fat_show = abc_fat.copy()
        abc_fat_show["Valor total"] = abc_fat_show["Valor total"].apply(lambda x: f"R$ {format_brl(x)}")
        abc_fat_show["%"] = abc_fat_show["%"].apply(pct_br)
        abc_fat_show["% Acum"] = abc_fat_show["% Acum"].apply(pct_br)
        st.dataframe(abc_fat_show[["Produto", "Valor total", "%", "% Acum", "Curva"]], use_container_width=True, hide_index=True)

        st.subheader("Ranking de Produtos (Faturamento, Custo e Margem)")

        prod_rank = df_prod_f.groupby("Produto", as_index=False).agg(
            FATURAMENTO=("Valor total", "sum"),
            CUSTO=("Custo total", "sum"),
            QTD=("Quantidade", "sum"),
        )

        prod_rank["MARGEM_BRUTA_R$"] = prod_rank["FATURAMENTO"] - prod_rank["CUSTO"]
        prod_rank["MARGEM_BRUTA_%"] = prod_rank.apply(
            lambda r: (r["MARGEM_BRUTA_R$"] / r["FATURAMENTO"]) if r["FATURAMENTO"] else 0.0,
            axis=1
        )

        prod_rank = prod_rank.sort_values("FATURAMENTO", ascending=False)

        prod_rank_show = prod_rank.copy()
        prod_rank_show["QTD"] = prod_rank_show["QTD"].apply(
            lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        prod_rank_show["FATURAMENTO"] = prod_rank_show["FATURAMENTO"].apply(lambda x: f"R$ {format_brl(x)}")
        prod_rank_show["CUSTO"] = prod_rank_show["CUSTO"].apply(lambda x: f"R$ {format_brl(x)}")
        prod_rank_show["MARGEM_BRUTA_R$"] = prod_rank_show["MARGEM_BRUTA_R$"].apply(lambda x: f"R$ {format_brl(x)}")
        prod_rank_show["MARGEM_BRUTA_%"] = prod_rank_show["MARGEM_BRUTA_%"].apply(pct_br)

        st.dataframe(
            prod_rank_show[["Produto", "QTD", "FATURAMENTO", "CUSTO", "MARGEM_BRUTA_R$", "MARGEM_BRUTA_%"]],
            use_container_width=True,
            hide_index=True
        )

        st.caption("Em Produtos, o filtro é por MÊS (meses contidos no período selecionado em Vendas).")
