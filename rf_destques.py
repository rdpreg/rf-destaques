import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO

st.set_page_config(page_title="Seleção de Crédito Bancário", layout="wide")

SHEET_NAME = "Crédito bancário"

# -----------------------------
# Helpers
# -----------------------------
def normalize_colname(c: str) -> str:
    return (
        str(c)
        .strip()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("  ", " ")
    )

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = list(df.columns)
    lower_map = {c.lower(): c for c in cols}
    for cand in candidates:
        cand = cand.lower()
        for c in cols:
            if cand == c.lower():
                return c
        for c in cols:
            if cand in c.lower():
                return c
    return None

def to_numeric_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    s2 = s.astype(str).str.strip()

    # troca vírgula por ponto quando for decimal brasileiro
    s2 = s2.str.replace(".", "", regex=False)
    s2 = s2.str.replace(",", ".", regex=False)

    # extrai primeiro número, caso venha como "360 dias"
    extracted = s2.str.extract(r"(-?\d+(\.\d+)?)", expand=True)[0]
    return pd.to_numeric(extracted, errors="coerce")

def to_date_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def categorize_horizon(days: float) -> str | None:
    if pd.isna(days):
        return None
    if days <= 360:
        return "Curto (até 360d)"
    if 360 < days <= 1080:
        return "Médio (361 a 1080d)"
    return "Longo (acima de 1080d)"

def classify_indexer(raw: str) -> str | None:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    t = str(raw).strip().upper()

    # pós
    if "CDI" in t or "PÓS" in t or "POS" in t or "DI" in t:
        return "Pós (CDI)"
    # ipca
    if "IPCA" in t:
        return "IPCA"
    # pré
    if "PRÉ" in t or "PRE" in t or "FIXA" in t:
        return "Pré"

    return None

def parse_rate_value(x) -> float | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip().upper()

    # remove símbolos comuns
    s = s.replace("%", "").replace("A.A.", "").replace("AA", "").replace("A A", "")
    s = s.replace(" ", "")

    # padroniza separador decimal
    s = s.replace(".", "").replace(",", ".")

    # casos comuns:
    # 110CDI, 110%CDI, 1.2CDI etc
    # IPCA+7, PRE13.5
    import re
    m = re.search(r"(-?\d+(\.\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except:
        return None

def make_block(df: pd.DataFrame, idx_label: str, horizon_label: str, top_n: int) -> pd.DataFrame:
    sub = df[(df["indexador_pad"] == idx_label) & (df["horizonte"] == horizon_label)].copy()
    sub = sub.sort_values(by=["taxa_num"], ascending=False, na_position="last")
    return sub.head(top_n)

def download_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


# -----------------------------
# UI
# -----------------------------
st.title("Seleção diária de emissores bancários (Top 5 por indexador e prazo)")

st.markdown(
    """
Upload da planilha diária. O app lê apenas a aba **Crédito bancário** e monta 9 blocos:
**IPCA, Pré, Pós** × **Curto, Médio, Longo**, com **Top 5 por maior taxa**.
"""
)

uploaded = st.file_uploader("Envie a planilha (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

with st.sidebar:
    st.header("Configurações")
    top_n = st.number_input("Top N por bloco", min_value=1, max_value=20, value=5, step=1)

    st.subheader("Filtros mínimos (opcional)")
    use_rating_filter = st.checkbox("Aplicar filtro de rating mínimo", value=False)
    rating_min = st.text_input("Rating mínimo (ex: A-, A, AA-)", value="A-")

    use_min_app_filter = st.checkbox("Filtrar por aplicação mínima máxima", value=False)
    max_min_app = st.number_input("Aplicação mínima máxima (R$)", min_value=0, value=0, step=1000)

    st.caption("Dica: se você quiser só destacar e não excluir, deixe os filtros desligados.")


if not uploaded:
    st.info("Envie o arquivo para começar.")
    st.stop()

# -----------------------------
# Read Excel, only the sheet
# -----------------------------
try:
    raw = pd.read_excel(uploaded, sheet_name=SHEET_NAME, engine="openpyxl")
except ValueError:
    st.error(f'Não encontrei a aba "{SHEET_NAME}". Confira o nome exato na planilha.')
    st.stop()
except Exception as e:
    st.error(f"Erro ao ler o arquivo: {e}")
    st.stop()

# normalize columns
raw.columns = [normalize_colname(c) for c in raw.columns]

# Try to detect important columns
col_emissor = find_col(raw, ["Emissor", "Banco", "Instituição", "Instituicao"])
col_produto = find_col(raw, ["Produto", "Ativo", "Tipo"])
col_indexador = find_col(raw, ["Indexador", "Indexador / Remuneração", "Remuneração", "Remuneracao", "Benchmark"])
col_taxa = find_col(raw, ["Tx. Máxima", "Taxa Máxima", "Tx Máxima", "Taxa", "Tx.", "Taxa Portal", "Tx. Portal"])
col_prazo = find_col(raw, ["Prazo", "Prazo (dias)", "Dias", "Duration"])
col_venc = find_col(raw, ["Vencimento", "Data Vencimento", "Dt. Vencimento", "Data de vencimento"])
col_min_app = find_col(raw, ["Aplicação mínima", "Aplicacao minima", "Mínimo", "Minimo", "Aplicação Min"])
col_rating = find_col(raw, ["Rating", "Classificação", "Classificacao", "Nota"])

missing = []
if col_indexador is None:
    missing.append("Indexador")
if col_taxa is None:
    missing.append("Tx. Máxima (ou Taxa)")
if col_prazo is None and col_venc is None:
    missing.append("Prazo (ou Vencimento)")

if missing:
    st.error(
        "Não consegui identificar as colunas necessárias: "
        + ", ".join(missing)
        + ".\n\n"
        + "Se quiser, me diga os nomes exatos dessas colunas na sua planilha que eu ajusto o código."
    )
    st.stop()

df = raw.copy()

# Build prazo em dias
prazo_days = None
if col_prazo is not None:
    prazo_days = to_numeric_series(df[col_prazo])
else:
    venc_dt = to_date_series(df[col_venc])
    today = pd.Timestamp(date.today())
    prazo_days = (venc_dt - today).dt.days

df["prazo_dias"] = pd.to_numeric(prazo_days, errors="coerce")

# Horizon category
df["horizonte"] = df["prazo_dias"].apply(categorize_horizon)

# Normalize indexer
df["indexador_pad"] = df[col_indexador].apply(classify_indexer)

# Numeric rate
df["taxa_num"] = df[col_taxa].apply(parse_rate_value)

# Optional rating filter (simple, lexicographic fallback)
# If your rating has a consistent scale, we can improve this later.
if use_rating_filter and col_rating is not None:
    # Keep rows with rating not null and rating >= rating_min in a rough way.
    # This is a simple guardrail. If you want exact ordering, we map rating to scores.
    df = df[df[col_rating].notna()].copy()

    # A crude comparator: keep those that start with the same or "better" prefix.
    # We'll implement a mapping that generally works for common scales.
    rating_map = {
        "AAA": 1, "AA+": 2, "AA": 3, "AA-": 4,
        "A+": 5, "A": 6, "A-": 7,
        "BBB+": 8, "BBB": 9, "BBB-": 10,
        "BB+": 11, "BB": 12, "BB-": 13,
        "B+": 14, "B": 15, "B-": 16,
        "CCC": 17, "CC": 18, "C": 19, "D": 20
    }
    def rating_to_score(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        t = str(x).strip().upper().replace(" ", "")
        return rating_map.get(t)

    df["_rating_score"] = df[col_rating].apply(rating_to_score)
    min_score = rating_to_score(rating_min)
    if min_score is None:
        st.warning("Rating mínimo não reconhecido no mapeamento. Filtro de rating foi ignorado.")
        df.drop(columns=["_rating_score"], inplace=True, errors="ignore")
    else:
        df = df[df["_rating_score"].notna() & (df["_rating_score"] <= min_score)].copy()

# Optional aplicação mínima filter
if use_min_app_filter and max_min_app > 0 and col_min_app is not None:
    df["_min_app_num"] = to_numeric_series(df[col_min_app])
    df = df[df["_min_app_num"].notna() & (df["_min_app_num"] <= float(max_min_app))].copy()

# Keep only rows we can classify
df = df[df["indexador_pad"].notna() & df["horizonte"].notna() & df["taxa_num"].notna()].copy()

st.subheader("Base tratada (após filtros)")
st.caption(f"Linhas úteis: {len(df):,}".replace(",", "."))

# Show a compact preview
preview_cols = []
for c in [col_emissor, col_produto, col_indexador, col_taxa, col_prazo, col_venc, col_min_app, col_rating]:
    if c is not None and c in df.columns and c not in preview_cols:
        preview_cols.append(c)
preview_cols += [c for c in ["prazo_dias", "horizonte", "indexador_pad", "taxa_num"] if c in df.columns]
st.dataframe(df[preview_cols].head(50), use_container_width=True, height=320)

# -----------------------------
# Build the 9 blocks
# -----------------------------
indexers_order = ["Pós (CDI)", "IPCA", "Pré"]
horizons_order = ["Curto (até 360d)", "Médio (361 a 1080d)", "Longo (acima de 1080d)"]

blocks = []
for idx in indexers_order:
    for hz in horizons_order:
        b = make_block(df, idx, hz, int(top_n))
        b = b.copy()
        b["Bloco"] = f"{idx} | {hz}"
        blocks.append(b)

result = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame()

st.divider()
st.subheader("Top do dia por bloco")

tabs = st.tabs(indexers_order)

for t_idx, idx in enumerate(indexers_order):
    with tabs[t_idx]:
        cols = st.columns(3)
        for j, hz in enumerate(horizons_order):
            with cols[j]:
                st.markdown(f"### {hz}")
                b = make_block(df, idx, hz, int(top_n))

                # Columns to show, keep useful ones
                show_cols = []
                for c in [col_emissor, col_produto, col_indexador, col_taxa, col_min_app, col_rating, col_venc]:
                    if c is not None and c in b.columns and c not in show_cols:
                        show_cols.append(c)
                for c in ["prazo_dias", "taxa_num"]:
                    if c in b.columns and c not in show_cols:
                        show_cols.append(c)

                if b.empty:
                    st.info("Sem ativos nesse bloco.")
                else:
                    st.dataframe(b[show_cols], use_container_width=True, height=260)

st.divider()
st.subheader("Download")
if result.empty:
    st.info("Não há dados suficientes para gerar o consolidado.")
else:
    st.caption("CSV consolidado com todos os blocos (já com a coluna Bloco).")
    st.download_button(
        "Baixar CSV consolidado",
        data=download_csv(result),
        file_name=f"top_ativos_credito_bancario_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )
