import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import re
from openpyxl import load_workbook

SHEET_NAME = "Crédito bancário"

st.set_page_config(page_title="RF | Destaques Crédito Bancário", layout="wide")
st.title("RF | Destaques Crédito Bancário")
st.caption('Lê apenas a aba "Crédito bancário" (cabeçalho fixo na linha 6) e monta Top 5 por indexador e prazo.')

# -----------------------------
# Helpers
# -----------------------------
def normalize_colname(c) -> str:
    if c is None:
        return ""
    return str(c).strip().replace("\n", " ").replace("\r", " ").replace("  ", " ")

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = list(df.columns)
    for cand in candidates:
        cand_l = cand.lower()
        for c in cols:
            if cand_l == c.lower():
                return c
        for c in cols:
            if cand_l in c.lower():
                return c
    return None

def to_numeric_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    s2 = s.astype(str).str.strip()
    s2 = s2.str.replace(".", "", regex=False)
    s2 = s2.str.replace(",", ".", regex=False)
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
    if days <= 1080:
        return "Médio (361 a 1080d)"
    return "Longo (acima de 1080d)"

def classify_indexer(raw) -> str | None:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    t = str(raw).strip().upper()

    if "IPCA" in t:
        return "IPCA"

    if "CDI" in t or "PÓS" in t or "POS" in t or t == "DI":
        return "Pós (CDI)"

    if "PRÉ" in t or "PRE" in t or "FIXA" in t:
        return "Pré"

    return None

def parse_rate_value(x) -> float | None:
    """
    Converte taxa textual para número (para ordenar).
    Exemplos:
      "110% CDI" -> 110
      "IPCA + 7,20%" -> 7.2
      "13,45% a.a." -> 13.45
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None

    s = str(x).strip().upper()
    s = s.replace("%", "")
    s = s.replace("A.A.", "").replace("AA", "").replace("A A", "")
    s = s.replace(" ", "")

    # 13,45 -> 13.45 e 1.234,56 -> 1234.56
    s = s.replace(".", "")
    s = s.replace(",", ".")

    m = re.search(r"(-?\d+(\.\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except:
        return None

def download_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def top_n_block(df: pd.DataFrame, idx_label: str, horizon_label: str, n: int) -> pd.DataFrame:
    sub = df[(df["indexador_pad"] == idx_label) & (df["horizonte"] == horizon_label)].copy()
    sub = sub.sort_values("taxa_num", ascending=False, na_position="last")
    return sub.head(n)

# -----------------------------
# Fast Excel reader (openpyxl read_only) - header fixed at row 6
# -----------------------------
@st.cache_data(show_spinner=False)
def read_credito_bancario_fast(file_bytes: bytes) -> pd.DataFrame:
    bio = BytesIO(file_bytes)
    wb = load_workbook(bio, read_only=True, data_only=True)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Não encontrei a aba "{SHEET_NAME}". Abas disponíveis: {wb.sheetnames}')

    ws = wb[SHEET_NAME]

    HEADER_ROW = 6  # cabeçalho fixo na linha 6

    header = [normalize_colname(cell.value) for cell in ws[HEADER_ROW]]

    data = []
    empty_streak = 0

    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        if row is None or all((x is None or str(x).strip() == "") for x in row):
            empty_streak += 1
            if empty_streak >= 20:
                break
            continue
        empty_streak = 0
        data.append(list(row))

    df = pd.DataFrame(data, columns=header)
    df.columns = [normalize_colname(c) for c in df.columns]
    df = df.dropna(axis=1, how="all")
    return df

# -----------------------------
# UI
# -----------------------------
uploaded = st.file_uploader("Envie a planilha (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])

with st.sidebar:
    st.header("Configurações")
    top_n = st.number_input("Top N por bloco", min_value=1, max_value=20, value=5, step=1)

    st.subheader("Filtros opcionais")
    use_rating_filter = st.checkbox("Aplicar rating mínimo", value=False)
    rating_min = st.text_input("Rating mínimo (ex: A-, A, AA-)", value="A-")

    use_min_app_filter = st.checkbox("Filtrar por aplicação mínima máxima", value=False)
    max_min_app = st.number_input("Aplicação mínima máxima (R$)", min_value=0, value=0, step=1000)

if not uploaded:
    st.info("Envie o arquivo para começar.")
    st.stop()

file_bytes = uploaded.getvalue()

with st.spinner('Lendo a aba "Crédito bancário"...'):
    try:
        raw = read_credito_bancario_fast(file_bytes)
    except Exception as e:
        st.error(f"Erro ao ler planilha: {e}")
        st.stop()

raw.columns = [normalize_colname(c) for c in raw.columns]

# Detect columns
col_emissor = find_col(raw, ["Emissor", "Banco", "Instituição", "Instituicao"])
col_produto = find_col(raw, ["Produto", "Ativo", "Tipo"])
col_indexador = find_col(raw, ["Indexador", "Remuneração", "Remuneracao", "Benchmark"])
col_taxa = find_col(raw, ["Tx. Máxima", "Taxa Máxima", "Tx Máxima", "Tx Máx", "Taxa", "Tx.", "Taxa Portal", "Tx. Portal"])
col_prazo = find_col(raw, ["Prazo", "Prazo (dias)", "Dias"])
col_venc = find_col(raw, ["Vencimento", "Data Vencimento", "Dt. Vencimento", "Data de vencimento"])
col_min_app = find_col(raw, ["Aplicação mínima", "Aplicacao minima", "Mínimo", "Minimo", "Aplicação Min"])
col_rating = find_col(raw, ["Rating", "Classificação", "Classificacao", "Nota"])

missing = []
if col_indexador is None:
    missing.append("Indexador")
if col_taxa is None:
    missing.append("Tx. Máxima/Taxa")
if col_prazo is None and col_venc is None:
    missing.append("Prazo ou Vencimento")

if missing:
    st.error("Colunas obrigatórias não encontradas: " + ", ".join(missing))
    st.write("Colunas detectadas:", list(raw.columns))
    st.stop()

df = raw.copy()

# prazo em dias
if col_prazo is not None:
    df["prazo_dias"] = to_numeric_series(df[col_prazo])
else:
    venc_dt = to_date_series(df[col_venc])
    today = pd.Timestamp(date.today())
    df["prazo_dias"] = (venc_dt - today).dt.days

df["horizonte"] = df["prazo_dias"].apply(categorize_horizon)
df["indexador_pad"] = df[col_indexador].apply(classify_indexer)
df["taxa_num"] = df[col_taxa].apply(parse_rate_value)

# aplicação mínima numérica (se existir)
if col_min_app is not None:
    df["_min_app_num"] = to_numeric_series(df[col_min_app])

# rating filter (mapeamento simples)
if use_rating_filter and col_rating is not None:
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
        st.warning("Rating mínimo não reconhecido no mapeamento. Filtro de rating ignorado.")
    else:
        df = df[df["_rating_score"].notna() & (df["_rating_score"] <= min_score)].copy()

# aplicação mínima filter
if use_min_app_filter and max_min_app > 0 and col_min_app is not None:
    df = df[df["_min_app_num"].notna() & (df["_min_app_num"] <= float(max_min_app))].copy()

# mantém só linhas úteis
df = df[df["horizonte"].notna() & df["indexador_pad"].notna() & df["taxa_num"].notna()].copy()

st.subheader("Base tratada")
st.caption(f"Linhas úteis: {len(df):,}".replace(",", "."))

preview_cols = []
for c in [col_emissor, col_produto, col_indexador, col_taxa, col_prazo, col_venc, col_min_app, col_rating]:
    if c is not None and c in df.columns and c not in preview_cols:
        preview_cols.append(c)
preview_cols += [c for c in ["prazo_dias", "horizonte", "indexador_pad", "taxa_num"] if c in df.columns]

st.dataframe(df[preview_cols].head(80), use_container_width=True, height=340)

indexers_order = ["Pós (CDI)", "IPCA", "Pré"]
horizons_order = ["Curto (até 360d)", "Médio (361 a 1080d)", "Longo (acima de 1080d)"]

st.divider()
st.subheader("Top do dia (Top N por bloco)")

tabs = st.tabs(indexers_order)

for i, idx in enumerate(indexers_order):
    with tabs[i]:
        cols = st.columns(3)
        for j, hz in enumerate(horizons_order):
            with cols[j]:
                st.markdown(f"### {hz}")
                b = top_n_block(df, idx, hz, int(top_n))

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

# consolidado download
blocks = []
for idx in indexers_order:
    for hz in horizons_order:
        b = top_n_block(df, idx, hz, int(top_n)).copy()
        b["Bloco"] = f"{idx} | {hz}"
        blocks.append(b)

result = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame()

st.divider()
st.subheader("Download do consolidado")

if result.empty:
    st.info("Não há dados suficientes para gerar o consolidado.")
else:
    st.download_button(
        "Baixar CSV consolidado",
        data=download_csv(result),
        file_name=f"top_ativos_credito_bancario_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )
