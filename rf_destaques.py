import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import re
from openpyxl import load_workbook

SHEET_NAME = "Crédito bancário"

st.set_page_config(page_title="RF | Destaques Crédito Bancário", layout="wide")
st.title("RF | Destaques Crédito Bancário")
st.caption(
    'Lê apenas a aba "Crédito bancário" (cabeçalho fixo na linha 6) e monta Top 5 por indexador e prazo, '
    'com taxas, aplicação mínima e vencimento formatados.'
)

# =============================
# Helpers
# =============================
def normalize_colname(c):
    if c is None:
        return ""
    return str(c).strip().replace("\n", " ").replace("\r", " ")

def find_col(df, candidates):
    for cand in candidates:
        cand_l = cand.lower()
        for c in df.columns:
            if cand_l == c.lower() or cand_l in c.lower():
                return c
    return None

def to_numeric_series(s):
    if s is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    s = s.astype(str).str.strip()
    s = s.str.replace(".", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    s = s.str.extract(r"(-?\d+(\.\d+)?)", expand=True)[0]
    return pd.to_numeric(s, errors="coerce")

def to_date_series(s):
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def categorize_horizon(days):
    if pd.isna(days):
        return None
    if days <= 360:
        return "Curto (até 360d)"
    if days <= 1080:
        return "Médio (361 a 1080d)"
    return "Longo (acima de 1080d)"

def classify_indexer(raw):
    if raw is None or pd.isna(raw):
        return None
    t = str(raw).upper()
    if "IPCA" in t:
        return "IPCA"
    if "CDI" in t or "PÓS" in t or "POS" in t:
        return "Pós (CDI)"
    if "PRÉ" in t or "PRE" in t or "FIXA" in t:
        return "Pré"
    return None

def parse_rate_value(x):
    if x is None or pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).upper().replace("%", "").replace(" ", "")
    m = re.search(r"(-?\d[\d\.,]*)", s)
    if not m:
        return None
    num = m.group(1)
    if "." in num and "," in num:
        num = num.replace(".", "").replace(",", ".")
    elif "," in num:
        num = num.replace(",", ".")
    return float(num)

def format_rate_for_display(rate_num, indexador):
    if rate_num is None or pd.isna(rate_num):
        return ""
    val = float(rate_num)

    if indexador == "Pós (CDI)":
        val = val * 100 if val <= 2 else val
        return f"{val:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

    if val <= 1.5:
        val = val * 100
    return f"{val:.2f}%".replace(".", ",")

def format_currency_brl(value):
    if value is None or pd.isna(value):
        return ""
    return f"R$ {int(value):,}".replace(",", ".")

def format_date_br(dt):
    if dt is None or pd.isna(dt):
        return ""
    return pd.to_datetime(dt).strftime("%d/%m/%Y")

def top_n_block(df, idx, horizon, n):
    sub = df[(df["indexador_pad"] == idx) & (df["horizonte"] == horizon)]
    return sub.sort_values("taxa_num", ascending=False).head(n)

# =============================
# Fast Excel reader
# =============================
@st.cache_data(show_spinner=False)
def read_credito_bancario_fast(file_bytes):
    bio = BytesIO(file_bytes)
    wb = load_workbook(bio, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    HEADER_ROW = 6
    header = [normalize_colname(c.value) for c in ws[HEADER_ROW]]

    data = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        if row is None or all(x is None or str(x).strip() == "" for x in row):
            continue
        data.append(row)

    df = pd.DataFrame(data, columns=header)
    return df.dropna(axis=1, how="all")

# =============================
# UI
# =============================
uploaded = st.file_uploader("Envie a planilha (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])
if not uploaded:
    st.stop()

df = read_credito_bancario_fast(uploaded.getvalue())

col_emissor = find_col(df, ["Emissor"])
col_produto = find_col(df, ["Produto"])
col_indexador = find_col(df, ["Indexador"])
col_taxa = find_col(df, ["Tx"])
col_prazo = find_col(df, ["Prazo"])
col_venc = find_col(df, ["Vencimento"])
col_min = find_col(df, ["Aplicação"])
col_rating = find_col(df, ["Rating"])

df["prazo_dias"] = to_numeric_series(df[col_prazo]) if col_prazo else (
    to_date_series(df[col_venc]) - pd.Timestamp(date.today())
).dt.days

df["horizonte"] = df["prazo_dias"].apply(categorize_horizon)
df["indexador_pad"] = df[col_indexador].apply(classify_indexer)

df["taxa_num"] = df[col_taxa].apply(parse_rate_value)
df["taxa_fmt"] = df.apply(lambda r: format_rate_for_display(r["taxa_num"], r["indexador_pad"]), axis=1)

df["aplic_min_num"] = to_numeric_series(df[col_min])
df["aplic_min_fmt"] = df["aplic_min_num"].apply(format_currency_brl)

df["venc_fmt"] = to_date_series(df[col_venc]).apply(format_date_br)

df = df[df["horizonte"].notna() & df["indexador_pad"].notna()]

indexers = ["Pós (CDI)", "IPCA", "Pré"]
horizons = ["Curto (até 360d)", "Médio (361 a 1080d)", "Longo (acima de 1080d)"]

tabs = st.tabs(indexers)

for i, idx in enumerate(indexers):
    with tabs[i]:
        cols = st.columns(3)
        for j, hz in enumerate(horizons):
            with cols[j]:
                st.markdown(f"### {hz}")
                b = top_n_block(df, idx, hz, 5)
                if b.empty:
                    st.info("Sem ativos")
                else:
                    st.dataframe(
                        b[
                            [
                                col_emissor,
                                col_produto,
                                "taxa_fmt",
                                "aplic_min_fmt",
                                col_rating,
                                "venc_fmt",
                            ]
                        ],
                        use_container_width=True,
                        height=260,
                    )
