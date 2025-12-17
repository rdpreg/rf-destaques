import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import re
from openpyxl import load_workbook
import streamlit.components.v1 as components

SHEET_NAME = "Crédito bancário"

st.set_page_config(page_title="RF | Destaques Crédito Bancário", layout="wide")
st.title("RF | Destaques Crédito Bancário")
st.caption("Seleção e comunicação diária de ativos bancários")

# =============================
# Helpers gerais
# =============================
def normalize_colname(c):
    if c is None:
        return ""
    return str(c).strip().replace("\n", " ").replace("\r", " ")

def find_col(df, candidates):
    for cand in candidates:
        cl = cand.lower()
        for c in df.columns:
            if cl == c.lower() or cl in c.lower():
                return c
    return None

def to_numeric_series(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    s = s.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    s = s.str.extract(r"(-?\d+(\.\d+)?)", expand=True)[0]
    return pd.to_numeric(s, errors="coerce")

def to_date_series(s):
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
    v = float(rate_num)
    if indexador == "Pós (CDI)":
        v = v * 100 if v <= 2 else v
        return f"{v:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    if v <= 1.5:
        v = v * 100
    return f"{v:.2f}%".replace(".", ",")

def format_currency_brl(v):
    if v is None or pd.isna(v):
        return ""
    return f"R$ {int(v):,}".replace(",", ".")

def format_date_br(d):
    if d is None or pd.isna(d):
        return ""
    return pd.to_datetime(d).strftime("%d/%m/%Y")

# =============================
# Leitura rápida Excel
# =============================
@st.cache_data(show_spinner=False)
def read_credito_bancario_fast(file_bytes):
    wb = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    header = [normalize_colname(c.value) for c in ws[6]]
    data = []
    for row in ws.iter_rows(min_row=7, values_only=True):
        if row and not all(x is None or str(x).strip() == "" for x in row):
            data.append(row)
    return pd.DataFrame(data, columns=header).dropna(axis=1, how="all")

# =============================
# Upload
# =============================
uploaded = st.file_uploader("Envie a planilha (.xlsx ou .xlsm)", type=["xlsx", "xlsm"])
if not uploaded:
    st.stop()

df = read_credito_bancario_fast(uploaded.getvalue())

# =============================
# Colunas
# =============================
col_emissor = find_col(df, ["Emissor"])
col_produto = find_col(df, ["Produto"])
col_indexador = find_col(df, ["Indexador"])
col_taxa = find_col(df, ["Tx. Portal", "Taxa Portal"])
col_venc = find_col(df, ["Vencimento"])
col_min = find_col(df, ["Aplicação", "Aplicacao"])

if not all([col_emissor, col_produto, col_indexador, col_taxa, col_venc, col_min]):
    st.error("Colunas obrigatórias não encontradas.")
    st.stop()

# =============================
# Transformações
# =============================
df["indexador_pad"] = df[col_indexador].apply(classify_indexer)
df["_venc_dt"] = to_date_series(df[col_venc])
df["prazo_dias"] = (df["_venc_dt"] - pd.Timestamp(date.today())).dt.days
df["horizonte"] = df["prazo_dias"].apply(categorize_horizon)

df["taxa_num"] = df[col_taxa].apply(parse_rate_value)
df["taxa_fmt"] = df.apply(lambda r: format_rate_for_display(r["taxa_num"], r["indexador_pad"]), axis=1)

df["aplic_min_num"] = to_numeric_series(df[col_min])
df["aplic_min_fmt"] = df["aplic_min_num"].apply(format_currency_brl)
df["venc_fmt"] = df["_venc_dt"].apply(format_date_br)

df = df[df["indexador_pad"].notna() & df["horizonte"].notna() & df["taxa_num"].notna()].copy()

# =============================
# Mensagem WhatsApp (todos os ativos)
# =============================
def build_whatsapp_message_all(df):
    hoje = datetime.now().strftime("%d/%m/%Y")

    indexadores = [
        ("Pós (CDI)", "📍*PÓS-FIXADOS*"),
        ("Pré", "📍*PRÉ-FIXADOS*"),
        ("IPCA", "📍*IPCA*"),
    ]
    prazos = [
        ("Curto (até 360d)", "⏱ *Curto (até 360d)*"),
        ("Médio (361 a 1080d)", "⏱ *Médio (361 a 1080d)*"),
        ("Longo (acima de 1080d)", "⏱ *Longo (acima de 1080d)*"),
    ]

    def card(row, prefixo=""):
        titulo = f"{row[col_produto]} {row[col_emissor]}"
        taxa = f"{prefixo}{row['taxa_fmt']}"
        return (
            f"🏦*{titulo}*\n"
            f"⏰ Vencimento: {row['venc_fmt']}\n"
            f"📈 Taxa: {taxa}\n"
            f"💰mínimo: {row['aplic_min_fmt']}\n"
        )

    parts = [
        "*Destaques de ativos Bancários*",
        f"🚨*TAXAS DE HOJE ({hoje})*\n"
    ]

    for idx, idx_title in indexadores:
        parts.append(idx_title)
        for hz, hz_title in prazos:
            sub = df[(df["indexador_pad"] == idx) & (df["horizonte"] == hz)].sort_values("taxa_num", ascending=False)
            if sub.empty:
                continue
            parts.append(hz_title)
            prefixo = "IPCA+ " if idx == "IPCA" else ""
            for _, r in sub.iterrows():
                parts.append(card(r, prefixo))
    return "\n".join(parts)

msg = build_whatsapp_message_all(df)

# =============================
# UI Mensagem + copiar
# =============================
st.divider()
st.subheader("Mensagem pronta para WhatsApp")

st.text_area("Mensagem", value=msg, height=600)

def copy_button(text):
    safe = text.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    html = f"""
    <button onclick="navigator.clipboard.writeText(`{safe}`)"
    style="cursor:pointer;padding:10px 14px;border-radius:10px;border:1px solid #ddd;background:white;">
    📋 Copiar mensagem
    </button>
    """
    components.html(html, height=60)

copy_button(msg)
