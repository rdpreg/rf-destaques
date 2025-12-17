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
st.caption(
    'Lê apenas a aba "Crédito bancário" e gera mensagens prontas para WhatsApp, '
    'organizadas por indexador e prazo.'
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
    if "PRÉ" in t or "PRE" in t:
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

def top_n_block(df, idx, horizon, n):
    sub = df[(df["indexador_pad"] == idx) & (df["horizonte"] == horizon)]
    return sub.sort_values("taxa_num", ascending=False).head(int(n))

# =============================
# Excel reader
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
# Columns
# =============================
col_emissor = find_col(df, ["Emissor"])
col_produto = find_col(df, ["Produto"])
col_indexador = find_col(df, ["Indexador"])
col_taxa = find_col(df, ["Tx. Portal", "Taxa Portal"])
col_venc = find_col(df, ["Vencimento"])
col_min = find_col(df, ["Aplicação", "Aplicacao"])

# =============================
# Transform
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
# WhatsApp config
# =============================
with st.sidebar:
    st.subheader("WhatsApp")
    mostrar_apenas_blocos_com_ativos = st.checkbox(
        "Mostrar apenas blocos com ativos", value=True
    )

TOP_WA = 5

def format_card(row, prefixo=""):
    titulo = f"{row[col_produto]} {row[col_emissor]}"
    taxa = f"{prefixo}{row['taxa_fmt']}"
    return (
        f"🏦*{titulo}*\n"
        f"⏰ Vencimento: {row['venc_fmt']}\n"
        f"📈 Taxa: {taxa}\n"
        f"💰mínimo: {row['aplic_min_fmt']}\n"
    )

def build_message(indexador, titulo, prefixo=""):
    hoje = datetime.now().strftime("%d/%m/%Y")
    msg = (
        "*Destaques de ativos Bancários*\n"
        f"🚨*TAXAS DE HOJE ({hoje})*\n\n"
        f"📍*{titulo}*\n\n"
    )

    for hz_label, hz_title in [
        ("Curto (até 360d)", "Curto Prazo (até 360d)"),
        ("Médio (361 a 1080d)", "Médio Prazo (361 a 1080d)"),
        ("Longo (acima de 1080d)", "Longo Prazo (acima de 1080d)"),
    ]:
        sub = top_n_block(df, indexador, hz_label, TOP_WA)
        if sub.empty and mostrar_apenas_blocos_com_ativos:
            continue

        msg += f"*{hz_title}*\n\n"
        if sub.empty:
            msg += "- (sem ativos hoje)\n\n"
        else:
            for _, r in sub.iterrows():
                msg += format_card(r, prefixo) + "\n"

    return msg

msg_pos = build_message("Pós (CDI)", "PÓS-FIXADOS")
msg_pre = build_message("Pré", "PRÉ-FIXADOS")
msg_ipca = build_message("IPCA", "IPCA", prefixo="IPCA+ ")

def copy_button(text):
    safe = text.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    html = f"""
    <button onclick="navigator.clipboard.writeText(`{safe}`)"
    style="cursor:pointer;padding:8px 12px;border-radius:8px;border:1px solid #ddd;background:white;">
    📋 Copiar
    </button>
    """
    components.html(html, height=45)

# =============================
# UI
# =============================
st.divider()
st.subheader("Mensagens prontas para WhatsApp")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("### Pós-fixados")
    st.text_area("Mensagem Pós", msg_pos, height=560)
    copy_button(msg_pos)

with c2:
    st.markdown("### Pré-fixados")
    st.text_area("Mensagem Pré", msg_pre, height=560)
    copy_button(msg_pre)

with c3:
    st.markdown("### IPCA")
    st.text_area("Mensagem IPCA", msg_ipca, height=560)
    copy_button(msg_ipca)
