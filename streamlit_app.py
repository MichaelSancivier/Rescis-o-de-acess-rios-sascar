import json
import difflib
from io import BytesIO
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import numpy as np
import pandas as pd
import streamlit as st

# ===================== Config =====================
st.set_page_config(page_title="Cálculo de Rescisão de Acessórios", layout="wide")

PRIMARY = "#27509b"
SECONDARY = "#fce500"

st.markdown(f"""
<style>
:root {{ --primary: {PRIMARY}; --secondary: {SECONDARY}; }}
h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {{ color: var(--primary) !important; }}
div[data-testid="metric-container"] {{
  background: rgba(252,229,0,0.15);
  border-left: 6px solid var(--primary);
  border-radius: 12px; padding:.5rem .75rem;
}}
div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {{ color: var(--primary) !important; }}
.stButton>button {{ border:1px solid var(--primary); border-radius:10px; }}
[data-testid="stDataFrame"] div[role="columnheader"] {{ background: var(--primary) !important; color:#fff !important; }}
</style>
""", unsafe_allow_html=True)

st.title("Calculo de Rescisão de acessórios.")
st.caption("Ferramenta de cálculo para rescisão de acessórios (com/sem devolução).")

# ===================== Utils =====================
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    mapa = {"ã":"a","â":"a","á":"a","à":"a","é":"e","ê":"e","í":"i","ó":"o","ô":"o","ú":"u","ç":"c"}
    for k, v in mapa.items():
        s = s.replace(k, v)
    while "  " in s:
        s = s.replace("  ", " ")
    return s

def parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.to_datetime(x)
    x = str(x).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return pd.to_datetime(x, format=fmt)
        except Exception:
            pass
    return pd.to_datetime(x, errors="coerce")

def read_any(uploaded, sheet=None):
    """Lê CSV ou XLSX. Para XLSX usa openpyxl; se faltar, mostra erro amigável."""
    if uploaded is None:
        return None
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded)
    if name.endswith(".xlsx"):
        try:
            import openpyxl  # garante engine
            if sheet:
                return pd.read_excel(uploaded, engine="openpyxl", sheet_name=sheet)
            return pd.read_excel(uploaded, engine="openpyxl")
        except Exception:
            st.error("Para ler Excel (.xlsx), é necessário **openpyxl**. Ele já está no requirements.txt. "
                     "Se persistir, envie a base em CSV.")
            st.stop()
    # fallback: tenta CSV
    return pd.read_csv(uploaded)

def brl(x) -> str:
    if pd.isna(x):
        x = 0
    q = Decimal(str(float(x))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{q:.2f}"
    inteiro, frac = s.split(".")
    inteiro_rev = inteiro[::-1]
    partes = [inteiro_rev[i:i+3] for i in range(0, len(inteiro_rev), 3)]
    inteiro_pt = ".".join(p[::-1] for p in partes[::-1])
    return f"R$ {inteiro_pt},{frac}"

def to_number(series: pd.Series, decimal_sep: str, thousand_sep: str) -> pd.Series:
    """Converte string numérica com separadores locais para float."""
    if series.dtype.kind in "biufc":
        return pd.to_numeric(series, errors="coerce")
    s = series.astype(str).str.strip()
    if thousand_sep:
        s = s.str.replace(thousand_sep, "", regex=False)
    if decimal_sep and decimal_sep != ".":
        s = s.str.replace(decimal_sep, ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

# ===================== Sidebar: upload + perfil =====================
st.sidebar.header("Upload")
file = st.sidebar.file_uploader("Base (CSV ou XLSX)", type=["csv", "xlsx"])

st.sidebar.header("Perfil (JSON)")
profile_upload = st.sidebar.file_uploader("Carregar perfil", type=["json"], key="profile_up")
loaded_profile = None
if profile_upload is not None:
    try:
        loaded_profile = json.load(profile_upload)
        st.sidebar.success("Perfil carregado.")
    except Exception:
        st.sidebar.error("JSON inválido.")

sheet_name = None
df_raw = None

if file is not None and file.name.lower().endswith(".xlsx"):
    try:
        import openpyxl
        xls = pd.ExcelFile(file, engine="openpyxl")
        default_sheet = (loaded_profile or {}).get("sheet", "")
        idx = xls.sheet_names.index(default_sheet) if default_sheet in xls.sheet_names else 0
        sheet_name = st.sidebar.selectbox("Planilha", options=xls.sheet_names, index=idx)
        df_raw = pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)
    except Exception:
        st.sidebar.error("Falhou ao ler XLSX. Verifique o arquivo.")
        st.stop()
elif file is not None:
    df_raw = pd.read_csv(file)

if df_raw is None:
    st.info("Faça o upload da base para iniciar.")
    st.stop()

# ===================== Wizard de Mapeamento =====================
st.subheader("Passo 1 — Mapeamento de colunas")

cols_raw = list(df_raw.columns)
cols_norm = {c: normalize_text(c) for c in cols_raw}

required_fields = ["cliente", "classe", "termo", "servico_acessorio", "valor_mensalidade"]
optional_fields = [
    "inicio_vigencia", "fim_vigencia", "meses_restantes",
    "valor_taxa_cancelamento", "valor_multa_nao_devolucao", "taxa_multa_25pct",
    "status_do_contrato", "instalado",
]

def suggest(field):
    guess = difflib.get_close_matches(field, [normalize_text(c) for c in cols_raw], n=1)
    if guess:
        for original, norm in cols_norm.items():
            if norm == guess[0]:
                return original
    return None

profile_cols = (loaded_profile or {}).get("columns", {})

sel_required = {}
req_cols = st.columns(len(required_fields))
for i, field in enumerate(required_fields):
    with req_cols[i]:
        default = profile_cols.get(field) or suggest(field)
        sel_required[field] = st.selectbox(
            f"Coluna para **{field}**",
            options=["(selecione)"] + cols_raw,
            index=(cols_raw.index(default) + 1 if default in cols_raw else 0),
            key=f"req_{field}",
        )

with st.expander("Mapeamento adicional (opcional)", expanded=False):
    sel_optional = {}
    opt_cols = st.columns(4)
    for idx, field in enumerate(optional_fields):
        with opt_cols[idx % 4]:
            default = profile_cols.get(field)
            sel_optional[field] = st.selectbox(
                f"{field}",
                options=["(não usar)"] + cols_raw,
                index=(cols_raw.index(default) + 1 if default in cols_raw else 0),
                key=f"opt_{field}",
            )

# ===================== Status do contrato - modos =====================
st.subheader("Passo 2 — Como obter o **status do contrato**")

status_mode_default = (loaded_profile or {}).get("status_mode", "by_mapping")
status_mode = st.radio(
    "Escolha o modo:",
    options=["by_mapping", "by_flags", "by_dates"],
    format_func=lambda x: {
        "by_mapping": "A) A partir de uma coluna de status (mapeando os valores)",
        "by_flags":   "B) Derivar de meses_restantes + flag de instalado",
        "by_dates":   "C) Derivar de fim_vigencia + flag de instalado (usa data de corte)",
    }[x],
    index=["by_mapping", "by_flags", "by_dates"].index(status_mode_default),
    key="status_mode",
)

status_config = (loaded_profile or {})
installed_cfg = (status_config.get("columns", {}) or {}).get("installed_flag", {})
norm_values_default = {
    "true": installed_cfg.get("true_values", ["SIM", "INSTALADO", "TRUE", "1"]),
    "false": installed_cfg.get("false_values", ["NAO", "NÃO", "FALSE", "0"]),
}

# valores padrão para uso no bloco de salvar perfil
status_col = ""
map_cvi = map_cvn = map_svi = map_svn = []
installed_col = ""
true_vals = ",".join(norm_values_default["true"])
false_vals = ",".join(norm_values_default["false"])
data_corte = date.today()

if status_mode == "by_mapping":
    default_col = profile_cols.get("status_column")
    status_col = st.selectbox(
        "Coluna que contém o status",
        options=["(selecione)"] + cols_raw,
        index=(cols_raw.index(default_col) + 1 if default_col in cols_raw else 0),
    )
    vals = df_raw[status_col] if status_col and status_col != "(selecione)" else pd.Series([], dtype=object)
    sample_vals = sorted(pd.Series(vals).astype(str).str.strip().unique().tolist())[:100]

    st.caption("Mapeie os valores da coluna de status para as 4 categorias oficiais.")
    c1, c2 = st.columns(2)
    with c1:
        map_cvi = st.multiselect(
            "Com vigência e instalado",
            sample_vals,
            default=(status_config.get("status_map", {}).get("Com vigência e instalado", []) if loaded_profile else []),
        )
        map_cvn = st.multiselect(
            "Com vigência e não instalado",
            sample_vals,
            default=(status_config.get("status_map", {}).get("Com vigência e não instalado", []) if loaded_profile else []),
        )
    with c2:
        map_svi = st.multiselect(
            "Sem vigência e instalado",
            sample_vals,
            default=(status_config.get("status_map", {}).get("Sem vigência e instalado", []) if loaded_profile else []),
        )
        map_svn = st.multiselect(
            "Sem vigência e não instalado",
            sample_vals,
            default=(status_config.get("status_map", {}).get("Sem vigência e não instalado", []) if loaded_profile else []),
        )

elif status_mode in ("by_flags", "by_dates"):
    inst_default = profile_cols.get("installed")
    installed_col = st.selectbox(
        "Coluna que indica instalado (sim/não)",
        options=["(selecione)"] + cols_raw,
        index=(cols_raw.index(inst_default) + 1 if inst_default in cols_raw else 0),
    )
    st.caption("Defina quais valores serão interpretados como Verdadeiro (instalado) e Falso.")
    c1, c2 = st.columns(2)
    with c1:
        true_vals = st.text_input(
            "Valores 'verdadeiro' (separados por vírgula)",
            value=",".join(norm_values_default["true"]),
        )
    with c2:
        false_vals = st.text_input(
            "Valores 'falso' (separados por vírgula)",
            value=",".join(norm_values_default["false"]),
        )
    if status_mode == "by_dates":
        cutoff_default = (status_config.get("calculo", {}) or {}).get("data_corte", "today")
        if cutoff_default == "today":
            data_corte = date.today()
        else:
            try:
                data_corte = pd.to_datetime(cutoff_default).date()
            except Exception:
                data_corte = date.today()

# ===================== Passo 3 — Regras e parsing =====================
st.subheader("Passo 3 — Regras de cálculo e parsing")

percent_default = (loaded_profile or {}).get("calculo", {}).get("percent_multa", 0.25)
percent_multa = st.number_input(
    "Percentual padrão da multa (25% = 0.25)",
    min_value=0.0, max_value=1.0, step=0.01, value=float(percent_default),
)

parsing_default = (loaded_profile or {}).get("parsing", {}) or {}
decimal_sep = st.selectbox("Separador decimal", options=[",", "."],
                           index=(0 if parsing_default.get("decimal", ",") == "," else 1))

idx_map = {".": 0, ",": 1, " ": 2, "(nenhum)": 3}
thousand_sep_val = parsing_default.get("thousands", ".")
thousand_choice = st.selectbox(
    "Separador de milhar",
    options=[".", ",", " ", "(nenhum)"],
    index=idx_map.get(thousand_sep_val, 0),
)
thousand_sep = "" if thousand_choice == "(nenhum)" else thousand_choice

st.markdown("---")

# ===================== Gerar DataFrame mapeado =====================
def can_proceed_required():
    return all(sel_required[f] and sel_required[f] != "(selecione)" for f in required_fields)

proceed = st.button("Aplicar mapeamento e **gerar dados**", disabled=not can_proceed_required())
if not proceed:
    st.stop()

# 1) Renomeia colunas escolhidas para nomes internos
df = df_raw.copy()
columns_map = {sel_required[k]: k for k in required_fields if sel_required[k] != "(selecione)"}
columns_map.update({sel_optional[k]: k for k in optional_fields
                    if sel_optional.get(k) and sel_optional[k] != "(não usar)"})
df = df.rename(columns=columns_map)

# 2) Parsing: números e datas
num_cols = ["valor_mensalidade", "valor_taxa_cancelamento", "valor_multa_nao_devolucao",
            "taxa_multa_25pct", "meses_restantes"]
for c in num_cols:
    if c in df.columns:
        df[c] = to_number(df[c], decimal_sep=decimal_sep, thousand_sep=thousand_sep)

for c in ["inicio_vigencia", "fim_vigencia"]:
    if c in df.columns:
        df[c] = df[c].apply(parse_date_any)

# 3) Meses restantes (se não vier)
if "meses_restantes" not in df.columns or df["meses_restantes"].isna().any():
    if "fim_vigencia" in df.columns:
        ref_date = pd.Timestamp(data_corte) if status_mode == "by_dates" else pd.Timestamp(date.today())
        dias = (df["fim_vigencia"] - ref_date).dt.days.fillna(0).clip(lower=0)
        df["meses_restantes"] = np.ceil(dias / 30.0).astype(int)
    else:
        df["meses_restantes"] = 0

# 4) Multa 25% (se não vier)
if "taxa_multa_25pct" not in df.columns or df["taxa_multa_25pct"].isna().any():
    df["taxa_multa_25pct"] = df["valor_mensalidade"].fillna(0) * df["meses_restantes"].fillna(0) * float(percent_multa)

# 5) Status do contrato
def make_installed_flag(series: pd.Series, true_list, false_list):
    s = series.astype(str).str.strip().str.upper()
    tset = {normalize_text(x).upper() for x in true_list}
    fset = {normalize_text(x).upper() for x in false_list}
    def conv(v):
        vn = normalize_text(v).upper()
        if vn in tset:
            return True
        if vn in fset:
            return False
        return vn in {"1", "TRUE", "SIM", "INSTALADO", "INSTALLED", "YES", "Y"}
    return s.map(conv)

if status_mode == "by_mapping":
    if status_col and status_col != "(selecione)":
        raw = df[status_col].astype(str).str.strip()
        def map_status(v):
            v_norm = normalize_text(v)
            if v in (map_cvi or []) or v_norm in [normalize_text(x) for x in (map_cvi or [])]:
                return "Com vigência e instalado"
            if v in (map_cvn or []) or v_norm in [normalize_text(x) for x in (map_cvn or [])]:
                return "Com vigência e não instalado"
            if v in (map_svi or []) or v_norm in [normalize_text(x) for x in (map_svi or [])]:
                return "Sem vigência e instalado"
            if v in (map_svn or []) or v_norm in [normalize_text(x) for x in (map_svn or [])]:
                return "Sem vigência e não instalado"
            return None
        df["status_do_contrato"] = raw.map(map_status)
    else:
        st.error("Selecione a coluna de status para o modo A.")
        st.stop()

elif status_mode == "by_flags":
    if installed_col and installed_col != "(selecione)":
        true_list = [x.strip() for x in true_vals.split(",") if x.strip()]
        false_list = [x.strip() for x in false_vals.split(",") if x.strip()]
        installed_bool = make_installed_flag(df[installed_col], true_list, false_list)
        df["status_do_contrato"] = np.select(
            [
                (df["meses_restantes"] > 0) & installed_bool,
                (df["meses_restantes"] > 0) & (~installed_bool),
                (df["meses_restantes"] == 0) & installed_bool,
            ],
            [
                "Com vigência e instalado",
                "Com vigência e não instalado",
                "Sem vigência e instalado",
            ],
            default="Sem vigência e não instalado",
        )
    else:
        st.error("Selecione a coluna de instalado para o modo B.")
        st.stop()

elif status_mode == "by_dates":
    if ("fim_vigencia" in df.columns) and installed_col and installed_col != "(selecione)":
        true_list = [x.strip() for x in true_vals.split(",") if x.strip()]
        false_list = [x.strip() for x in false_vals.split(",") if x.strip()]
        installed_bool = make_installed_flag(df[installed_col], true_list, false_list)
        ref_date = pd.Timestamp(data_corte)
        meses_pos = (df["fim_vigencia"] - ref_date).dt.days.fillna(0) > 0
        df["status_do_contrato"] = np.select(
            [
                meses_pos & installed_bool,
                meses_pos & (~installed_bool),
                (~meses_pos) & installed_bool,
            ],
            [
                "Com vigência e instalado",
                "Com vigência e não instalado",
                "Sem vigência e instalado",
            ],
            default="Sem vigência e não instalado",
        )
    else:
        st.error("Para o modo C, informe 'fim_vigencia' e a coluna 'instalado'.")
        st.stop()

# Preenche taxas ausentes como 0
for c in ["valor_taxa_cancelamento", "valor_multa_nao_devolucao"]:
    if c not in df.columns:
        df[c] = 0.0
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

# 6) Cálculos finais
def valor_com_devolucao(r):
    stt, m25, taxa = r["status_do_contrato"], r["taxa_multa_25pct"], r["valor_taxa_cancelamento"]
    if stt.startswith("Com vigência"):
        return float(m25) + float(taxa)
    if stt == "Sem vigência e instalado":
        return float(taxa)
    return 0.0

def valor_sem_devolucao(r):
    stt, m25, multa = r["status_do_contrato"], r["taxa_multa_25pct"], r["valor_multa_nao_devolucao"]
    if stt.startswith("Com vigência"):
        return float(m25) + float(multa)
    if stt == "Sem vigência e instalado":
        return float(multa)
    return 0.0

df["valor_cobrar_com_devolucao"] = df.apply(valor_com_devolucao, axis=1)
df["valor_cobrar_sem_devolucao"]  = df.apply(valor_sem_devolucao,  axis=1)

# ===================== Salvar perfil =====================
st.subheader("Salvar/Carregar perfil")

profile = {
    "sheet": sheet_name or "",
    "columns": {
        **{k: v for k, v in {f: sel_required[f] for f in required_fields}.items() if v and v != "(selecione)"},
        **{k: v for k, v in {f: sel_optional[f] for f in optional_fields}.items() if v and v != "(não usar)"},
        "installed_flag": {
            "source": (installed_col if status_mode in ("by_flags", "by_dates") else ""),
            "true_values": ([x.strip() for x in true_vals.split(",")] if status_mode in ("by_flags", "by_dates") else []),
            "false_values": ([x.strip() for x in false_vals.split(",")] if status_mode in ("by_flags", "by_dates") else []),
        },
    },
    "status_mode": status_mode,
    "status_column": (status_col if status_mode == "by_mapping" else ""),
    "status_map": ({
        "Com vigência e instalado": map_cvi,
        "Com vigência e não instalado": map_cvn,
        "Sem vigência e instalado": map_svi,
        "Sem vigência e não instalado": map_svn,
    } if status_mode == "by_mapping" else {}),
    "calculo": {
        "percent_multa": float(percent_multa),
        "meses_restantes": "from_fim_vigencia" if "fim_vigencia" in df.columns else "provided",
        "data_corte": str(date.today()) if status_mode != "by_dates" else str(data_corte),
    },
    "parsing": {"decimal": decimal_sep, "thousands": thousand_sep or ""},
}
st.download_button(
    "Baixar perfil (JSON)",
    data=json.dumps(profile, ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="perfil_mapeamento.json",
    mime="application/json",
)

st.markdown("---")

# ===================== Observação sliders =====================
st.subheader("Filtros")
st.info(
    "💡 **Sobre os sliders**: as **pontas** mostram o **mínimo** e o **máximo** que existem no arquivo carregado. "
    "As **duas alças** definem o **intervalo selecionado**. Esses filtros **combinam** com os campos acima."
)

# ===================== Filtros =====================
c1, c2, c3 = st.columns(3)
with c1:
    sel_clientes = st.multiselect("Cliente", sorted(df["cliente"].astype(str).unique().tolist()))
with c2:
    sel_classes = st.multiselect("Classe", sorted(df["classe"].astype(str).unique().tolist()))
with c3:
    sel_termos = st.multiselect("Termo", sorted(df["termo"].astype(str).unique().tolist()))

c4, c5 = st.columns(2)
with c4:
    sel_servicos = st.multiselect("Serviço/Acessório", sorted(df["servico_acessorio"].astype(str).unique().tolist()))
with c5:
    status_opts = [
        "Com vigência e instalado",
        "Com vigência e não instalado",
        "Sem vigência e instalado",
        "Sem vigência e não instalado",
    ]
    sel_status = st.multiselect("Status do contrato", status_opts)

c6, c7 = st.columns(2)

min_cdev, max_cdev = float(df["valor_cobrar_com_devolucao"].min()), float(df["valor_cobrar_com_devolucao"].max())
min_sdev, max_sdev = float(df["valor_cobrar_sem_devolucao"].min()), float(df["valor_cobrar_sem_devolucao"].max())

with c6:
    faixa_cdev = st.slider(
        label="Faixa de valores (Com Devolução)",

