
import io
import re
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Aurora Pricing Cockpit",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Styling ----------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #ffffff 0%, #fafafa 100%);
    border: 1px solid #e8e8e8;
    padding: 14px 16px;
    border-radius: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,.04);
}
.badge {
    display:inline-block; padding:4px 10px; border-radius:999px;
    font-size:12px; font-weight:600; border:1px solid transparent;
}
.badge-red{background:#fff1f2; color:#b42318; border-color:#fecdd3;}
.badge-orange{background:#fff7ed; color:#c2410c; border-color:#fdba74;}
.badge-yellow{background:#fefce8; color:#a16207; border-color:#fde047;}
.badge-green{background:#ecfdf3; color:#067647; border-color:#86efac;}
.badge-blue{background:#eff6ff; color:#1d4ed8; border-color:#93c5fd;}
.badge-gray{background:#f4f4f5; color:#52525b; border-color:#d4d4d8;}
.section-card{
    border:1px solid #ececec; border-radius:18px; padding:16px 18px;
    background:#ffffff; box-shadow:0 1px 4px rgba(0,0,0,.04); margin-bottom:10px;
}
.small-note{font-size:12px; color:#666;}
.big-title{font-size:28px; font-weight:700; margin-bottom:0px;}
.subtle{color:#666;}
</style>
""", unsafe_allow_html=True)

TARGET_SHEETS = ["MAESTRA de precios", "MLC -SKU", "CONTROL DE PROMOCIONES"]
CAMPAIGN_COLS = ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]


# ---------- Helpers ----------
def safe_int(v, default=0):
    try:
        if pd.isna(v):
            return default
        return int(float(v))
    except Exception:
        return default


def safe_float(v, default=np.nan):
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default


def money(v):
    if pd.isna(v):
        return "—"
    try:
        return f"${float(v):,.0f}".replace(",", ".")
    except Exception:
        return str(v)


def pct(v):
    if pd.isna(v):
        return "—"
    try:
        fv = float(v)
        if abs(fv) <= 1.2:
            fv *= 100
        return f"{fv:.1f}%"
    except Exception:
        return str(v)


def normalize_sku(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = s.replace(".0", "")
    digits = re.sub(r"\D", "", s)
    return digits or s


def normalize_mlc(x):
    if pd.isna(x):
        return None
    s = str(x).upper().strip()
    digits = re.sub(r"\D", "", s)
    return f"MLC{digits}" if digits else None


def extract_mlcs(raw):
    if pd.isna(raw):
        return []
    s = str(raw).upper()
    hits = re.findall(r"MLC\s*-?\s*(\d+)|\b(\d{7,})\b", s)
    out = []
    for a, b in hits:
        num = a or b
        if num:
            out.append(f"MLC{num}")
    # dedupe preserve order
    ded = []
    for v in out:
        if v not in ded:
            ded.append(v)
    return ded


def urgency_info(days):
    if pd.isna(days):
        return ("Sin fecha", "badge-gray", 99)
    days = int(days)
    if days < 0:
        return (f"Vencida hace {abs(days)}d", "badge-red", 0)
    if days == 0:
        return ("Vence hoy", "badge-red", 0)
    if days == 1:
        return ("Vence mañana", "badge-orange", 1)
    if days == 2:
        return ("Vence en 2 días", "badge-yellow", 2)
    if days <= 7:
        return (f"Vence en {days} días", "badge-blue", 3)
    return ("Vigente", "badge-green", 4)


def read_excel_all(uploaded_file):
    raw_bytes = uploaded_file.getvalue()
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheets = {sh: pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sh) for sh in xls.sheet_names}
    return sheets


def df_series(df, col, default=""):
    if col in df.columns:
        return df[col]
    return pd.Series([default] * len(df), index=df.index, dtype=object)


@st.cache_data(show_spinner=False)
def build_model(master_bytes, compras_bytes=None):
    sheets = read_excel_all(master_bytes)

    master = sheets.get("MAESTRA de precios", pd.DataFrame()).copy()
    bridge = sheets.get("MLC -SKU", pd.DataFrame()).copy()
    promos = sheets.get("CONTROL DE PROMOCIONES", pd.DataFrame()).copy()

    # master prep
    if "Unnamed: 12" in master.columns and "MLC_aux" not in master.columns:
        master = master.rename(columns={"Unnamed: 12": "MLC_aux"})
    master["SKU_norm"] = master.get("SKU", pd.Series(dtype=object)).apply(normalize_sku)
    mlc_cols = [c for c in master.columns if str(c).startswith("MLC")]
    primary = None
    for col in ["MLC", "MLC_aux", "MLC.1"]:
        if col in master.columns:
            primary = col
            break
    if primary is None and mlc_cols:
        primary = mlc_cols[0]
    master["MLC_norm"] = master.get(primary, pd.Series(dtype=object)).apply(normalize_mlc) if primary else None

    # bridge prep
    bridge["SKU_norm"] = bridge.iloc[:, 0].apply(normalize_sku) if len(bridge.columns) > 0 else None
    if len(bridge.columns) > 1:
        bridge["MLC_norm"] = bridge.iloc[:, 1].apply(normalize_mlc)
    else:
        bridge["MLC_norm"] = None
    bridge = bridge[["SKU_norm", "MLC_norm"]].dropna().drop_duplicates()

    # promos prep
    if not promos.empty:
        sku_col = promos.columns[0]
        mlc_col = "N° Publicación" if "N° Publicación" in promos.columns else promos.columns[2]
        promos["SKU_norm"] = promos[sku_col].apply(normalize_sku)
        promos["mlc_list"] = promos[mlc_col].apply(extract_mlcs)
        promos = promos.explode("mlc_list")
        promos["MLC_norm"] = promos["mlc_list"]
        promos["Precio promocional"] = pd.to_numeric(promos.get("Precio promocional"), errors="coerce")
        for c in CAMPAIGN_COLS:
            if c in promos.columns:
                promos[c] = pd.to_datetime(promos[c], errors="coerce")
        promos["next_campaign_date"] = promos[CAMPAIGN_COLS].min(axis=1) if set(CAMPAIGN_COLS).intersection(promos.columns) else pd.NaT
        promos["days_to_next"] = (promos["next_campaign_date"].dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    else:
        for c in ["SKU_norm", "MLC_norm", "Precio promocional", "next_campaign_date", "days_to_next"]:
            promos[c] = np.nan

    # build SKU <-> MLC universe
    sku_mlc = pd.concat([
        master[["SKU_norm", "MLC_norm"]].dropna(),
        bridge[["SKU_norm", "MLC_norm"]].dropna(),
        promos[["SKU_norm", "MLC_norm"]].dropna(),
    ], ignore_index=True).drop_duplicates()

    # aggregated product table
    product = master.copy()
    bridge_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index()
    promo_agg = promos.groupby("SKU_norm").agg(
        total_promos=("MLC_norm", "count"),
        promo_mlcs=("MLC_norm", lambda s: list(pd.unique([x for x in s if pd.notna(x)]))),
        next_campaign_date=("next_campaign_date", "min"),
        min_promo_price=("Precio promocional", "min"),
        max_promo_price=("Precio promocional", "max"),
        ads_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
    ).reset_index() if not promos.empty else pd.DataFrame(columns=["SKU_norm"])
    product = product.merge(bridge_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product["days_to_next"] = (pd.to_datetime(product["next_campaign_date"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    product["search_text"] = (
        df_series(product, "SKU").astype(str).fillna("") + " | " +
        df_series(product, "DESCRIPCIÓN").astype(str).fillna("") + " | " +
        df_series(product, "MLC_norm").astype(str).fillna("")
    ).str.lower()

    # compras prep
    compras = pd.DataFrame()
    compras_summary = pd.DataFrame()
    if compras_bytes is not None:
        try:
            c_sheets = read_excel_all(compras_bytes)
            # choose largest sheet
            best_name = max(c_sheets, key=lambda k: len(c_sheets[k]))
            compras = c_sheets[best_name].copy()
            norm_cols = {c: str(c).strip().lower() for c in compras.columns}
            # heuristics
            sku_col = next((c for c in compras.columns if "sku" in norm_cols[c] or "codigo" in norm_cols[c] or "art" in norm_cols[c]), None)
            date_col = next((c for c in compras.columns if "fecha" in norm_cols[c]), None)
            prov_col = next((c for c in compras.columns if "prove" in norm_cols[c]), None)
            price_col = next((c for c in compras.columns if "precio" in norm_cols[c] or "costo" in norm_cols[c] or "neto" in norm_cols[c]), None)
            qty_col = next((c for c in compras.columns if "cant" in norm_cols[c]), None)

            if sku_col:
                compras["SKU_norm"] = compras[sku_col].apply(normalize_sku)
            if date_col:
                compras["fecha_compra"] = pd.to_datetime(compras[date_col], errors="coerce")
            if prov_col:
                compras["proveedor"] = compras[prov_col]
            if price_col:
                compras["precio_compra"] = pd.to_numeric(compras[price_col], errors="coerce")
            if qty_col:
                compras["cantidad"] = pd.to_numeric(compras[qty_col], errors="coerce")
            compras_summary = compras.groupby("SKU_norm").agg(
                ultima_compra=("fecha_compra", "max"),
                ultimo_precio_compra=("precio_compra", "last"),
                proveedor_ultimo=("proveedor", "last"),
                proveedores=("proveedor", lambda s: " | ".join(pd.unique([str(x) for x in s if pd.notna(x)]))),
                precio_min_hist=("precio_compra", "min"),
                precio_max_hist=("precio_compra", "max"),
                compras_registros=("SKU_norm", "count"),
            ).reset_index()
            product = product.merge(compras_summary, on="SKU_norm", how="left")
        except Exception:
            compras = pd.DataFrame()
            compras_summary = pd.DataFrame()

    return {
        "sheets": sheets,
        "master": master,
        "bridge": bridge,
        "promos": promos,
        "sku_mlc": sku_mlc,
        "product": product,
        "compras": compras,
        "compras_summary": compras_summary,
    }


def compute_decision_row(row):
    margin = safe_float(row.get("MARGEN MELI 2"))
    if pd.isna(margin):
        margin = safe_float(row.get("MARGEN MELI 1"))
    ads_txt = str(row.get("ads_comment") or "").lower()
    days = row.get("days_to_next")
    total_promos = safe_int(row.get("total_promos"), 0)
    if total_promos == 0:
        return "Sin promo", "badge-gray"
    if not pd.isna(days) and days <= 1:
        return "Urgente renovar", "badge-red"
    if not pd.isna(margin) and margin < 0:
        return "Margen negativo", "badge-red"
    if "paus" in ads_txt:
        return "Revisar Ads", "badge-orange"
    if not pd.isna(days) and days <= 7:
        return "Revisar esta semana", "badge-yellow"
    return "Bajo control", "badge-green"


def promo_editable_columns(df):
    preferred = [
        "SKU_norm", "MLC_norm", "Descripción", "% F", "% F.1", "Precio promocional",
        "Motivo promoción", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"
    ]
    return [c for c in preferred if c in df.columns]


def make_download_workbook(all_sheets, master_df, bridge_df, promo_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            if name == "MAESTRA de precios":
                master_df.drop(columns=[c for c in master_df.columns if c.endswith("_norm") or c in ["search_text"]], errors="ignore").to_excel(writer, sheet_name=name, index=False)
            elif name == "MLC -SKU":
                base = bridge_df.copy()
                # restore conventional columns
                if "SKU_norm" in base.columns:
                    base = base.rename(columns={"SKU_norm": "SKU"})
                if "MLC_norm" in base.columns:
                    base = base.rename(columns={"MLC_norm": "Número de publicación"})
                base.to_excel(writer, sheet_name=name, index=False)
            elif name == "CONTROL DE PROMOCIONES":
                base = promo_df.copy()
                base = base.drop(columns=[c for c in ["mlc_list", "SKU_norm", "MLC_norm", "next_campaign_date", "days_to_next"] if c in base.columns], errors="ignore")
                base.to_excel(writer, sheet_name=name, index=False)
            else:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    return out.getvalue()


def render_badge(text, cls):
    st.markdown(f'<span class="badge {cls}">{text}</span>', unsafe_allow_html=True)


# ---------- Sidebar ----------
st.sidebar.title("Aurora Pricing Cockpit")
master_file = st.sidebar.file_uploader("Maestra integrada", type=["xlsx"], key="master")
compras_file = st.sidebar.file_uploader("Compras históricas", type=["xlsx"], key="compras")
st.sidebar.caption("Carga manual para trabajar siempre con el último archivo.")

if master_file is None:
    st.info("Carga la maestra integrada para empezar.")
    st.stop()

model = build_model(master_file, compras_file)
all_sheets = model["sheets"]

# Session data for edits
if "master_df" not in st.session_state:
    st.session_state.master_df = model["master"].copy()
    st.session_state.bridge_df = model["bridge"].copy()
    st.session_state.promos_df = model["promos"].copy()
    st.session_state.source_name = master_file.name

master_df = st.session_state.master_df
bridge_df = st.session_state.bridge_df
promos_df = st.session_state.promos_df

# Refresh model from edited data
edited_model = {
    "product": build_model(master_file, compras_file)["product"] if False else None
}

# re-build product with current session edits (lightweight)
tmp_sheets = dict(all_sheets)
tmp_sheets["MAESTRA de precios"] = master_df.copy()
tmp_sheets["MLC -SKU"] = bridge_df.copy()
tmp_sheets["CONTROL DE PROMOCIONES"] = promos_df.copy()

# manual rebuild without cache
def rebuild_from_session(sheets, compras_file):
    raw = sheets
    master = raw["MAESTRA de precios"].copy()
    bridge = raw["MLC -SKU"].copy()
    promos = raw["CONTROL DE PROMOCIONES"].copy()

    if "Unnamed: 12" in master.columns and "MLC_aux" not in master.columns:
        master = master.rename(columns={"Unnamed: 12": "MLC_aux"})
    master["SKU_norm"] = master.get("SKU", pd.Series(dtype=object)).apply(normalize_sku)
    primary = "MLC" if "MLC" in master.columns else ("MLC_aux" if "MLC_aux" in master.columns else None)
    if primary is None and any(str(c).startswith("MLC") for c in master.columns):
        primary = [c for c in master.columns if str(c).startswith("MLC")][0]
    master["MLC_norm"] = master.get(primary, pd.Series(dtype=object)).apply(normalize_mlc) if primary else None

    if len(bridge.columns) >= 2:
        if "SKU_norm" not in bridge.columns:
            bridge["SKU_norm"] = bridge.iloc[:, 0].apply(normalize_sku)
        if "MLC_norm" not in bridge.columns:
            bridge["MLC_norm"] = bridge.iloc[:, 1].apply(normalize_mlc)
    bridge = bridge[[c for c in bridge.columns if c in ["SKU_norm", "MLC_norm"] or c not in ["SKU_norm", "MLC_norm"]]]

    if not promos.empty:
        if "SKU_norm" not in promos.columns:
            promos["SKU_norm"] = promos.iloc[:, 0].apply(normalize_sku)
        if "MLC_norm" not in promos.columns:
            mlc_col = "N° Publicación" if "N° Publicación" in promos.columns else promos.columns[2]
            promos["mlc_list"] = promos[mlc_col].apply(extract_mlcs)
            promos = promos.explode("mlc_list")
            promos["MLC_norm"] = promos["mlc_list"]
        for c in CAMPAIGN_COLS:
            if c in promos.columns:
                promos[c] = pd.to_datetime(promos[c], errors="coerce")
        promos["Precio promocional"] = pd.to_numeric(promos.get("Precio promocional"), errors="coerce")
        promos["next_campaign_date"] = promos[CAMPAIGN_COLS].min(axis=1) if set(CAMPAIGN_COLS).intersection(promos.columns) else pd.NaT
        promos["days_to_next"] = (promos["next_campaign_date"].dt.normalize() - pd.Timestamp.today().normalize()).dt.days

    sku_mlc = pd.concat([
        master[["SKU_norm", "MLC_norm"]].dropna(),
        bridge[["SKU_norm", "MLC_norm"]].dropna() if set(["SKU_norm", "MLC_norm"]).issubset(bridge.columns) else pd.DataFrame(columns=["SKU_norm", "MLC_norm"]),
        promos[["SKU_norm", "MLC_norm"]].dropna() if set(["SKU_norm", "MLC_norm"]).issubset(promos.columns) else pd.DataFrame(columns=["SKU_norm", "MLC_norm"]),
    ], ignore_index=True).drop_duplicates()

    product = master.copy()
    bridge_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index()
    promo_agg = promos.groupby("SKU_norm").agg(
        total_promos=("MLC_norm", "count"),
        promo_mlcs=("MLC_norm", lambda s: list(pd.unique([x for x in s if pd.notna(x)]))),
        next_campaign_date=("next_campaign_date", "min"),
        min_promo_price=("Precio promocional", "min"),
        max_promo_price=("Precio promocional", "max"),
        ads_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])) if "Ads/Comentario" in promos.columns else ("MLC_norm", lambda s: ""),
        motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])) if "Motivo promoción" in promos.columns else ("MLC_norm", lambda s: ""),
    ).reset_index() if not promos.empty else pd.DataFrame(columns=["SKU_norm"])
    product = product.merge(bridge_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product["days_to_next"] = (pd.to_datetime(product["next_campaign_date"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    product["search_text"] = (
        df_series(product, "SKU").astype(str).fillna("") + " | " +
        df_series(product, "DESCRIPCIÓN").astype(str).fillna("") + " | " +
        df_series(product, "MLC_norm").astype(str).fillna("")
    ).str.lower()

    compras = model["compras"]
    if not compras.empty:
        compras_summary = model["compras_summary"]
        product = product.merge(compras_summary, on="SKU_norm", how="left")
    return product, promos

product_df, promos_df_view = rebuild_from_session(tmp_sheets, compras_file)

# Sidebar actions
download_bytes = make_download_workbook(all_sheets, st.session_state.master_df, st.session_state.bridge_df, st.session_state.promos_df)
st.sidebar.download_button(
    "Descargar Excel actualizado",
    data=download_bytes,
    file_name=f"MAESTRA_ACTUALIZADA_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

page = st.sidebar.radio(
    "Módulos",
    ["Cockpit por producto", "Operador de promos", "Centro de control", "Alta de producto"],
)

# ---------- Cockpit ----------
if page == "Cockpit por producto":
    st.markdown('<div class="big-title">Ficha de decisión por producto</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Una vista clara para decidir precio, promo, vencimientos y contexto de compra.</div>', unsafe_allow_html=True)

    q = st.text_input("Buscar por SKU, descripción o MLC")
    filtered = product_df.copy()
    if q.strip():
        ql = q.lower().strip()
        filtered = filtered[filtered["search_text"].str.contains(ql, na=False)]

    if filtered.empty:
        st.warning("No encontré coincidencias.")
        st.stop()

    options = filtered.apply(lambda r: f"{r.get('SKU', '')} · {str(r.get('DESCRIPCIÓN', ''))[:90]}", axis=1).tolist()
    selected = st.selectbox("Selecciona producto", options)
    row = filtered.iloc[options.index(selected)]
    sku = row["SKU_norm"]

    decision_text, decision_cls = compute_decision_row(row)
    urgency_text, urgency_cls, _ = urgency_info(row.get("days_to_next"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Precio bruto", money(row.get("PRECIO BRUTO")))
    c2.metric("Último costo", money(row.get("ÚLTIMO COSTO")))
    c3.metric("Margen Meli 2", pct(row.get("MARGEN MELI 2")))
    with c4:
        st.markdown("**Estado**")
        render_badge(decision_text, decision_cls)
        st.markdown(" ")
        render_badge(urgency_text, urgency_cls)

    left, mid, right = st.columns([1.1, 1, 1])

    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Identidad")
        st.write(f"**SKU:** {row.get('SKU', '—')}")
        st.write(f"**Descripción:** {row.get('DESCRIPCIÓN', '—')}")
        st.write(f"**Ubicación:** {row.get('UBIC', '—')}")
        st.write(f"**MLC principal:** {row.get('MLC_norm', '—')}")
        mlcs = row.get("MLC_norm_y") if "MLC_norm_y" in row else row.get("MLC_norm")
        if isinstance(row.get("MLC_norm"), list):
            st.write("**MLCs asociadas:** " + ", ".join(row.get("MLC_norm")))
        elif isinstance(row.get("promo_mlcs"), list):
            st.write("**MLCs con promo:** " + ", ".join(row.get("promo_mlcs")))
        st.write(f"**Comentario maestra:** {row.get('COMENTARIO', '—')}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Compras")
        st.write(f"**Última compra:** {pd.to_datetime(row.get('ultima_compra')).strftime('%d-%m-%Y') if pd.notna(row.get('ultima_compra')) else '—'}")
        st.write(f"**Último precio compra:** {money(row.get('ultimo_precio_compra'))}")
        st.write(f"**Proveedor último:** {row.get('proveedor_ultimo', '—')}")
        st.write(f"**Rango histórico:** {money(row.get('precio_min_hist'))} a {money(row.get('precio_max_hist'))}")
        st.markdown("</div>", unsafe_allow_html=True)

    with mid:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Precio y rentabilidad")
        m1, m2 = st.columns(2)
        m1.metric("Precio neto", money(row.get("PRECIO NETO")))
        m2.metric("Cambio precio", money(row.get("CAMBIO DE PRECIO")))
        m1.metric("Margen local", pct(row.get("MARGEN LOCAL")))
        m2.metric("Margen Meli 1", pct(row.get("MARGEN MELI 1")))
        st.write(f"**Precio promo mínimo:** {money(row.get('min_promo_price'))}")
        st.write(f"**Precio promo máximo:** {money(row.get('max_promo_price'))}")
        st.write(f"**Ads / comentario:** {row.get('ads_comment', '—')}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Lectura automática")
        bullets = []
        if safe_int(row.get("total_promos"), 0) == 0:
            bullets.append("Producto sin promo activa registrada.")
        if pd.notna(row.get("days_to_next")) and row.get("days_to_next") <= 1:
            bullets.append("Urgente: una promo vence hoy o mañana.")
        if pd.notna(row.get("MARGEN MELI 2")) and safe_float(row.get("MARGEN MELI 2")) < 0:
            bullets.append("Margen Meli 2 negativo. Revisar precio o campaña.")
        if pd.notna(row.get("ultima_compra")) and (pd.Timestamp.today().normalize() - pd.to_datetime(row.get("ultima_compra")).normalize()).days > 180:
            bullets.append("Última compra antigua. Confirmar costo vigente.")
        if not bullets:
            bullets.append("Producto bajo control. No veo alertas críticas inmediatas.")
        for b in bullets:
            st.markdown(f"- {b}")
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        sku_promos = promos_df_view[promos_df_view["SKU_norm"] == sku].copy()
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Promos asociadas")
        st.metric("Promos asociadas", safe_int(row.get("total_promos"), 0))
        if not sku_promos.empty:
            display_cols = [c for c in ["MLC_norm", "Precio promocional", "Motivo promoción", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"] if c in sku_promos.columns]
            st.dataframe(sku_promos[display_cols], use_container_width=True, hide_index=True, height=280)
        else:
            st.info("Este SKU no tiene promos asociadas.")
        st.markdown("</div>", unsafe_allow_html=True)

    if not model["compras"].empty:
        sku_hist = model["compras"][model["compras"]["SKU_norm"] == sku].copy()
        if not sku_hist.empty and "fecha_compra" in sku_hist.columns and "precio_compra" in sku_hist.columns:
            sku_hist = sku_hist.sort_values("fecha_compra")
            chart_df = sku_hist[["fecha_compra", "precio_compra"]].dropna()
            if not chart_df.empty:
                st.subheader("Evolución precio de compra")
                st.line_chart(chart_df.set_index("fecha_compra"))

# ---------- Operator ----------
elif page == "Operador de promos":
    st.markdown('<div class="big-title">Operador de promociones</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Actualización rápida de fechas, precios promo y comentarios sin perder contexto.</div>', unsafe_allow_html=True)

    promo_work = promos_df.copy()
    promo_work["urgency_text"] = promo_work["days_to_next"].apply(lambda x: urgency_info(x)[0] if not pd.isna(x) else "Sin fecha")
    c1, c2, c3 = st.columns(3)
    only_urgent = c1.checkbox("Solo hoy/mañana")
    only_no_date = c2.checkbox("Solo sin fecha")
    search = c3.text_input("Buscar SKU / MLC / descripción", key="promo_search")

    if only_urgent:
        promo_work = promo_work[promo_work["days_to_next"].fillna(99) <= 1]
    if only_no_date:
        promo_work = promo_work[promo_work["next_campaign_date"].isna()]
    if search.strip():
        mask = (
            promo_work.get("SKU_norm", "").astype(str).str.contains(search, case=False, na=False) |
            promo_work.get("MLC_norm", "").astype(str).str.contains(search, case=False, na=False) |
            promo_work.get("Descripción", "").astype(str).str.contains(search, case=False, na=False)
        )
        promo_work = promo_work[mask]

    st.caption("Puedes editar directamente las fechas de campaña, el precio promo, motivo y comentario.")
    edit_cols = promo_editable_columns(promo_work)
    edited = st.data_editor(
        promo_work[edit_cols].copy(),
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        height=520,
        key="promo_editor",
    )

    a1, a2, a3 = st.columns([1, 1, 2])
    with a1:
        extend_days = st.number_input("Extender Campaña 1 por días", min_value=1, max_value=30, value=7)
    with a2:
        target_sku = st.text_input("SKU para extensión rápida")
    with a3:
        st.write("")
        if st.button("Aplicar extensión rápida", use_container_width=True):
            mask = st.session_state.promos_df["SKU_norm"].astype(str) == target_sku.strip()
            if mask.any() and "Campaña 1" in st.session_state.promos_df.columns:
                st.session_state.promos_df.loc[mask, "Campaña 1"] = pd.to_datetime(
                    st.session_state.promos_df.loc[mask, "Campaña 1"], errors="coerce"
                ).fillna(pd.Timestamp.today().normalize()) + pd.to_timedelta(int(extend_days), unit="D")
                st.success("Extensión aplicada.")
                st.rerun()
            else:
                st.warning("No encontré ese SKU en promociones.")

    if st.button("Guardar cambios del editor", type="primary", use_container_width=True):
        # merge back by row order of filtered subset
        base = st.session_state.promos_df.copy()
        idxs = promo_work.index.tolist()
        for col in edited.columns:
            base.loc[idxs, col] = edited[col].values
        # recompute derived cols
        for c in CAMPAIGN_COLS:
            if c in base.columns:
                base[c] = pd.to_datetime(base[c], errors="coerce")
        base["next_campaign_date"] = base[CAMPAIGN_COLS].min(axis=1)
        base["days_to_next"] = (base["next_campaign_date"].dt.normalize() - pd.Timestamp.today().normalize()).dt.days
        st.session_state.promos_df = base
        st.success("Promos actualizadas en memoria. Descarga el Excel para consolidar cambios.")
        st.rerun()

# ---------- Control Center ----------
elif page == "Centro de control":
    st.markdown('<div class="big-title">Centro de control comercial</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Prioriza lo urgente y lo económicamente riesgoso.</div>', unsafe_allow_html=True)

    control = product_df.copy()
    control["decision"], control["decision_cls"] = zip(*control.apply(compute_decision_row, axis=1))
    control["urgency_order"] = control["days_to_next"].fillna(99)
    control["risk_score"] = (
        np.where(control["days_to_next"].fillna(99) <= 1, 50, np.where(control["days_to_next"].fillna(99) <= 7, 20, 0)) +
        np.where(pd.to_numeric(control["MARGEN MELI 2"], errors="coerce").fillna(0) < 0, 40, 0) +
        np.where(control["total_promos"].fillna(0) == 0, 10, 0)
    )
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Promos hoy/mañana", int((control["days_to_next"].fillna(99) <= 1).sum()))
    k2.metric("Sin promo", int((control["total_promos"].fillna(0) == 0).sum()))
    k3.metric("Margen Meli 2 negativo", int((pd.to_numeric(control["MARGEN MELI 2"], errors="coerce").fillna(0) < 0).sum()))
    k4.metric("Con compra histórica", int(control["ultima_compra"].notna().sum()))

    left, right = st.columns([1.1, 1])

    with left:
        st.subheader("Top urgencias")
        urgent = control.sort_values(["risk_score", "days_to_next"], ascending=[False, True]).head(20).copy()
        show = urgent[["SKU", "DESCRIPCIÓN", "PRECIO BRUTO", "MARGEN MELI 2", "total_promos", "next_campaign_date", "decision"]].copy()
        show["next_campaign_date"] = pd.to_datetime(show["next_campaign_date"], errors="coerce").dt.strftime("%d-%m-%Y")
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)

    with right:
        st.subheader("Agenda rápida")
        buckets = {
            "Vence hoy": control[control["days_to_next"] == 0],
            "Mañana": control[control["days_to_next"] == 1],
            "En 2 días": control[control["days_to_next"] == 2],
            "Semana": control[(control["days_to_next"] >= 3) & (control["days_to_next"] <= 7)],
        }
        for label, dfb in buckets.items():
            st.markdown(f"**{label}** · {len(dfb)}")
            if len(dfb):
                st.caption(" / ".join(dfb["SKU"].astype(str).head(5).tolist()))

# ---------- Create product ----------
elif page == "Alta de producto":
    st.markdown('<div class="big-title">Alta de producto nuevo</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Crea SKU nuevo y, si quieres, deja promo inicial cargada desde aquí.</div>', unsafe_allow_html=True)

    with st.form("new_product"):
        c1, c2, c3 = st.columns(3)
        new_sku = c1.text_input("SKU *")
        desc = c2.text_input("Descripción")
        ubic = c3.text_input("Ubicación")
        cost = c1.number_input("Último costo", min_value=0.0, value=0.0, step=1.0)
        bruto = c2.number_input("Precio bruto", min_value=0.0, value=0.0, step=1.0)
        mlc = c3.text_input("MLC principal")
        comment = st.text_input("Comentario maestra")
        st.markdown("**Promo inicial opcional**")
        p1, p2, p3 = st.columns(3)
        promo_price = p1.number_input("Precio promocional", min_value=0.0, value=0.0, step=1.0)
        motivo = p2.text_input("Motivo")
        ads = p3.text_input("Ads / comentario")
        camp1 = st.date_input("Campaña 1", value=None)
        submitted = st.form_submit_button("Crear producto", type="primary")

    if submitted:
        sku_norm = normalize_sku(new_sku)
        if not sku_norm:
            st.error("SKU inválido.")
        elif sku_norm in set(st.session_state.master_df["SKU_norm"].dropna().astype(str)):
            st.error("Ese SKU ya existe.")
        else:
            new_master = {c: np.nan for c in st.session_state.master_df.columns}
            for k, v in {
                "SKU": new_sku,
                "DESCRIPCIÓN": desc,
                "UBIC": ubic,
                "ÚLTIMO COSTO": cost,
                "PRECIO BRUTO": bruto,
                "COMENTARIO": comment,
                "SKU_norm": sku_norm,
                "MLC_norm": normalize_mlc(mlc),
            }.items():
                if k in new_master:
                    new_master[k] = v
            st.session_state.master_df = pd.concat([st.session_state.master_df, pd.DataFrame([new_master])], ignore_index=True)

            if mlc:
                st.session_state.bridge_df = pd.concat([
                    st.session_state.bridge_df,
                    pd.DataFrame([{"SKU_norm": sku_norm, "MLC_norm": normalize_mlc(mlc)}])
                ], ignore_index=True)

            if promo_price > 0 or motivo or ads or camp1:
                promo_row = {c: np.nan for c in st.session_state.promos_df.columns}
                base_assign = {
                    "SKU_norm": sku_norm,
                    "MLC_norm": normalize_mlc(mlc),
                    "Descripción": desc,
                    "Precio promocional": promo_price if promo_price > 0 else np.nan,
                    "Motivo promoción": motivo,
                    "Ads/Comentario": ads,
                    "Campaña 1": pd.to_datetime(camp1) if camp1 else pd.NaT,
                }
                for k, v in base_assign.items():
                    if k in promo_row:
                        promo_row[k] = v
                st.session_state.promos_df = pd.concat([st.session_state.promos_df, pd.DataFrame([promo_row])], ignore_index=True)

            st.success("Producto creado en memoria. Descarga el Excel actualizado.")
            st.rerun()
