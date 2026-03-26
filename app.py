
import io
import re
import hashlib
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Aurora Pricing Cockpit",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem;}
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
.big-title{font-size:28px; font-weight:700; margin-bottom:0px;}
.subtle{color:#666;}
div[data-testid="stButton"] > button {
    white-space: pre-wrap;
    height: auto;
    min-height: 108px;
    border-radius: 18px;
    border: 1px solid #e8e8e8;
    box-shadow: 0 1px 4px rgba(0,0,0,.04);
    background: #ffffff;
    text-align: left;
    font-size: 12px;
    line-height: 1.35;
    padding: 14px;
}
div[data-testid="stButton"] > button:hover {
    border-color:#d0d7ff;
    background:#fafbff;
}
</style>
""", unsafe_allow_html=True)

PROMO_SLOTS = [
    {
        "slot_key": "promo_1",
        "label": "Promo 1",
        "mlc_col": "MLC",
        "published_col": "PRECIO B2C PUBLICADO ",
        "discount_col": "% DCTO",
        "date_col": "FECHA VENCI",
        "comment_col": "COMENTARIO",
    },
    {
        "slot_key": "promo_2",
        "label": "Promo 2",
        "mlc_col": "MLC.1",
        "published_col": "PRECIO B2C",
        "discount_col": "% DCTO.1",
        "date_col": "FECHA VENCI.1",
        "comment_col": "COMENTARIO.1",
    },
]
RELAMPAGO_COLS = ["SKU", "Descripción", "Precio promocional", "Extra", "Motivo promoción", "COMENTARIO"]


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
    if not s:
        return None
    s = s.replace(".0", "")
    digits = re.sub(r"\D", "", s)
    return digits or s


def normalize_mlc(x):
    if pd.isna(x):
        return None
    s = str(x).upper().strip()
    if not s:
        return None
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
            mlc = f"MLC{num}"
            if mlc not in out:
                out.append(mlc)
    return out


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
    if days == 3:
        return ("Vence en 3 días", "badge-blue", 3)
    if days <= 7:
        return (f"Vence en {days} días", "badge-blue", 4)
    return ("Vigente", "badge-green", 5)


def render_badge(text, cls):
    st.markdown(f'<span class="badge {cls}">{text}</span>', unsafe_allow_html=True)


def df_series(df, col, default=""):
    if col in df.columns:
        return df[col]
    return pd.Series([default] * len(df), index=df.index, dtype=object)


def first_nonempty(*values):
    for value in values:
        if isinstance(value, pd.Series):
            value = value.iloc[0] if not value.empty else None
        try:
            if pd.isna(value):
                continue
        except Exception:
            pass
        if isinstance(value, str) and not value.strip():
            continue
        if value is not None:
            return value
    return None


def file_bytes_and_sig(uploaded_file):
    if uploaded_file is None:
        return None, None
    raw = uploaded_file.getvalue()
    sig = hashlib.md5(raw).hexdigest()
    return raw, sig


@st.cache_data(show_spinner=False)
def read_excel_all_from_bytes(raw_bytes):
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheets = {}
    for sh in xls.sheet_names:
        if sh == "Relampago mi pagina":
            tmp = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sh, header=None)
            if tmp.shape[1] < len(RELAMPAGO_COLS):
                for i in range(tmp.shape[1], len(RELAMPAGO_COLS)):
                    tmp[i] = np.nan
            tmp = tmp.iloc[:, :len(RELAMPAGO_COLS)].copy()
            tmp.columns = RELAMPAGO_COLS
            sheets[sh] = tmp
        else:
            sheets[sh] = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sh)
    return sheets


def read_excel_all(uploaded_file):
    raw_bytes, _ = file_bytes_and_sig(uploaded_file)
    if raw_bytes is None:
        return {}
    return read_excel_all_from_bytes(raw_bytes)


def pick_first_existing(columns_lower, preferred_names=(), contains_terms=()):
    for name in preferred_names:
        for original, lowered in columns_lower.items():
            if lowered == name.lower():
                return original
    for term in contains_terms:
        for original, lowered in columns_lower.items():
            if term.lower() in lowered:
                return original
    return None


def normalize_desc_for_match(x):
    if pd.isna(x):
        return ""
    s = str(x).upper()
    s = re.sub(r"\[UBC:.*?\]", " ", s)
    s = s.replace("N/GENESIS", "N GENESIS")
    replacements = {
        " BCO ": " BLANCO ",
        "BCA ": "BLANCA ",
        " BCA": " BLANCA",
        " MOD. ": " MODULO ",
        " MODULO ": " ",
        " EMBUTIDO ": " EMB ",
        " EMB. ": " EMB ",
        " S/P ": " ",
        " C/ ": " ",
        "/": " ",
        "-": " ",
    }
    s = f" {s} "
    for a, b in replacements.items():
        s = s.replace(a, b)
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


STOP_TOKENS = {
    "", "DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "EN", "CON", "SIN", "PARA", "POR", "UN", "UNA",
    "UND", "UNID", "PACK", "EMB", "MODULO", "ART", "ARTICULO", "TIPO", "COLOR", "N"
}


def desc_tokens(x):
    s = normalize_desc_for_match(x)
    toks = []
    for t in s.split():
        if t in STOP_TOKENS or len(t) == 1:
            continue
        toks.append(t)
    return toks


def prepare_compras_dataframe(raw_compras: pd.DataFrame):
    compras = raw_compras.copy()
    if compras.empty:
        return compras, pd.DataFrame(), {}

    columns_lower = {c: str(c).strip().lower() for c in compras.columns}
    sku_col = pick_first_existing(columns_lower, preferred_names=["SKU", "Codigo", "Código"], contains_terms=["sku"])
    date_col = pick_first_existing(columns_lower, preferred_names=["Fecha"], contains_terms=["fecha"])
    provider_col = pick_first_existing(columns_lower, preferred_names=["Razón Social", "Proveedor", "Proveedor / Razón Social"], contains_terms=["razón social", "proveedor"])
    price_col = pick_first_existing(columns_lower, preferred_names=["Precio Un.", "Precio Unitario", "Costo Unitario"], contains_terms=["precio un", "precio", "costo unit"])
    qty_col = pick_first_existing(columns_lower, preferred_names=["Cantidad"], contains_terms=["cantidad", "cant"])
    desc_col = pick_first_existing(columns_lower, preferred_names=["Concepto / Artículo", "Descripción", "Detalle Concepto Compra"], contains_terms=["concepto", "artículo", "articulo", "descripción", "descripcion", "detalle"])
    doc_col = pick_first_existing(columns_lower, preferred_names=["Documento"], contains_terms=["documento"])
    folio_col = pick_first_existing(columns_lower, preferred_names=["Folio"], contains_terms=["folio"])
    total_col = pick_first_existing(columns_lower, preferred_names=["Total Línea", "Total"], contains_terms=["total"])

    meta = {"sku_col": sku_col, "date_col": date_col, "provider_col": provider_col, "price_col": price_col, "qty_col": qty_col, "desc_col": desc_col}
    if sku_col is None:
        return pd.DataFrame(), pd.DataFrame(), meta

    compras["SKU_norm"] = compras[sku_col].apply(normalize_sku)
    compras = compras[compras["SKU_norm"].notna()].copy()
    compras["fecha_compra"] = pd.to_datetime(compras[date_col], errors="coerce", dayfirst=True) if date_col else pd.NaT
    compras["proveedor"] = compras[provider_col].astype(str).replace("nan", "").str.strip() if provider_col else ""
    compras["precio_compra"] = pd.to_numeric(compras[price_col], errors="coerce") if price_col else np.nan
    compras["cantidad"] = pd.to_numeric(compras[qty_col], errors="coerce") if qty_col else np.nan
    compras["descripcion_compra"] = compras[desc_col].astype(str).replace("nan", "").str.strip() if desc_col else ""
    compras["documento_compra"] = compras[doc_col].astype(str).replace("nan", "").str.strip() if doc_col else ""
    compras["folio_compra"] = compras[folio_col].astype(str).replace("nan", "").str.strip() if folio_col else ""
    compras["total_linea_compra"] = pd.to_numeric(compras[total_col], errors="coerce") if total_col else np.nan

    compras = compras.sort_values(["SKU_norm", "fecha_compra"]).reset_index(drop=True)
    compras["desc_match_tokens"] = compras["descripcion_compra"].apply(desc_tokens)
    compras["precio_anterior"] = compras.groupby("SKU_norm")["precio_compra"].shift(1)
    compras["variacion_vs_anterior_pct"] = np.where(
        compras["precio_anterior"].notna() & (compras["precio_anterior"] != 0) & compras["precio_compra"].notna(),
        (compras["precio_compra"] / compras["precio_anterior"] - 1.0) * 100.0,
        np.nan,
    )

    latest = compras.dropna(subset=["fecha_compra"]).groupby("SKU_norm", as_index=False).tail(1)
    if latest.empty:
        latest = compras.groupby("SKU_norm", as_index=False).tail(1)

    summary = compras.groupby("SKU_norm").agg(
        ultima_compra=("fecha_compra", "max"),
        precio_min_hist=("precio_compra", "min"),
        precio_max_hist=("precio_compra", "max"),
        compras_registros=("SKU_norm", "count"),
        proveedores=("proveedor", lambda s: " | ".join(pd.unique([str(x) for x in s if pd.notna(x) and str(x).strip()]))),
    ).reset_index()

    latest_fields = latest[["SKU_norm", "fecha_compra", "precio_compra", "proveedor", "cantidad", "variacion_vs_anterior_pct"]].rename(columns={
        "fecha_compra": "ultima_compra_registro",
        "precio_compra": "ultimo_precio_compra",
        "proveedor": "proveedor_ultimo",
        "cantidad": "cantidad_ultima_compra",
        "variacion_vs_anterior_pct": "variacion_ultima_vs_anterior_pct",
    })
    summary = summary.merge(latest_fields, on="SKU_norm", how="left")
    summary["ultima_compra"] = summary["ultima_compra"].combine_first(summary["ultima_compra_registro"])
    summary = summary.drop(columns=["ultima_compra_registro"], errors="ignore")
    return compras, summary, meta


def compras_candidates_for_row(row, compras_df: pd.DataFrame):
    if compras_df is None or compras_df.empty:
        return pd.DataFrame(), "sin_archivo", 0.0

    sku_norm = row.get("SKU_norm")
    exact = compras_df[compras_df["SKU_norm"] == sku_norm].copy()
    if not exact.empty:
        exact["match_score"] = 1.0
        exact["match_method"] = "SKU exacto"
        return exact.sort_values(["fecha_compra", "precio_compra"]), "SKU exacto", 1.0

    target_tokens = set(desc_tokens(row.get("DESCRIPCIÓN", "")))
    if not target_tokens:
        return pd.DataFrame(), "sin_match", 0.0

    tmp = compras_df.copy()
    token_sets = tmp["desc_match_tokens"].apply(set)
    inter = token_sets.apply(lambda s: len(s & target_tokens))
    union = token_sets.apply(lambda s: len(s | target_tokens) if (s | target_tokens) else 0)
    coverage = token_sets.apply(lambda s: len(s & target_tokens) / max(len(target_tokens), 1))
    jaccard = pd.Series(np.where(union > 0, inter / union, 0.0), index=tmp.index)
    score = 0.65 * coverage + 0.35 * jaccard
    tmp["match_score"] = score
    tmp["shared_tokens"] = inter
    cand = tmp[(tmp["shared_tokens"] >= 3) & (tmp["match_score"] >= 0.52)].copy()
    if cand.empty:
        return pd.DataFrame(), "sin_match", 0.0

    cand["match_method"] = "Descripción aproximada"
    best_skus = cand.groupby("SKU_norm")["match_score"].max().sort_values(ascending=False).head(3).index.tolist()
    cand = cand[cand["SKU_norm"].isin(best_skus)].copy()
    return cand.sort_values(["match_score", "fecha_compra"], ascending=[False, True]), "Descripción aproximada", float(cand["match_score"].max())


def compras_summary_from_rows(rows: pd.DataFrame):
    if rows is None or rows.empty:
        return {}
    rows = rows.sort_values(["fecha_compra", "precio_compra"]).copy()
    latest = rows.dropna(subset=["fecha_compra"]).tail(1)
    if latest.empty:
        latest = rows.tail(1)
    latest = latest.iloc[0]
    precios_validos = rows["precio_compra"].dropna()
    return {
        "ultima_compra": latest.get("fecha_compra"),
        "ultimo_precio_compra": latest.get("precio_compra"),
        "proveedor_ultimo": latest.get("proveedor"),
        "cantidad_ultima_compra": latest.get("cantidad"),
        "variacion_ultima_vs_anterior_pct": latest.get("variacion_vs_anterior_pct"),
        "precio_min_hist": precios_validos.min() if not precios_validos.empty else np.nan,
        "precio_max_hist": precios_validos.max() if not precios_validos.empty else np.nan,
        "proveedores": " | ".join(pd.unique([str(x) for x in rows["proveedor"] if pd.notna(x) and str(x).strip()])),
        "compras_registros": len(rows),
    }


def prep_relampago(df: pd.DataFrame):
    rel = df.copy()
    if rel.empty:
        for c in RELAMPAGO_COLS:
            if c not in rel.columns:
                rel[c] = np.nan
    rel = rel.iloc[:, :len(RELAMPAGO_COLS)].copy()
    rel.columns = RELAMPAGO_COLS
    rel["SKU_norm"] = rel["SKU"].apply(normalize_sku)
    rel["Precio promocional"] = pd.to_numeric(rel["Precio promocional"], errors="coerce")
    rel["COMENTARIO"] = rel["COMENTARIO"].astype(str).replace("nan", "").str.strip()
    rel = rel[rel["SKU_norm"].notna()].copy()
    rel["source"] = "Relámpago mi página"
    return rel


def build_master_promo_cards(master: pd.DataFrame):
    cards = []
    if master.empty:
        return pd.DataFrame()
    for idx, row in master.iterrows():
        sku_norm = normalize_sku(row.get("SKU"))
        desc = row.get("DESCRIPCIÓN")
        for slot in PROMO_SLOTS:
            mlcs = extract_mlcs(row.get(slot["mlc_col"])) if slot["mlc_col"] in master.columns else []
            published = safe_float(row.get(slot["published_col"]))
            discount = safe_float(row.get(slot["discount_col"]))
            promo_price = np.nan
            if pd.notna(published) and pd.notna(discount):
                promo_price = published * (1 - discount)
            fe = pd.to_datetime(row.get(slot["date_col"]), errors="coerce") if slot["date_col"] in master.columns else pd.NaT
            comment = first_nonempty(row.get(slot["comment_col"]), "")
            should_create = bool(mlcs) or pd.notna(fe) or pd.notna(promo_price) or (isinstance(comment, str) and comment.strip())
            if not should_create:
                continue
            for mlc in mlcs or [None]:
                cards.append({
                    "source_kind": "maestra",
                    "slot_key": slot["slot_key"],
                    "slot_label": slot["label"],
                    "source_index": idx,
                    "SKU_norm": sku_norm,
                    "Descripción": desc,
                    "MLC_norm": mlc,
                    "published_price": published,
                    "Precio promocional": promo_price,
                    "FECHA VENCI": fe,
                    "COMENTARIO": comment if comment is not None else "",
                })
    cards_df = pd.DataFrame(cards)
    if cards_df.empty:
        return pd.DataFrame(columns=["source_kind", "slot_key", "source_index", "SKU_norm", "Descripción", "MLC_norm", "published_price", "Precio promocional", "FECHA VENCI", "COMENTARIO", "days_to_next", "urgency_text", "urgency_cls", "urgency_rank"])
    cards_df["days_to_next"] = (pd.to_datetime(cards_df["FECHA VENCI"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    urg = cards_df["days_to_next"].apply(urgency_info)
    cards_df["urgency_text"] = urg.apply(lambda x: x[0])
    cards_df["urgency_cls"] = urg.apply(lambda x: x[1])
    cards_df["urgency_rank"] = urg.apply(lambda x: x[2])
    return cards_df


def aggregate_product(master: pd.DataFrame, bridge: pd.DataFrame, promo_cards: pd.DataFrame, relampago: pd.DataFrame):
    product = master.copy()
    if "Unnamed: 12" in product.columns and "MLC_aux" not in product.columns:
        product = product.rename(columns={"Unnamed: 12": "MLC_aux"})
    product["SKU_norm"] = product.get("SKU", pd.Series(dtype=object)).apply(normalize_sku)

    bridge = bridge.copy()
    bridge["SKU_norm"] = bridge.iloc[:, 0].apply(normalize_sku) if len(bridge.columns) > 0 else pd.Series(dtype=object)
    bridge["MLC_norm"] = bridge.iloc[:, 1].apply(normalize_mlc) if len(bridge.columns) > 1 else pd.Series(dtype=object)
    bridge = bridge[["SKU_norm", "MLC_norm"]].dropna().drop_duplicates()

    master_mlcs = []
    for col in ["MLC", "MLC.1", "MLC_aux"]:
        if col in product.columns:
            tmp = product[["SKU_norm", col]].copy()
            tmp["MLC_norm"] = tmp[col].apply(extract_mlcs)
            tmp = tmp.explode("MLC_norm")
            tmp = tmp[["SKU_norm", "MLC_norm"]].dropna()
            master_mlcs.append(tmp)
    sku_mlc = pd.concat(master_mlcs + [bridge], ignore_index=True) if master_mlcs else bridge.copy()
    sku_mlc = sku_mlc.drop_duplicates()

    bridge_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index() if not sku_mlc.empty else pd.DataFrame(columns=["SKU_norm", "MLC_norm"])

    if not promo_cards.empty:
        promo_agg = promo_cards.groupby("SKU_norm").agg(
            total_promos=("SKU_norm", "count"),
            promo_mlcs=("MLC_norm", lambda s: list(pd.unique([x for x in s if pd.notna(x)]))),
            next_campaign_date=("FECHA VENCI", "min"),
            min_promo_price=("Precio promocional", "min"),
            max_promo_price=("Precio promocional", "max"),
            promo_comments=("COMENTARIO", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        ).reset_index()
    else:
        promo_agg = pd.DataFrame(columns=["SKU_norm", "total_promos", "promo_mlcs", "next_campaign_date", "min_promo_price", "max_promo_price", "promo_comments"])

    if not relampago.empty:
        rel_agg = relampago.groupby("SKU_norm").agg(
            relampago_count=("SKU_norm", "count"),
            relampago_min_price=("Precio promocional", "min"),
            relampago_max_price=("Precio promocional", "max"),
            relampago_comment=("COMENTARIO", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        ).reset_index()
    else:
        rel_agg = pd.DataFrame(columns=["SKU_norm", "relampago_count", "relampago_min_price", "relampago_max_price", "relampago_comment"])

    product = product.merge(bridge_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product = product.merge(rel_agg, on="SKU_norm", how="left")
    product["total_promos"] = product["total_promos"].fillna(0)
    product["relampago_count"] = product["relampago_count"].fillna(0)
    product["days_to_next"] = (pd.to_datetime(product["next_campaign_date"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    product["search_text"] = (
        df_series(product, "SKU").astype(str).fillna("") + " | " +
        df_series(product, "DESCRIPCIÓN").astype(str).fillna("") + " | " +
        df_series(product, "MLC_norm").astype(str).fillna("") + " | " +
        df_series(product, "promo_mlcs").astype(str).fillna("")
    ).str.lower()
    return product, bridge


def compute_decision_row(row):
    margin = safe_float(row.get("MARGEN MELI 1"))
    days = row.get("days_to_next")
    total_promos = safe_int(row.get("total_promos"), 0)
    relampago_count = safe_int(row.get("relampago_count"), 0)
    if total_promos == 0 and relampago_count == 0:
        return "Sin promo", "badge-gray"
    if not pd.isna(days) and days <= 1:
        return "Urgente renovar", "badge-red"
    if not pd.isna(margin) and margin < 0:
        return "Margen negativo", "badge-red"
    if relampago_count > 0:
        return "Tiene relámpago", "badge-blue"
    if not pd.isna(days) and days <= 7:
        return "Revisar esta semana", "badge-yellow"
    return "Bajo control", "badge-green"


@st.cache_data(show_spinner=False)
def cached_prepare_compras_from_bytes(compras_bytes):
    if compras_bytes is None:
        return pd.DataFrame(), pd.DataFrame()
    try:
        c_sheets = read_excel_all_from_bytes(compras_bytes)
        best_name = max(c_sheets, key=lambda k: len(c_sheets[k]))
        compras, compras_summary, _ = prepare_compras_dataframe(c_sheets[best_name])
        return compras, compras_summary
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


@st.cache_data(show_spinner=False)
def cached_rebuild_from_frames(master_df, bridge_df, relampago_df, compras_bytes=None):
    relampago = prep_relampago(relampago_df.copy())
    promo_cards = build_master_promo_cards(master_df)
    product, bridge = aggregate_product(master_df, bridge_df, promo_cards, relampago)
    compras = pd.DataFrame()
    if compras_bytes is not None:
        compras, compras_summary = cached_prepare_compras_from_bytes(compras_bytes)
        if not compras_summary.empty:
            product = product.merge(compras_summary, on="SKU_norm", how="left")
    return product, promo_cards, relampago, compras


@st.cache_data(show_spinner=False)
def build_model(master_bytes, compras_bytes=None):
    sheets = read_excel_all_from_bytes(master_bytes)
    master = sheets.get("MAESTRA de precios", pd.DataFrame()).copy()
    bridge = sheets.get("MLC -SKU", pd.DataFrame()).copy()
    relampago_sheet = sheets.get("Relampago mi pagina") if "Relampago mi pagina" in sheets else sheets.get("Relámpago mi página", pd.DataFrame())
    relampago = prep_relampago(relampago_sheet.copy())
    promo_cards = build_master_promo_cards(master)
    product, bridge = aggregate_product(master, bridge, promo_cards, relampago)

    compras, compras_summary = cached_prepare_compras_from_bytes(compras_bytes)
    if not compras_summary.empty:
        product = product.merge(compras_summary, on="SKU_norm", how="left")

    return {
        "sheets": sheets,
        "master": master,
        "bridge": bridge,
        "relampago": relampago,
        "promo_cards": promo_cards,
        "product": product,
        "compras": compras,
    }


def rebuild_from_session(compras_bytes, master_df, bridge_df, relampago_df):
    return cached_rebuild_from_frames(master_df, bridge_df, relampago_df, compras_bytes)


@st.cache_data(show_spinner=False)
def make_download_workbook(all_sheets, master_df, bridge_df, relampago_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            if name == "MAESTRA de precios":
                master_df.to_excel(writer, sheet_name=name, index=False)
            elif name == "MLC -SKU":
                bridge_df.to_excel(writer, sheet_name=name, index=False)
            elif name == "Relampago mi pagina":
                base = relampago_df.copy()
                base = base.iloc[:, :len(RELAMPAGO_COLS)].copy()
                base.columns = RELAMPAGO_COLS
                base.to_excel(writer, sheet_name=name, index=False, header=False)
            else:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    return out.getvalue()


def display_date(v):
    dt = pd.to_datetime(v, errors="coerce")
    return dt.strftime("%d/%m/%Y") if pd.notna(dt) else "—"


def promo_price_from_row(row, slot):
    published = safe_float(row.get(slot["published_col"]))
    dcto = safe_float(row.get(slot["discount_col"]))
    if pd.notna(published) and pd.notna(dcto):
        return published * (1 - dcto)
    return np.nan


def update_master_promo(master_df, source_index, slot_key, new_price, new_date, new_comment):
    slot = next(s for s in PROMO_SLOTS if s["slot_key"] == slot_key)
    row = master_df.loc[source_index].copy()
    published = safe_float(row.get(slot["published_col"]))
    if pd.notna(published) and published > 0 and pd.notna(new_price) and new_price > 0:
        master_df.loc[source_index, slot["discount_col"]] = 1 - (float(new_price) / published)
    else:
        master_df.loc[source_index, slot["discount_col"]] = np.nan
    master_df.loc[source_index, slot["date_col"]] = pd.to_datetime(new_date) if new_date else pd.NaT
    master_df.loc[source_index, slot["comment_col"]] = new_comment if str(new_comment).strip() else np.nan
    return master_df


def dialog_save_and_close():
    st.session_state.dialog_card = None
    st.rerun()


@st.dialog("Editar promo")
def promo_dialog(card):
    master_df = st.session_state.master_df.copy()
    source_index = int(card["source_index"])
    slot_key = card["slot_key"]
    slot = next(s for s in PROMO_SLOTS if s["slot_key"] == slot_key)
    row = master_df.loc[source_index]
    current_price = promo_price_from_row(row, slot)
    current_date = pd.to_datetime(row.get(slot["date_col"]), errors="coerce")
    current_comment = first_nonempty(row.get(slot["comment_col"]), "")

    st.write(f"**SKU:** {card.get('SKU_norm')}")
    st.write(f"**Descripción:** {card.get('Descripción')}")
    if card.get("MLC_norm"):
        st.write(f"**MLC:** {card.get('MLC_norm')}")
    st.caption(f"Precio B2C publicado: {money(row.get(slot['published_col']))}")
    st.markdown("#### Fecha de vencimiento")

    with st.form(f"promo_edit_{source_index}_{slot_key}"):
        new_date = st.date_input(
            "Fecha",
            value=current_date.date() if pd.notna(current_date) else date.today(),
            format="DD/MM/YYYY",
            help="Este es el cambio principal de la promo.",
        )
        with st.expander("Precio promocional y comentario", expanded=False):
            c1, c2 = st.columns([1, 1])
            new_price = c1.number_input("Precio promocional", min_value=0.0, value=float(safe_float(current_price, 0.0)), step=1.0)
            new_comment = c2.text_input("Comentario", value=str(current_comment))
        save = st.form_submit_button("Guardar cambios", type="primary", use_container_width=True)
    if save:
        st.session_state.master_df = update_master_promo(master_df, source_index, slot_key, new_price if new_price > 0 else np.nan, new_date, new_comment)
        dialog_save_and_close()

    if st.button("Cerrar", use_container_width=True):
        dialog_save_and_close()


st.sidebar.title("Aurora Pricing Cockpit")
master_file = st.sidebar.file_uploader("Maestra integrada", type=["xlsx"], key="master")
compras_file = st.sidebar.file_uploader("Compras históricas", type=["xlsx"], key="compras")
st.sidebar.caption("Carga manual para trabajar siempre con el último archivo.")

if master_file is None:
    st.info("Carga la maestra integrada para empezar.")
    st.stop()

master_bytes, master_sig = file_bytes_and_sig(master_file)
compras_bytes, compras_sig = file_bytes_and_sig(compras_file)
model = build_model(master_bytes, compras_bytes)
all_sheets = model["sheets"]

state_needs_refresh = (st.session_state.get("master_sig") != master_sig) or (st.session_state.get("compras_sig") != compras_sig)
if state_needs_refresh:
    st.session_state.master_sig = master_sig
    st.session_state.compras_sig = compras_sig
    st.session_state.source_name = master_file.name

if state_needs_refresh or "master_df" not in st.session_state:
    st.session_state.master_df = model["master"].copy()
if state_needs_refresh or "bridge_df" not in st.session_state:
    st.session_state.bridge_df = model["bridge"].copy()
if state_needs_refresh or "relampago_df" not in st.session_state:
    st.session_state.relampago_df = model["relampago"].copy()
if "dialog_card" not in st.session_state:
    st.session_state.dialog_card = None

product_df, promo_cards_df, relampago_view, compras_df = rebuild_from_session(
    compras_bytes, st.session_state.master_df, st.session_state.bridge_df, st.session_state.relampago_df
)

download_bytes = make_download_workbook(all_sheets, st.session_state.master_df, st.session_state.bridge_df, st.session_state.relampago_df)
st.sidebar.download_button(
    "Descargar Excel actualizado",
    data=download_bytes,
    file_name=f"MAESTRA_ACTUALIZADA_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

page = st.sidebar.radio("Módulos", ["Cockpit por producto", "Operador de promos", "Alta de producto"])

if page == "Cockpit por producto":
    st.markdown('<div class="big-title">Ficha de decisión por producto</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Vista limpia para decidir precio, promo, relámpago y contexto de compra.</div>', unsafe_allow_html=True)

    q = st.text_input("Buscar por SKU, descripción o MLC")
    filtered = product_df.copy()
    if q.strip():
        filtered = filtered[filtered["search_text"].str.contains(q.lower().strip(), na=False)]

    if filtered.empty:
        st.warning("No encontré coincidencias.")
        st.stop()

    options = filtered.apply(lambda r: f"{normalize_sku(r.get('SKU'))} · {str(r.get('DESCRIPCIÓN', ''))[:100]}", axis=1).tolist()
    selected = st.selectbox("Selecciona producto", options)
    row = filtered.iloc[options.index(selected)]
    sku = row["SKU_norm"]

    decision_text, decision_cls = compute_decision_row(row)
    urgency_text, urgency_cls, _ = urgency_info(row.get("days_to_next"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Precio bruto", money(row.get("PRECIO BRUTO")))
    c2.metric("Último costo", money(row.get("ÚLTIMO COSTO")))
    c3.metric("Margen Meli 1", pct(row.get("MARGEN MELI 1")))
    c4.metric("Promos asociadas", safe_int(row.get("total_promos"), 0))

    b1, b2 = st.columns([1, 1])
    with b1:
        render_badge(decision_text, decision_cls)
    with b2:
        render_badge(urgency_text, urgency_cls)

    left, mid, right = st.columns([1.15, 1, 1])

    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Identidad")
        st.write(f"**SKU:** {normalize_sku(row.get('SKU')) or '—'}")
        st.write(f"**Descripción:** {row.get('DESCRIPCIÓN', '—')}")
        st.write(f"**Ubicación:** {row.get('UBIC', '—')}")
        all_mlcs = row.get("MLC_norm")
        if isinstance(all_mlcs, list) and all_mlcs:
            st.write(f"**MLCs asociados:** {', '.join(all_mlcs)}")
        elif pd.notna(all_mlcs):
            st.write(f"**MLCs asociados:** {all_mlcs}")
        else:
            st.write("**MLCs asociados:** —")
        st.write(f"**Comentario maestra:** {first_nonempty(row.get('COMENTARIO'), row.get('COMENTARIO.1'), '—')}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Compras")
        compras_rows, compras_match_method, compras_match_score = compras_candidates_for_row(row, compras_df)
        compras_info = compras_summary_from_rows(compras_rows)
        if compras_rows.empty:
            st.write("**Última compra:** —")
            st.write("**Último precio compra:** —")
            st.write("**Proveedor último:** —")
            st.write("**Cantidad última compra:** —")
            st.write("**Variación vs compra anterior:** —")
            st.write("**Rango histórico:** — a —")
        else:
            render_badge("Compras por SKU exacto" if compras_match_method == "SKU exacto" else f"Compras por descripción ({compras_match_score:.2f})", "badge-green" if compras_match_method == "SKU exacto" else "badge-yellow")
            st.write(f"**Última compra:** {display_date(compras_info.get('ultima_compra'))}")
            st.write(f"**Último precio compra:** {money(compras_info.get('ultimo_precio_compra'))}")
            st.write(f"**Proveedor último:** {compras_info.get('proveedor_ultimo', '—')}")
            st.write(f"**Cantidad última compra:** {safe_int(compras_info.get('cantidad_ultima_compra'), 0) if pd.notna(compras_info.get('cantidad_ultima_compra')) else '—'}")
            st.write(f"**Variación vs compra anterior:** {pct((compras_info.get('variacion_ultima_vs_anterior_pct') / 100.0) if pd.notna(compras_info.get('variacion_ultima_vs_anterior_pct')) else np.nan)}")
            st.write(f"**Rango histórico:** {money(compras_info.get('precio_min_hist'))} a {money(compras_info.get('precio_max_hist'))}")
            if compras_info.get("proveedores"):
                st.write(f"**Proveedores históricos:** {compras_info.get('proveedores')}")
        st.markdown("</div>", unsafe_allow_html=True)

    with mid:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Precio y rentabilidad")
        m1, m2 = st.columns(2)
        m1.metric("Precio neto", money(row.get("PRECIO NETO")))
        m2.metric("Cambio precio", money(row.get("CAMBIO DE PRECIO")))
        m1.metric("Margen local", pct(row.get("MARGEN LOCAL")))
        m2.metric("Monto en simulación", money(row.get("MONTO EN SIMULACIÓN")))
        st.write(f"**Neto Meli 1:** {money(row.get(' NETO MELI 1'))}")
        st.write(f"**Precio promo mínimo:** {money(row.get('min_promo_price'))}")
        st.write(f"**Precio promo máximo:** {money(row.get('max_promo_price'))}")
        st.write(f"**Precio relámpago mínimo:** {money(row.get('relampago_min_price'))}")
        st.write(f"**Precio relámpago máximo:** {money(row.get('relampago_max_price'))}")
        st.write(f"**Comentario promo:** {row.get('promo_comments') or '—'}")
        st.write(f"**Comentario relámpago:** {row.get('relampago_comment') or '—'}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Lectura automática")
        bullets = []
        if safe_int(row.get("total_promos"), 0) == 0 and safe_int(row.get("relampago_count"), 0) == 0:
            bullets.append("Producto sin promo activa registrada.")
        if pd.notna(row.get("days_to_next")) and row.get("days_to_next") <= 1:
            bullets.append("Urgente: una promo vence hoy o mañana.")
        if pd.notna(row.get("MARGEN MELI 1")) and safe_float(row.get("MARGEN MELI 1")) < 0:
            bullets.append("Margen Meli 1 negativo. Revisar precio o descuento.")
        if safe_int(row.get("relampago_count"), 0) > 0:
            bullets.append("Producto presente en relámpago mi página.")
        if pd.notna(row.get("ultima_compra")) and (pd.Timestamp.today().normalize() - pd.to_datetime(row.get("ultima_compra")).normalize()).days > 180:
            bullets.append("Última compra antigua. Confirmar costo vigente.")
        if not bullets:
            bullets.append("Producto bajo control. No veo alertas críticas inmediatas.")
        for b in bullets:
            st.markdown(f"- {b}")
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        sku_promos = promo_cards_df[promo_cards_df["SKU_norm"] == sku].copy()
        sku_rel = relampago_view[relampago_view["SKU_norm"] == sku].copy()

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Promos maestra")
        st.metric("Filas promo", len(sku_promos))
        if not sku_promos.empty:
            display_cols = [c for c in ["slot_label", "MLC_norm", "published_price", "Precio promocional", "FECHA VENCI", "COMENTARIO"] if c in sku_promos.columns]
            tmp = sku_promos[display_cols].copy()
            tmp = tmp.rename(columns={"slot_label": "Bloque", "published_price": "B2C publicado"})
            st.dataframe(tmp, use_container_width=True, hide_index=True, height=220)
        else:
            st.info("Sin promos en maestra.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Relámpago mi página")
        st.metric("Filas relámpago", len(sku_rel))
        if not sku_rel.empty:
            display_cols = [c for c in ["Descripción", "Precio promocional", "COMENTARIO"] if c in sku_rel.columns]
            st.dataframe(sku_rel[display_cols], use_container_width=True, hide_index=True, height=180)
        else:
            st.info("No está en relámpago mi página.")
        st.markdown("</div>", unsafe_allow_html=True)

elif page == "Operador de promos":
    st.markdown('<div class="big-title">Operador de promos</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Bandeja por tarjeta para editar desde la maestra actual. Cada tarjeta abre un popup.</div>', unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Bandeja de promos", "Relámpago mi página"])

    with tab1:
        cards = promo_cards_df.copy()
        cards["search_text"] = (
            cards["SKU_norm"].astype(str).fillna("") + " | " +
            cards["Descripción"].astype(str).fillna("") + " | " +
            cards["MLC_norm"].astype(str).fillna("")
        ).str.lower()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Vencen hoy", int((cards["days_to_next"] == 0).sum()))
        k2.metric("Vencen mañana", int((cards["days_to_next"] == 1).sum()))
        k3.metric("Vencen en 3 días", int((cards["days_to_next"] == 3).sum()))
        k4.metric("Sin precio promo", int(cards["Precio promocional"].isna().sum()))

        f1, f2, f3 = st.columns([1, 1, 2])
        prioridad = f1.selectbox("Prioridad", ["Todas", "Hoy", "Mañana", "En 3 días", "Sin precio promocional", "Vigentes"], key="prio_cards")
        q = f3.text_input("Buscar SKU / MLC / descripción", key="search_cards")
        filtered = cards.copy()
        if prioridad == "Hoy":
            filtered = filtered[filtered["days_to_next"] == 0]
        elif prioridad == "Mañana":
            filtered = filtered[filtered["days_to_next"] == 1]
        elif prioridad == "En 3 días":
            filtered = filtered[filtered["days_to_next"] == 3]
        elif prioridad == "Sin precio promocional":
            filtered = filtered[filtered["Precio promocional"].isna()]
        elif prioridad == "Vigentes":
            filtered = filtered[(filtered["days_to_next"].notna()) & (filtered["days_to_next"] >= 0)]
        if q.strip():
            filtered = filtered[filtered["search_text"].str.contains(q.lower().strip(), na=False)]

        s1, s2, s3 = st.columns(3)
        s1.metric("Tarjetas filtradas", len(filtered))
        s2.metric("Con fecha crítica", int(filtered["days_to_next"].isin([0, 1, 2, 3]).sum()))
        s3.metric("Bloques afectados", int(filtered[["source_index", "slot_key"]].drop_duplicates().shape[0]))

        bulk_date = st.date_input("Nueva fecha para todas las promos filtradas", value=None, format="DD/MM/YYYY", key="bulk_date_maestra")
        if st.button("Aplicar fecha masiva a filtradas", use_container_width=True):
            if bulk_date:
                master = st.session_state.master_df.copy()
                for _, r in filtered[["source_index", "slot_key"]].drop_duplicates().iterrows():
                    master = update_master_promo(master, int(r["source_index"]), r["slot_key"], promo_price_from_row(master.loc[int(r["source_index"])], next(s for s in PROMO_SLOTS if s["slot_key"] == r["slot_key"])), bulk_date, first_nonempty(master.loc[int(r["source_index"])].get(next(s for s in PROMO_SLOTS if s["slot_key"] == r["slot_key"])["comment_col"]), ""))
                st.session_state.master_df = master
                st.success(f"Fecha aplicada a {len(filtered[['source_index', 'slot_key']].drop_duplicates())} bloques de promo.")
                st.rerun()
            else:
                st.warning("Elige una fecha antes de aplicar.")

        if filtered.empty:
            st.info("No hay promos que coincidan con los filtros.")
        else:
            cols = st.columns(4)
            preview = filtered.sort_values(["urgency_rank", "FECHA VENCI", "SKU_norm"], na_position="last")
            for n, (_, r) in enumerate(preview.iterrows()):
                fecha_txt = display_date(r.get("FECHA VENCI"))
                label = f"{r.get('SKU_norm')}\n{str(r.get('Descripción') or '—')[:58]}\n{r.get('MLC_norm') or '—'}\n{fecha_txt}"
                with cols[n % 4]:
                    render_badge(r.get("urgency_text", "Sin fecha"), r.get("urgency_cls", "badge-gray"))
                    if st.button(label, key=f"card_{int(r['source_index'])}_{r['slot_key']}_{n}", use_container_width=True):
                        st.session_state.dialog_card = {
                            "source_index": int(r["source_index"]),
                            "slot_key": r["slot_key"],
                            "SKU_norm": r.get("SKU_norm"),
                            "Descripción": r.get("Descripción"),
                            "MLC_norm": r.get("MLC_norm"),
                        }
                        st.rerun()

        if st.session_state.dialog_card is not None:
            promo_dialog(st.session_state.dialog_card)

    with tab2:
        st.caption("Lista simple para agregar, sacar o modificar relámpago mi página.")
        rel_cols = [c for c in ["SKU", "Descripción", "Precio promocional", "COMENTARIO"] if c in st.session_state.relampago_df.columns]
        rel_edit = st.data_editor(
            st.session_state.relampago_df[rel_cols].copy(),
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=520,
            key="rel_editor_simple",
        )
        if st.button("Guardar lista relámpago", type="primary", use_container_width=True):
            base = st.session_state.relampago_df.copy()
            for c in rel_cols:
                base[c] = rel_edit[c]
            st.session_state.relampago_df = base
            st.success("Lista relámpago actualizada.")
            st.rerun()

elif page == "Alta de producto":
    st.markdown('<div class="big-title">Alta de producto nuevo</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Crea un SKU nuevo directamente sobre la maestra actual.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    sku = c1.text_input("SKU", key="alta_sku")
    desc = c2.text_input("Descripción", key="alta_desc")
    ubic = c3.text_input("Ubicación", key="alta_ubic")
    costo = c1.number_input("Último costo", min_value=0.0, value=0.0, step=1.0, key="alta_costo")
    bruto_tienda = c2.number_input("Precio bruto en tienda", min_value=0.0, value=0.0, step=1.0, key="alta_bruto")
    monto_sim = c3.number_input("MONTO EN SIMULACIÓN", min_value=0.0, value=0.0, step=1.0, key="alta_monto")

    st.markdown("#### Promo base")
    p1, p2, p3 = st.columns(3)
    mlc = p1.text_input("MLC", key="alta_mlc")
    b2c = p2.number_input("Precio B2C publicado", min_value=0.0, value=0.0, step=1.0, key="alta_b2c")
    promo_date = p3.date_input("Fecha vencimiento", value=None, format="DD/MM/YYYY", key="alta_fecha")
    promo_comment = p1.text_input("Comentario", key="alta_comentario")
    add_relampago = p2.checkbox("Agregar a relámpago", key="alta_relampago")

    precio_neto = bruto_tienda / 1.19 if bruto_tienda else np.nan
    margen_local = ((precio_neto - costo) / precio_neto) if pd.notna(precio_neto) and precio_neto != 0 else np.nan
    neto_meli_1 = monto_sim / 1.19 if monto_sim else np.nan
    margen_meli_1 = ((neto_meli_1 - costo) / neto_meli_1) if pd.notna(neto_meli_1) and neto_meli_1 != 0 else np.nan

    s1, s2 = st.columns(2)
    s1.metric("Margen local proyectado", pct(margen_local))
    s2.metric("Margen Meli 1 proyectado", pct(margen_meli_1))

    create = st.button("Crear producto", type="primary", use_container_width=True)

    if create:
        if not normalize_sku(sku):
            st.error("Debes ingresar un SKU válido.")
        else:
            master = st.session_state.master_df.copy()
            new_row = {c: np.nan for c in master.columns}
            new_row["SKU"] = normalize_sku(sku)
            new_row["DESCRIPCIÓN"] = desc
            new_row["UBIC"] = ubic
            new_row["ÚLTIMO COSTO"] = costo
            new_row["PRECIO BRUTO"] = bruto_tienda
            new_row["PRECIO NETO"] = precio_neto
            new_row["MARGEN LOCAL"] = margen_local
            new_row["MONTO EN SIMULACIÓN"] = monto_sim
            new_row[" NETO MELI 1"] = neto_meli_1
            new_row["MARGEN MELI 1"] = margen_meli_1
            if mlc:
                new_row["MLC"] = mlc
            if b2c > 0:
                new_row["PRECIO B2C PUBLICADO "] = b2c
            if promo_date:
                new_row["FECHA VENCI"] = pd.to_datetime(promo_date)
            if promo_comment.strip():
                new_row["COMENTARIO"] = promo_comment
            master = pd.concat([master, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state.master_df = master

            if mlc:
                bridge = st.session_state.bridge_df.copy()
                bridge = pd.concat([bridge, pd.DataFrame([{"SKU_norm": normalize_sku(sku), "MLC_norm": normalize_mlc(mlc)}])], ignore_index=True)
                st.session_state.bridge_df = bridge

            if add_relampago:
                rel = st.session_state.relampago_df.copy()
                rel_row = {c: np.nan for c in RELAMPAGO_COLS}
                rel_row["SKU"] = normalize_sku(sku)
                rel_row["Descripción"] = desc
                rel_row["Precio promocional"] = b2c if b2c > 0 else np.nan
                rel_row["COMENTARIO"] = promo_comment if promo_comment.strip() else np.nan
                rel = pd.concat([rel, pd.DataFrame([rel_row])], ignore_index=True)
                st.session_state.relampago_df = rel

            st.success("Producto creado. Descarga el Excel actualizado desde la barra lateral.")
            st.rerun()
