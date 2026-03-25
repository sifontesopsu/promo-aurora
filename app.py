
import io
import re
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
</style>
""", unsafe_allow_html=True)

CAMPAIGN_COLS = ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]
RELAMPAGO_COLS = ["SKU", "Descripción", "Precio promocional", "Extra", "Motivo promoción", "Ads/Comentario"]


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


def read_excel_all(uploaded_file):
    raw_bytes = uploaded_file.getvalue()
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
    provider_col = pick_first_existing(
        columns_lower,
        preferred_names=["Razón Social", "Proveedor", "Proveedor / Razón Social"],
        contains_terms=["razón social", "proveedor"],
    )
    price_col = pick_first_existing(
        columns_lower,
        preferred_names=["Precio Un.", "Precio Unitario", "Costo Unitario"],
        contains_terms=["precio un", "precio", "costo unit"],
    )
    qty_col = pick_first_existing(columns_lower, preferred_names=["Cantidad"], contains_terms=["cantidad", "cant"])
    desc_col = pick_first_existing(
        columns_lower,
        preferred_names=["Concepto / Artículo", "Descripción", "Detalle Concepto Compra"],
        contains_terms=["concepto", "artículo", "articulo", "descripción", "descripcion", "detalle"],
    )
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

    latest_fields = latest[["SKU_norm", "fecha_compra", "precio_compra", "proveedor", "cantidad", "variacion_vs_anterior_pct", "descripcion_compra", "documento_compra", "folio_compra"]].rename(columns={
        "fecha_compra": "ultima_compra_registro",
        "precio_compra": "ultimo_precio_compra",
        "proveedor": "proveedor_ultimo",
        "cantidad": "cantidad_ultima_compra",
        "variacion_vs_anterior_pct": "variacion_ultima_vs_anterior_pct",
        "descripcion_compra": "descripcion_ultima_compra",
        "documento_compra": "documento_ultima_compra",
        "folio_compra": "folio_ultima_compra",
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
    rel["Motivo promoción"] = rel["Motivo promoción"].astype(str).replace("nan", "").str.strip()
    rel["Ads/Comentario"] = rel["Ads/Comentario"].astype(str).replace("nan", "").str.strip()
    rel = rel[rel["SKU_norm"].notna()].copy()
    rel["source"] = "Relámpago mi página"
    return rel


@st.cache_data(show_spinner=False)
def build_model(master_bytes, compras_bytes=None):
    sheets = read_excel_all(master_bytes)
    master = sheets.get("MAESTRA de precios", pd.DataFrame()).copy()
    bridge = sheets.get("MLC -SKU", pd.DataFrame()).copy()
    promos = sheets.get("CONTROL DE PROMOCIONES", pd.DataFrame()).copy()
    relampago = prep_relampago((sheets.get("Relampago mi pagina") if "Relampago mi pagina" in sheets else sheets.get("Relámpago mi página", pd.DataFrame())).copy())

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
    master["MLC_norm"] = master.get(primary, pd.Series(dtype=object)).apply(normalize_mlc) if primary else pd.Series([None] * len(master), index=master.index)

    bridge["SKU_norm"] = bridge.iloc[:, 0].apply(normalize_sku) if len(bridge.columns) > 0 else pd.Series(dtype=object)
    bridge["MLC_norm"] = bridge.iloc[:, 1].apply(normalize_mlc) if len(bridge.columns) > 1 else pd.Series(dtype=object)
    bridge = bridge[["SKU_norm", "MLC_norm"]].dropna().drop_duplicates()

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
        promos["source"] = "Control promociones"
    else:
        promos = pd.DataFrame(columns=["SKU_norm", "MLC_norm", "Precio promocional", "next_campaign_date", "days_to_next", "source"])

    sku_mlc = pd.concat([
        master[["SKU_norm", "MLC_norm"]].dropna(),
        bridge[["SKU_norm", "MLC_norm"]].dropna(),
        promos[["SKU_norm", "MLC_norm"]].dropna(),
    ], ignore_index=True).drop_duplicates()

    product = master.copy()
    bridge_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index()
    promo_agg = promos.groupby("SKU_norm").agg(
        total_promos_control=("source", "count"),
        promo_mlcs=("MLC_norm", lambda s: list(pd.unique([x for x in s if pd.notna(x)]))),
        next_campaign_date=("next_campaign_date", "min"),
        min_promo_price=("Precio promocional", "min"),
        max_promo_price=("Precio promocional", "max"),
        ads_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
    ).reset_index() if not promos.empty else pd.DataFrame(columns=["SKU_norm"])

    rel_agg = relampago.groupby("SKU_norm").agg(
        relampago_count=("SKU_norm", "count"),
        relampago_min_price=("Precio promocional", "min"),
        relampago_max_price=("Precio promocional", "max"),
        relampago_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        relampago_motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
    ).reset_index() if not relampago.empty else pd.DataFrame(columns=["SKU_norm"])

    product = product.merge(bridge_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product = product.merge(rel_agg, on="SKU_norm", how="left")
    product["total_promos_control"] = product["total_promos_control"].fillna(0)
    product["relampago_count"] = product["relampago_count"].fillna(0)
    product["total_promos"] = product["total_promos_control"] + product["relampago_count"]
    product["days_to_next"] = (pd.to_datetime(product["next_campaign_date"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    product["search_text"] = (
        df_series(product, "SKU").astype(str).fillna("") + " | " +
        df_series(product, "DESCRIPCIÓN").astype(str).fillna("") + " | " +
        df_series(product, "MLC_norm").astype(str).fillna("") + " | " +
        df_series(product, "promo_mlcs").astype(str).fillna("")
    ).str.lower()

    compras = pd.DataFrame()
    compras_summary = pd.DataFrame()
    if compras_bytes is not None:
        try:
            c_sheets = read_excel_all(compras_bytes)
            best_name = max(c_sheets, key=lambda k: len(c_sheets[k]))
            compras, compras_summary, _ = prepare_compras_dataframe(c_sheets[best_name])
            if not compras_summary.empty:
                product = product.merge(compras_summary, on="SKU_norm", how="left")
        except Exception:
            compras = pd.DataFrame()
            compras_summary = pd.DataFrame()

    return {
        "sheets": sheets,
        "master": master,
        "bridge": bridge,
        "promos": promos,
        "relampago": relampago,
        "product": product,
        "compras": compras,
    }


def compute_decision_row(row):
    margin = safe_float(row.get("MARGEN MELI 1"))
    ads_txt = " ".join([str(row.get("ads_comment") or ""), str(row.get("relampago_comment") or "")]).lower()
    days = row.get("days_to_next")
    total_promos = safe_int(row.get("total_promos"), 0)
    relampago_count = safe_int(row.get("relampago_count"), 0)
    if total_promos == 0 and relampago_count == 0:
        return "Sin promo", "badge-gray"
    if not pd.isna(days) and days <= 1:
        return "Urgente renovar", "badge-red"
    if not pd.isna(margin) and margin < 0:
        return "Margen negativo", "badge-red"
    if "paus" in ads_txt:
        return "Revisar publicación", "badge-orange"
    if relampago_count > 0:
        return "Tiene relámpago", "badge-blue"
    if not pd.isna(days) and days <= 7:
        return "Revisar esta semana", "badge-yellow"
    return "Bajo control", "badge-green"


def make_download_workbook(all_sheets, master_df, bridge_df, promo_df, relampago_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            if name == "MAESTRA de precios":
                master_df.drop(columns=[c for c in master_df.columns if c.endswith("_norm") or c in ["search_text"]], errors="ignore").to_excel(writer, sheet_name=name, index=False)
            elif name == "MLC -SKU":
                base = bridge_df.copy()
                if "SKU_norm" in base.columns:
                    base = base.rename(columns={"SKU_norm": "SKU"})
                if "MLC_norm" in base.columns:
                    base = base.rename(columns={"MLC_norm": "Número de publicación"})
                base.to_excel(writer, sheet_name=name, index=False)
            elif name == "CONTROL DE PROMOCIONES":
                base = promo_df.copy()
                base = base.drop(columns=[c for c in ["mlc_list", "SKU_norm", "MLC_norm", "next_campaign_date", "days_to_next", "source"] if c in base.columns], errors="ignore")
                base.to_excel(writer, sheet_name=name, index=False)
            elif name == "Relampago mi pagina":
                base = relampago_df.copy()
                base = base[[c for c in RELAMPAGO_COLS if c in base.columns]].copy()
                base.to_excel(writer, sheet_name=name, index=False, header=False)
            else:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    return out.getvalue()


def rebuild_from_session(sheets, compras_file, master_df, bridge_df, promos_df, relampago_df):
    tmp_sheets = dict(sheets)
    tmp_sheets["MAESTRA de precios"] = master_df.copy()
    tmp_sheets["MLC -SKU"] = bridge_df.copy()
    tmp_sheets["CONTROL DE PROMOCIONES"] = promos_df.copy()
    tmp_sheets["Relampago mi pagina"] = relampago_df.copy()

    master = tmp_sheets["MAESTRA de precios"].copy()
    bridge = tmp_sheets["MLC -SKU"].copy()
    promos = tmp_sheets["CONTROL DE PROMOCIONES"].copy()
    relampago = prep_relampago(tmp_sheets["Relampago mi pagina"].copy())

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
    master["MLC_norm"] = master.get(primary, pd.Series(dtype=object)).apply(normalize_mlc) if primary else pd.Series([None] * len(master), index=master.index)

    bridge["SKU_norm"] = bridge.iloc[:, 0].apply(normalize_sku) if len(bridge.columns) > 0 else pd.Series(dtype=object)
    bridge["MLC_norm"] = bridge.iloc[:, 1].apply(normalize_mlc) if len(bridge.columns) > 1 else pd.Series(dtype=object)
    bridge = bridge[["SKU_norm", "MLC_norm"]].dropna().drop_duplicates()

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
        promos["source"] = "Control promociones"
    else:
        promos = pd.DataFrame(columns=["SKU_norm", "MLC_norm", "Precio promocional", "next_campaign_date", "days_to_next", "source"])

    sku_mlc = pd.concat([
        master[["SKU_norm", "MLC_norm"]].dropna(),
        bridge[["SKU_norm", "MLC_norm"]].dropna(),
        promos[["SKU_norm", "MLC_norm"]].dropna(),
    ], ignore_index=True).drop_duplicates()

    product = master.copy()
    bridge_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index()
    promo_agg = promos.groupby("SKU_norm").agg(
        total_promos_control=("source", "count"),
        promo_mlcs=("MLC_norm", lambda s: list(pd.unique([x for x in s if pd.notna(x)]))),
        next_campaign_date=("next_campaign_date", "min"),
        min_promo_price=("Precio promocional", "min"),
        max_promo_price=("Precio promocional", "max"),
        ads_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
    ).reset_index() if not promos.empty else pd.DataFrame(columns=["SKU_norm"])
    rel_agg = relampago.groupby("SKU_norm").agg(
        relampago_count=("SKU_norm", "count"),
        relampago_min_price=("Precio promocional", "min"),
        relampago_max_price=("Precio promocional", "max"),
        relampago_comment=("Ads/Comentario", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
        relampago_motivo=("Motivo promoción", lambda s: " | ".join([str(x) for x in s if pd.notna(x) and str(x).strip()][:5])),
    ).reset_index() if not relampago.empty else pd.DataFrame(columns=["SKU_norm"])
    product = product.merge(bridge_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product = product.merge(rel_agg, on="SKU_norm", how="left")
    product["total_promos_control"] = product["total_promos_control"].fillna(0)
    product["relampago_count"] = product["relampago_count"].fillna(0)
    product["total_promos"] = product["total_promos_control"] + product["relampago_count"]
    product["days_to_next"] = (pd.to_datetime(product["next_campaign_date"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    product["search_text"] = (
        df_series(product, "SKU").astype(str).fillna("") + " | " +
        df_series(product, "DESCRIPCIÓN").astype(str).fillna("") + " | " +
        df_series(product, "MLC_norm").astype(str).fillna("") + " | " +
        df_series(product, "promo_mlcs").astype(str).fillna("")
    ).str.lower()

    compras = pd.DataFrame()
    if compras_file is not None:
        try:
            c_sheets = read_excel_all(compras_file)
            best_name = max(c_sheets, key=lambda k: len(c_sheets[k]))
            compras, compras_summary, _ = prepare_compras_dataframe(c_sheets[best_name])
            if not compras_summary.empty:
                product = product.merge(compras_summary, on="SKU_norm", how="left")
        except Exception:
            compras = pd.DataFrame()

    return product, promos, relampago, compras


st.sidebar.title("Aurora Pricing Cockpit")
master_file = st.sidebar.file_uploader("Maestra integrada", type=["xlsx"], key="master")
compras_file = st.sidebar.file_uploader("Compras históricas", type=["xlsx"], key="compras")
st.sidebar.caption("Carga manual para trabajar siempre con el último archivo.")

if master_file is None:
    st.info("Carga la maestra integrada para empezar.")
    st.stop()

model = build_model(master_file, compras_file)
all_sheets = model["sheets"]

state_needs_refresh = st.session_state.get("source_name") != master_file.name
if state_needs_refresh:
    st.session_state.source_name = master_file.name

if state_needs_refresh or "master_df" not in st.session_state:
    st.session_state.master_df = model["master"].copy()
if state_needs_refresh or "bridge_df" not in st.session_state:
    st.session_state.bridge_df = model["bridge"].copy()
if state_needs_refresh or "promos_df" not in st.session_state:
    st.session_state.promos_df = model["promos"].copy()
if state_needs_refresh or "relampago_df" not in st.session_state:
    st.session_state.relampago_df = model.get("relampago", pd.DataFrame()).copy()

master_df = st.session_state.get("master_df", model["master"].copy())
bridge_df = st.session_state.get("bridge_df", model["bridge"].copy())
promos_df = st.session_state.get("promos_df", model["promos"].copy())
relampago_df = st.session_state.get("relampago_df", model.get("relampago", pd.DataFrame()).copy())

product_df, promos_df_view, relampago_view, compras_df = rebuild_from_session(
    all_sheets, compras_file, master_df, bridge_df, promos_df, relampago_df
)

download_bytes = make_download_workbook(all_sheets, st.session_state.master_df, st.session_state.bridge_df, st.session_state.promos_df, st.session_state.relampago_df)
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

    options = filtered.apply(lambda r: f"{r.get('SKU', '')} · {str(r.get('DESCRIPCIÓN', ''))[:100]}", axis=1).tolist()
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

    st.markdown("")
    b1, b2 = st.columns([1, 1])
    with b1:
        render_badge(decision_text, decision_cls)
    with b2:
        render_badge(urgency_text, urgency_cls)

    left, mid, right = st.columns([1.15, 1, 1])

    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Identidad")
        st.write(f"**SKU:** {row.get('SKU', '—')}")
        st.write(f"**Descripción:** {row.get('DESCRIPCIÓN', '—')}")
        st.write(f"**Ubicación:** {row.get('UBIC', '—')}")
        mlcs = row.get("MLC_norm")
        if isinstance(mlcs, list) and mlcs:
            st.write("**MLCs asociados:** " + ", ".join(mlcs))
        elif isinstance(row.get("promo_mlcs"), list) and row.get("promo_mlcs"):
            st.write("**MLCs asociados:** " + ", ".join(row.get("promo_mlcs")))
        else:
            st.write("**MLCs asociados:** —")
        st.write(f"**Comentario maestra:** {row.get('COMENTARIO', '—')}")
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
            st.write(f"**Última compra:** {pd.to_datetime(compras_info.get('ultima_compra')).strftime('%d-%m-%Y') if pd.notna(compras_info.get('ultima_compra')) else '—'}")
            st.write(f"**Último precio compra:** {money(compras_info.get('ultimo_precio_compra'))}")
            st.write(f"**Proveedor último:** {compras_info.get('proveedor_ultimo', '—')}")
            st.write(f"**Cantidad última compra:** {safe_int(compras_info.get('cantidad_ultima_compra'), 0) if pd.notna(compras_info.get('cantidad_ultima_compra')) else '—'}")
            st.write(f"**Variación vs compra anterior:** {pct((compras_info.get('variacion_ultima_vs_anterior_pct') / 100.0) if pd.notna(compras_info.get('variacion_ultima_vs_anterior_pct')) else np.nan)}")
            st.write(f"**Rango histórico:** {money(compras_info.get('precio_min_hist'))} a {money(compras_info.get('precio_max_hist'))}")
            if compras_info.get("proveedores"):
                st.write(f"**Proveedores históricos:** {compras_info.get('proveedores')}")
            with st.expander("Ver historial de compras"):
                hist = compras_rows.sort_values("fecha_compra", ascending=False)[[
                    c for c in ["fecha_compra", "proveedor", "precio_compra", "cantidad", "descripcion_compra", "SKU_norm", "match_method", "match_score"] if c in compras_rows.columns
                ]].copy()
                hist = hist.rename(columns={
                    "fecha_compra": "Fecha",
                    "proveedor": "Proveedor",
                    "precio_compra": "Precio compra",
                    "cantidad": "Cantidad",
                    "descripcion_compra": "Descripción compra",
                    "SKU_norm": "SKU compra",
                    "match_method": "Método match",
                    "match_score": "Score",
                })
                st.dataframe(hist, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with mid:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Precio y rentabilidad")
        m1, m2 = st.columns(2)
        m1.metric("Precio neto", money(row.get("PRECIO NETO")))
        m2.metric("Cambio precio", money(row.get("CAMBIO DE PRECIO")))
        m1.metric("Margen local", pct(row.get("MARGEN LOCAL")))
        m2.metric("Margen Meli 1", pct(row.get("MARGEN MELI 1")))
        st.write(f"**Precio promo mínimo control:** {money(row.get('min_promo_price'))}")
        st.write(f"**Precio promo máximo control:** {money(row.get('max_promo_price'))}")
        st.write(f"**Precio relámpago mínimo:** {money(row.get('relampago_min_price'))}")
        st.write(f"**Precio relámpago máximo:** {money(row.get('relampago_max_price'))}")
        st.write(f"**Ads / comentario control:** {row.get('ads_comment', '—')}")
        st.write(f"**Ads / comentario relámpago:** {row.get('relampago_comment', '—')}")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Lectura automática")
        bullets = []
        if safe_int(row.get("total_promos_control"), 0) == 0 and safe_int(row.get("relampago_count"), 0) == 0:
            bullets.append("Producto sin promo activa registrada.")
        if pd.notna(row.get("days_to_next")) and row.get("days_to_next") <= 1:
            bullets.append("Urgente: una promo del control vence hoy o mañana.")
        if pd.notna(row.get("MARGEN MELI 1")) and safe_float(row.get("MARGEN MELI 1")) < 0:
            bullets.append("Margen Meli 1 negativo. Revisar precio o campaña.")
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
        sku_promos = promos_df_view[promos_df_view["SKU_norm"] == sku].copy()
        sku_rel = relampago_view[relampago_view["SKU_norm"] == sku].copy()

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Promos control")
        st.metric("Filas control", len(sku_promos))
        if not sku_promos.empty:
            display_cols = [c for c in ["MLC_norm", "Precio promocional", "Motivo promoción", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"] if c in sku_promos.columns]
            st.dataframe(sku_promos[display_cols], use_container_width=True, hide_index=True, height=220)
        else:
            st.info("Sin promos en control.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Relámpago mi página")
        st.metric("Filas relámpago", len(sku_rel))
        if not sku_rel.empty:
            display_cols = [c for c in ["Descripción", "Precio promocional", "Motivo promoción", "Ads/Comentario"] if c in sku_rel.columns]
            st.dataframe(sku_rel[display_cols], use_container_width=True, hide_index=True, height=180)
        else:
            st.info("No está en relámpago mi página.")
        st.markdown("</div>", unsafe_allow_html=True)

    sku_hist = compras_df[compras_df["SKU_norm"] == sku].copy() if not compras_df.empty else pd.DataFrame()
    if not sku_hist.empty and "fecha_compra" in sku_hist.columns and "precio_compra" in sku_hist.columns:
        sku_hist = sku_hist.sort_values("fecha_compra")
        chart_df = sku_hist[["fecha_compra", "precio_compra"]].dropna()
        if not chart_df.empty:
            st.subheader("Evolución precio de compra")
            st.line_chart(chart_df.set_index("fecha_compra"))


elif page == "Operador de promos":
    st.markdown('<div class="big-title">Operador de promociones</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Bandeja operativa + edición guiada para control y relámpago en un solo lugar.</div>', unsafe_allow_html=True)

    control_summary = product_df.copy()
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Vencen hoy", int((control_summary["days_to_next"] == 0).sum()))
    k2.metric("Vencen mañana", int((control_summary["days_to_next"] == 1).sum()))
    k3.metric("Vencen en 3 días", int((control_summary["days_to_next"] == 3).sum()))
    k4.metric("Sin precio promo", int(promos_df_view["Precio promocional"].isna().sum()) + int(relampago_view["Precio promocional"].isna().sum()))
    k5.metric("En relámpago", int((control_summary["relampago_count"].fillna(0) > 0).sum()))

    op_tab1, op_tab2, op_tab3 = st.tabs(["Gestión por producto", "Bandeja operativa", "Relámpago mi página"])

    with op_tab1:
        filt = product_df.copy()
        search_prod = st.text_input("Buscar SKU o descripción", key="op_prod_search")
        if search_prod.strip():
            filt = filt[filt["search_text"].str.contains(search_prod.lower().strip(), na=False)]
        if filt.empty:
            st.info("Sin coincidencias.")
        else:
            prod_options = filt.apply(lambda r: f"{r.get('SKU', '')} · {str(r.get('DESCRIPCIÓN', ''))[:100]}", axis=1).tolist()
            selected_prod = st.selectbox("Producto", prod_options, key="op_prod_select")
            prow = filt.iloc[prod_options.index(selected_prod)]
            psku = prow["SKU_norm"]

            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.write(f"**SKU:** {prow.get('SKU', '—')}")
            st.write(f"**Descripción:** {prow.get('DESCRIPCIÓN', '—')}")
            st.write(f"**Margen Meli 1:** {pct(prow.get('MARGEN MELI 1'))}")
            st.write(f"**Precio bruto tienda:** {money(prow.get('PRECIO BRUTO'))}")
            mlcs = prow.get("MLC_norm")
            st.write(f"**MLCs asociados:** {', '.join(mlcs) if isinstance(mlcs, list) and mlcs else '—'}")
            st.markdown("</div>", unsafe_allow_html=True)

            sku_control = promos_df_view[promos_df_view["SKU_norm"] == psku].copy()
            sku_rel = relampago_view[relampago_view["SKU_norm"] == psku].copy()

            sub1, sub2 = st.columns(2)
            with sub1:
                st.markdown("### Control de promociones")
                if sku_control.empty:
                    st.info("Este SKU no tiene filas en control.")
                else:
                    control_options = sku_control.apply(lambda r: f"{r.get('MLC_norm', 'Sin MLC')} · {money(r.get('Precio promocional'))}", axis=1).tolist()
                    chosen_control = st.selectbox("Fila control", control_options, key="control_row_select")
                    ctrl_row = sku_control.iloc[control_options.index(chosen_control)]
                    ctrl_idx = ctrl_row.name

                    with st.form("control_form"):
                        col1, col2 = st.columns(2)
                        new_mlc = col1.text_input("MLC", value=str(ctrl_row.get("MLC_norm") or ""))
                        new_price = col2.number_input("Precio promocional", min_value=0.0, value=float(safe_float(ctrl_row.get("Precio promocional"), 0.0)), step=1.0)
                        new_motivo = col1.text_input("Motivo promoción", value=str(ctrl_row.get("Motivo promoción") or ""))
                        new_ads = col2.text_input("Ads / comentario", value=str(ctrl_row.get("Ads/Comentario") or ""))
                        c1, c2, c3, c4 = st.columns(4)
                        camp_values = []
                        for i, col in enumerate(CAMPAIGN_COLS):
                            cur = pd.to_datetime(ctrl_row.get(col), errors="coerce")
                            camp_values.append([c1, c2, c3, c4][i].date_input(col, value=cur.date() if pd.notna(cur) else None, key=f"{col}_{ctrl_idx}"))
                        save_ctrl = st.form_submit_button("Guardar fila control", type="primary", use_container_width=True)

                    if save_ctrl:
                        base = st.session_state.promos_df.copy()
                        base.loc[ctrl_idx, "N° Publicación"] = new_mlc
                        if "Precio promocional" in base.columns:
                            base.loc[ctrl_idx, "Precio promocional"] = new_price if new_price > 0 else np.nan
                        if "Motivo promoción" in base.columns:
                            base.loc[ctrl_idx, "Motivo promoción"] = new_motivo
                        if "Ads/Comentario" in base.columns:
                            base.loc[ctrl_idx, "Ads/Comentario"] = new_ads
                        for col, val in zip(CAMPAIGN_COLS, camp_values):
                            if col in base.columns:
                                base.loc[ctrl_idx, col] = pd.to_datetime(val) if val else pd.NaT
                        st.session_state.promos_df = base
                        st.success("Fila control actualizada.")
                        st.rerun()

                    if st.button("Eliminar fila control", key=f"del_control_{ctrl_idx}", use_container_width=True):
                        st.session_state.promos_df = st.session_state.promos_df.drop(index=ctrl_idx).reset_index(drop=True)
                        st.success("Fila control eliminada.")
                        st.rerun()

                with st.expander("Agregar nueva fila a control", expanded=False):
                    with st.form("add_control_form"):
                        a1, a2 = st.columns(2)
                        add_mlc = a1.text_input("MLC nuevo")
                        add_price = a2.number_input("Precio promocional nuevo", min_value=0.0, value=0.0, step=1.0)
                        add_motivo = a1.text_input("Motivo")
                        add_ads = a2.text_input("Ads / comentario")
                        d1, d2, d3, d4 = st.columns(4)
                        add_dates = [
                            d1.date_input("Campaña 1 nueva", value=None, key="add_c1"),
                            d2.date_input("Campaña 2 nueva", value=None, key="add_c2"),
                            d3.date_input("Campaña 3 nueva", value=None, key="add_c3"),
                            d4.date_input("Campaña 4 nueva", value=None, key="add_c4"),
                        ]
                        add_submit = st.form_submit_button("Agregar a control", use_container_width=True)
                    if add_submit:
                        row_new = {c: np.nan for c in st.session_state.promos_df.columns}
                        first_col = st.session_state.promos_df.columns[0]
                        row_new[first_col] = prow.get("SKU")
                        if "N° Publicación" in row_new:
                            row_new["N° Publicación"] = add_mlc
                        if "Descripción" in row_new:
                            row_new["Descripción"] = prow.get("DESCRIPCIÓN")
                        if "Precio promocional" in row_new:
                            row_new["Precio promocional"] = add_price if add_price > 0 else np.nan
                        if "Motivo promoción" in row_new:
                            row_new["Motivo promoción"] = add_motivo
                        if "Ads/Comentario" in row_new:
                            row_new["Ads/Comentario"] = add_ads
                        for col, val in zip(CAMPAIGN_COLS, add_dates):
                            if col in row_new:
                                row_new[col] = pd.to_datetime(val) if val else pd.NaT
                        st.session_state.promos_df = pd.concat([st.session_state.promos_df, pd.DataFrame([row_new])], ignore_index=True)
                        if add_mlc:
                            new_bridge = pd.DataFrame([{"SKU_norm": normalize_sku(prow.get("SKU")), "MLC_norm": normalize_mlc(add_mlc)}])
                            st.session_state.bridge_df = pd.concat([st.session_state.bridge_df, new_bridge], ignore_index=True).drop_duplicates()
                        st.success("Fila agregada a control.")
                        st.rerun()

            with sub2:
                st.markdown("### Relámpago mi página")
                if sku_rel.empty:
                    st.info("Este SKU no está en relámpago.")
                else:
                    rel_options = sku_rel.apply(lambda r: f"{money(r.get('Precio promocional'))} · {r.get('Motivo promoción', '')}", axis=1).tolist()
                    chosen_rel = st.selectbox("Fila relámpago", rel_options, key="rel_row_select")
                    rel_row = sku_rel.iloc[rel_options.index(chosen_rel)]
                    rel_idx = rel_row.name

                    with st.form("rel_form"):
                        rr1, rr2 = st.columns(2)
                        rel_desc = rr1.text_input("Descripción", value=str(rel_row.get("Descripción") or ""))
                        rel_price = rr2.number_input("Precio promocional", min_value=0.0, value=float(safe_float(rel_row.get("Precio promocional"), 0.0)), step=1.0)
                        rel_motivo = rr1.text_input("Motivo promoción", value=str(rel_row.get("Motivo promoción") or ""))
                        rel_ads = rr2.text_input("Ads / comentario", value=str(rel_row.get("Ads/Comentario") or ""))
                        rel_save = st.form_submit_button("Guardar fila relámpago", type="primary", use_container_width=True)
                    if rel_save:
                        base = st.session_state.relampago_df.copy()
                        base.loc[rel_idx, "Descripción"] = rel_desc
                        base.loc[rel_idx, "Precio promocional"] = rel_price if rel_price > 0 else np.nan
                        base.loc[rel_idx, "Motivo promoción"] = rel_motivo
                        base.loc[rel_idx, "Ads/Comentario"] = rel_ads
                        st.session_state.relampago_df = base
                        st.success("Fila relámpago actualizada.")
                        st.rerun()

                    if st.button("Eliminar fila relámpago", key=f"del_rel_{rel_idx}", use_container_width=True):
                        st.session_state.relampago_df = st.session_state.relampago_df.drop(index=rel_idx).reset_index(drop=True)
                        st.success("Fila relámpago eliminada.")
                        st.rerun()

                with st.expander("Agregar a relámpago mi página", expanded=False):
                    with st.form("add_rel_form"):
                        z1, z2 = st.columns(2)
                        add_rel_desc = z1.text_input("Descripción relámpago", value=str(prow.get("DESCRIPCIÓN") or ""))
                        add_rel_price = z2.number_input("Precio promocional relámpago", min_value=0.0, value=0.0, step=1.0)
                        add_rel_motivo = z1.text_input("Motivo relámpago", value="LIQUIDACION")
                        add_rel_ads = z2.text_input("Estado / comentario")
                        add_rel_submit = st.form_submit_button("Agregar a relámpago", use_container_width=True)
                    if add_rel_submit:
                        row_new = {
                            "SKU": prow.get("SKU"),
                            "Descripción": add_rel_desc,
                            "Precio promocional": add_rel_price if add_rel_price > 0 else np.nan,
                            "Extra": np.nan,
                            "Motivo promoción": add_rel_motivo,
                            "Ads/Comentario": add_rel_ads,
                        }
                        st.session_state.relampago_df = pd.concat([st.session_state.relampago_df, pd.DataFrame([row_new])], ignore_index=True)
                        st.success("Producto agregado a relámpago.")
                        st.rerun()

    with op_tab2:
        promo_universe = promos_df_view.copy()
        promo_universe["Origen"] = "Control"
        promo_universe["Tipo"] = "Control"
        rel_table = relampago_view.copy()
        rel_table["Origen"] = "Relámpago"
        rel_table["Tipo"] = "Relámpago"
        rel_table["MLC_norm"] = ""
        rel_table["next_campaign_date"] = pd.NaT
        rel_table["days_to_next"] = np.nan
        if "Campaña 1" not in rel_table.columns:
            rel_table["Campaña 1"] = pd.NaT

        all_ops = pd.concat([
            promo_universe[[c for c in ["SKU_norm", "MLC_norm", "Descripción", "Precio promocional", "Motivo promoción", "Ads/Comentario", "next_campaign_date", "days_to_next", "Origen", "Tipo"] if c in promo_universe.columns]],
            rel_table[[c for c in ["SKU_norm", "MLC_norm", "Descripción", "Precio promocional", "Motivo promoción", "Ads/Comentario", "next_campaign_date", "days_to_next", "Origen", "Tipo"] if c in rel_table.columns]],
        ], ignore_index=True)

        f1, f2, f3, f4, f5 = st.columns(5)
        only_today = f1.checkbox("Vencen hoy")
        only_tomorrow = f2.checkbox("Vencen mañana")
        only_3d = f3.checkbox("Vencen en 3 días")
        without_price = f4.checkbox("Sin precio promocional")
        rel_only = f5.checkbox("Solo relámpago")

        f6, f7 = st.columns([1, 2])
        source_filter = f6.selectbox("Origen", ["Todos", "Control", "Relámpago"])
        search = f7.text_input("Buscar SKU / MLC / descripción", key="promo_search")

        work = all_ops.copy()
        if only_today:
            work = work[work["days_to_next"] == 0]
        if only_tomorrow:
            work = work[work["days_to_next"] == 1]
        if only_3d:
            work = work[work["days_to_next"] == 3]
        if without_price:
            work = work[work["Precio promocional"].isna()]
        if rel_only or source_filter == "Relámpago":
            work = work[work["Origen"] == "Relámpago"]
        elif source_filter == "Control":
            work = work[work["Origen"] == "Control"]
        if search.strip():
            mask = (
                df_series(work, "SKU_norm").astype(str).str.contains(search, case=False, na=False) |
                df_series(work, "MLC_norm").astype(str).str.contains(search, case=False, na=False) |
                df_series(work, "Descripción").astype(str).str.contains(search, case=False, na=False)
            )
            work = work[mask]
        st.dataframe(work.sort_values(["Origen", "days_to_next"], ascending=[True, True]), use_container_width=True, hide_index=True, height=500)

    with op_tab3:
        st.caption("Lista dedicada para agregar, sacar o modificar filas de relámpago mi página.")
        rel_edit = st.data_editor(
            st.session_state.relampago_df[[c for c in RELAMPAGO_COLS if c in st.session_state.relampago_df.columns]].copy(),
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=520,
            key="rel_editor",
        )
        if st.button("Guardar lista relámpago", type="primary", use_container_width=True):
            st.session_state.relampago_df = rel_edit.copy()
            st.success("Lista relámpago actualizada.")
            st.rerun()


elif page == "Alta de producto":
    st.markdown('<div class="big-title">Alta de producto nuevo</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtle">Crea SKU, MLCs y promo inicial con una ficha parecida al operador.</div>', unsafe_allow_html=True)

    with st.form("new_product"):
        c1, c2, c3 = st.columns(3)
        new_sku = c1.text_input("SKU *")
        desc = c2.text_input("Descripción")
        ubic = c3.text_input("Ubicación")
        costo = c1.number_input("Último costo", min_value=0.0, value=0.0, step=1.0)
        precio_tienda = c2.number_input("Precio bruto en tienda", min_value=0.0, value=0.0, step=1.0)
        margen_meli1 = c3.number_input("Margen Meli 1", value=0.0, step=0.1)
        mlcs_text = st.text_input("MLCs asociados (separados por coma)")
        comentario = st.text_input("Comentario maestra")

        precio_neto_calc = precio_tienda / 1.19 if precio_tienda else np.nan
        margen_local_calc = ((precio_neto_calc - costo) / costo * 100.0) if costo and precio_neto_calc == precio_neto_calc else np.nan

        p1, p2 = st.columns(2)
        p1.write(f"**Precio neto calculado:** {money(precio_neto_calc)}")
        p2.write(f"**Margen local calculado:** {pct(margen_local_calc)}")

        st.markdown("**Promo inicial opcional**")
        p1, p2, p3 = st.columns(3)
        promo_price = p1.number_input("Precio promocional", min_value=0.0, value=0.0, step=1.0)
        motivo = p2.text_input("Motivo")
        ads = p3.text_input("Ads / comentario")
        camp1 = st.date_input("Campaña 1", value=None)
        add_relampago = st.checkbox("Agregar también a relámpago mi página")
        submitted = st.form_submit_button("Crear producto", type="primary")

    if submitted:
        sku_norm = normalize_sku(new_sku)
        if not sku_norm:
            st.error("SKU inválido.")
        elif sku_norm in set(st.session_state.master_df["SKU_norm"].dropna().astype(str)):
            st.error("Ese SKU ya existe.")
        else:
            new_master = {c: np.nan for c in st.session_state.master_df.columns}
            assignments = {
                "SKU": new_sku,
                "DESCRIPCIÓN": desc,
                "UBIC": ubic,
                "ÚLTIMO COSTO": costo,
                "PRECIO BRUTO": precio_tienda,
                "PRECIO NETO": precio_neto_calc,
                "MARGEN LOCAL": margen_local_calc,
                "MARGEN MELI 1": margen_meli1,
                "COMENTARIO": comentario,
                "SKU_norm": sku_norm,
            }
            for k, v in assignments.items():
                if k in new_master:
                    new_master[k] = v
            st.session_state.master_df = pd.concat([st.session_state.master_df, pd.DataFrame([new_master])], ignore_index=True)

            mlcs = [normalize_mlc(x) for x in mlcs_text.split(",") if normalize_mlc(x)]
            if mlcs:
                st.session_state.bridge_df = pd.concat([
                    st.session_state.bridge_df,
                    pd.DataFrame([{"SKU_norm": sku_norm, "MLC_norm": mlc} for mlc in mlcs])
                ], ignore_index=True).drop_duplicates()

            if promo_price > 0 or motivo or ads or camp1:
                row_new = {c: np.nan for c in st.session_state.promos_df.columns}
                first_col = st.session_state.promos_df.columns[0]
                row_new[first_col] = new_sku
                if "N° Publicación" in row_new and mlcs:
                    row_new["N° Publicación"] = mlcs[0]
                if "Descripción" in row_new:
                    row_new["Descripción"] = desc
                if "Precio promocional" in row_new:
                    row_new["Precio promocional"] = promo_price if promo_price > 0 else np.nan
                if "Motivo promoción" in row_new:
                    row_new["Motivo promoción"] = motivo
                if "Ads/Comentario" in row_new:
                    row_new["Ads/Comentario"] = ads
                if "Campaña 1" in row_new:
                    row_new["Campaña 1"] = pd.to_datetime(camp1) if camp1 else pd.NaT
                st.session_state.promos_df = pd.concat([st.session_state.promos_df, pd.DataFrame([row_new])], ignore_index=True)

            if add_relampago:
                row_rel = {
                    "SKU": new_sku,
                    "Descripción": desc,
                    "Precio promocional": promo_price if promo_price > 0 else np.nan,
                    "Extra": np.nan,
                    "Motivo promoción": motivo if motivo else "LIQUIDACION",
                    "Ads/Comentario": ads,
                }
                st.session_state.relampago_df = pd.concat([st.session_state.relampago_df, pd.DataFrame([row_rel])], ignore_index=True)

            st.success("Producto creado en memoria. Descarga el Excel actualizado.")
            st.rerun()
