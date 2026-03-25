import io
import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "Aurora Pricing & Promos"
MASTER_SHEETS = [
    "BASE DIMASOFT MELI4 (no tocar)",
    "MLC -SKU",
    "ART KAME",
    "MAESTRA de precios",
    "Relampago mi pagina",
    "CALCULADORA",
    "CONTROL DE PROMOCIONES",
    "cambio precio max",
    "cambio nombre full max",
    "Stock 3112",
    "base inactivas",
    "Base kame",
]

CORE_MAESTRA_COLUMNS = [
    "SKU",
    "DESCRIPCIÓN",
    "UBIC",
    "CAMBIO DE PRECIO",
    "ÚLTIMO COSTO",
    "MARGEN LOCAL",
    "PRECIO NETO",
    "PRECIO BRUTO",
    "MARGEN MELI 1",
    " NETO MELI 1",
    "MONTO EN SIMULACIÓN",
    "MLC",
    "CAMPAÑA PADS",
    "PRECIO B2C PUBLICADO ",
    "% DCTO",
    "FECHA VENCI",
    "COMENTARIO",
    "MARGEN MELI 2",
    "NETO MELI 2",
    "VENTA BRUTO MELI 2",
    "MLC.1",
    "CAMPAÑA PADS.1",
    "PRECIO B2C",
    "% DCTO.1",
    "FECHA VENCI.1",
    "COMENTARIO.1",
]

PROMO_DATE_COLS = ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]
PROMO_CORE_COLUMNS = [
    " ",
    "% F",
    "N° Publicación",
    "Descripción",
    "% F.1",
    "Precio promocional",
    "Motivo promoción",
    "Unnamed: 7",
    "margen",
    "Ads/Comentario",
    *PROMO_DATE_COLS,
]

st.set_page_config(page_title=APP_TITLE, layout="wide")


# ---------- helpers ----------
def normalize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    return df


def looks_like_empty_col(name: str) -> bool:
    s = str(name).strip().lower()
    return s.startswith("unnamed:")


def clean_for_editor(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = pd.to_datetime(out[c], errors="coerce").dt.date
        elif out[c].dtype == object:
            out[c] = out[c].replace({np.nan: None})
    return out


def sanitize_numeric_text(value):
    if pd.isna(value):
        return None
    s = str(value).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def normalize_sku(value):
    s = sanitize_numeric_text(value)
    if s is None or s == "":
        return None
    nums = re.findall(r"\d+", s)
    return nums[0] if nums else s


def normalize_mlc(value):
    if pd.isna(value):
        return None
    s = str(value).upper().strip()
    nums = re.findall(r"(\d{7,})", s)
    if nums:
        return f"MLC{nums[0]}"
    return None


def extract_mlcs(value):
    if pd.isna(value):
        return []
    s = str(value).upper()
    nums = re.findall(r"(?:MLC)?\s*(\d{7,})", s)
    unique = []
    for n in nums:
        mlc = f"MLC{n}"
        if mlc not in unique:
            unique.append(mlc)
    return unique


def first_existing(df: pd.DataFrame, cols):
    for c in cols:
        if c in df.columns:
            return c
    return None


def safe_to_numeric(series):
    return pd.to_numeric(series, errors="coerce")


def upcoming_status(next_date: pd.Timestamp | None):
    if pd.isna(next_date) or next_date is None:
        return "Sin fecha"
    today = pd.Timestamp(date.today())
    diff = (pd.Timestamp(next_date).normalize() - today).days
    if diff < 0:
        return f"Vencida hace {-diff} día(s)"
    if diff == 0:
        return "Vence hoy"
    if diff == 1:
        return "Vence mañana"
    if diff == 2:
        return "Vence pasado mañana"
    return f"Vence en {diff} día(s)"


def build_download_workbook(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df_to_save = df.copy()
            for col in df_to_save.columns:
                if pd.api.types.is_datetime64_any_dtype(df_to_save[col]):
                    df_to_save[col] = pd.to_datetime(df_to_save[col], errors="coerce")
            df_to_save.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()


def parse_purchase_file(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    df = pd.read_excel(uploaded)
    df = normalize_colnames(df)
    if "SKU" not in df.columns:
        return pd.DataFrame()
    df["SKU_norm"] = df["SKU"].apply(normalize_sku)
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)
    if "Precio Un." in df.columns:
        df["Precio Un."] = pd.to_numeric(df["Precio Un."], errors="coerce")
    return df[df["SKU_norm"].notna()].copy()


@st.cache_data(show_spinner=False)
def load_workbook(uploaded_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(uploaded_bytes))
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = normalize_colnames(pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=name))
    return sheets


def prepare_state(sheets: dict[str, pd.DataFrame]):
    if "sheets" not in st.session_state:
        st.session_state.sheets = {k: v.copy() for k, v in sheets.items()}
    if "loaded_signature" not in st.session_state:
        st.session_state.loaded_signature = None


def rebuild_views():
    maestra = st.session_state.sheets.get("MAESTRA de precios", pd.DataFrame()).copy()
    promo = st.session_state.sheets.get("CONTROL DE PROMOCIONES", pd.DataFrame()).copy()
    mlcsku = st.session_state.sheets.get("MLC -SKU", pd.DataFrame()).copy()

    maestra = normalize_colnames(maestra)
    promo = normalize_colnames(promo)
    mlcsku = normalize_colnames(mlcsku)

    if "SKU" in maestra.columns:
        maestra["SKU_norm"] = maestra["SKU"].apply(normalize_sku)
    else:
        maestra["SKU_norm"] = None

    mlc_col = first_existing(mlcsku, ["Número de publicación", "Numero de publicacion", "MLC"])
    if "SKU" in mlcsku.columns:
        mlcsku["SKU_norm"] = mlcsku["SKU"].apply(normalize_sku)
    else:
        mlcsku["SKU_norm"] = None
    if mlc_col:
        mlcsku["MLC_norm"] = mlcsku[mlc_col].apply(normalize_mlc)
    else:
        mlcsku["MLC_norm"] = None

    promo_sku_col = first_existing(promo, [" ", "SKU", "Sku"])
    promo_pub_col = first_existing(promo, ["N° Publicación", "N° Publicacion", "N Publicación", "N Publicacion"])
    if promo_sku_col:
        promo["SKU_norm"] = promo[promo_sku_col].apply(normalize_sku)
    else:
        promo["SKU_norm"] = None
    promo["promo_row_id"] = np.arange(len(promo))
    promo["promo_mlcs"] = promo[promo_pub_col].apply(extract_mlcs) if promo_pub_col else [[] for _ in range(len(promo))]

    for col in PROMO_DATE_COLS:
        if col in promo.columns:
            promo[col] = pd.to_datetime(promo[col], errors="coerce")
        else:
            promo[col] = pd.NaT

    # explode promotions by mlc for linking
    promo_exp = promo.explode("promo_mlcs").rename(columns={"promo_mlcs": "MLC_norm"})
    promo_exp["MLC_norm"] = promo_exp["MLC_norm"].replace({"": None})

    # collect mlcs per SKU from mapping + maestra columns
    maestra_mlc_frames = []
    for c in ["MLC", "MLC.1", "Unnamed: 12"]:
        if c in maestra.columns:
            tmp = maestra[["SKU_norm", c]].copy()
            tmp["MLC_norm"] = tmp[c].apply(normalize_mlc)
            maestra_mlc_frames.append(tmp[["SKU_norm", "MLC_norm"]])
    maestra_mlc = pd.concat(maestra_mlc_frames, ignore_index=True) if maestra_mlc_frames else pd.DataFrame(columns=["SKU_norm", "MLC_norm"])

    sku_mlc_map = pd.concat([
        mlcsku[["SKU_norm", "MLC_norm"]] if {"SKU_norm", "MLC_norm"}.issubset(mlcsku.columns) else pd.DataFrame(columns=["SKU_norm", "MLC_norm"]),
        maestra_mlc,
    ], ignore_index=True).dropna().drop_duplicates()

    # add promo rows that have sku only even without mlc mapping
    sku_from_promo = promo[["promo_row_id", "SKU_norm"]].copy()

    # fusion summary per SKU
    grouped_mlc = sku_mlc_map.groupby("SKU_norm")["MLC_norm"].apply(lambda x: sorted(set([v for v in x if pd.notna(v)]))).reset_index(name="MLCs") if not sku_mlc_map.empty else pd.DataFrame(columns=["SKU_norm", "MLCs"])

    promo_exp_by_sku = promo_exp.merge(sku_mlc_map, on="MLC_norm", how="left", suffixes=("", "_map"))
    promo_exp_by_sku["SKU_join"] = promo_exp_by_sku["SKU_norm_map"].combine_first(promo_exp_by_sku["SKU_norm"])
    promo_exp_by_sku = promo_exp_by_sku.drop(columns=[c for c in ["SKU_norm_map"] if c in promo_exp_by_sku.columns])

    # summary of dates/promos
    def compute_next_campaign(row):
        vals = []
        for c in PROMO_DATE_COLS:
            v = row.get(c)
            if pd.notna(v):
                vals.append(pd.Timestamp(v).normalize())
        return min(vals) if vals else pd.NaT

    promo["next_campaign_date"] = promo.apply(compute_next_campaign, axis=1)
    promo["campaign_status"] = promo["next_campaign_date"].apply(upcoming_status)

    promo_per_sku = promo_exp_by_sku.groupby("SKU_join").agg(
        total_promos=("promo_row_id", "nunique"),
        min_promo_price=("Precio promocional", lambda s: pd.to_numeric(s, errors="coerce").min()),
        max_promo_price=("Precio promocional", lambda s: pd.to_numeric(s, errors="coerce").max()),
    ).reset_index().rename(columns={"SKU_join": "SKU_norm"}) if not promo_exp_by_sku.empty else pd.DataFrame(columns=["SKU_norm", "total_promos", "min_promo_price", "max_promo_price"])

    next_dates = []
    for sku, grp in promo_exp_by_sku.groupby("SKU_join") if not promo_exp_by_sku.empty else []:
        dates = []
        for c in PROMO_DATE_COLS:
            if c in grp.columns:
                dates.extend([pd.Timestamp(v).normalize() for v in grp[c].dropna().tolist()])
        next_date = min([d for d in dates if pd.notna(d)], default=pd.NaT)
        ads_vals = grp.get("Ads/Comentario", pd.Series(dtype=object)).dropna().astype(str).unique().tolist() if "Ads/Comentario" in grp.columns else []
        next_dates.append({
            "SKU_norm": sku,
            "next_campaign_date": next_date,
            "campaign_status": upcoming_status(next_date),
            "ads_resumen": " | ".join(ads_vals[:3]),
        })
    next_dates = pd.DataFrame(next_dates)

    fusion = maestra.merge(grouped_mlc, on="SKU_norm", how="left") if not maestra.empty else maestra
    if not promo_per_sku.empty:
        fusion = fusion.merge(promo_per_sku, on="SKU_norm", how="left")
    if not next_dates.empty:
        fusion = fusion.merge(next_dates, on="SKU_norm", how="left")
    if "MLCs" in fusion.columns:
        fusion["MLCs"] = fusion["MLCs"].apply(lambda x: ", ".join(x) if isinstance(x, list) else "")

    # calc decision helper
    for col in ["ÚLTIMO COSTO", "PRECIO BRUTO", "PRECIO B2C PUBLICADO ", "PRECIO B2C"]:
        if col in fusion.columns:
            fusion[col] = pd.to_numeric(fusion[col], errors="coerce")
    fusion["precio_base_decision"] = fusion.get("PRECIO B2C PUBLICADO ", pd.Series(index=fusion.index, dtype=float)).combine_first(
        fusion.get("PRECIO B2C", pd.Series(index=fusion.index, dtype=float))
    ).combine_first(fusion.get("PRECIO BRUTO", pd.Series(index=fusion.index, dtype=float)))
    fusion["promo_vs_base_%"] = np.where(
        fusion["precio_base_decision"].notna() & fusion.get("min_promo_price", pd.Series(index=fusion.index, dtype=float)).notna() & (fusion["precio_base_decision"] != 0),
        (fusion.get("min_promo_price", pd.Series(index=fusion.index, dtype=float)) / fusion["precio_base_decision"] - 1) * 100,
        np.nan,
    )

    st.session_state.views = {
        "maestra": maestra,
        "promo": promo,
        "mlcsku": mlcsku,
        "sku_mlc_map": sku_mlc_map,
        "promo_exp_by_sku": promo_exp_by_sku,
        "fusion": fusion,
    }


# ---------- app ----------
st.title(APP_TITLE)
st.caption("Nueva app base sobre la maestra integrada: operación de promos + ficha de decisión por producto.")

with st.sidebar:
    st.header("Archivos")
    master_file = st.file_uploader("Maestra integrada", type=["xlsx"], key="master_file")
    purchase_file = st.file_uploader("Compras históricas", type=["xlsx"], key="purchase_file")

    st.markdown("---")
    st.subheader("Descarga")
    if st.session_state.get("sheets"):
        export_sheets = {k: v.copy() for k, v in st.session_state.sheets.items()}
        if "fusion" in st.session_state.get("views", {}):
            export_sheets["MAESTRA APP FUSIONADA"] = st.session_state.views["fusion"].drop(columns=[c for c in ["SKU_norm"] if c in st.session_state.views["fusion"].columns])
        out = build_download_workbook(export_sheets)
        st.download_button(
            "Descargar Excel actualizado",
            data=out,
            file_name=f"MAESTRA_PRECIOS_PROMOS_APP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.info("La app genera un Excel nuevo con tus cambios. No sobrescribe automáticamente tu archivo local.")

if not master_file:
    st.info("Carga la maestra integrada para empezar.")
    st.stop()

master_bytes = master_file.getvalue()
signature = (master_file.name, len(master_bytes))
if st.session_state.get("loaded_signature") != signature:
    sheets = load_workbook(master_bytes)
    st.session_state.sheets = {k: v.copy() for k, v in sheets.items()}
    st.session_state.loaded_signature = signature
    rebuild_views()

if "views" not in st.session_state:
    rebuild_views()

purchase_df = parse_purchase_file(purchase_file) if purchase_file is not None else pd.DataFrame()
views = st.session_state.views
fusion = views["fusion"].copy()
promo_all = views["promo"].copy()
maestra_all = views["maestra"].copy()
mlcsku_all = views["mlcsku"].copy()
promo_exp_by_sku = views["promo_exp_by_sku"].copy()
sku_mlc_map = views["sku_mlc_map"].copy()

# Search controls
left, right = st.columns([2, 1])
with left:
    search = st.text_input("Buscar por SKU, descripción o MLC")
with right:
    search_mode = st.selectbox("Modo", ["contiene", "empieza con"], index=0)

if search:
    token = search.strip().upper()
    mask = pd.Series(False, index=fusion.index)
    for col in ["SKU", "DESCRIPCIÓN", "MLCs"]:
        if col in fusion.columns:
            ser = fusion[col].astype(str).str.upper().fillna("")
            if search_mode == "contiene":
                mask = mask | ser.str.contains(re.escape(token), regex=True, na=False)
            else:
                mask = mask | ser.str.startswith(token, na=False)
    search_results = fusion.loc[mask].copy()
else:
    search_results = fusion.copy()

# Summary cards
sum_cols = st.columns(4)
with sum_cols[0]:
    st.metric("SKU en maestra", len(fusion))
with sum_cols[1]:
    vence_hoy = int((pd.to_datetime(fusion.get("next_campaign_date"), errors="coerce").dt.normalize() == pd.Timestamp(date.today())).sum()) if "next_campaign_date" in fusion.columns else 0
    st.metric("Promos vencen hoy", vence_hoy)
with sum_cols[2]:
    manana = pd.Timestamp(date.today() + timedelta(days=1))
    vence_manana = int((pd.to_datetime(fusion.get("next_campaign_date"), errors="coerce").dt.normalize() == manana).sum()) if "next_campaign_date" in fusion.columns else 0
    st.metric("Promos vencen mañana", vence_manana)
with sum_cols[3]:
    st.metric("Resultados búsqueda", len(search_results))

# ---------- tabs ----------
tab1, tab2, tab3 = st.tabs(["Ficha de decisión", "Operación de promos", "Edición total / alta"])

with tab1:
    st.subheader("Ficha de decisión por producto")
    base_cols = [c for c in ["SKU", "DESCRIPCIÓN", "MLCs", "next_campaign_date", "campaign_status", "total_promos"] if c in search_results.columns]
    if base_cols:
        st.dataframe(search_results[base_cols].head(50), use_container_width=True, hide_index=True)

    sku_options = search_results["SKU"].dropna().astype(str).tolist() if "SKU" in search_results.columns else []
    selected_sku = st.selectbox("Selecciona producto", options=sku_options, index=0 if sku_options else None)

    if selected_sku:
        sku_norm = normalize_sku(selected_sku)
        product = fusion[fusion["SKU_norm"] == sku_norm].head(1)
        if product.empty:
            st.warning("No se encontró el producto.")
        else:
            p = product.iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("SKU", sanitize_numeric_text(p.get("SKU")) or "-")
            c2.metric("Ubicación", sanitize_numeric_text(p.get("UBIC")) or "-")
            c3.metric("Estado promo", p.get("campaign_status") or "Sin datos")
            c4.metric("Promos asociadas", int(p.get("total_promos", 0) or 0))
            st.markdown(f"### {p.get('DESCRIPCIÓN', '-')}")

            l1, l2, l3 = st.columns([1.2, 1.2, 1.6])
            with l1:
                st.markdown("#### Precio base")
                st.write({
                    "Último costo": p.get("ÚLTIMO COSTO"),
                    "Precio neto": p.get("PRECIO NETO"),
                    "Precio bruto": p.get("PRECIO BRUTO"),
                    "Precio B2C publicado": p.get("PRECIO B2C PUBLICADO "),
                    "Precio B2C 2": p.get("PRECIO B2C"),
                    "Margen local": p.get("MARGEN LOCAL"),
                    "Margen Meli 1": p.get("MARGEN MELI 1"),
                    "Margen Meli 2": p.get("MARGEN MELI 2"),
                })
            with l2:
                st.markdown("#### Lectura comercial")
                lectura = []
                if pd.notna(p.get("next_campaign_date")):
                    lectura.append(p.get("campaign_status"))
                if pd.notna(p.get("promo_vs_base_%")):
                    lectura.append(f"Mejor promo vs base: {p.get('promo_vs_base_%'):.1f}%")
                if p.get("ads_resumen"):
                    lectura.append(f"Ads/comentarios: {p.get('ads_resumen')}")
                if pd.notna(p.get("CAMBIO DE PRECIO")):
                    lectura.append(f"Cambio de precio sugerido: {p.get('CAMBIO DE PRECIO')}")
                if p.get("COMENTARIO"):
                    lectura.append(f"Comentario: {p.get('COMENTARIO')}")
                for item in lectura or ["Sin alertas generadas."]:
                    st.write(f"• {item}")
            with l3:
                st.markdown("#### Promos y campañas asociadas")
                promos = promo_exp_by_sku[promo_exp_by_sku["SKU_join"] == sku_norm].copy()
                if not promos.empty:
                    show_cols = [c for c in ["N° Publicación", "Precio promocional", "% F", "Motivo promoción", "Ads/Comentario", *PROMO_DATE_COLS] if c in promos.columns]
                    st.dataframe(clean_for_editor(promos[show_cols].drop_duplicates()), use_container_width=True, hide_index=True)
                else:
                    st.info("Sin promos asociadas.")

            if not purchase_df.empty:
                st.markdown("#### Historial de compras")
                pur = purchase_df[purchase_df["SKU_norm"] == sku_norm].copy()
                if not pur.empty:
                    pur = pur.sort_values("Fecha", ascending=False)
                    latest = pur.head(1).iloc[0]
                    p1, p2, p3, p4 = st.columns(4)
                    p1.metric("Última compra", latest.get("Fecha").date().isoformat() if pd.notna(latest.get("Fecha")) else "-")
                    p2.metric("Último precio compra", latest.get("Precio Un."))
                    p3.metric("Proveedor último", latest.get("Razón Social") or "-")
                    p4.metric("Compras históricas", len(pur))
                    hist_cols = [c for c in ["Fecha", "Razón Social", "Documento", "Folio", "Cantidad", "Precio Un.", "Total Línea"] if c in pur.columns]
                    st.dataframe(clean_for_editor(pur[hist_cols]), use_container_width=True, hide_index=True)
                else:
                    st.info("No hay historial de compras para este SKU en el archivo cargado.")

    st.markdown("---")
    st.subheader("Agenda comercial")
    agenda = fusion[[c for c in ["SKU", "DESCRIPCIÓN", "MLCs", "next_campaign_date", "campaign_status", "total_promos", "ads_resumen"] if c in fusion.columns]].copy()
    agenda["next_campaign_date"] = pd.to_datetime(agenda.get("next_campaign_date"), errors="coerce")
    today = pd.Timestamp(date.today())
    buckets = {
        "Vence hoy": agenda[agenda["next_campaign_date"].dt.normalize() == today],
        "Vence mañana": agenda[agenda["next_campaign_date"].dt.normalize() == today + pd.Timedelta(days=1)],
        "Vence pasado mañana": agenda[agenda["next_campaign_date"].dt.normalize() == today + pd.Timedelta(days=2)],
        "Próximos 7 días": agenda[(agenda["next_campaign_date"].dt.normalize() > today + pd.Timedelta(days=2)) & (agenda["next_campaign_date"].dt.normalize() <= today + pd.Timedelta(days=7))],
    }
    a1, a2 = st.columns(2)
    for i, (title, dfb) in enumerate(buckets.items()):
        with (a1 if i % 2 == 0 else a2):
            st.markdown(f"##### {title}")
            st.dataframe(clean_for_editor(dfb), use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Operación de promos")
    sku_ops = st.text_input("SKU a operar", value=(selected_sku if 'selected_sku' in locals() and selected_sku else ""), key="sku_ops")
    sku_norm_ops = normalize_sku(sku_ops)
    if sku_norm_ops:
        maestra_row = maestra_all[maestra_all["SKU_norm"] == sku_norm_ops].copy()
        promo_rows = promo_all[promo_all["SKU_norm"] == sku_norm_ops].copy()
        # plus rows matched by mlc
        matched_mlcs = sku_mlc_map[sku_mlc_map["SKU_norm"] == sku_norm_ops]["MLC_norm"].dropna().tolist()
        if matched_mlcs:
            promo_rows = pd.concat([
                promo_rows,
                promo_all[promo_all["promo_mlcs"].apply(lambda xs: any(m in xs for m in matched_mlcs))]
            ], ignore_index=True).drop_duplicates(subset=["promo_row_id"])

        st.markdown("#### Resumen del producto")
        if not maestra_row.empty:
            st.dataframe(clean_for_editor(maestra_row[[c for c in CORE_MAESTRA_COLUMNS if c in maestra_row.columns]].head(1)), use_container_width=True, hide_index=True)
        else:
            st.warning("SKU no existe en MAESTRA de precios.")

        st.markdown("#### Editar promos y fechas")
        if promo_rows.empty:
            promo_rows = pd.DataFrame(columns=[c for c in PROMO_CORE_COLUMNS if c in promo_all.columns])
            if " " in promo_all.columns:
                promo_rows.loc[0, " "] = sku_norm_ops
        promo_edit = clean_for_editor(promo_rows[[c for c in promo_rows.columns if c in promo_all.columns]].copy())
        promo_edit = st.data_editor(
            promo_edit,
            use_container_width=True,
            num_rows="dynamic",
            key=f"promo_editor_{sku_norm_ops}",
            hide_index=True,
        )
        if st.button("Guardar promos del SKU", key=f"save_promos_{sku_norm_ops}"):
            original = st.session_state.sheets["CONTROL DE PROMOCIONES"].copy()
            original = normalize_colnames(original)
            if "promo_row_id" not in original.columns:
                original["promo_row_id"] = np.arange(len(original))
            old_ids = set(promo_rows.get("promo_row_id", pd.Series(dtype=int)).dropna().astype(int).tolist()) if "promo_row_id" in promo_rows.columns else set()
            remaining = original[~original["promo_row_id"].isin(old_ids)].copy()
            updated = promo_edit.copy()
            # ensure columns match original
            for c in original.columns:
                if c not in updated.columns:
                    updated[c] = np.nan
            updated = updated[original.columns]
            # assign ids to new rows
            max_id = int(original["promo_row_id"].max()) if len(original) else -1
            if "promo_row_id" in updated.columns:
                new_mask = pd.to_numeric(updated["promo_row_id"], errors="coerce").isna()
                count_new = int(new_mask.sum())
                if count_new:
                    updated.loc[new_mask, "promo_row_id"] = list(range(max_id + 1, max_id + 1 + count_new))
            st.session_state.sheets["CONTROL DE PROMOCIONES"] = pd.concat([remaining, updated], ignore_index=True).drop(columns=[c for c in ["promo_row_id", "SKU_norm", "promo_mlcs", "next_campaign_date", "campaign_status"] if c in remaining.columns or c in updated.columns], errors="ignore")
            rebuild_views()
            st.success("Promos actualizadas en la sesión.")
    else:
        st.info("Ingresa un SKU para editar promos.")

with tab3:
    st.subheader("Edición total / alta de producto")
    mode = st.radio("Modo", ["Editar SKU existente", "Crear SKU nuevo"], horizontal=True)
    if mode == "Editar SKU existente":
        sku_edit = st.text_input("SKU a editar", value=(selected_sku if 'selected_sku' in locals() and selected_sku else ""), key="sku_edit_total")
        sku_norm_edit = normalize_sku(sku_edit)
        if sku_norm_edit:
            maestra_rows = maestra_all[maestra_all["SKU_norm"] == sku_norm_edit].copy()
            st.markdown("#### MAESTRA de precios")
            if maestra_rows.empty:
                st.warning("No existe en la maestra.")
            else:
                maestra_view = clean_for_editor(maestra_rows.drop(columns=[c for c in ["SKU_norm"] if c in maestra_rows.columns]))
                maestra_edit = st.data_editor(maestra_view, use_container_width=True, num_rows="fixed", hide_index=True, key=f"maestra_edit_{sku_norm_edit}")
                if st.button("Guardar maestra del SKU", key=f"save_maestra_{sku_norm_edit}"):
                    original = normalize_colnames(st.session_state.sheets["MAESTRA de precios"].copy())
                    original["SKU_norm"] = original["SKU"].apply(normalize_sku)
                    remaining = original[original["SKU_norm"] != sku_norm_edit].drop(columns=["SKU_norm"])
                    st.session_state.sheets["MAESTRA de precios"] = pd.concat([remaining, maestra_edit], ignore_index=True)
                    rebuild_views()
                    st.success("Maestra actualizada.")

            st.markdown("#### MLC - SKU")
            map_rows = mlcsku_all[mlcsku_all["SKU_norm"] == sku_norm_edit].copy()
            map_view = clean_for_editor(map_rows.drop(columns=[c for c in ["SKU_norm", "MLC_norm"] if c in map_rows.columns]))
            map_edit = st.data_editor(map_view if not map_view.empty else pd.DataFrame(columns=[c for c in ["SKU", "Número de publicación"] if c in mlcsku_all.columns]), use_container_width=True, num_rows="dynamic", hide_index=True, key=f"map_edit_{sku_norm_edit}")
            if st.button("Guardar mapeo MLC", key=f"save_map_{sku_norm_edit}"):
                original = normalize_colnames(st.session_state.sheets["MLC -SKU"].copy())
                original["SKU_norm"] = original["SKU"].apply(normalize_sku)
                remaining = original[original["SKU_norm"] != sku_norm_edit].drop(columns=["SKU_norm"], errors="ignore")
                st.session_state.sheets["MLC -SKU"] = pd.concat([remaining, map_edit], ignore_index=True)
                rebuild_views()
                st.success("Mapeo actualizado.")
    else:
        st.markdown("#### Crear SKU nuevo")
        st.write("Crea el registro base en la maestra y, si quieres, agrega publicaciones y promos iniciales.")
        new_maestra = pd.DataFrame([{c: None for c in maestra_all.columns if c != "SKU_norm"}])
        if "SKU" in new_maestra.columns:
            new_maestra.loc[0, "SKU"] = ""
        new_maestra_edit = st.data_editor(clean_for_editor(new_maestra), use_container_width=True, num_rows="fixed", hide_index=True, key="new_maestra")
        new_map = pd.DataFrame([{c: None for c in ["SKU", "Número de publicación"] if c in mlcsku_all.columns}])
        new_map_edit = st.data_editor(clean_for_editor(new_map), use_container_width=True, num_rows="dynamic", hide_index=True, key="new_map")
        new_promo = pd.DataFrame([{c: None for c in [c for c in PROMO_CORE_COLUMNS if c in promo_all.columns]}])
        new_promo_edit = st.data_editor(clean_for_editor(new_promo), use_container_width=True, num_rows="dynamic", hide_index=True, key="new_promo")
        if st.button("Crear producto nuevo"):
            if new_maestra_edit.empty or pd.isna(new_maestra_edit.iloc[0].get("SKU")) or str(new_maestra_edit.iloc[0].get("SKU")).strip() == "":
                st.error("El SKU es obligatorio.")
            else:
                new_sku_norm = normalize_sku(new_maestra_edit.iloc[0].get("SKU"))
                existing = maestra_all[maestra_all["SKU_norm"] == new_sku_norm]
                if not existing.empty:
                    st.error("Ese SKU ya existe. Usa el modo editar.")
                else:
                    st.session_state.sheets["MAESTRA de precios"] = pd.concat([
                        st.session_state.sheets["MAESTRA de precios"],
                        new_maestra_edit,
                    ], ignore_index=True)
                    if not new_map_edit.empty:
                        st.session_state.sheets["MLC -SKU"] = pd.concat([
                            st.session_state.sheets["MLC -SKU"],
                            new_map_edit,
                        ], ignore_index=True)
                    promo_to_add = new_promo_edit.copy()
                    if " " in promo_to_add.columns:
                        promo_to_add[" "] = promo_to_add[" "].fillna(new_sku_norm)
                    if not promo_to_add.empty:
                        st.session_state.sheets["CONTROL DE PROMOCIONES"] = pd.concat([
                            st.session_state.sheets["CONTROL DE PROMOCIONES"],
                            promo_to_add,
                        ], ignore_index=True)
                    rebuild_views()
                    st.success("Producto creado en la sesión.")
