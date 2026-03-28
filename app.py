
import io
import hashlib
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st


# =========================================================
# Config
# =========================================================
st.set_page_config(page_title="Centro de Control Comercial Aurora", layout="wide")


# =========================================================
# Helpers
# =========================================================
VAT_RATE = 1.19


def file_signature(uploaded_file) -> str:
    data = uploaded_file.getvalue()
    return hashlib.md5(data).hexdigest()


def re_is_numberlike(s: str) -> bool:
    import re
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", str(s).strip()))


def norm_sku(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if re_is_numberlike(s):
        try:
            f = float(s.replace(",", "."))
            if float(int(f)) == f:
                return str(int(f))
        except Exception:
            pass
    return s


def norm_mlc(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    s = str(value).strip().upper()
    if not s or s == "NAN":
        return ""
    if s.endswith(".0") and re_is_numberlike(s):
        s = s[:-2]
    return s


def safe_float(value, default=np.nan):
    try:
        if value is None:
            return default
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return default
            value = text.replace("$", "").replace(".", "").replace(",", ".")
        out = float(value)
        return out
    except Exception:
        return default


def to_date_only(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return pd.NaT
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce").normalize()
    except Exception:
        return pd.NaT


def fmt_date(value) -> str:
    dt = to_date_only(value)
    if pd.isna(dt):
        return "—"
    return dt.strftime("%d/%m/%Y")


def fmt_money(value) -> str:
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return "—"
    return f"${x:,.0f}".replace(",", ".")


def fmt_int(value) -> str:
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return "—"
    return f"{int(round(x)):,}".replace(",", ".")


def margin_display(value) -> str:
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return "—"
    if abs(x) <= 2:
        x = x * 100
    return f"{x:.1f}%"


def calc_margin_from_neto(cost_neto, sale_neto):
    cost = safe_float(cost_neto, np.nan)
    sale = safe_float(sale_neto, np.nan)
    if np.isnan(cost) or np.isnan(sale) or sale == 0:
        return np.nan
    return ((sale - cost) / sale) * 100


def calc_margin_from_bruto(cost_neto, sale_bruto):
    bruto = safe_float(sale_bruto, np.nan)
    if np.isnan(bruto):
        return np.nan
    return calc_margin_from_neto(cost_neto, bruto / VAT_RATE)


def days_since(value):
    dt = to_date_only(value)
    if pd.isna(dt):
        return np.nan
    return int((pd.Timestamp(date.today()) - dt).days)


def first_existing(df: pd.DataFrame, candidates, default=np.nan):
    for col in candidates:
        if col in df.columns:
            return df[col]
    return pd.Series([default] * len(df), index=df.index)


def _find_sheet(sheet_names, wanted):
    wanted = wanted.lower().strip()
    for name in sheet_names:
        if name.lower().strip() == wanted:
            return name
    for name in sheet_names:
        if wanted in name.lower().strip():
            return name
    return None


def promo_status(dt):
    dt = to_date_only(dt)
    if pd.isna(dt):
        return "Sin fecha", 999
    today = pd.Timestamp(date.today())
    delta = int((dt - today).days)
    if delta < 0:
        return "Vencida", -1
    if delta == 0:
        return "Vence hoy", 0
    if delta == 1:
        return "Vence mañana", 1
    if delta <= 7:
        return "Vence en 7 días", 7
    if delta <= 15:
        return "Vence en 15 días", 15
    return "Vigente", 30


def parse_multi_mlc(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return []
    text = str(value).strip()
    if not text:
        return []
    text = text.replace("–", "-").replace("—", "-")
    parts = []
    for token in text.replace("/", "-").split("-"):
        token = token.strip()
        if token:
            parts.append(norm_mlc(token))
    return [p for p in parts if p]


# =========================================================
# Loaders
# =========================================================
@st.cache_data(show_spinner=False)
def load_workbook_cached(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xls.sheet_names

    maestra_name = _find_sheet(names, "MAESTRA de precios")
    bridge_name = _find_sheet(names, "MLC -SKU")
    rel_name = _find_sheet(names, "Relampago mi pagina")
    control_name = _find_sheet(names, "CONTROL DE PROMOCIONES")

    if not maestra_name:
        raise ValueError(f"No encontré la hoja MAESTRA de precios. Disponibles: {', '.join(names)}")

    master_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=maestra_name)
    bridge_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=bridge_name) if bridge_name else pd.DataFrame()
    rel_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=rel_name, header=None) if rel_name else pd.DataFrame()
    control_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=control_name) if control_name else pd.DataFrame()

    return {
        "sheet_names": names,
        "maestra_name": maestra_name,
        "bridge_name": bridge_name,
        "rel_name": rel_name,
        "control_name": control_name,
        "master_df": master_df,
        "bridge_df": bridge_df,
        "rel_df": rel_df,
        "control_df": control_df,
        "file_bytes": file_bytes,
    }


@st.cache_data(show_spinner=False)
def load_purchases_cached(file_bytes: bytes):
    if not file_bytes:
        return {"raw": pd.DataFrame(), "summary": pd.DataFrame(), "by_sku": {}}

    raw = pd.read_excel(io.BytesIO(file_bytes))
    if raw.empty:
        return {"raw": pd.DataFrame(), "summary": pd.DataFrame(), "by_sku": {}}

    raw = raw.copy()
    for col in ["SKU", "Fecha", "Razón Social", "Precio Un.", "Cantidad"]:
        if col not in raw.columns:
            raw[col] = np.nan

    raw["SKU_norm"] = raw["SKU"].map(norm_sku)
    raw["Fecha_dt"] = pd.to_datetime(raw["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
    raw["Precio_Un_Neto"] = raw["Precio Un."].map(lambda x: safe_float(x, np.nan))
    raw["Cantidad_num"] = raw["Cantidad"].map(lambda x: safe_float(x, np.nan))
    raw = raw[raw["SKU_norm"] != ""].copy()
    raw = raw.sort_values(["SKU_norm", "Fecha_dt"])

    if raw.empty:
        return {"raw": raw, "summary": pd.DataFrame(), "by_sku": {}}

    by_sku = {sku: grp.copy() for sku, grp in raw.groupby("SKU_norm", sort=False)}

    summary_rows = []
    for sku, grp in by_sku.items():
        grp = grp.sort_values("Fecha_dt")
        last = grp.iloc[-1]
        prev_price = safe_float(grp.iloc[-2]["Precio_Un_Neto"], np.nan) if len(grp) >= 2 else np.nan
        last_price = safe_float(last["Precio_Un_Neto"], np.nan)
        variation = np.nan
        if not np.isnan(last_price) and not np.isnan(prev_price) and prev_price != 0:
            variation = ((last_price - prev_price) / prev_price) * 100

        summary_rows.append(
            {
                "SKU_norm": sku,
                "ultima_compra_fecha": last["Fecha_dt"],
                "costo_actual_neto": last_price,
                "ultimo_proveedor": last.get("Razón Social", ""),
                "ultima_cantidad": safe_float(last.get("Cantidad_num"), np.nan),
                "costo_anterior_neto": prev_price,
                "variacion_costo_pct": variation,
                "compras_total": len(grp),
                "dias_sin_compra": days_since(last["Fecha_dt"]),
                "costo_promedio_neto": grp["Precio_Un_Neto"].mean(),
                "costo_min_neto": grp["Precio_Un_Neto"].min(),
                "costo_max_neto": grp["Precio_Un_Neto"].max(),
            }
        )
    summary = pd.DataFrame(summary_rows)
    return {"raw": raw, "summary": summary, "by_sku": by_sku}


@st.cache_data(show_spinner=False)
def load_sales_cached(file_bytes: bytes):
    if not file_bytes:
        return {
            "raw": pd.DataFrame(),
            "summary": pd.DataFrame(),
            "summary_30d": pd.DataFrame(),
            "summary_90d": pd.DataFrame(),
            "by_sku": {},
        }

    raw = pd.read_excel(io.BytesIO(file_bytes))
    if raw.empty:
        return {
            "raw": raw,
            "summary": pd.DataFrame(),
            "summary_30d": pd.DataFrame(),
            "summary_90d": pd.DataFrame(),
            "by_sku": {},
        }

    raw = raw.copy()
    for col in ["SKU", "Fecha", "Cantidad", "Precio Un.", "Total Línea", "Familia", "Producto"]:
        if col not in raw.columns:
            raw[col] = np.nan

    raw["SKU_norm"] = raw["SKU"].map(norm_sku)
    raw["Fecha_dt"] = pd.to_datetime(raw["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
    raw["Cantidad_num"] = raw["Cantidad"].map(lambda x: safe_float(x, np.nan))
    raw["Precio_Un_Bruto"] = raw["Precio Un."].map(lambda x: safe_float(x, np.nan))
    # En este informe Total Línea viene neto
    raw["Total_Linea_Neto"] = raw["Total Línea"].map(lambda x: safe_float(x, np.nan))
    raw["Precio_Prom_Neto"] = raw["Total_Linea_Neto"] / raw["Cantidad_num"].replace(0, np.nan)
    raw = raw[raw["SKU_norm"] != ""].copy()

    by_sku = {sku: grp.copy() for sku, grp in raw.groupby("SKU_norm", sort=False)}

    def summarize(df_in: pd.DataFrame):
        if df_in.empty:
            return pd.DataFrame(columns=[
                "SKU_norm", "ventas_docs", "unidades_vendidas", "venta_neta_total",
                "precio_promedio_neto", "precio_promedio_bruto", "ultima_venta_fecha",
                "dias_sin_venta", "familia_venta", "producto_venta"
            ])
        grp = (
            df_in.groupby("SKU_norm", dropna=False)
            .agg(
                ventas_docs=("SKU_norm", "size"),
                unidades_vendidas=("Cantidad_num", "sum"),
                venta_neta_total=("Total_Linea_Neto", "sum"),
                precio_promedio_neto=("Precio_Prom_Neto", "mean"),
                precio_promedio_bruto=("Precio_Un_Bruto", "mean"),
                ultima_venta_fecha=("Fecha_dt", "max"),
                familia_venta=("Familia", "last"),
                producto_venta=("Producto", "last"),
            )
            .reset_index()
        )
        grp["dias_sin_venta"] = grp["ultima_venta_fecha"].map(days_since)
        return grp

    today = pd.Timestamp(date.today())
    sales_30d = raw[raw["Fecha_dt"] >= (today - pd.Timedelta(days=30))].copy()
    sales_90d = raw[raw["Fecha_dt"] >= (today - pd.Timedelta(days=90))].copy()

    return {
        "raw": raw,
        "summary": summarize(raw),
        "summary_30d": summarize(sales_30d),
        "summary_90d": summarize(sales_90d),
        "by_sku": by_sku,
    }


# =========================================================
# Normalizers
# =========================================================
def normalize_master(master_df: pd.DataFrame) -> pd.DataFrame:
    df = master_df.copy()

    for col in [
        "SKU", "DESCRIPCIÓN", "UBIC", "ÚLTIMO COSTO", "PRECIO NETO", "PRECIO BRUTO",
        "MARGEN LOCAL", "MARGEN MELI 1", " NETO MELI 1", "MONTO EN SIMULACIÓN",
        "CAMPAÑA PADS", "MLC", "MLC SINCRONIZADO", " DCTO", "PRECIO B2C PUBLICADO ",
        "FECHA VENCI", "COMENTARIO", "MARGEN MELI 2", "NETO MELI 2", "VENTA BRUTO MELI 2",
        "MLC.1", "MLC SINCRONIZADO.1", "CAMPAÑA PADS.1", " DCTO.1", "PRECIO B2C",
        "FECHA VENCI.1", "COMENTARIO.1", "CAMBIO DE PRECIO", "T-L"
    ]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU"].map(norm_sku)
    df["DESCRIPCION_clean"] = df["DESCRIPCIÓN"].fillna("").astype(str)
    df["UBIC_clean"] = df["UBIC"].fillna("")

    numeric_cols = [
        "ÚLTIMO COSTO", "PRECIO NETO", "PRECIO BRUTO", "MARGEN LOCAL", "MARGEN MELI 1",
        " NETO MELI 1", "MONTO EN SIMULACIÓN", " DCTO", "PRECIO B2C PUBLICADO ",
        "MARGEN MELI 2", "NETO MELI 2", "VENTA BRUTO MELI 2", " DCTO.1", "PRECIO B2C"
    ]
    for col in numeric_cols:
        df[col] = df[col].map(lambda x: safe_float(x, np.nan))

    for col in ["FECHA VENCI", "FECHA VENCI.1"]:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce").dt.normalize()

    df["MLC_1"] = df["MLC"].map(norm_mlc)
    df["MLC_SYNC_1"] = df["MLC SINCRONIZADO"].map(norm_mlc)
    df["MLC_2"] = df["MLC.1"].map(norm_mlc)
    df["MLC_SYNC_2"] = df["MLC SINCRONIZADO.1"].map(norm_mlc)

    df["ADS_1"] = df["CAMPAÑA PADS"].fillna("").astype(str).str.strip()
    df["ADS_2"] = df["CAMPAÑA PADS.1"].fillna("").astype(str).str.strip()
    df["tiene_ads_master"] = (df["ADS_1"] != "") | (df["ADS_2"] != "")

    return df[df["SKU_norm"] != ""].copy()


def normalize_bridge(bridge_df: pd.DataFrame) -> pd.DataFrame:
    if bridge_df is None or bridge_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "MLC_norm"])

    df = bridge_df.copy()
    sku_col = "SKU" if "SKU" in df.columns else df.columns[0]
    mlc_candidates = [c for c in df.columns if "publicación" in str(c).lower() or "publicacion" in str(c).lower()]
    mlc_col = mlc_candidates[0] if mlc_candidates else df.columns[-1]

    df["SKU_norm"] = df[sku_col].map(norm_sku)
    df["MLC_norm"] = df[mlc_col].map(norm_mlc)
    df = df[(df["SKU_norm"] != "") & (df["MLC_norm"] != "")].copy()
    return df[["SKU_norm", "MLC_norm"]].drop_duplicates()


def normalize_rel(rel_df: pd.DataFrame) -> pd.DataFrame:
    if rel_df is None or rel_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"])

    df = rel_df.copy()
    while df.shape[1] < 6:
        df[df.shape[1]] = np.nan
    df = df.iloc[:, :6]
    df.columns = ["SKU_raw", "DESCRIPCION", "PRECIO_B2C", "EXTRA", "TIPO", "ESTADO"]
    df["SKU_norm"] = df["SKU_raw"].map(norm_sku)
    df["PRECIO_B2C"] = df["PRECIO_B2C"].map(lambda x: safe_float(x, np.nan))
    df = df[df["SKU_norm"] != ""].copy()
    return df[["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"]]


def normalize_control_promos(control_df: pd.DataFrame) -> pd.DataFrame:
    if control_df is None or control_df.empty:
        return pd.DataFrame(columns=[
            "SKU_norm", "MLC", "Descripcion", "Pct_F", "Precio_promocional",
            "Motivo", "Ads_Comentario", "Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4",
            "Campaña_min", "Campaña_status", "Campaña_order"
        ])

    df = control_df.copy()

    mapping = {}
    for col in df.columns:
        low = str(col).strip().lower()
        if low == "":
            mapping[col] = "SKU_raw"
        elif "publicación" in low or "publicacion" in low:
            mapping[col] = "N_Publicacion"
        elif low == "descripción" or low == "descripcion":
            mapping[col] = "Descripcion"
        elif low == "% f":
            mapping[col] = "Pct_F"
        elif "precio promocional" in low:
            mapping[col] = "Precio_promocional"
        elif "motivo" in low:
            mapping[col] = "Motivo"
        elif "ads/comentario" in low:
            mapping[col] = "Ads_Comentario"
        elif low == "campaña 1":
            mapping[col] = "Campaña_1"
        elif low == "campaña 2":
            mapping[col] = "Campaña_2"
        elif low == "campaña 3":
            mapping[col] = "Campaña_3"
        elif low == "campaña 4":
            mapping[col] = "Campaña_4"

    df = df.rename(columns=mapping)
    for col in [
        "SKU_raw", "N_Publicacion", "Descripcion", "Pct_F", "Precio_promocional",
        "Motivo", "Ads_Comentario", "Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"
    ]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU_raw"].map(norm_sku)
    df["MLC"] = df["N_Publicacion"].fillna("").astype(str)
    df["Precio_promocional"] = df["Precio_promocional"].map(lambda x: safe_float(x, np.nan))
    df["Pct_F"] = df["Pct_F"].map(lambda x: safe_float(x, np.nan))
    df["Ads_Comentario"] = df["Ads_Comentario"].fillna("").astype(str).str.strip()
    for c in ["Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"]:
        df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce").dt.normalize()

    df["Campaña_min"] = df[["Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"]].min(axis=1)
    status_info = df["Campaña_min"].map(lambda x: promo_status(x))
    df["Campaña_status"] = status_info.map(lambda x: x[0] if isinstance(x, tuple) else "Sin fecha")
    df["Campaña_order"] = status_info.map(lambda x: x[1] if isinstance(x, tuple) else 999)
    df["tiene_ads_control"] = df["Ads_Comentario"] != ""

    return df[df["SKU_norm"] != ""].copy()


# =========================================================
# Model builders
# =========================================================
def build_price_events(master: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in master.iterrows():
        sku = row["SKU_norm"]
        desc = row["DESCRIPCION_clean"]
        rows.append({
            "SKU_norm": sku,
            "fecha": pd.Timestamp(date.today()),
            "tipo": "snapshot_precio",
            "detalle": "Precio actual tienda",
            "valor_neto": safe_float(row.get("PRECIO NETO"), np.nan),
            "valor_bruto": safe_float(row.get("PRECIO BRUTO"), np.nan),
            "extra": desc,
        })
        if not pd.isna(row.get("PRECIO B2C PUBLICADO ")):
            rows.append({
                "SKU_norm": sku,
                "fecha": to_date_only(row.get("FECHA VENCI")) if pd.notna(row.get("FECHA VENCI")) else pd.Timestamp(date.today()),
                "tipo": "promo_master_1",
                "detalle": "Promo / precio B2C Meli 1",
                "valor_neto": safe_float(row.get("PRECIO B2C PUBLICADO "), np.nan) / VAT_RATE,
                "valor_bruto": safe_float(row.get("PRECIO B2C PUBLICADO "), np.nan),
                "extra": row.get("COMENTARIO", ""),
            })
        if not pd.isna(row.get("PRECIO B2C")):
            rows.append({
                "SKU_norm": sku,
                "fecha": to_date_only(row.get("FECHA VENCI.1")) if pd.notna(row.get("FECHA VENCI.1")) else pd.Timestamp(date.today()),
                "tipo": "promo_master_2",
                "detalle": "Promo / precio B2C Meli 2",
                "valor_neto": safe_float(row.get("PRECIO B2C"), np.nan) / VAT_RATE,
                "valor_bruto": safe_float(row.get("PRECIO B2C"), np.nan),
                "extra": row.get("COMENTARIO.1", ""),
            })
    return pd.DataFrame(rows)


def build_promos_unified(master: pd.DataFrame, control_promos: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for idx, row in master.iterrows():
        for slot, mlc_col, ads_col, price_col, dt_col, comment_col, dcto_col in [
            (1, "MLC_1", "ADS_1", "PRECIO B2C PUBLICADO ", "FECHA VENCI", "COMENTARIO", " DCTO"),
            (2, "MLC_2", "ADS_2", "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1", " DCTO.1"),
        ]:
            mlc = row.get(mlc_col, "")
            promo_price = safe_float(row.get(price_col), np.nan)
            promo_dt = row.get(dt_col)
            comment = row.get(comment_col, "")
            ads = row.get(ads_col, "")
            dcto = safe_float(row.get(dcto_col), np.nan)

            if mlc or not np.isnan(promo_price) or pd.notna(promo_dt) or str(comment).strip() or str(ads).strip():
                status, order = promo_status(promo_dt)
                rows.append({
                    "origen": "maestra",
                    "master_index": idx,
                    "SKU_norm": row["SKU_norm"],
                    "DESCRIPCION": row["DESCRIPCION_clean"],
                    "slot": slot,
                    "MLC": mlc,
                    "Precio_promocional_bruto": promo_price,
                    "Fecha_evento": promo_dt,
                    "Comentario": comment,
                    "Ads": ads,
                    "Descuento_pct": dcto,
                    "Motivo": "",
                    "STATUS": status,
                    "STATUS_ORDER": order,
                })

    if control_promos is not None and not control_promos.empty:
        for idx, row in control_promos.iterrows():
            rows.append({
                "origen": "control_promos",
                "master_index": idx,
                "SKU_norm": row["SKU_norm"],
                "DESCRIPCION": row.get("Descripcion", ""),
                "slot": "CP",
                "MLC": row.get("MLC", ""),
                "Precio_promocional_bruto": safe_float(row.get("Precio_promocional"), np.nan),
                "Fecha_evento": row.get("Campaña_min"),
                "Comentario": row.get("Ads_Comentario", ""),
                "Ads": row.get("Ads_Comentario", ""),
                "Descuento_pct": safe_float(row.get("Pct_F"), np.nan),
                "Motivo": row.get("Motivo", ""),
                "STATUS": row.get("Campaña_status", "Sin fecha"),
                "STATUS_ORDER": row.get("Campaña_order", 999),
            })

    promos = pd.DataFrame(rows)
    if promos.empty:
        promos = pd.DataFrame(columns=[
            "origen", "master_index", "SKU_norm", "DESCRIPCION", "slot", "MLC",
            "Precio_promocional_bruto", "Fecha_evento", "Comentario", "Ads",
            "Descuento_pct", "Motivo", "STATUS", "STATUS_ORDER"
        ])
    return promos


def build_mlc_map(master: pd.DataFrame, bridge: pd.DataFrame, control_promos: pd.DataFrame):
    mlc_map = {}
    for _, row in master.iterrows():
        sku = row["SKU_norm"]
        mlcs = [
            row.get("MLC_1", ""), row.get("MLC_SYNC_1", ""),
            row.get("MLC_2", ""), row.get("MLC_SYNC_2", "")
        ]
        mlc_map[sku] = sorted({m for m in mlcs if m})

    if bridge is not None and not bridge.empty:
        for sku, grp in bridge.groupby("SKU_norm"):
            mlc_map.setdefault(sku, [])
            mlc_map[sku] = sorted(set(mlc_map[sku]).union(set(grp["MLC_norm"].tolist())))

    if control_promos is not None and not control_promos.empty:
        for _, row in control_promos.iterrows():
            sku = row["SKU_norm"]
            more = parse_multi_mlc(row.get("MLC", ""))
            mlc_map.setdefault(sku, [])
            mlc_map[sku] = sorted(set(mlc_map[sku]).union(set(more)))

    return mlc_map


def build_product_base(master: pd.DataFrame, purchases: dict, sales: dict, rel: pd.DataFrame, control_promos: pd.DataFrame, bridge: pd.DataFrame):
    purchases_summary = purchases["summary"].copy()
    sales_summary = sales["summary"].copy()
    sales_summary_30d = sales["summary_30d"].copy()
    sales_summary_90d = sales["summary_90d"].copy()

    base = master.copy()

    if not purchases_summary.empty:
        base = base.merge(purchases_summary, on="SKU_norm", how="left")
    if not sales_summary.empty:
        base = base.merge(
            sales_summary.add_suffix("_hist").rename(columns={"SKU_norm_hist": "SKU_norm"}),
            on="SKU_norm",
            how="left",
        )
    if not sales_summary_30d.empty:
        base = base.merge(
            sales_summary_30d.add_suffix("_30d").rename(columns={"SKU_norm_30d": "SKU_norm"}),
            on="SKU_norm",
            how="left",
        )
    if not sales_summary_90d.empty:
        base = base.merge(
            sales_summary_90d.add_suffix("_90d").rename(columns={"SKU_norm_90d": "SKU_norm"}),
            on="SKU_norm",
            how="left",
        )

    if rel is not None and not rel.empty:
        rel_flag = rel[["SKU_norm"]].drop_duplicates().copy()
        rel_flag["en_relampago"] = True
        base = base.merge(rel_flag, on="SKU_norm", how="left")
    else:
        base["en_relampago"] = False

    control_ads = pd.DataFrame()
    if control_promos is not None and not control_promos.empty:
        control_ads = control_promos.groupby("SKU_norm").agg(
            tiene_ads_control=("tiene_ads_control", "max"),
            promos_control=("SKU_norm", "size"),
            primera_campania_control=("Campaña_min", "min"),
        ).reset_index()
        base = base.merge(control_ads, on="SKU_norm", how="left")
    else:
        base["tiene_ads_control"] = False
        base["promos_control"] = 0
        base["primera_campania_control"] = pd.NaT

    base["en_relampago"] = base["en_relampago"].fillna(False)
    base["tiene_ads_control"] = base["tiene_ads_control"].fillna(False)
    base["promos_control"] = base["promos_control"].fillna(0)
    base["tiene_ads"] = base["tiene_ads_master"].fillna(False) | base["tiene_ads_control"].fillna(False)

    base["costo_base_neto"] = base["costo_actual_neto"].fillna(base["ÚLTIMO COSTO"])
    base["margen_local_actual_pct"] = base.apply(
        lambda r: calc_margin_from_bruto(r["costo_base_neto"], r["PRECIO BRUTO"]),
        axis=1,
    )
    base["margen_meli1_actual_pct"] = base.apply(
        lambda r: calc_margin_from_bruto(r["costo_base_neto"], r["MONTO EN SIMULACIÓN"]),
        axis=1,
    )
    base["margen_real_30d_pct"] = base.apply(
        lambda r: calc_margin_from_neto(r["costo_base_neto"], r.get("precio_promedio_neto_30d", np.nan)),
        axis=1,
    )
    base["margen_real_90d_pct"] = base.apply(
        lambda r: calc_margin_from_neto(r["costo_base_neto"], r.get("precio_promedio_neto_90d", np.nan)),
        axis=1,
    )

    base["dias_sin_ultima_compra"] = base["ultima_compra_fecha"].map(days_since)
    base["dias_sin_ultima_venta"] = base["ultima_venta_fecha_hist"].map(days_since)

    base["alza_fuerte_costo"] = base["variacion_costo_pct"] >= 15
    base["alza_critica_costo"] = base["variacion_costo_pct"] >= 20
    base["sin_ventas_30d"] = base["unidades_vendidas_30d"].fillna(0) <= 0
    base["sin_compra_90d"] = base["dias_sin_ultima_compra"] >= 90
    base["promo_vigente_master"] = (
        (base["FECHA VENCI"].map(lambda x: promo_status(x)[1]) <= 15)
        | (base["FECHA VENCI.1"].map(lambda x: promo_status(x)[1]) <= 15)
    )

    def classify(row):
        ventas_30 = safe_float(row.get("unidades_vendidas_30d"), 0)
        margen_real = safe_float(row.get("margen_real_30d_pct"), np.nan)
        margen_local = safe_float(row.get("margen_local_actual_pct"), np.nan)
        alza = safe_float(row.get("variacion_costo_pct"), np.nan)
        ads = bool(row.get("tiene_ads"))
        promo = bool(row.get("promo_vigente_master")) or safe_float(row.get("promos_control"), 0) > 0

        if ventas_30 > 0 and not np.isnan(margen_real) and margen_real < 0:
            return "CRÍTICO", "Vendes con margen real negativo"
        if ads and ventas_30 <= 0:
            return "CRÍTICO", "Tiene Ads pero no registra ventas 30d"
        if not np.isnan(alza) and alza >= 20 and ((not np.isnan(margen_real) and margen_real < 15) or (not np.isnan(margen_local) and margen_local < 20)):
            return "CRÍTICO", "Alza fuerte de costo con margen apretado"
        if promo and not np.isnan(margen_real) and margen_real < 10 and ventas_30 > 0:
            return "CRÍTICO", "Promo activa con margen real muy bajo"
        if ventas_30 > 0 and not np.isnan(alza) and alza >= 15:
            return "ALERTA", "Vende y su costo viene subiendo"
        if ventas_30 <= 0 and safe_float(row.get("venta_neta_total_90d"), 0) <= 0:
            return "MUERTO", "Sin ventas en 90 días"
        if ventas_30 > 0 and not np.isnan(margen_real) and margen_real >= 20 and not ads:
            return "OPORTUNIDAD", "Buen margen real y venta sin Ads"
        if ventas_30 > 0 and not np.isnan(margen_real) and margen_real >= 15:
            return "SANO", "Producto estable"
        return "MONITOREAR", "Revisar comportamiento"

    states = base.apply(classify, axis=1, result_type="expand")
    base["estado_comercial"] = states[0]
    base["diagnostico"] = states[1]

    def score_priority(row):
        score = 0.0
        ventas_30 = safe_float(row.get("venta_neta_total_30d"), 0)
        ventas_scaled = min(100, ventas_30 / 10000)
        score += ventas_scaled
        if row.get("estado_comercial") == "CRÍTICO":
            score += 120
        elif row.get("estado_comercial") == "ALERTA":
            score += 80
        elif row.get("estado_comercial") == "OPORTUNIDAD":
            score += 60
        elif row.get("estado_comercial") == "MUERTO":
            score += 20
        if bool(row.get("tiene_ads")):
            score += 20
        if bool(row.get("promo_vigente_master")) or safe_float(row.get("promos_control"), 0) > 0:
            score += 15
        alza = safe_float(row.get("variacion_costo_pct"), 0)
        if not np.isnan(alza):
            score += max(0, alza)
        margen_real = safe_float(row.get("margen_real_30d_pct"), np.nan)
        if not np.isnan(margen_real):
            score += max(0, 20 - margen_real)
        return round(score, 2)

    base["priority_score"] = base.apply(score_priority, axis=1)

    mlc_map = build_mlc_map(master, bridge, control_promos)

    return {
        "products": base,
        "mlc_map": mlc_map,
        "purchase_map": purchases["by_sku"],
        "sales_map": sales["by_sku"],
        "sales_raw": sales["raw"],
        "purchases_raw": purchases["raw"],
        "price_events": build_price_events(master),
    }


def build_timeline_for_sku(model, sku: str) -> pd.DataFrame:
    events = []

    purchases_df = model["purchase_map"].get(sku, pd.DataFrame())
    if purchases_df is not None and not purchases_df.empty:
        for _, row in purchases_df.iterrows():
            events.append({
                "fecha": row.get("Fecha_dt"),
                "tipo": "compra",
                "detalle": f"Compra a {row.get('Razón Social', '—')}",
                "cantidad": safe_float(row.get("Cantidad_num"), np.nan),
                "valor_neto": safe_float(row.get("Precio_Un_Neto"), np.nan),
                "valor_bruto": np.nan,
                "extra": "Costo neto unitario",
            })

    sales_df = model["sales_map"].get(sku, pd.DataFrame())
    if sales_df is not None and not sales_df.empty:
        grouped_sales = (
            sales_df.groupby("Fecha_dt", dropna=False)
            .agg(
                cantidad=("Cantidad_num", "sum"),
                venta_neta=("Total_Linea_Neto", "sum"),
                precio_prom_neto=("Precio_Prom_Neto", "mean"),
            )
            .reset_index()
        )
        for _, row in grouped_sales.iterrows():
            events.append({
                "fecha": row.get("Fecha_dt"),
                "tipo": "venta",
                "detalle": "Venta del día",
                "cantidad": safe_float(row.get("cantidad"), np.nan),
                "valor_neto": safe_float(row.get("precio_prom_neto"), np.nan),
                "valor_bruto": np.nan,
                "extra": f"Venta neta del día: {fmt_money(row.get('venta_neta'))}",
            })

    price_events = model["price_events"]
    if price_events is not None and not price_events.empty:
        sub = price_events[price_events["SKU_norm"] == sku].copy()
        for _, row in sub.iterrows():
            events.append({
                "fecha": row.get("fecha"),
                "tipo": row.get("tipo"),
                "detalle": row.get("detalle"),
                "cantidad": np.nan,
                "valor_neto": row.get("valor_neto"),
                "valor_bruto": row.get("valor_bruto"),
                "extra": row.get("extra", ""),
            })

    promos = model["promos"]
    if promos is not None and not promos.empty:
        sub = promos[promos["SKU_norm"] == sku].copy()
        for _, row in sub.iterrows():
            events.append({
                "fecha": row.get("Fecha_evento"),
                "tipo": f"promo_{row.get('origen')}",
                "detalle": f"Promo {row.get('slot')} / {row.get('STATUS')}",
                "cantidad": np.nan,
                "valor_neto": safe_float(row.get("Precio_promocional_bruto"), np.nan) / VAT_RATE,
                "valor_bruto": safe_float(row.get("Precio_promocional_bruto"), np.nan),
                "extra": row.get("Comentario", "") or row.get("Motivo", ""),
            })

    timeline = pd.DataFrame(events)
    if timeline.empty:
        return pd.DataFrame(columns=["fecha", "tipo", "detalle", "cantidad", "valor_neto", "valor_bruto", "extra"])
    timeline["fecha"] = pd.to_datetime(timeline["fecha"], errors="coerce")
    timeline = timeline.sort_values(["fecha", "tipo"], ascending=[False, True]).reset_index(drop=True)
    return timeline


# =========================================================
# Session init
# =========================================================
def build_model_from_uploads(master_up, purchases_up, sales_up):
    wb = load_workbook_cached(master_up.getvalue())
    purchases = load_purchases_cached(purchases_up.getvalue() if purchases_up else b"")
    sales = load_sales_cached(sales_up.getvalue() if sales_up else b"")

    master = normalize_master(wb["master_df"])
    bridge = normalize_bridge(wb["bridge_df"])
    rel = normalize_rel(wb["rel_df"])
    control_promos = normalize_control_promos(wb["control_df"])
    promos = build_promos_unified(master, control_promos)
    bundle = build_product_base(master, purchases, sales, rel, control_promos, bridge)

    sku_desc = {row["SKU_norm"]: row["DESCRIPCION_clean"] for _, row in master.iterrows()}
    sku_options = sorted(master["SKU_norm"].dropna().unique().tolist())

    return {
        "master": master,
        "bridge": bridge,
        "rel": rel,
        "control_promos": control_promos,
        "promos": promos,
        "products": bundle["products"],
        "mlc_map": bundle["mlc_map"],
        "purchase_map": bundle["purchase_map"],
        "sales_map": bundle["sales_map"],
        "sales_raw": bundle["sales_raw"],
        "purchases_raw": bundle["purchases_raw"],
        "price_events": bundle["price_events"],
        "sku_desc": sku_desc,
        "sku_options": sku_options,
        "workbook_meta": wb,
    }


def init_state_from_upload(master_up, purchases_up, sales_up):
    master_sig = file_signature(master_up)
    purchases_sig = file_signature(purchases_up) if purchases_up else ""
    sales_sig = file_signature(sales_up) if sales_up else ""
    combined = f"{master_sig}|{purchases_sig}|{sales_sig}"

    if st.session_state.get("model_sig") == combined:
        return

    st.session_state.model = build_model_from_uploads(master_up, purchases_up, sales_up)
    st.session_state.model_sig = combined
    st.session_state.download_bytes = None


# =========================================================
# Download writers
# =========================================================
def rel_to_sheet_df(rel_df: pd.DataFrame) -> pd.DataFrame:
    if rel_df is None or rel_df.empty:
        return pd.DataFrame(columns=list(range(6)))
    out = pd.DataFrame({
        0: rel_df["SKU_norm"],
        1: rel_df["DESCRIPCION"],
        2: rel_df["PRECIO_B2C"],
        3: np.nan,
        4: rel_df["TIPO"],
        5: rel_df["ESTADO"],
    })
    return out


def control_promos_to_sheet_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            " ", "% F", "N° Publicación", "Descripción", "% F.1", "Precio promocional",
            "Motivo promoción", "Unnamed: 7", "margen", "Ads/Comentario",
            "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"
        ])
    out = pd.DataFrame({
        " ": df["SKU_norm"],
        "% F": df["Pct_F"],
        "N° Publicación": df["MLC"],
        "Descripción": df["Descripcion"],
        "% F.1": df["Pct_F"],
        "Precio promocional": df["Precio_promocional"],
        "Motivo promoción": df["Motivo"],
        "Unnamed: 7": np.nan,
        "margen": np.nan,
        "Ads/Comentario": df["Ads_Comentario"],
        "Campaña 1": df["Campaña_1"],
        "Campaña 2": df["Campaña_2"],
        "Campaña 3": df["Campaña_3"],
        "Campaña 4": df["Campaña_4"],
    })
    return out


@st.cache_data(show_spinner=False)
def build_download_bytes(master_df: pd.DataFrame, rel_df: pd.DataFrame, control_df: pd.DataFrame, original_bytes: bytes, maestra_name: str, rel_name: str, control_name: str):
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            if sheet == maestra_name:
                clean = master_df.copy()
                drop_cols = [c for c in clean.columns if c.endswith("_clean") or c in [
                    "SKU_norm", "MLC_1", "MLC_SYNC_1", "MLC_2", "MLC_SYNC_2", "ADS_1", "ADS_2", "tiene_ads_master"
                ]]
                clean = clean.drop(columns=[c for c in drop_cols if c in clean.columns], errors="ignore")
                clean.to_excel(writer, sheet_name=sheet, index=False)
            elif rel_name and sheet == rel_name:
                rel_to_sheet_df(rel_df).to_excel(writer, sheet_name=sheet, index=False, header=False)
            elif control_name and sheet == control_name:
                control_promos_to_sheet_df(control_df).to_excel(writer, sheet_name=sheet, index=False)
            else:
                original = pd.read_excel(io.BytesIO(original_bytes), sheet_name=sheet, header=None if rel_name and sheet == rel_name else 0)
                original.to_excel(writer, sheet_name=sheet, index=False, header=not (rel_name and sheet == rel_name))
    return out.getvalue()


# =========================================================
# UI helpers
# =========================================================
def state_badge(state: str) -> str:
    mapping = {
        "CRÍTICO": "🔴",
        "ALERTA": "🟠",
        "OPORTUNIDAD": "🟢",
        "SANO": "🟢",
        "MUERTO": "⚫",
        "MONITOREAR": "🔵",
    }
    return f"{mapping.get(state, '🔵')} {state}"


def get_selected_row(products: pd.DataFrame, sku: str):
    sub = products[products["SKU_norm"] == sku]
    if sub.empty:
        return None
    return sub.iloc[0]


# =========================================================
# UI
# =========================================================
st.title("Centro de Control Comercial Aurora")
st.caption("Control comercial integrado: costos, ventas, precios, promos, Ads y trazabilidad por SKU.")

with st.sidebar:
    st.subheader("Archivos")
    master_up = st.file_uploader("Maestra de precios y promos", type=["xlsx"], key="master")
    purchases_up = st.file_uploader("Compras", type=["xlsx"], key="purchases")
    sales_up = st.file_uploader("Ventas", type=["xlsx"], key="sales")
    st.caption("La app soporta la maestra actualizada, control de promociones, relámpago, compras y ventas.")

if not master_up:
    st.info("Sube la maestra actualizada para comenzar.")
    st.stop()

try:
    init_state_from_upload(master_up, purchases_up, sales_up)
except Exception as e:
    st.error(f"No pude construir el modelo comercial: {e}")
    st.stop()

model = st.session_state.model
products = model["products"].copy()

search_text = st.text_input("Buscar SKU / descripción / MLC")
if search_text:
    q = search_text.lower().strip()
    products = products[
        products["SKU_norm"].astype(str).str.lower().str.contains(q, na=False)
        | products["DESCRIPCION_clean"].astype(str).str.lower().str.contains(q, na=False)
        | products["SKU_norm"].map(lambda sku: " ".join(model["mlc_map"].get(sku, []))).str.lower().str.contains(q, na=False)
    ]

product_options = [f"{sku} — {model['sku_desc'].get(sku, '')}" for sku in products["SKU_norm"].dropna().unique().tolist()]
selected_label = st.selectbox("Seleccionar producto", product_options, index=0 if product_options else None)
selected_sku = selected_label.split(" — ")[0] if selected_label else ""
row = get_selected_row(model["products"], selected_sku)

tabs = st.tabs([
    "Centro de Control Comercial",
    "Ficha de Producto",
    "Timeline",
    "Promociones",
    "Relámpago",
    "Descargar",
])

# =========================================================
# Tab 1: Centro de Control Comercial
# =========================================================
with tabs[0]:
    df = model["products"].copy()
    if df.empty:
        st.warning("No encontré productos.")
    else:
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("SKUs totales", fmt_int(len(df)))
        k2.metric("SKUs críticos", fmt_int((df["estado_comercial"] == "CRÍTICO").sum()))
        k3.metric("SKUs con alza fuerte", fmt_int(df["alza_fuerte_costo"].sum()))
        k4.metric("SKUs con Ads", fmt_int(df["tiene_ads"].sum()))
        k5.metric("Venta neta 30d", fmt_money(df["venta_neta_total_30d"].fillna(0).sum()))

        st.subheader("Radar ejecutivo")
        c1, c2, c3 = st.columns(3)
        with c1:
            critical = df[df["estado_comercial"] == "CRÍTICO"].sort_values("priority_score", ascending=False).head(12)
            st.markdown("**Críticos hoy**")
            if critical.empty:
                st.success("No hay críticos con la lógica actual.")
            else:
                show = critical[[
                    "SKU_norm", "DESCRIPCION_clean", "diagnostico", "variacion_costo_pct",
                    "margen_real_30d_pct", "venta_neta_total_30d"
                ]].copy()
                show.columns = ["SKU", "Descripción", "Diagnóstico", "Var. costo", "Margen real 30d", "Venta neta 30d"]
                show["Var. costo"] = show["Var. costo"].map(margin_display)
                show["Margen real 30d"] = show["Margen real 30d"].map(margin_display)
                show["Venta neta 30d"] = show["Venta neta 30d"].map(fmt_money)
                st.dataframe(show, use_container_width=True, hide_index=True, height=360)

        with c2:
            rises = df[df["alza_fuerte_costo"]].sort_values(["variacion_costo_pct", "venta_neta_total_30d"], ascending=[False, False]).head(12)
            st.markdown("**Alzas de costo**")
            if rises.empty:
                st.info("No hay alzas fuertes detectadas.")
            else:
                show = rises[[
                    "SKU_norm", "DESCRIPCION_clean", "costo_actual_neto", "costo_anterior_neto",
                    "variacion_costo_pct", "margen_real_30d_pct"
                ]].copy()
                show.columns = ["SKU", "Descripción", "Costo actual", "Costo anterior", "Var. costo", "Margen real 30d"]
                show["Costo actual"] = show["Costo actual"].map(fmt_money)
                show["Costo anterior"] = show["Costo anterior"].map(fmt_money)
                show["Var. costo"] = show["Var. costo"].map(margin_display)
                show["Margen real 30d"] = show["Margen real 30d"].map(margin_display)
                st.dataframe(show, use_container_width=True, hide_index=True, height=360)

        with c3:
            opp = df[df["estado_comercial"] == "OPORTUNIDAD"].sort_values("priority_score", ascending=False).head(12)
            st.markdown("**Oportunidades**")
            if opp.empty:
                st.info("Todavía no hay oportunidades claras.")
            else:
                show = opp[[
                    "SKU_norm", "DESCRIPCION_clean", "margen_real_30d_pct", "venta_neta_total_30d",
                    "tiene_ads", "diagnostico"
                ]].copy()
                show.columns = ["SKU", "Descripción", "Margen real 30d", "Venta neta 30d", "Ads", "Diagnóstico"]
                show["Margen real 30d"] = show["Margen real 30d"].map(margin_display)
                show["Venta neta 30d"] = show["Venta neta 30d"].map(fmt_money)
                show["Ads"] = show["Ads"].map(lambda x: "Sí" if bool(x) else "No")
                st.dataframe(show, use_container_width=True, hide_index=True, height=360)

        st.subheader("Prioridad comercial")
        filt1, filt2, filt3 = st.columns([1, 1, 2])
        with filt1:
            states = st.multiselect(
                "Estado comercial",
                options=sorted(df["estado_comercial"].dropna().unique().tolist()),
                default=["CRÍTICO", "ALERTA", "OPORTUNIDAD"] if set(["CRÍTICO", "ALERTA", "OPORTUNIDAD"]).intersection(set(df["estado_comercial"].unique())) else sorted(df["estado_comercial"].dropna().unique().tolist()),
            )
        with filt2:
            ads_filter = st.selectbox("Filtro Ads", ["Todos", "Solo con Ads", "Solo sin Ads"])
        with filt3:
            min_sales = st.number_input("Venta neta mínima 30d", min_value=0.0, value=0.0, step=10000.0)

        table = df.copy()
        if states:
            table = table[table["estado_comercial"].isin(states)]
        if ads_filter == "Solo con Ads":
            table = table[table["tiene_ads"]]
        elif ads_filter == "Solo sin Ads":
            table = table[~table["tiene_ads"]]
        table = table[table["venta_neta_total_30d"].fillna(0) >= min_sales]
        table = table.sort_values(["priority_score", "venta_neta_total_30d"], ascending=[False, False])

        show = table[[
            "SKU_norm", "DESCRIPCION_clean", "estado_comercial", "diagnostico", "priority_score",
            "costo_base_neto", "variacion_costo_pct", "margen_local_actual_pct",
            "margen_real_30d_pct", "unidades_vendidas_30d", "venta_neta_total_30d", "tiene_ads",
            "dias_sin_ultima_compra", "dias_sin_ultima_venta"
        ]].copy()
        show.columns = [
            "SKU", "Descripción", "Estado", "Diagnóstico", "Prioridad",
            "Costo neto", "Var. costo", "Margen local", "Margen real 30d",
            "Unid. 30d", "Venta neta 30d", "Ads", "Días sin compra", "Días sin venta"
        ]
        show["Estado"] = show["Estado"].map(state_badge)
        show["Costo neto"] = show["Costo neto"].map(fmt_money)
        show["Var. costo"] = show["Var. costo"].map(margin_display)
        show["Margen local"] = show["Margen local"].map(margin_display)
        show["Margen real 30d"] = show["Margen real 30d"].map(margin_display)
        show["Unid. 30d"] = show["Unid. 30d"].map(fmt_int)
        show["Venta neta 30d"] = show["Venta neta 30d"].map(fmt_money)
        show["Ads"] = show["Ads"].map(lambda x: "Sí" if bool(x) else "No")
        show["Días sin compra"] = show["Días sin compra"].map(fmt_int)
        show["Días sin venta"] = show["Días sin venta"].map(fmt_int)
        st.dataframe(show, use_container_width=True, hide_index=True, height=520)

# =========================================================
# Tab 2: Ficha de Producto
# =========================================================
with tabs[1]:
    if row is None:
        st.warning("Selecciona un producto.")
    else:
        st.subheader(f"{row['SKU_norm']} — {row['DESCRIPCION_clean']}")

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Estado", state_badge(row["estado_comercial"]))
        m2.metric("Costo neto actual", fmt_money(row["costo_base_neto"]))
        m3.metric("Margen local", margin_display(row["margen_local_actual_pct"]))
        m4.metric("Margen real 30d", margin_display(row["margen_real_30d_pct"]))
        m5.metric("Var. costo", margin_display(row["variacion_costo_pct"]))
        m6.metric("Venta neta 30d", fmt_money(row["venta_neta_total_30d"]))

        st.info(f"**Diagnóstico:** {row['diagnostico']}")

        a, b = st.columns([1.15, 1])
        with a:
            st.markdown("### Información micro del producto")
            st.write(f"**SKU:** {row['SKU_norm']}")
            st.write(f"**Descripción:** {row['DESCRIPCION_clean']}")
            st.write(f"**Ubicación:** {row.get('UBIC', '—') if pd.notna(row.get('UBIC')) else '—'}")
            st.write(f"**Precio neto tienda:** {fmt_money(row.get('PRECIO NETO'))}")
            st.write(f"**Precio bruto tienda:** {fmt_money(row.get('PRECIO BRUTO'))}")
            st.write(f"**Monto en simulación Meli 1:** {fmt_money(row.get('MONTO EN SIMULACIÓN'))}")
            st.write(f"**Neto Meli 1:** {fmt_money(row.get(' NETO MELI 1'))}")
            st.write(f"**Venta bruto Meli 2:** {fmt_money(row.get('VENTA BRUTO MELI 2'))}")
            st.write(f"**Neto Meli 2:** {fmt_money(row.get('NETO MELI 2'))}")
            st.write(f"**Campaña Ads 1:** {row.get('ADS_1') or '—'}")
            st.write(f"**Campaña Ads 2:** {row.get('ADS_2') or '—'}")
            st.write(f"**MLC asociados:** {', '.join(model['mlc_map'].get(selected_sku, [])) if model['mlc_map'].get(selected_sku) else '—'}")
            st.write(f"**En Relámpago:** {'Sí' if bool(row.get('en_relampago')) else 'No'}")
            st.write(f"**Ads detectado:** {'Sí' if bool(row.get('tiene_ads')) else 'No'}")

        with b:
            st.markdown("### Lectura comercial")
            diag = pd.DataFrame([
                {"Indicador": "Ventas 30d (unidades)", "Valor": fmt_int(row.get("unidades_vendidas_30d"))},
                {"Indicador": "Ventas 90d (unidades)", "Valor": fmt_int(row.get("unidades_vendidas_90d"))},
                {"Indicador": "Venta neta 90d", "Valor": fmt_money(row.get("venta_neta_total_90d"))},
                {"Indicador": "Precio prom. neto 30d", "Valor": fmt_money(row.get("precio_promedio_neto_30d"))},
                {"Indicador": "Última venta", "Valor": fmt_date(row.get("ultima_venta_fecha_hist"))},
                {"Indicador": "Última compra", "Valor": fmt_date(row.get("ultima_compra_fecha"))},
                {"Indicador": "Días sin venta", "Valor": fmt_int(row.get("dias_sin_ultima_venta"))},
                {"Indicador": "Días sin compra", "Valor": fmt_int(row.get("dias_sin_ultima_compra"))},
                {"Indicador": "Proveedor último", "Valor": str(row.get("ultimo_proveedor", "—") or "—")},
                {"Indicador": "Promos control", "Valor": fmt_int(row.get("promos_control"))},
            ])
            st.dataframe(diag, use_container_width=True, hide_index=True, height=380)

        st.markdown("### Promociones y publicaciones asociadas")
        promos_sku = model["promos"][model["promos"]["SKU_norm"] == selected_sku].sort_values(["STATUS_ORDER", "Fecha_evento"])
        if promos_sku.empty:
            st.info("No encontré promos asociadas.")
        else:
            show = promos_sku[[
                "origen", "slot", "MLC", "Precio_promocional_bruto", "Fecha_evento",
                "Descuento_pct", "Motivo", "Ads", "Comentario", "STATUS"
            ]].copy()
            show.columns = ["Origen", "Slot", "MLC", "Precio promo", "Fecha", "Dcto %", "Motivo", "Ads", "Comentario", "Estado"]
            show["Precio promo"] = show["Precio promo"].map(fmt_money)
            show["Fecha"] = show["Fecha"].map(fmt_date)
            show["Dcto %"] = show["Dcto %"].map(margin_display)
            st.dataframe(show, use_container_width=True, hide_index=True, height=260)

        sub1, sub2 = st.columns(2)
        with sub1:
            st.markdown("### Compras")
            hist = model["purchase_map"].get(selected_sku, pd.DataFrame()).copy()
            if hist.empty:
                st.info("No encontré compras para este SKU.")
            else:
                show = hist[["Fecha_dt", "Razón Social", "Cantidad_num", "Precio_Un_Neto"]].sort_values("Fecha_dt", ascending=False)
                show.columns = ["Fecha", "Proveedor", "Cantidad", "Costo neto"]
                show["Fecha"] = show["Fecha"].map(fmt_date)
                show["Cantidad"] = show["Cantidad"].map(fmt_int)
                show["Costo neto"] = show["Costo neto"].map(fmt_money)
                st.dataframe(show, use_container_width=True, hide_index=True, height=260)

        with sub2:
            st.markdown("### Ventas")
            sales_hist = model["sales_map"].get(selected_sku, pd.DataFrame()).copy()
            if sales_hist.empty:
                st.info("No encontré ventas para este SKU.")
            else:
                grouped = (
                    sales_hist.groupby("Fecha_dt", dropna=False)
                    .agg(
                        unidades=("Cantidad_num", "sum"),
                        venta_neta=("Total_Linea_Neto", "sum"),
                        precio_prom_neto=("Precio_Prom_Neto", "mean"),
                    )
                    .reset_index()
                    .sort_values("Fecha_dt", ascending=False)
                )
                grouped.columns = ["Fecha", "Unidades", "Venta neta", "Precio prom. neto"]
                grouped["Fecha"] = grouped["Fecha"].map(fmt_date)
                grouped["Unidades"] = grouped["Unidades"].map(fmt_int)
                grouped["Venta neta"] = grouped["Venta neta"].map(fmt_money)
                grouped["Precio prom. neto"] = grouped["Precio prom. neto"].map(fmt_money)
                st.dataframe(grouped, use_container_width=True, hide_index=True, height=260)

# =========================================================
# Tab 3: Timeline
# =========================================================
with tabs[2]:
    if row is None:
        st.warning("Selecciona un producto.")
    else:
        st.subheader(f"Timeline del producto — {row['SKU_norm']}")
        timeline = build_timeline_for_sku(model, selected_sku)

        if timeline.empty:
            st.info("No hay eventos suficientes para este SKU.")
        else:
            chart_df = timeline.copy()
            chart_df = chart_df.dropna(subset=["fecha"])
            chart_df = chart_df.sort_values("fecha")
            cost_curve = chart_df[chart_df["tipo"] == "compra"][["fecha", "valor_neto"]].rename(columns={"valor_neto": "Costo neto"})
            sale_curve = chart_df[chart_df["tipo"] == "venta"][["fecha", "valor_neto"]].rename(columns={"valor_neto": "Precio venta neto"})
            promo_curve = chart_df[chart_df["tipo"].astype(str).str.contains("promo")][["fecha", "valor_neto"]].rename(columns={"valor_neto": "Precio promo neto"})

            c1, c2, c3 = st.columns(3)
            with c1:
                if not cost_curve.empty:
                    st.markdown("**Evolución de costo neto**")
                    st.line_chart(cost_curve.set_index("fecha"))
                else:
                    st.info("Sin datos de costo histórico.")
            with c2:
                if not sale_curve.empty:
                    st.markdown("**Precio venta neto realizado**")
                    st.line_chart(sale_curve.set_index("fecha"))
                else:
                    st.info("Sin ventas históricas.")
            with c3:
                if not promo_curve.empty:
                    st.markdown("**Precio promo neto**")
                    st.line_chart(promo_curve.set_index("fecha"))
                else:
                    st.info("Sin eventos de promo.")

            st.markdown("### Eventos cronológicos")
            show = timeline.copy()
            show["fecha"] = show["fecha"].map(fmt_date)
            show["cantidad"] = show["cantidad"].map(fmt_int)
            show["valor_neto"] = show["valor_neto"].map(fmt_money)
            show["valor_bruto"] = show["valor_bruto"].map(fmt_money)
            show.columns = ["Fecha", "Tipo", "Detalle", "Cantidad", "Valor neto", "Valor bruto", "Extra"]
            st.dataframe(show, use_container_width=True, hide_index=True, height=540)

# =========================================================
# Tab 4: Promociones
# =========================================================
with tabs[3]:
    st.subheader("Gestión de promociones")
    promos = model["promos"].copy()
    if promos.empty:
        st.info("No encontré promos en maestra ni en control.")
    else:
        left, right = st.columns([1, 2])

        with left:
            status_options = sorted(promos["STATUS"].dropna().unique().tolist(), key=lambda x: ["Vencida", "Vence hoy", "Vence mañana", "Vence en 7 días", "Vence en 15 días", "Vigente", "Sin fecha"].index(x) if x in ["Vencida", "Vence hoy", "Vence mañana", "Vence en 7 días", "Vence en 15 días", "Vigente", "Sin fecha"] else 999)
            status_filter = st.multiselect("Estado promo", status_options, default=status_options)
            origin_filter = st.multiselect("Origen", sorted(promos["origen"].dropna().unique().tolist()), default=sorted(promos["origen"].dropna().unique().tolist()))
            q = st.text_input("Buscar promo por SKU / descripción / MLC", key="promo_search")

            subset = promos.copy()
            if status_filter:
                subset = subset[subset["STATUS"].isin(status_filter)]
            if origin_filter:
                subset = subset[subset["origen"].isin(origin_filter)]
            if q:
                s = q.lower().strip()
                subset = subset[
                    subset["SKU_norm"].astype(str).str.lower().str.contains(s, na=False)
                    | subset["DESCRIPCION"].astype(str).str.lower().str.contains(s, na=False)
                    | subset["MLC"].astype(str).str.lower().str.contains(s, na=False)
                ]

            st.caption(f"Promos filtradas: {len(subset)}")

        with right:
            show = subset.sort_values(["STATUS_ORDER", "Fecha_evento", "SKU_norm"]).copy()
            show = show[[
                "origen", "SKU_norm", "DESCRIPCION", "slot", "MLC", "Precio_promocional_bruto",
                "Fecha_evento", "Descuento_pct", "Motivo", "Ads", "STATUS"
            ]]
            show.columns = ["Origen", "SKU", "Descripción", "Slot", "MLC", "Precio promo", "Fecha", "Dcto %", "Motivo", "Ads", "Estado"]
            show["Precio promo"] = show["Precio promo"].map(fmt_money)
            show["Fecha"] = show["Fecha"].map(fmt_date)
            show["Dcto %"] = show["Dcto %"].map(margin_display)
            st.dataframe(show, use_container_width=True, hide_index=True, height=620)

        st.markdown("### Editor rápido de Control de Promociones")
        cp = model["control_promos"].copy()
        if cp.empty:
            st.info("No encontré filas en CONTROL DE PROMOCIONES.")
        else:
            editable = cp[[
                "SKU_norm", "MLC", "Descripcion", "Pct_F", "Precio_promocional",
                "Motivo", "Ads_Comentario", "Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"
            ]].copy()
            editable.columns = [
                "SKU", "MLC", "Descripción", "% F", "Precio promocional",
                "Motivo", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"
            ]
            edited = st.data_editor(
                editable,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key="control_promos_editor",
                column_config={
                    "% F": st.column_config.NumberColumn(step=0.01),
                    "Precio promocional": st.column_config.NumberColumn(step=100),
                },
                height=360,
            )
            if st.button("Guardar editor de CONTROL DE PROMOCIONES"):
                new_df = edited.rename(columns={
                    "SKU": "SKU_norm", "MLC": "MLC", "Descripción": "Descripcion", "% F": "Pct_F",
                    "Precio promocional": "Precio_promocional", "Motivo": "Motivo",
                    "Ads/Comentario": "Ads_Comentario", "Campaña 1": "Campaña_1", "Campaña 2": "Campaña_2",
                    "Campaña 3": "Campaña_3", "Campaña 4": "Campaña_4"
                }).copy()
                new_df["SKU_norm"] = new_df["SKU_norm"].map(norm_sku)
                for c in ["Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"]:
                    new_df[c] = pd.to_datetime(new_df[c], dayfirst=True, errors="coerce").dt.normalize()
                new_df["Ads_Comentario"] = new_df["Ads_Comentario"].fillna("").astype(str)
                new_df["Pct_F"] = new_df["Pct_F"].map(lambda x: safe_float(x, np.nan))
                new_df["Precio_promocional"] = new_df["Precio_promocional"].map(lambda x: safe_float(x, np.nan))
                new_df["tiene_ads_control"] = new_df["Ads_Comentario"].str.strip() != ""
                new_df["Campaña_min"] = new_df[["Campaña_1", "Campaña_2", "Campaña_3", "Campaña_4"]].min(axis=1)
                status_info = new_df["Campaña_min"].map(lambda x: promo_status(x))
                new_df["Campaña_status"] = status_info.map(lambda x: x[0] if isinstance(x, tuple) else "Sin fecha")
                new_df["Campaña_order"] = status_info.map(lambda x: x[1] if isinstance(x, tuple) else 999)
                st.session_state.model["control_promos"] = new_df[new_df["SKU_norm"] != ""].copy()
                st.session_state.model["promos"] = build_promos_unified(st.session_state.model["master"], st.session_state.model["control_promos"])
                st.success("Editor de promociones actualizado en memoria.")

# =========================================================
# Tab 5: Relámpago
# =========================================================
with tabs[4]:
    st.subheader("Relámpago mi página")
    rel = model["rel"].copy()
    if rel.empty:
        rel = pd.DataFrame(columns=["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"])

    editable = rel.rename(columns={
        "SKU_norm": "SKU",
        "DESCRIPCION": "Descripción",
        "PRECIO_B2C": "Precio B2C",
        "TIPO": "Tipo",
        "ESTADO": "Estado",
    }).copy()

    edited_rel = st.data_editor(
        editable,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="relampago_editor",
        column_config={
            "Precio B2C": st.column_config.NumberColumn(step=100),
        },
        height=420,
    )

    if st.button("Guardar Relámpago"):
        new_rel = edited_rel.rename(columns={
            "SKU": "SKU_norm",
            "Descripción": "DESCRIPCION",
            "Precio B2C": "PRECIO_B2C",
            "Tipo": "TIPO",
            "Estado": "ESTADO",
        }).copy()
        new_rel["SKU_norm"] = new_rel["SKU_norm"].map(norm_sku)
        new_rel["PRECIO_B2C"] = new_rel["PRECIO_B2C"].map(lambda x: safe_float(x, np.nan))
        st.session_state.model["rel"] = new_rel[new_rel["SKU_norm"] != ""].copy()
        st.success("Relámpago actualizado en memoria.")

# =========================================================
# Tab 6: Descargar
# =========================================================
with tabs[5]:
    st.subheader("Descargar maestra actualizada")
    st.write("Descarga la maestra con los cambios hechos en Relámpago y CONTROL DE PROMOCIONES.")
    if st.button("Preparar Excel actualizado"):
        wb = model["workbook_meta"]
        payload = build_download_bytes(
            st.session_state.model["master"],
            st.session_state.model["rel"],
            st.session_state.model["control_promos"],
            wb["file_bytes"],
            wb["maestra_name"],
            wb["rel_name"],
            wb["control_name"],
        )
        st.session_state.download_bytes = payload

    if st.session_state.get("download_bytes"):
        st.download_button(
            "Descargar Excel actualizado",
            data=st.session_state.download_bytes,
            file_name="CENTRO_CONTROL_COMERCIAL_AURORA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
