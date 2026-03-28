
import io
import hashlib
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Centro de Control Comercial Aurora", layout="wide")

IVA = 1.19


# =============================
# Helpers
# =============================
def file_signature(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    data = uploaded_file.getvalue()
    return hashlib.md5(data).hexdigest()


def re_is_numberlike(s: str) -> bool:
    import re
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", s))


def safe_float(value, default=np.nan):
    try:
        if value is None:
            return default
        if isinstance(value, str):
            value = value.strip().replace("$", "").replace("%", "").replace(".", "").replace(",", ".")
            if value in {"", "-", "nan", "NaN"}:
                return default
        return float(value)
    except Exception:
        return default


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
    s = str(value).strip().upper().replace(" ", "")
    if not s or s == "NAN":
        return ""
    if s.startswith("MLC"):
        return s
    if s.isdigit():
        return f"MLC{s}"
    return s


def extract_mlcs(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return []
    s = str(value).upper().replace(" ", "")
    import re
    found = re.findall(r"(MLC\d+|\d{7,})", s)
    out = []
    for item in found:
        out.append(norm_mlc(item))
    return [x for x in dict.fromkeys(out) if x]


def to_date_only(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return pd.NaT
    try:
        return pd.to_datetime(value, errors="coerce", dayfirst=True).normalize()
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


def fmt_pct(value) -> str:
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return "—"
    if abs(x) <= 2:
        x *= 100
    return f"{x:.1f}%"


def pct_value(value):
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return np.nan
    if abs(x) <= 2:
        x *= 100
    return x


def div0(a, b):
    a = safe_float(a, np.nan)
    b = safe_float(b, np.nan)
    if np.isnan(a) or np.isnan(b) or b == 0:
        return np.nan
    return a / b


def promo_status(dt):
    dt = to_date_only(dt)
    if pd.isna(dt):
        return "Sin fecha", 999
    today = pd.Timestamp(date.today())
    delta = (dt - today).days
    if delta < 0:
        return "Vencida", delta
    if delta == 0:
        return "Vence hoy", 0
    if delta == 1:
        return "Vence mañana", 1
    if delta <= 7:
        return "Vence <= 7d", delta
    if delta <= 30:
        return "Vence <= 30d", delta
    return "Vence > 30d", delta


def detect_channel(vendor):
    s = str(vendor).upper().strip()
    if "MERCADO LIBRE" in s:
        return "ML"
    return "TIENDA"


def detect_customer_type(documento):
    s = str(documento).upper().strip()
    if "FACTURA" in s:
        return "EMPRESA"
    if "BOLETA" in s:
        return "PERSONA"
    return "OTRO"


def classify_purchase_pattern(median_qty, p75, p90, cv):
    if np.isnan(median_qty):
        return "Sin datos"
    if median_qty <= 1 and (np.isnan(p75) or p75 <= 2):
        return "Unitario"
    if not np.isnan(median_qty) and median_qty >= 4:
        return "Volumen"
    if not np.isnan(p90) and p90 >= max(6, median_qty * 3):
        return "Mixto"
    if not np.isnan(cv) and cv >= 1.5:
        return "Errático"
    return "Regular"


def find_sheet(sheet_names, wanted):
    wanted_low = wanted.lower().strip()
    for name in sheet_names:
        if name.lower().strip() == wanted_low:
            return name
    for name in sheet_names:
        if wanted_low in name.lower().strip():
            return name
    return None


# =============================
# Loaders
# =============================
@st.cache_data(show_spinner=False)
def load_master_bundle(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xls.sheet_names

    maestra_name = find_sheet(names, "MAESTRA de precios")
    bridge_name = find_sheet(names, "MLC -SKU")
    rel_name = find_sheet(names, "Relampago mi pagina")
    control_name = find_sheet(names, "CONTROL DE PROMOCIONES")

    if not maestra_name:
        raise ValueError("No encontré la hoja 'MAESTRA de precios'.")

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
def load_sales(file_bytes: bytes):
    raw = pd.read_excel(io.BytesIO(file_bytes))
    if raw.empty:
        return {"raw": pd.DataFrame()}

    df = raw.copy()
    for col in ["SKU", "Fecha", "Cantidad", "Precio Un.", "Total Línea", "Vendedor", "Documento", "Razón Social", "Producto", "Familia"]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU"].map(norm_sku)
    df["Fecha_dt"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
    df["Cantidad_num"] = df["Cantidad"].map(lambda x: safe_float(x, 0))
    df["Precio_Unitario_Bruto"] = df["Precio Un."].map(lambda x: safe_float(x, np.nan))
    df["Total_Linea_Bruto"] = df["Total Línea"].map(lambda x: safe_float(x, np.nan))
    df["Precio_Unitario_Neto"] = df["Precio_Unitario_Bruto"] / IVA
    df["Total_Linea_Neto"] = df["Total_Linea_Bruto"] / IVA
    df["canal"] = df["Vendedor"].apply(detect_channel)
    df["tipo_cliente"] = df["Documento"].apply(detect_customer_type)
    df["descripcion_venta"] = df["Producto"].fillna("").astype(str)
    df = df[df["SKU_norm"] != ""].copy()
    return {"raw": df}


@st.cache_data(show_spinner=False)
def load_purchases(file_bytes: bytes):
    if not file_bytes:
        return {"raw": pd.DataFrame(), "summary": pd.DataFrame(), "by_sku": {}}

    raw = pd.read_excel(io.BytesIO(file_bytes))
    if raw.empty:
        return {"raw": pd.DataFrame(), "summary": pd.DataFrame(), "by_sku": {}}

    df = raw.copy()
    for col in ["SKU", "Fecha", "Razón Social", "Precio Un.", "Cantidad"]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU"].map(norm_sku)
    df["Fecha_dt"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
    df["Precio_Un_Num"] = df["Precio Un."].map(lambda x: safe_float(x, np.nan))
    df["Cantidad_num"] = df["Cantidad"].map(lambda x: safe_float(x, np.nan))
    df = df[df["SKU_norm"] != ""].copy()
    df = df.sort_values(["SKU_norm", "Fecha_dt"])

    by_sku = {sku: grp.copy() for sku, grp in df.groupby("SKU_norm", sort=False)}

    rows = []
    for sku, grp in by_sku.items():
        grp = grp.sort_values("Fecha_dt")
        last = grp.iloc[-1]
        prev_price = safe_float(grp.iloc[-2]["Precio_Un_Num"], np.nan) if len(grp) >= 2 else np.nan
        last_price = safe_float(last["Precio_Un_Num"], np.nan)
        variation = np.nan
        if not np.isnan(last_price) and not np.isnan(prev_price) and prev_price != 0:
            variation = ((last_price - prev_price) / prev_price) * 100
        days_since = (pd.Timestamp(date.today()) - last["Fecha_dt"]).days if pd.notna(last["Fecha_dt"]) else np.nan
        rows.append({
            "SKU_norm": sku,
            "ultima_fecha": last["Fecha_dt"],
            "ultimo_precio": last_price,
            "ultimo_proveedor": last.get("Razón Social", ""),
            "ultima_cantidad": safe_float(last.get("Cantidad_num"), np.nan),
            "compra_anterior": prev_price,
            "variacion_costo_pct": variation,
            "compras_total": len(grp),
            "dias_sin_compra": days_since,
        })
    summary = pd.DataFrame(rows)
    return {"raw": df, "summary": summary, "by_sku": by_sku}


@st.cache_data(show_spinner=False)
def load_product_ads(file_bytes: bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Reporte por anuncios", header=0)
    if df.empty:
        return {"raw": pd.DataFrame(), "summary": pd.DataFrame()}

    first_row = df.iloc[0].tolist()
    df = df.iloc[1:].copy()
    df.columns = first_row
    df = df.rename(columns={
        "Desde": "Desde",
        "Hasta": "Hasta",
        "Campaña": "Campaña",
        "Título de anuncio": "Titulo",
        "Número de \npublicación": "MLC",
        "Estado": "Estado",
        "Impresiones": "Impresiones",
        "Clics": "Clics",
        "CPC \n(Costo por clic)": "CPC",
        "CTR\n(Click Through Rate)": "CTR",
        "CVR\n(Convertion rate)": "CVR",
        "Ingresos\n(Moneda local)": "Ingresos_Bruto",
        "Inversión\n(Moneda local)": "Inversion",
        "ACOS\n(Inversión / Ingresos)": "ACOS",
        "ROAS\n(Ingresos / Inversión)": "ROAS",
        "Ventas directas": "Ventas_Directas",
        "Ventas indirectas": "Ventas_Indirectas",
        "Ventas por publicidad\n(Directas + Indirectas)": "Ventas_Ads",
        "Ingresos por ventas directas\n(Moneda local)": "Ingresos_Directos_Bruto",
        "Ingresos por ventas indirectas\n(Moneda local)": "Ingresos_Indirectos_Bruto",
    })
    needed = ["Desde", "Hasta", "Campaña", "Titulo", "MLC", "Estado", "Impresiones", "Clics", "CPC", "CTR", "CVR", "Ingresos_Bruto", "Inversion", "ACOS", "ROAS", "Ventas_Directas", "Ventas_Indirectas", "Ventas_Ads"]
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    df["Desde_dt"] = pd.to_datetime(df["Desde"], dayfirst=True, errors="coerce").dt.normalize()
    df["Hasta_dt"] = pd.to_datetime(df["Hasta"], dayfirst=True, errors="coerce").dt.normalize()
    df["MLC_norm"] = df["MLC"].map(norm_mlc)
    numeric_cols = ["Impresiones", "Clics", "CPC", "CTR", "CVR", "Ingresos_Bruto", "Inversion", "ACOS", "ROAS", "Ventas_Directas", "Ventas_Indirectas", "Ventas_Ads"]
    for col in numeric_cols:
        df[col] = df[col].map(lambda x: safe_float(x, 0))
    df["Ingresos_Neto"] = df["Ingresos_Bruto"] / IVA
    summary = df.groupby("MLC_norm", dropna=False).agg({
        "Impresiones": "sum",
        "Clics": "sum",
        "Ingresos_Bruto": "sum",
        "Ingresos_Neto": "sum",
        "Inversion": "sum",
        "Ventas_Ads": "sum",
        "Ventas_Directas": "sum",
        "Ventas_Indirectas": "sum"
    }).reset_index()
    summary["ACOS_calc"] = summary.apply(lambda r: div0(r["Inversion"], r["Ingresos_Bruto"]) * 100 if r["Ingresos_Bruto"] else np.nan, axis=1)
    summary["ROAS_calc"] = summary.apply(lambda r: div0(r["Ingresos_Bruto"], r["Inversion"]), axis=1)
    return {"raw": df, "summary": summary}


@st.cache_data(show_spinner=False)
def load_brand_ads(file_bytes: bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Reporte por palabras clave", header=0)
    if df.empty:
        return {"raw": pd.DataFrame(), "campaign_summary": pd.DataFrame(), "top_keywords": pd.DataFrame()}

    first_row = df.iloc[0].tolist()
    df = df.iloc[1:].copy()
    df.columns = first_row
    rename_map = {
        "Campaña": "Campaña",
        "Estado de campaña": "Estado_Campaña",
        "Palabra clave": "Palabra_Clave",
        "Segmentación": "Segmentacion",
        "Estado de palabra clave": "Estado_Palabra",
        "CPC máximo": "CPC_Max",
        "Desde": "Desde",
        "Hasta": "Hasta",
        "Impresiones": "Impresiones",
        "Clics": "Clics",
        "CPC \n(Costo por clic)": "CPC",
        "CTR\n(Click through rate)": "CTR",
        "CVR\n(Conversion rate)": "CVR",
        "Ingresos\n(Moneda local)": "Ingresos_Bruto",
        "Inversión\n(Moneda local)": "Inversion",
        "ACOS\n(Inversión / Ingresos)": "ACOS",
        "ROAS\n(Ingresos / Inversión)": "ROAS",
        "Ventas por publicidad": "Ventas_Ads",
        "Unidades vendidas por publicidad": "Unidades_Ads",
    }
    df = df.rename(columns=rename_map)
    for col in ["Campaña", "Palabra_Clave", "Impresiones", "Clics", "CTR", "CVR", "Ingresos_Bruto", "Inversion", "ACOS", "ROAS", "Ventas_Ads", "Unidades_Ads"]:
        if col not in df.columns:
            df[col] = np.nan
    for col in ["Impresiones", "Clics", "CTR", "CVR", "Ingresos_Bruto", "Inversion", "ACOS", "ROAS", "Ventas_Ads", "Unidades_Ads"]:
        df[col] = df[col].map(lambda x: safe_float(x, 0))
    campaign_summary = df.groupby("Campaña", dropna=False).agg({
        "Impresiones": "sum",
        "Clics": "sum",
        "Ingresos_Bruto": "sum",
        "Inversion": "sum",
        "Ventas_Ads": "sum",
        "Unidades_Ads": "sum",
    }).reset_index()
    campaign_summary["ACOS_calc"] = campaign_summary.apply(lambda r: div0(r["Inversion"], r["Ingresos_Bruto"]) * 100 if r["Ingresos_Bruto"] else np.nan, axis=1)
    campaign_summary["ROAS_calc"] = campaign_summary.apply(lambda r: div0(r["Ingresos_Bruto"], r["Inversion"]), axis=1)
    top_keywords = df.sort_values(["Ventas_Ads", "Ingresos_Bruto", "Clics"], ascending=False).head(200).copy()
    return {"raw": df, "campaign_summary": campaign_summary, "top_keywords": top_keywords}


# =============================
# Normalization
# =============================
def normalize_master(master_df: pd.DataFrame) -> pd.DataFrame:
    df = master_df.copy()

    wanted = [
        "SKU", "DESCRIPCIÓN", "UBIC", "ÚLTIMO COSTO", "MARGEN LOCAL", "PRECIO NETO", "PRECIO BRUTO",
        "MARGEN MELI 1", " NETO MELI 1", "MONTO EN SIMULACIÓN", "CAMPAÑA PADS", "MLC", "MLC SINCRONIZADO",
        " DCTO", "PRECIO B2C PUBLICADO ", "FECHA VENCI", "COMENTARIO",
        "MARGEN MELI 2", "NETO MELI 2", "VENTA BRUTO MELI 2", "MLC.1", "MLC SINCRONIZADO.1", "CAMPAÑA PADS.1",
        " DCTO.1", "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1"
    ]
    for col in wanted:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU"].map(norm_sku)
    df["DESCRIPCION"] = df["DESCRIPCIÓN"].fillna("").astype(str)
    df["UBICACION"] = df["UBIC"]
    df["COSTO_NETO"] = df["ÚLTIMO COSTO"].map(lambda x: safe_float(x, np.nan))
    df["PRECIO_TIENDA_BRUTO"] = df["PRECIO BRUTO"].map(lambda x: safe_float(x, np.nan))
    df["PRECIO_TIENDA_NETO"] = df["PRECIO NETO"].map(lambda x: safe_float(x, np.nan))
    df["MONTO_SIM_BRUTO"] = df["MONTO EN SIMULACIÓN"].map(lambda x: safe_float(x, np.nan))
    df["MONTO_SIM_NETO"] = df["MONTO_SIM_BRUTO"] / IVA
    df["MARGEN_ML_SIM_PCT"] = df.apply(lambda r: ((r["MONTO_SIM_NETO"] - r["COSTO_NETO"]) / r["MONTO_SIM_NETO"] * 100)
                                       if pd.notna(r["MONTO_SIM_NETO"]) and r["MONTO_SIM_NETO"] != 0 and pd.notna(r["COSTO_NETO"]) else np.nan, axis=1)
    df["MARGEN_TIENDA_PCT"] = df.apply(lambda r: ((r["PRECIO_TIENDA_NETO"] - r["COSTO_NETO"]) / r["PRECIO_TIENDA_NETO"] * 100)
                                       if pd.notna(r["PRECIO_TIENDA_NETO"]) and r["PRECIO_TIENDA_NETO"] != 0 and pd.notna(r["COSTO_NETO"]) else np.nan, axis=1)

    df["MLC_1"] = df["MLC"].map(norm_mlc)
    df["MLC_SYNC_1"] = df["MLC SINCRONIZADO"].map(norm_mlc)
    df["MLC_2"] = df["MLC.1"].map(norm_mlc)
    df["MLC_SYNC_2"] = df["MLC SINCRONIZADO.1"].map(norm_mlc)

    df["ADS_FLAG_1"] = df["CAMPAÑA PADS"].fillna("").astype(str).str.strip()
    df["ADS_FLAG_2"] = df["CAMPAÑA PADS.1"].fillna("").astype(str).str.strip()
    df["TIENE_ADS"] = (df["ADS_FLAG_1"] != "") | (df["ADS_FLAG_2"] != "")

    for col in ["FECHA VENCI", "FECHA VENCI.1"]:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()

    return df[df["SKU_norm"] != ""].copy()


def normalize_bridge(bridge_df):
    if bridge_df is None or bridge_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "MLC_norm"])
    df = bridge_df.copy()
    sku_col = "SKU" if "SKU" in df.columns else df.columns[0]
    mlc_col = "Número de publicación" if "Número de publicación" in df.columns else df.columns[-1]
    df["SKU_norm"] = df[sku_col].map(norm_sku)
    df["MLC_norm"] = df[mlc_col].map(norm_mlc)
    df = df[(df["SKU_norm"] != "") & (df["MLC_norm"] != "")]
    return df[["SKU_norm", "MLC_norm"]].drop_duplicates()


def normalize_rel(rel_df):
    if rel_df is None or rel_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"])
    df = rel_df.copy()
    while df.shape[1] < 6:
        df[df.shape[1]] = np.nan
    df = df.iloc[:, :6].copy()
    df.columns = ["SKU_raw", "DESCRIPCION", "PRECIO_B2C", "EXTRA", "TIPO", "ESTADO"]
    df["SKU_norm"] = df["SKU_raw"].map(norm_sku)
    df["PRECIO_B2C"] = df["PRECIO_B2C"].map(lambda x: safe_float(x, np.nan))
    return df[df["SKU_norm"] != ""][["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"]].copy()


def normalize_control_promos(control_df):
    if control_df is None or control_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "MLC_norm", "Descripcion", "Precio_Promo", "Motivo", "Ads_Comentario", "Fecha_Vencimiento", "Estado_Promo"])
    df = control_df.copy()
    sku_col = " " if " " in df.columns else df.columns[0]
    for col in [sku_col, "N° Publicación", "Descripción", "Precio promocional", "Motivo promoción", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df[sku_col].map(norm_sku)
    df["MLCs"] = df["N° Publicación"].apply(extract_mlcs)
    df["MLC_norm"] = df["MLCs"].apply(lambda xs: xs[0] if xs else "")
    df["Precio_Promo"] = df["Precio promocional"].map(lambda x: safe_float(x, np.nan))
    date_cols = ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()
    df["Fecha_Vencimiento"] = df[date_cols].min(axis=1)
    tmp = df["Fecha_Vencimiento"].apply(promo_status)
    df["Estado_Promo"] = tmp.apply(lambda x: x[0] if isinstance(x, tuple) else "Sin fecha")
    out = df[df["SKU_norm"] != ""].copy()
    out = out.rename(columns={"Descripción": "Descripcion", "Motivo promoción": "Motivo", "Ads/Comentario": "Ads_Comentario"})
    return out[["SKU_norm", "MLC_norm", "Descripcion", "Precio_Promo", "Motivo", "Ads_Comentario", "Fecha_Vencimiento", "Estado_Promo"]].copy()


# =============================
# Model building
# =============================
def summarize_sales_window(sales_raw: pd.DataFrame, days: int | None):
    if sales_raw is None or sales_raw.empty:
        empty = pd.DataFrame(columns=[
            "SKU_norm", "ventas_n", "unidades_total", "ingresos_bruto_total", "ingresos_neto_total",
            "precio_promedio_bruto", "precio_promedio_neto", "ventas_ml_n", "ventas_tienda_n",
            "ingresos_ml_neto", "ingresos_tienda_neto", "unidades_ml", "unidades_tienda",
            "participacion_ml_pct", "participacion_tienda_pct", "cliente_empresa_pct", "cliente_persona_pct",
            "mediana_cantidad", "p75_cantidad", "p90_cantidad", "cv_cantidad", "patron_compra"
        ])
        return {"raw": sales_raw, "summary": empty, "timeline_daily": pd.DataFrame()}

    df = sales_raw.copy()
    if days is not None:
        cutoff = pd.Timestamp(date.today()) - pd.Timedelta(days=days)
        df = df[df["Fecha_dt"] >= cutoff].copy()

    if df.empty:
        return summarize_sales_window(pd.DataFrame(columns=sales_raw.columns), None)

    rows = []
    for sku, grp in df.groupby("SKU_norm", sort=False):
        ventas_n = len(grp)
        unidades = grp["Cantidad_num"].sum()
        bruto = grp["Total_Linea_Bruto"].sum()
        neto = grp["Total_Linea_Neto"].sum()
        ml = grp[grp["canal"] == "ML"]
        tienda = grp[grp["canal"] == "TIENDA"]
        emp = grp[grp["tipo_cliente"] == "EMPRESA"]
        per = grp[grp["tipo_cliente"] == "PERSONA"]

        mediana = grp["Cantidad_num"].median()
        p75 = grp["Cantidad_num"].quantile(0.75) if len(grp) > 1 else mediana
        p90 = grp["Cantidad_num"].quantile(0.90) if len(grp) > 1 else mediana
        mean_qty = grp["Cantidad_num"].mean()
        std_qty = grp["Cantidad_num"].std()
        cv = div0(std_qty, mean_qty)
        patron = classify_purchase_pattern(mediana, p75, p90, cv)

        rows.append({
            "SKU_norm": sku,
            "ventas_n": ventas_n,
            "unidades_total": unidades,
            "ingresos_bruto_total": bruto,
            "ingresos_neto_total": neto,
            "precio_promedio_bruto": div0(bruto, unidades),
            "precio_promedio_neto": div0(neto, unidades),
            "ventas_ml_n": len(ml),
            "ventas_tienda_n": len(tienda),
            "ingresos_ml_neto": ml["Total_Linea_Neto"].sum(),
            "ingresos_tienda_neto": tienda["Total_Linea_Neto"].sum(),
            "unidades_ml": ml["Cantidad_num"].sum(),
            "unidades_tienda": tienda["Cantidad_num"].sum(),
            "participacion_ml_pct": div0(ml["Total_Linea_Neto"].sum(), neto) * 100 if neto else np.nan,
            "participacion_tienda_pct": div0(tienda["Total_Linea_Neto"].sum(), neto) * 100 if neto else np.nan,
            "cliente_empresa_pct": div0(emp["Total_Linea_Neto"].sum(), neto) * 100 if neto else np.nan,
            "cliente_persona_pct": div0(per["Total_Linea_Neto"].sum(), neto) * 100 if neto else np.nan,
            "mediana_cantidad": mediana,
            "p75_cantidad": p75,
            "p90_cantidad": p90,
            "cv_cantidad": cv,
            "patron_compra": patron,
            "ticket_empresa_neto": div0(emp["Total_Linea_Neto"].sum(), len(emp)),
            "ticket_persona_neto": div0(per["Total_Linea_Neto"].sum(), len(per)),
        })

    summary = pd.DataFrame(rows)

    timeline_daily = df.groupby(["Fecha_dt", "SKU_norm", "canal"], dropna=False).agg({
        "Cantidad_num": "sum",
        "Total_Linea_Neto": "sum"
    }).reset_index()

    return {"raw": df, "summary": summary, "timeline_daily": timeline_daily}


def build_model(bundle, sales_data, purchases_data, product_ads_data=None, brand_ads_data=None):
    master = normalize_master(bundle["master_df"])
    bridge = normalize_bridge(bundle["bridge_df"])
    rel = normalize_rel(bundle["rel_df"])
    control_promos = normalize_control_promos(bundle["control_df"])

    # MLC map
    mlc_map = bridge.groupby("SKU_norm")["MLC_norm"].apply(list).to_dict() if not bridge.empty else {}
    for idx, row in master.iterrows():
        sku = row["SKU_norm"]
        extra = []
        for mlc in [row["MLC_1"], row["MLC_SYNC_1"], row["MLC_2"], row["MLC_SYNC_2"]]:
            if mlc:
                extra.append(mlc)
        if sku in mlc_map:
            extra.extend(mlc_map[sku])
        mlc_map[sku] = sorted([m for m in dict.fromkeys(extra) if m])

    # Promote control promo rows by SKU fallback
    if not control_promos.empty:
        control_promos["SKU_norm"] = control_promos["SKU_norm"].replace("", np.nan)
        # fallback from bridge if empty SKU and mlc exists
        mlc_to_sku = bridge.drop_duplicates("MLC_norm").set_index("MLC_norm")["SKU_norm"].to_dict() if not bridge.empty else {}
        control_promos.loc[control_promos["SKU_norm"].isna() & control_promos["MLC_norm"].notna(), "SKU_norm"] = control_promos.loc[
            control_promos["SKU_norm"].isna() & control_promos["MLC_norm"].notna(), "MLC_norm"
        ].map(mlc_to_sku)
        control_promos["SKU_norm"] = control_promos["SKU_norm"].fillna("")

    sales_30 = summarize_sales_window(sales_data["raw"], 30)
    sales_90 = summarize_sales_window(sales_data["raw"], 90)
    sales_all = summarize_sales_window(sales_data["raw"], None)

    impact = master[[
        "SKU_norm", "DESCRIPCION", "UBICACION", "COSTO_NETO", "PRECIO_TIENDA_BRUTO", "PRECIO_TIENDA_NETO",
        "MONTO_SIM_BRUTO", "MONTO_SIM_NETO", "MARGEN_ML_SIM_PCT", "MARGEN_TIENDA_PCT",
        "TIENE_ADS", "ADS_FLAG_1", "ADS_FLAG_2", "MLC_1", "MLC_SYNC_1", "MLC_2", "MLC_SYNC_2",
        "PRECIO B2C PUBLICADO ", "PRECIO B2C", "FECHA VENCI", "FECHA VENCI.1", "COMENTARIO", "COMENTARIO.1"
    ]].copy()

    impact = impact.merge(sales_30["summary"], on="SKU_norm", how="left", suffixes=("", "_30"))
    if "ventas_n" in impact.columns:
        impact = impact.rename(columns={
            "ventas_n": "ventas_30d",
            "unidades_total": "unidades_30d",
            "ingresos_neto_total": "ingresos_neto_30d",
            "ingresos_bruto_total": "ingresos_bruto_30d",
            "precio_promedio_neto": "precio_promedio_neto_30d",
            "precio_promedio_bruto": "precio_promedio_bruto_30d",
            "ingresos_ml_neto": "ingresos_ml_neto_30d",
            "ingresos_tienda_neto": "ingresos_tienda_neto_30d",
            "participacion_ml_pct": "participacion_ml_pct_30d",
            "participacion_tienda_pct": "participacion_tienda_pct_30d",
            "cliente_empresa_pct": "cliente_empresa_pct_30d",
            "cliente_persona_pct": "cliente_persona_pct_30d",
            "mediana_cantidad": "mediana_cantidad_30d",
            "p75_cantidad": "p75_cantidad_30d",
            "p90_cantidad": "p90_cantidad_30d",
            "cv_cantidad": "cv_cantidad_30d",
            "patron_compra": "patron_compra_30d",
            "ticket_empresa_neto": "ticket_empresa_neto_30d",
            "ticket_persona_neto": "ticket_persona_neto_30d",
            "ventas_ml_n": "ventas_ml_n_30d",
            "ventas_tienda_n": "ventas_tienda_n_30d",
            "unidades_ml": "unidades_ml_30d",
            "unidades_tienda": "unidades_tienda_30d",
        })

    impact = impact.merge(
        sales_90["summary"][["SKU_norm", "ventas_n", "unidades_total", "ingresos_neto_total", "mediana_cantidad", "patron_compra"]],
        on="SKU_norm", how="left"
    ).rename(columns={
        "ventas_n": "ventas_90d",
        "unidades_total": "unidades_90d",
        "ingresos_neto_total": "ingresos_neto_90d",
        "mediana_cantidad": "mediana_cantidad_90d",
        "patron_compra": "patron_compra_90d",
    })

    impact = impact.merge(purchases_data["summary"], on="SKU_norm", how="left")

    # promos summary
    promo_summary = control_promos.groupby("SKU_norm", dropna=False).agg({
        "MLC_norm": "count",
        "Precio_Promo": "min",
        "Fecha_Vencimiento": "min"
    }).reset_index().rename(columns={"MLC_norm": "promos_control_n", "Precio_Promo": "precio_promo_min", "Fecha_Vencimiento": "promo_prox_venci"})
    impact = impact.merge(promo_summary, on="SKU_norm", how="left")
    impact["promo_estado"] = impact["promo_prox_venci"].apply(lambda x: promo_status(x)[0] if pd.notna(x) else "Sin promo")

    # Ads by SKU through MLC map
    if product_ads_data and not product_ads_data["summary"].empty:
        ads_summary = product_ads_data["summary"].copy()
        sku_ads_rows = []
        for sku, mlcs in mlc_map.items():
            sub = ads_summary[ads_summary["MLC_norm"].isin(mlcs)]
            if sub.empty:
                continue
            sku_ads_rows.append({
                "SKU_norm": sku,
                "ads_mlc_activos": sub["MLC_norm"].nunique(),
                "ads_impresiones": sub["Impresiones"].sum(),
                "ads_clicks": sub["Clics"].sum(),
                "ads_ingresos_neto": sub["Ingresos_Neto"].sum(),
                "ads_inversion": sub["Inversion"].sum(),
                "ads_ventas": sub["Ventas_Ads"].sum(),
                "ads_acos_pct": div0(sub["Inversion"].sum(), sub["Ingresos_Bruto"].sum()) * 100 if sub["Ingresos_Bruto"].sum() else np.nan,
                "ads_roas": div0(sub["Ingresos_Bruto"].sum(), sub["Inversion"].sum()),
            })
        sku_ads = pd.DataFrame(sku_ads_rows)
        impact = impact.merge(sku_ads, on="SKU_norm", how="left")
    else:
        for col in ["ads_mlc_activos", "ads_impresiones", "ads_clicks", "ads_ingresos_neto", "ads_inversion", "ads_ventas", "ads_acos_pct", "ads_roas"]:
            impact[col] = np.nan

    # contribution and profitability
    impact["utilidad_ml_sim_unit"] = impact["MONTO_SIM_NETO"] - impact["COSTO_NETO"]
    impact["margen_real_ml_pct"] = impact["MARGEN_ML_SIM_PCT"]
    impact["margen_real_tienda_pct"] = impact["MARGEN_TIENDA_PCT"]

    impact["ads_sobre_ingreso_ml_pct_30d"] = impact.apply(
        lambda r: div0(r["ads_inversion"], r["ingresos_ml_neto_30d"]) * 100 if pd.notna(r.get("ingresos_ml_neto_30d")) and r.get("ingresos_ml_neto_30d", 0) > 0 else np.nan,
        axis=1
    )

    impact["margen_ml_despues_ads_pct"] = impact.apply(
        lambda r: ((r["ingresos_ml_neto_30d"] - (r["unidades_ml_30d"] * r["COSTO_NETO"]) - r["ads_inversion"]) / r["ingresos_ml_neto_30d"] * 100)
        if pd.notna(r.get("ingresos_ml_neto_30d")) and r.get("ingresos_ml_neto_30d", 0) > 0 and pd.notna(r.get("COSTO_NETO")) else np.nan,
        axis=1
    )

    impact["delta_habito_30_vs_90"] = impact["mediana_cantidad_30d"] - impact["mediana_cantidad_90d"]

    def classify_row(r):
        if safe_float(r.get("ventas_30d"), 0) == 0 and safe_float(r.get("ventas_90d"), 0) == 0:
            return "Sin tracción"
        if pd.notna(r.get("variacion_costo_pct")) and r.get("variacion_costo_pct") >= 20 and (pd.notna(r.get("margen_real_ml_pct")) and r.get("margen_real_ml_pct") <= 20):
            return "Crítico por costo"
        if pd.notna(r.get("promo_prox_venci")) and promo_status(r.get("promo_prox_venci"))[1] <= 7 and promo_status(r.get("promo_prox_venci"))[1] >= 0:
            return "Revisar promo"
        if pd.notna(r.get("ads_sobre_ingreso_ml_pct_30d")) and pd.notna(r.get("margen_real_ml_pct")) and r.get("ads_sobre_ingreso_ml_pct_30d") >= r.get("margen_real_ml_pct"):
            return "Ads en riesgo"
        if pd.notna(r.get("margen_real_ml_pct")) and r.get("margen_real_ml_pct") < 10 and safe_float(r.get("ventas_30d"), 0) > 0:
            return "Margen crítico"
        if safe_float(r.get("ventas_30d"), 0) > 0 and pd.notna(r.get("margen_real_ml_pct")) and r.get("margen_real_ml_pct") >= 20:
            return "Oportunidad"
        return "Monitorear"

    impact["estado_comercial"] = impact.apply(classify_row, axis=1)

    def score_row(r):
        score = 0
        ventas = safe_float(r.get("ingresos_neto_30d"), 0)
        score += min(40, ventas / 100000)
        var = safe_float(r.get("variacion_costo_pct"), 0)
        if not np.isnan(var) and var > 0:
            score += min(25, var)
        margin_ml = safe_float(r.get("margen_real_ml_pct"), np.nan)
        if not np.isnan(margin_ml):
            if margin_ml < 0:
                score += 30
            elif margin_ml < 10:
                score += 20
            elif margin_ml < 20:
                score += 10
        if str(r.get("estado_comercial")) == "Revisar promo":
            score += 12
        if str(r.get("estado_comercial")) == "Ads en riesgo":
            score += 15
        if pd.notna(r.get("dias_sin_compra")) and r.get("dias_sin_compra") > 90:
            score += 8
        return score

    impact["score_prioridad"] = impact.apply(score_row, axis=1)

    product_options = impact["SKU_norm"].tolist()
    sku_desc = dict(zip(impact["SKU_norm"], impact["DESCRIPCION"]))

    return {
        "bundle": bundle,
        "master": master,
        "bridge": bridge,
        "rel": rel,
        "control_promos": control_promos,
        "sales_raw": sales_data["raw"],
        "sales_30": sales_30,
        "sales_90": sales_90,
        "sales_all": sales_all,
        "purchases_summary": purchases_data["summary"],
        "purchase_map": purchases_data["by_sku"],
        "product_ads": product_ads_data if product_ads_data else {"raw": pd.DataFrame(), "summary": pd.DataFrame()},
        "brand_ads": brand_ads_data if brand_ads_data else {"raw": pd.DataFrame(), "campaign_summary": pd.DataFrame(), "top_keywords": pd.DataFrame()},
        "impact": impact.sort_values(["score_prioridad", "ingresos_neto_30d"], ascending=[False, False]),
        "mlc_map": mlc_map,
        "sku_options": product_options,
        "sku_desc": sku_desc,
    }


# =============================
# Download helpers
# =============================
def rel_to_sheet_df(rel_df: pd.DataFrame) -> pd.DataFrame:
    if rel_df is None or rel_df.empty:
        return pd.DataFrame(columns=list(range(6)))
    df = rel_df.copy()
    out = pd.DataFrame({
        0: df["SKU_norm"],
        1: df["DESCRIPCION"],
        2: df["PRECIO_B2C"],
        3: np.nan,
        4: df["TIPO"],
        5: df["ESTADO"],
    })
    return out


def control_to_sheet_df(control_df: pd.DataFrame) -> pd.DataFrame:
    if control_df is None or control_df.empty:
        return pd.DataFrame(columns=[" ", "% F", "N° Publicación", "Descripción", "% F.1", "Precio promocional", "Motivo promoción", "Unnamed: 7", "margen", "Ads/Comentario", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"])
    df = control_df.copy()
    out = pd.DataFrame({
        " ": df["SKU_norm"],
        "% F": np.nan,
        "N° Publicación": df["MLC_norm"],
        "Descripción": df["Descripcion"],
        "% F.1": np.nan,
        "Precio promocional": df["Precio_Promo"],
        "Motivo promoción": df["Motivo"],
        "Unnamed: 7": np.nan,
        "margen": np.nan,
        "Ads/Comentario": df["Ads_Comentario"],
        "Campaña 1": df["Fecha_Vencimiento"],
        "Campaña 2": np.nan,
        "Campaña 3": np.nan,
        "Campaña 4": np.nan,
    })
    return out


@st.cache_data(show_spinner=False)
def build_download_bytes(master_df, rel_df, control_df, original_bytes, maestra_name, rel_name, control_name):
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            if sheet == maestra_name:
                drop_cols = ["SKU_norm", "DESCRIPCION", "UBICACION", "COSTO_NETO", "PRECIO_TIENDA_BRUTO", "PRECIO_TIENDA_NETO", "MONTO_SIM_BRUTO", "MONTO_SIM_NETO", "MARGEN_ML_SIM_PCT", "MARGEN_TIENDA_PCT", "MLC_1", "MLC_SYNC_1", "MLC_2", "MLC_SYNC_2", "ADS_FLAG_1", "ADS_FLAG_2", "TIENE_ADS"]
                master_df.drop(columns=[c for c in drop_cols if c in master_df.columns], errors="ignore").to_excel(writer, sheet_name=sheet, index=False)
            elif rel_name and sheet == rel_name:
                rel_to_sheet_df(rel_df).to_excel(writer, sheet_name=sheet, index=False, header=False)
            elif control_name and sheet == control_name:
                control_to_sheet_df(control_df).to_excel(writer, sheet_name=sheet, index=False)
            else:
                pd.read_excel(io.BytesIO(original_bytes), sheet_name=sheet, header=None if "relampago" in sheet.lower() else 0).to_excel(
                    writer, sheet_name=sheet, index=False, header=not ("relampago" in sheet.lower())
                )
    return out.getvalue()


# =============================
# UI init
# =============================
st.title("Centro de Control Comercial Aurora")
st.caption("Enfoque principal: optimizar cambio de precios con compras, ventas, promos y ads.")

with st.sidebar:
    st.subheader("Fuentes")
    master_up = st.file_uploader("Maestra precios y promos", type=["xlsx"], key="master")
    sales_up = st.file_uploader("Informe de ventas", type=["xlsx"], key="sales")
    purchases_up = st.file_uploader("Informe de compras (opcional)", type=["xlsx"], key="purchases")
    product_ads_up = st.file_uploader("Product Ads (opcional)", type=["xlsx"], key="pads")
    brand_ads_up = st.file_uploader("Brand Ads / Keywords (opcional)", type=["xlsx"], key="bads")
    st.divider()

    st.subheader("Parámetros")
    analisis_dias = st.selectbox("Ventana principal", [30, 90], index=0)
    margen_objetivo = st.slider("Margen ML objetivo %", min_value=5, max_value=40, value=20, step=1)
    alza_fuerte = st.slider("Alza de costo fuerte %", min_value=5, max_value=50, value=20, step=1)
    producto_muerto_dias = st.slider("Sin ventas para alertar (días)", min_value=30, max_value=180, value=90, step=30)

if not master_up or not sales_up:
    st.info("Sube al menos la maestra y el informe de ventas para comenzar.")
    st.stop()

sig = "|".join([
    file_signature(master_up),
    file_signature(sales_up),
    file_signature(purchases_up),
    file_signature(product_ads_up),
    file_signature(brand_ads_up),
    str(analisis_dias),
    str(margen_objetivo),
    str(alza_fuerte),
    str(producto_muerto_dias),
])

if st.session_state.get("model_sig") != sig:
    try:
        bundle = load_master_bundle(master_up.getvalue())
        sales_data = load_sales(sales_up.getvalue())
        purchases_data = load_purchases(purchases_up.getvalue() if purchases_up else b"")
        product_ads_data = load_product_ads(product_ads_up.getvalue()) if product_ads_up else {"raw": pd.DataFrame(), "summary": pd.DataFrame()}
        brand_ads_data = load_brand_ads(brand_ads_up.getvalue()) if brand_ads_up else {"raw": pd.DataFrame(), "campaign_summary": pd.DataFrame(), "top_keywords": pd.DataFrame()}
        model = build_model(bundle, sales_data, purchases_data, product_ads_data, brand_ads_data)
        st.session_state.model = model
        st.session_state.model_sig = sig
    except Exception as e:
        st.error(f"No pude construir el modelo: {e}")
        st.stop()

model = st.session_state.model
impact = model["impact"].copy()

# dynamic thresholds
impact["alerta_alza_fuerte"] = impact["variacion_costo_pct"] >= alza_fuerte
impact["alerta_margen_bajo"] = impact["margen_real_ml_pct"] < margen_objetivo
impact["alerta_sin_ventas_90d"] = impact["ventas_90d"].fillna(0) == 0
impact["alerta_ads_riesgo"] = impact["ads_sobre_ingreso_ml_pct_30d"] >= impact["margen_real_ml_pct"]
impact["alerta_promo_prox"] = impact["promo_prox_venci"].apply(lambda x: pd.notna(x) and 0 <= promo_status(x)[1] <= 7)

period_col_map = {
    30: {
        "ventas": "ventas_30d",
        "unidades": "unidades_30d",
        "ingresos": "ingresos_neto_30d",
        "participacion_ml": "participacion_ml_pct_30d",
        "cliente_empresa": "cliente_empresa_pct_30d",
        "cliente_persona": "cliente_persona_pct_30d",
        "patron": "patron_compra_30d",
        "mediana": "mediana_cantidad_30d",
        "ticket_empresa": "ticket_empresa_neto_30d",
        "ticket_persona": "ticket_persona_neto_30d",
    },
    90: {
        "ventas": "ventas_90d",
        "unidades": "unidades_90d",
        "ingresos": "ingresos_neto_90d",
        "participacion_ml": "participacion_ml_pct_30d",
        "cliente_empresa": "cliente_empresa_pct_30d",
        "cliente_persona": "cliente_persona_pct_30d",
        "patron": "patron_compra_90d",
        "mediana": "mediana_cantidad_90d",
        "ticket_empresa": "ticket_empresa_neto_30d",
        "ticket_persona": "ticket_persona_neto_30d",
    },
}
pcols = period_col_map[analisis_dias]

product_options = [f"{sku} — {model['sku_desc'].get(sku, '')}" for sku in model["sku_options"]]
default_idx = 0 if product_options else None
selected_label = st.selectbox("Buscar producto", product_options, index=default_idx)
selected_sku = selected_label.split(" — ")[0] if selected_label else ""

tabs = st.tabs([
    "Centro de Control Comercial",
    "Ficha de Producto",
    "Promociones",
    "Ads y Demanda",
    "Descargar"
])

# =============================
# TAB 1
# =============================
with tabs[0]:
    st.subheader("Resumen ejecutivo")

    total_skus = len(impact)
    criticos = len(impact[(impact["alerta_alza_fuerte"] & impact["alerta_margen_bajo"]) | (impact["estado_comercial"] == "Crítico por costo")])
    promo_prox = len(impact[impact["alerta_promo_prox"]])
    ads_riesgo = len(impact[impact["alerta_ads_riesgo"].fillna(False)])
    oportunidad = len(impact[impact["estado_comercial"] == "Oportunidad"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("SKUs analizados", total_skus)
    c2.metric("Críticos precio/costo", criticos)
    c3.metric("Promos por revisar", promo_prox)
    c4.metric("Ads en riesgo", ads_riesgo)
    c5.metric("Oportunidades", oportunidad)

    st.divider()
    st.subheader("Prioridad de cambio de precios")

    view = impact.copy()
    filtro_estado = st.multiselect(
        "Estado comercial",
        sorted(view["estado_comercial"].dropna().unique().tolist()),
        default=["Crítico por costo", "Revisar promo", "Ads en riesgo", "Margen crítico", "Oportunidad"] if len(view) else []
    )
    texto = st.text_input("Filtro SKU / descripción")
    if filtro_estado:
        view = view[view["estado_comercial"].isin(filtro_estado)]
    if texto:
        q = texto.lower().strip()
        view = view[
            view["SKU_norm"].astype(str).str.lower().str.contains(q, na=False) |
            view["DESCRIPCION"].astype(str).str.lower().str.contains(q, na=False)
        ]

    prioridad = view[[
        "SKU_norm", "DESCRIPCION", "estado_comercial", "score_prioridad", "variacion_costo_pct",
        "margen_real_ml_pct", "margen_real_tienda_pct", pcols["ingresos"], pcols["ventas"], "promo_prox_venci",
        "ads_sobre_ingreso_ml_pct_30d", "ads_roas", "TIENE_ADS"
    ]].copy().sort_values(["score_prioridad", pcols["ingresos"]], ascending=[False, False])

    prioridad = prioridad.rename(columns={
        "SKU_norm": "SKU",
        "DESCRIPCION": "Descripción",
        "estado_comercial": "Estado",
        "score_prioridad": "Score",
        "variacion_costo_pct": "Alza costo %",
        "margen_real_ml_pct": "Margen ML %",
        "margen_real_tienda_pct": "Margen tienda %",
        pcols["ingresos"]: f"Ingresos {analisis_dias}d",
        pcols["ventas"]: f"Ventas {analisis_dias}d",
        "promo_prox_venci": "Vencimiento promo",
        "ads_sobre_ingreso_ml_pct_30d": "Ads sobre ingreso ML %",
        "ads_roas": "ROAS ads",
        "TIENE_ADS": "Ads flag",
    })
    if not prioridad.empty:
        prioridad["Alza costo %"] = prioridad["Alza costo %"].map(fmt_pct)
        prioridad["Margen ML %"] = prioridad["Margen ML %"].map(fmt_pct)
        prioridad["Margen tienda %"] = prioridad["Margen tienda %"].map(fmt_pct)
        prioridad[f"Ingresos {analisis_dias}d"] = prioridad[f"Ingresos {analisis_dias}d"].map(fmt_money)
        prioridad["Vencimiento promo"] = prioridad["Vencimiento promo"].map(fmt_date)
        prioridad["Ads sobre ingreso ML %"] = prioridad["Ads sobre ingreso ML %"].map(fmt_pct)
        prioridad["ROAS ads"] = prioridad["ROAS ads"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
        st.dataframe(prioridad, use_container_width=True, hide_index=True)
    else:
        st.info("No hay SKUs para esos filtros.")

    st.divider()
    a, b = st.columns(2)
    with a:
        st.subheader("Alertas fuertes")
        alertas = impact[
            (impact["alerta_alza_fuerte"] & impact["alerta_margen_bajo"]) |
            (impact["alerta_promo_prox"]) |
            (impact["alerta_ads_riesgo"].fillna(False))
        ][["SKU_norm", "DESCRIPCION", "variacion_costo_pct", "margen_real_ml_pct", "promo_prox_venci", "ads_sobre_ingreso_ml_pct_30d"]].copy()
        if alertas.empty:
            st.success("No encontré alertas fuertes con la parametrización actual.")
        else:
            alertas.columns = ["SKU", "Descripción", "Alza costo %", "Margen ML %", "Vence promo", "Ads sobre ingreso ML %"]
            alertas["Alza costo %"] = alertas["Alza costo %"].map(fmt_pct)
            alertas["Margen ML %"] = alertas["Margen ML %"].map(fmt_pct)
            alertas["Vence promo"] = alertas["Vence promo"].map(fmt_date)
            alertas["Ads sobre ingreso ML %"] = alertas["Ads sobre ingreso ML %"].map(fmt_pct)
            st.dataframe(alertas.head(20), use_container_width=True, hide_index=True)

    with b:
        st.subheader("Lecturas estratégicas")
        mensajes = []
        if criticos > 0:
            mensajes.append(f"Hay {criticos} SKU(s) donde el costo y el margen ya están presionando un ajuste de precio.")
        if promo_prox > 0:
            mensajes.append(f"Hay {promo_prox} promo(s) por vencer pronto; conviene revisar si sostienen margen o solo volumen.")
        if ads_riesgo > 0:
            mensajes.append(f"Hay {ads_riesgo} SKU(s) donde la presión de ads podría estar consumiendo demasiado ingreso ML.")
        if oportunidad > 0:
            mensajes.append(f"Hay {oportunidad} SKU(s) con espacio comercial para revisar precio al alza sin partir desde una situación débil.")
        if not mensajes:
            mensajes.append("La cartera se ve relativamente estable con los umbrales actuales.")
        for msg in mensajes:
            st.write(f"- {msg}")

# =============================
# TAB 2
# =============================
with tabs[1]:
    st.subheader("Ficha de Producto")
    row = impact[impact["SKU_norm"] == selected_sku]
    if row.empty:
        st.warning("No encontré el SKU.")
    else:
        row = row.iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("SKU", selected_sku)
        c2.metric("Costo neto", fmt_money(row["COSTO_NETO"]))
        c3.metric("Margen ML", fmt_pct(row["margen_real_ml_pct"]))
        c4.metric("Margen tienda", fmt_pct(row["margen_real_tienda_pct"]))
        c5.metric("Estado", row["estado_comercial"])

        st.write(f"**Descripción:** {row['DESCRIPCION']}")
        st.write(f"**Ubicación:** {row['UBICACION'] if pd.notna(row['UBICACION']) else '—'}")

        a, b, c = st.columns(3)
        with a:
            st.markdown("**Precios y estructura**")
            st.write(f"Precio tienda bruto: {fmt_money(row['PRECIO_TIENDA_BRUTO'])}")
            st.write(f"Precio tienda neto: {fmt_money(row['PRECIO_TIENDA_NETO'])}")
            st.write(f"Monto en simulación ML: {fmt_money(row['MONTO_SIM_BRUTO'])}")
            st.write(f"Ingreso ML neto real: {fmt_money(row['MONTO_SIM_NETO'])}")

        with b:
            st.markdown(f"**Ventas {analisis_dias}d**")
            st.write(f"Ingresos netos: {fmt_money(row.get(pcols['ingresos']))}")
            st.write(f"Ventas: {int(safe_float(row.get(pcols['ventas']), 0))}")
            st.write(f"Unidades: {safe_float(row.get(pcols['unidades']), 0):.0f}")
            st.write(f"Participación ML: {fmt_pct(row.get(pcols['participacion_ml']))}")

        with c:
            st.markdown("**Compras y costo**")
            st.write(f"Última compra: {fmt_date(row.get('ultima_fecha'))}")
            st.write(f"Último precio compra: {fmt_money(row.get('ultimo_precio'))}")
            st.write(f"Proveedor: {row.get('ultimo_proveedor') if pd.notna(row.get('ultimo_proveedor')) else '—'}")
            st.write(f"Variación costo: {fmt_pct(row.get('variacion_costo_pct'))}")
            st.write(f"Días sin compra: {int(row['dias_sin_compra']) if pd.notna(row['dias_sin_compra']) else '—'}")

        st.divider()
        x1, x2 = st.columns([1.2, 1])
        with x1:
            st.markdown("**Proceso de decisión**")
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Patrón compra", str(row.get(pcols["patron"], "—")))
            d2.metric("Compra típica", f"{safe_float(row.get(pcols['mediana']), np.nan):.1f}" if pd.notna(row.get(pcols["mediana"])) else "—")
            d3.metric("% empresas", fmt_pct(row.get(pcols["cliente_empresa"])))
            d4.metric("% personas", fmt_pct(row.get(pcols["cliente_persona"])))
            st.write(f"Ticket empresa neto: {fmt_money(row.get(pcols['ticket_empresa']))}")
            st.write(f"Ticket persona neto: {fmt_money(row.get(pcols['ticket_persona']))}")

            lectura = []
            if pd.notna(row.get("variacion_costo_pct")) and row["variacion_costo_pct"] >= alza_fuerte:
                lectura.append("Alza fuerte de costo detectada.")
            if pd.notna(row.get("margen_real_ml_pct")) and row["margen_real_ml_pct"] < margen_objetivo:
                lectura.append("Margen ML bajo el objetivo actual.")
            if pd.notna(row.get("promo_prox_venci")) and 0 <= promo_status(row["promo_prox_venci"])[1] <= 7:
                lectura.append("Promo próxima a vencer.")
            if pd.notna(row.get("ads_sobre_ingreso_ml_pct_30d")) and pd.notna(row.get("margen_real_ml_pct")) and row["ads_sobre_ingreso_ml_pct_30d"] >= row["margen_real_ml_pct"]:
                lectura.append("Ads están presionando demasiado el ingreso ML.")
            if not lectura:
                lectura.append("Producto relativamente estable con los parámetros actuales.")
            for item in lectura:
                st.write(f"- {item}")

        with x2:
            st.markdown("**Promos y ads**")
            mlcs = model["mlc_map"].get(selected_sku, [])
            st.write(f"MLC asociados: {', '.join(mlcs) if mlcs else '—'}")
            st.write(f"Ads flag maestra: {'Sí' if row['TIENE_ADS'] else 'No'}")
            if row["TIENE_ADS"]:
                flags = [x for x in [row.get("ADS_FLAG_1"), row.get("ADS_FLAG_2")] if isinstance(x, str) and x.strip()]
                if flags:
                    st.write("Campañas maestra:")
                    for f in flags:
                        st.write(f"- {f}")
            st.write(f"Ads ROAS: {row['ads_roas']:.2f}" if pd.notna(row.get("ads_roas")) else "Ads ROAS: —")
            st.write(f"Ads sobre ingreso ML 30d: {fmt_pct(row.get('ads_sobre_ingreso_ml_pct_30d'))}")
            st.write(f"Margen ML después ads: {fmt_pct(row.get('margen_ml_despues_ads_pct'))}")

        st.divider()
        p1, p2 = st.columns([1, 1])
        with p1:
            st.markdown("**Promociones asociadas**")
            promos_sku = model["control_promos"][model["control_promos"]["SKU_norm"] == selected_sku].copy()
            if promos_sku.empty:
                st.info("No encontré promos de control para este SKU.")
            else:
                promos_show = promos_sku[["MLC_norm", "Precio_Promo", "Motivo", "Ads_Comentario", "Fecha_Vencimiento", "Estado_Promo"]].copy()
                promos_show["Precio_Promo"] = promos_show["Precio_Promo"].map(fmt_money)
                promos_show["Fecha_Vencimiento"] = promos_show["Fecha_Vencimiento"].map(fmt_date)
                promos_show.columns = ["MLC", "Precio promo", "Motivo", "Ads/Comentario", "Vencimiento", "Estado"]
                st.dataframe(promos_show, use_container_width=True, hide_index=True)

        with p2:
            st.markdown("**Relámpago asociado**")
            rel_sku = model["rel"][model["rel"]["SKU_norm"] == selected_sku].copy()
            if rel_sku.empty:
                st.info("No está en relámpago.")
            else:
                rel_sku["PRECIO_B2C"] = rel_sku["PRECIO_B2C"].map(fmt_money)
                rel_sku.columns = ["SKU", "Descripción", "Precio B2C", "Tipo", "Estado"]
                st.dataframe(rel_sku, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("**Timeline del producto**")
        timeline_rows = []

        purch_hist = model["purchase_map"].get(selected_sku, pd.DataFrame())
        if not purch_hist.empty:
            for _, r in purch_hist.iterrows():
                timeline_rows.append({
                    "Fecha": r["Fecha_dt"],
                    "Tipo": "Compra",
                    "Detalle": f"{r.get('Razón Social', '—')} | {fmt_money(r.get('Precio_Un_Num'))} | Cant {safe_float(r.get('Cantidad_num'), 0):.0f}",
                    "Valor": safe_float(r.get("Precio_Un_Num"), np.nan),
                })

        sales_hist = model["sales_raw"][model["sales_raw"]["SKU_norm"] == selected_sku].copy()
        if not sales_hist.empty:
            for _, r in sales_hist.iterrows():
                timeline_rows.append({
                    "Fecha": r["Fecha_dt"],
                    "Tipo": f"Venta {r['canal']}",
                    "Detalle": f"{r.get('Razón Social', '—')} | {r['tipo_cliente']} | Cant {safe_float(r['Cantidad_num'], 0):.0f}",
                    "Valor": safe_float(r["Total_Linea_Neto"], np.nan),
                })

        promos_hist = model["control_promos"][model["control_promos"]["SKU_norm"] == selected_sku].copy()
        if not promos_hist.empty:
            for _, r in promos_hist.iterrows():
                timeline_rows.append({
                    "Fecha": r["Fecha_Vencimiento"],
                    "Tipo": "Promo",
                    "Detalle": f"{r['Motivo']} | {fmt_money(r['Precio_Promo'])}",
                    "Valor": safe_float(r["Precio_Promo"], np.nan),
                })

        timeline_rows.append({
            "Fecha": pd.Timestamp(date.today()),
            "Tipo": "Snapshot actual",
            "Detalle": f"Tienda {fmt_money(row['PRECIO_TIENDA_BRUTO'])} | ML sim {fmt_money(row['MONTO_SIM_BRUTO'])}",
            "Valor": safe_float(row["MONTO_SIM_BRUTO"], np.nan),
        })

        timeline = pd.DataFrame(timeline_rows).dropna(subset=["Fecha"]).sort_values("Fecha", ascending=False)
        if timeline.empty:
            st.info("No hay eventos suficientes para construir timeline.")
        else:
            st.dataframe(timeline.head(200), use_container_width=True, hide_index=True)

        st.markdown("**Serie de ventas**")
        daily = model[f"sales_{analisis_dias}"]["timeline_daily"]
        daily_sku = daily[daily["SKU_norm"] == selected_sku].copy()
        if daily_sku.empty:
            st.info("No hay serie de ventas para la ventana seleccionada.")
        else:
            chart = daily_sku.pivot_table(index="Fecha_dt", columns="canal", values="Total_Linea_Neto", aggfunc="sum").fillna(0)
            st.line_chart(chart)

# =============================
# TAB 3
# =============================
with tabs[2]:
    st.subheader("Promociones")
    t1, t2 = st.columns([1.2, 1])

    with t1:
        st.markdown("**Control de promociones**")
        cp = model["control_promos"].copy()
        if cp.empty:
            st.info("No hay datos de control de promociones.")
        else:
            cp_editor = cp[["SKU_norm", "MLC_norm", "Descripcion", "Precio_Promo", "Motivo", "Ads_Comentario", "Fecha_Vencimiento"]].rename(columns={
                "SKU_norm": "SKU",
                "MLC_norm": "MLC",
                "Descripcion": "Descripción",
                "Precio_Promo": "Precio promocional",
                "Motivo": "Motivo",
                "Ads_Comentario": "Ads/Comentario",
                "Fecha_Vencimiento": "Vencimiento",
            })
            edited_cp = st.data_editor(cp_editor, use_container_width=True, hide_index=True, num_rows="dynamic", key="cp_editor")
            if st.button("Guardar control promociones"):
                new_cp = edited_cp.rename(columns={
                    "SKU": "SKU_norm",
                    "MLC": "MLC_norm",
                    "Descripción": "Descripcion",
                    "Precio promocional": "Precio_Promo",
                    "Motivo": "Motivo",
                    "Ads/Comentario": "Ads_Comentario",
                    "Vencimiento": "Fecha_Vencimiento",
                }).copy()
                new_cp["SKU_norm"] = new_cp["SKU_norm"].map(norm_sku)
                new_cp["MLC_norm"] = new_cp["MLC_norm"].map(norm_mlc)
                new_cp["Fecha_Vencimiento"] = pd.to_datetime(new_cp["Fecha_Vencimiento"], errors="coerce").dt.normalize()
                new_cp["Estado_Promo"] = new_cp["Fecha_Vencimiento"].apply(lambda x: promo_status(x)[0] if pd.notna(x) else "Sin fecha")
                model["control_promos"] = new_cp
                st.success("Control de promociones actualizado en memoria. Descárgalo desde la pestaña final.")
                st.rerun()

    with t2:
        st.markdown("**Relámpago mi página**")
        rel = model["rel"].copy()
        if rel.empty:
            rel = pd.DataFrame(columns=["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"])
        rel_editor = rel.rename(columns={
            "SKU_norm": "SKU",
            "DESCRIPCION": "Descripción",
            "PRECIO_B2C": "Precio B2C",
            "TIPO": "Tipo",
            "ESTADO": "Estado",
        })
        edited_rel = st.data_editor(rel_editor, use_container_width=True, hide_index=True, num_rows="dynamic", key="rel_editor")
        if st.button("Guardar relámpago"):
            new_rel = edited_rel.rename(columns={
                "SKU": "SKU_norm",
                "Descripción": "DESCRIPCION",
                "Precio B2C": "PRECIO_B2C",
                "Tipo": "TIPO",
                "Estado": "ESTADO",
            }).copy()
            new_rel["SKU_norm"] = new_rel["SKU_norm"].map(norm_sku)
            new_rel["PRECIO_B2C"] = new_rel["PRECIO_B2C"].map(lambda x: safe_float(x, np.nan))
            model["rel"] = new_rel[new_rel["SKU_norm"] != ""]
            st.success("Relámpago actualizado en memoria. Descárgalo desde la pestaña final.")
            st.rerun()

# =============================
# TAB 4
# =============================
with tabs[3]:
    st.subheader("Ads y Demanda")

    x1, x2 = st.columns(2)
    with x1:
        st.markdown("**Product Ads por SKU / MLC**")
        if model["product_ads"]["summary"].empty:
            st.info("No cargaste reporte de Product Ads.")
        else:
            pads = impact[["SKU_norm", "DESCRIPCION", "ads_mlc_activos", "ads_impresiones", "ads_clicks", "ads_inversion", "ads_ingresos_neto", "ads_acos_pct", "ads_roas", "margen_ml_despues_ads_pct"]].copy()
            pads = pads[pads["ads_mlc_activos"].fillna(0) > 0].sort_values(["ads_inversion", "ads_ingresos_neto"], ascending=False)
            pads.columns = ["SKU", "Descripción", "MLC con ads", "Impresiones", "Clics", "Inversión", "Ingresos netos ads", "ACOS %", "ROAS", "Margen ML post ads %"]
            pads["Inversión"] = pads["Inversión"].map(fmt_money)
            pads["Ingresos netos ads"] = pads["Ingresos netos ads"].map(fmt_money)
            pads["ACOS %"] = pads["ACOS %"].map(fmt_pct)
            pads["ROAS"] = pads["ROAS"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
            pads["Margen ML post ads %"] = pads["Margen ML post ads %"].map(fmt_pct)
            st.dataframe(pads.head(100), use_container_width=True, hide_index=True)

    with x2:
        st.markdown("**Brand Ads / keywords**")
        if model["brand_ads"]["campaign_summary"].empty:
            st.info("No cargaste reporte de Brand Ads / keywords.")
        else:
            camp = model["brand_ads"]["campaign_summary"].copy().sort_values(["Inversion", "Ingresos_Bruto"], ascending=False)
            camp["Inversion"] = camp["Inversion"].map(fmt_money)
            camp["Ingresos_Bruto"] = camp["Ingresos_Bruto"].map(fmt_money)
            camp["ACOS_calc"] = camp["ACOS_calc"].map(fmt_pct)
            camp["ROAS_calc"] = camp["ROAS_calc"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
            st.dataframe(camp.rename(columns={
                "Campaña": "Campaña",
                "Impresiones": "Impresiones",
                "Clics": "Clics",
                "Ingresos_Bruto": "Ingresos",
                "Inversion": "Inversión",
                "Ventas_Ads": "Ventas Ads",
                "Unidades_Ads": "Unidades Ads",
                "ACOS_calc": "ACOS %",
                "ROAS_calc": "ROAS",
            }), use_container_width=True, hide_index=True)

        st.markdown("**Top palabras clave**")
        if model["brand_ads"]["top_keywords"].empty:
            st.info("No hay keywords disponibles.")
        else:
            kw = model["brand_ads"]["top_keywords"][["Campaña", "Palabra_Clave", "Impresiones", "Clics", "Ingresos_Bruto", "Inversion", "Ventas_Ads", "ROAS", "ACOS"]].copy()
            kw["Ingresos_Bruto"] = kw["Ingresos_Bruto"].map(fmt_money)
            kw["Inversion"] = kw["Inversion"].map(fmt_money)
            kw["ROAS"] = kw["ROAS"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "—")
            kw["ACOS"] = kw["ACOS"].map(fmt_pct)
            st.dataframe(kw.head(50), use_container_width=True, hide_index=True)

# =============================
# TAB 5
# =============================
with tabs[4]:
    st.subheader("Descargar maestra actualizada")
    st.caption("Descarga una copia con la maestra original y las hojas editables que ajustaste en esta app.")

    if st.button("Preparar Excel"):
        payload = build_download_bytes(
            model["master"],
            model["rel"],
            model["control_promos"],
            model["bundle"]["file_bytes"],
            model["bundle"]["maestra_name"],
            model["bundle"]["rel_name"],
            model["bundle"]["control_name"],
        )
        st.session_state.download_bytes = payload

    if st.session_state.get("download_bytes"):
        st.download_button(
            "Descargar Excel actualizado",
            data=st.session_state.download_bytes,
            file_name="AURORA_CONTROL_COMERCIAL_ACTUALIZADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
