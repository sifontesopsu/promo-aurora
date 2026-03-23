import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Aurora | Ficha Comercial",
    page_icon="📈",
    layout="wide",
)


DEFAULT_FILE = "MAESTRA PRECIOS Y PROMOS (2).xlsx"
MASTER_SHEET = "MAESTRA de precios"
MAP_SHEET = "MLC -SKU"
PROMO_SHEET = "CONTROL DE PROMOCIONES"


# ----------------------------
# Helpers
# ----------------------------
def clean_col(col: str) -> str:
    col = str(col).replace("\n", " ").strip()
    col = re.sub(r"\s+", " ", col)
    return col


def coerce_numeric(series: pd.Series) -> pd.Series:
    if series is None:
        return series
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def coerce_date(series: pd.Series, dayfirst: bool = False) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst).dt.normalize()


def parse_publication_ids(value) -> list[str]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    ids = re.findall(r"\d{8,14}", text)
    return [f"MLC{m}" for m in ids]



def fmt_money(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        return f"${int(round(float(value))):,}".replace(",", ".")
    except Exception:
        return str(value)



def fmt_percent(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        v = float(value)
        if abs(v) <= 1:
            v *= 100
        return f"{v:.1f}%"
    except Exception:
        return str(value)



def fmt_date(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    try:
        return pd.Timestamp(value).strftime("%d-%m-%Y")
    except Exception:
        return str(value)



def urgency_from_days(days):
    if pd.isna(days):
        return "Sin fecha", "secondary"
    days = int(days)
    if days < 0:
        return "Vencida", "secondary"
    if days == 0:
        return "Vence hoy", "red"
    if days == 1:
        return "Vence mañana", "orange"
    if days == 2:
        return "Vence pasado mañana", "yellow"
    if days <= 7:
        return f"Vence en {days} días", "blue"
    return f"Vence en {days} días", "green"



def pill(text: str, color: str = "blue") -> str:
    colors = {
        "red": ("#7f1d1d", "#fee2e2"),
        "orange": ("#9a3412", "#ffedd5"),
        "yellow": ("#854d0e", "#fef9c3"),
        "green": ("#166534", "#dcfce7"),
        "blue": ("#1d4ed8", "#dbeafe"),
        "secondary": ("#334155", "#e2e8f0"),
    }
    fg, bg = colors.get(color, colors["blue"])
    return f"""
    <span style="
        background:{bg};
        color:{fg};
        padding:0.25rem 0.55rem;
        border-radius:999px;
        font-size:0.82rem;
        font-weight:600;
        display:inline-block;
        margin-right:0.35rem;
    ">{text}</span>
    """



def find_default_file() -> Path | None:
    candidates = [
        Path.cwd() / DEFAULT_FILE,
        Path(__file__).resolve().parent / DEFAULT_FILE,
        Path("/mount/src") / DEFAULT_FILE,
    ]
    for p in candidates:
        if p.exists():
            return p
    return None



def style_metric_card(title: str, value: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div style="
            border:1px solid #e5e7eb;
            border-radius:18px;
            padding:16px 18px;
            background:white;
            min-height:110px;
        ">
            <div style="font-size:0.82rem;color:#64748b;margin-bottom:6px;">{title}</div>
            <div style="font-size:1.6rem;font-weight:700;line-height:1.1;">{value}</div>
            <div style="font-size:0.84rem;color:#64748b;margin-top:8px;">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ----------------------------
# Data loading and modeling
# ----------------------------
@st.cache_data(show_spinner=False)
def load_workbook(source_bytes: bytes):
    excel = pd.ExcelFile(BytesIO(source_bytes))

    master = pd.read_excel(excel, sheet_name=MASTER_SHEET)
    mlc_map = pd.read_excel(excel, sheet_name=MAP_SHEET)
    promo = pd.read_excel(excel, sheet_name=PROMO_SHEET)

    master.columns = [clean_col(c) for c in master.columns]
    mlc_map.columns = [clean_col(c) for c in mlc_map.columns]
    promo.columns = [clean_col(c) for c in promo.columns]

    if "SKU" not in master.columns:
        raise ValueError("La hoja 'MAESTRA de precios' no tiene la columna SKU.")
    master["SKU"] = coerce_numeric(master["SKU"]).astype("Int64")
    master = master[master["SKU"].notna()].copy()

    numeric_master_cols = [
        "ÚLTIMO COSTO",
        "MARGEN LOCAL",
        "PRECIO NETO",
        "PRECIO BRUTO",
        "MARGEN MELI 1",
        "NETO MELI 1",
        "PRECIO MELI REAL",
        "PRECIO B2C",
        "% DCTO",
        "MARGEN MELI 2",
        "NETO MELI 2",
        "VENTA BRUTO MELI 2",
        "UBIC",
        "CAMBIO DE PRECIO",
    ]
    for col in numeric_master_cols:
        if col in master.columns:
            master[col] = coerce_numeric(master[col])

    master["FECHA VENCI"] = coerce_date(master["FECHA VENCI"]) if "FECHA VENCI" in master.columns else pd.NaT

    if "Número de publicación" not in mlc_map.columns or "SKU" not in mlc_map.columns:
        raise ValueError("La hoja 'MLC -SKU' no tiene las columnas esperadas.")
    mlc_map["SKU"] = coerce_numeric(mlc_map["SKU"]).astype("Int64")
    mlc_map["MLC"] = (
        mlc_map["Número de publicación"]
        .astype(str)
        .str.extract(r"(\d{8,14})", expand=False)
        .map(lambda x: f"MLC{x}" if pd.notna(x) else np.nan)
    )
    mlc_map = mlc_map[mlc_map["SKU"].notna() & mlc_map["MLC"].notna()].copy()
    mlc_map = mlc_map[["SKU", "MLC"]].drop_duplicates()

    sku_col = next((c for c in promo.columns if c.strip() == ""), None)
    if sku_col is None:
        sku_col = promo.columns[0]
    promo = promo.rename(columns={sku_col: "SKU"})
    promo["SKU"] = coerce_numeric(promo["SKU"]).astype("Int64")
    promo["% F"] = coerce_numeric(promo["% F"]) if "% F" in promo.columns else np.nan
    promo["Precio promocional"] = (
        coerce_numeric(promo["Precio promocional"])
        if "Precio promocional" in promo.columns
        else np.nan
    )
    for c in ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]:
        promo[c] = coerce_date(promo[c]) if c in promo.columns else pd.NaT

    if "Ads/Comentario" not in promo.columns:
        promo["Ads/Comentario"] = np.nan
    if "Motivo promoción" not in promo.columns:
        promo["Motivo promoción"] = np.nan

    promo["publication_list"] = promo["N° Publicación"].apply(parse_publication_ids)
    promo["cant_publicaciones_raw"] = promo["publication_list"].apply(len)

    promo_expanded = promo.explode("publication_list").rename(columns={"publication_list": "MLC"})
    promo_expanded["MLC"] = promo_expanded["MLC"].fillna("")
    promo_expanded = promo_expanded.copy()

    campaign_cols = ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]

    def calc_next_date(row):
        vals = [row[c] for c in campaign_cols if c in row.index and pd.notna(row[c])]
        if not vals:
            return pd.NaT
        return min(vals)

    promo_expanded["proxima_campana"] = promo_expanded.apply(calc_next_date, axis=1)

    today = pd.Timestamp(date.today())
    promo_expanded["dias_para_vencer"] = (promo_expanded["proxima_campana"] - today).dt.days

    product_base = master.merge(
        mlc_map.groupby("SKU")["MLC"].agg(lambda x: sorted(set(x))).reset_index(name="mlc_list"),
        on="SKU",
        how="left",
    )
    product_base["mlc_list"] = product_base["mlc_list"].apply(lambda x: x if isinstance(x, list) else [])

    promo_by_sku = promo_expanded.groupby("SKU").size().reset_index(name="promo_rows_by_sku")
    product_base = product_base.merge(promo_by_sku, on="SKU", how="left")
    product_base["promo_rows_by_sku"] = product_base["promo_rows_by_sku"].fillna(0).astype(int)

    return {
        "master": master,
        "mlc_map": mlc_map,
        "promo_raw": promo,
        "promo_expanded": promo_expanded,
        "product_base": product_base,
    }


@st.cache_data(show_spinner=False)
def load_purchases(source_bytes: bytes):
    excel = pd.ExcelFile(BytesIO(source_bytes))
    sheet_name = excel.sheet_names[0]
    compras = pd.read_excel(excel, sheet_name=sheet_name)
    compras.columns = [clean_col(c) for c in compras.columns]

    required = ["Fecha", "Razón Social", "SKU", "Cantidad", "Precio Un."]
    missing = [c for c in required if c not in compras.columns]
    if missing:
        raise ValueError(f"El archivo de compras no tiene estas columnas: {', '.join(missing)}")

    compras["SKU"] = coerce_numeric(compras["SKU"]).astype("Int64")
    compras["Fecha"] = coerce_date(compras["Fecha"], dayfirst=True)
    compras["Cantidad"] = coerce_numeric(compras["Cantidad"])
    compras["Precio Un."] = coerce_numeric(compras["Precio Un."])
    if "Descuento" in compras.columns:
        compras["Descuento"] = coerce_numeric(compras["Descuento"])
    if "Total Línea" in compras.columns:
        compras["Total Línea"] = coerce_numeric(compras["Total Línea"])

    compras = compras[compras["SKU"].notna()].copy()
    compras = compras[compras["Fecha"].notna()].copy()
    compras = compras[compras["Precio Un."].notna()].copy()
    compras = compras.sort_values(["SKU", "Fecha", "#"], ascending=[True, True, True])

    supplier_summary = (
        compras.groupby(["SKU", "Razón Social"], dropna=False)
        .agg(
            compras=("SKU", "size"),
            ultima_fecha=("Fecha", "max"),
            primer_fecha=("Fecha", "min"),
            precio_min=("Precio Un.", "min"),
            precio_max=("Precio Un.", "max"),
            precio_ult=("Precio Un.", lambda s: s.iloc[-1]),
            cantidad_total=("Cantidad", "sum"),
        )
        .reset_index()
        .sort_values(["SKU", "ultima_fecha", "compras", "cantidad_total"], ascending=[True, False, False, False])
    )

    stats = (
        compras.groupby("SKU")
        .agg(
            compras_registradas=("SKU", "size"),
            proveedores_distintos=("Razón Social", pd.Series.nunique),
            fecha_ultima_compra=("Fecha", "max"),
            fecha_primera_compra=("Fecha", "min"),
            precio_ultimo=("Precio Un.", lambda s: s.iloc[-1]),
            precio_min=("Precio Un.", "min"),
            precio_max=("Precio Un.", "max"),
            precio_prom=("Precio Un.", "mean"),
        )
        .reset_index()
    )
    stats["variacion_vs_min"] = np.where(
        stats["precio_min"].gt(0),
        stats["precio_ultimo"] / stats["precio_min"] - 1,
        np.nan,
    )
    stats["variacion_vs_max"] = np.where(
        stats["precio_max"].gt(0),
        stats["precio_ultimo"] / stats["precio_max"] - 1,
        np.nan,
    )

    return {
        "compras": compras,
        "supplier_summary": supplier_summary,
        "stats": stats,
        "sheet_name": sheet_name,
    }



def build_search_index(product_base: pd.DataFrame) -> pd.DataFrame:
    out = product_base.copy()
    out["search_text"] = (
        out["SKU"].astype(str)
        + " "
        + out.get("DESCRIPCIÓN", pd.Series("", index=out.index)).fillna("").astype(str)
        + " "
        + out["mlc_list"].apply(lambda x: " ".join(x) if isinstance(x, list) else "")
    ).str.lower()
    return out



def get_product_view(sku: int, data: dict):
    base = data["product_base"]
    promo_expanded = data["promo_expanded"]

    product = base.loc[base["SKU"] == sku].copy()
    if product.empty:
        return None, pd.DataFrame()

    product_row = product.iloc[0].to_dict()
    promos = promo_expanded[promo_expanded["SKU"] == sku].copy()

    mapped_mlcs = set(product_row.get("mlc_list", []) or [])
    if promos.empty and mapped_mlcs:
        promos = promo_expanded[promo_expanded["MLC"].isin(mapped_mlcs)].copy()

    if not promos.empty:
        promos["dias_para_vencer"] = (promos["proxima_campana"] - pd.Timestamp(date.today())).dt.days
        promos = promos.sort_values(
            by=["dias_para_vencer", "Precio promocional"],
            ascending=[True, True],
            na_position="last",
        )
    return product_row, promos



def get_purchase_view(sku: int, purchases_data: dict | None):
    if purchases_data is None:
        return None

    compras = purchases_data["compras"]
    supplier_summary = purchases_data["supplier_summary"]
    stats = purchases_data["stats"]

    hist = compras[compras["SKU"] == sku].copy().sort_values("Fecha")
    if hist.empty:
        return {
            "history": hist,
            "stats": None,
            "suppliers": pd.DataFrame(),
            "last_row": None,
        }

    stat_row = stats[stats["SKU"] == sku]
    stat_row = stat_row.iloc[0].to_dict() if not stat_row.empty else None
    suppliers = supplier_summary[supplier_summary["SKU"] == sku].copy()
    last_row = hist.sort_values(["Fecha", "#"], ascending=[False, False]).iloc[0].to_dict()

    return {
        "history": hist,
        "stats": stat_row,
        "suppliers": suppliers,
        "last_row": last_row,
    }



def decision_rules(product: dict, promos: pd.DataFrame, purchase_pack: dict | None = None) -> list[str]:
    insights = []

    precio_meli = product.get("PRECIO MELI REAL")
    costo = product.get("ÚLTIMO COSTO")
    comentario = product.get("COMENTARIO")

    if promos.empty:
        insights.append("Sin promociones vinculadas en el control.")
        if pd.notna(product.get("FECHA VENCI")):
            delta = (product["FECHA VENCI"] - pd.Timestamp(date.today())).days
            if delta <= 2:
                insights.append("La maestra indica vencimiento próximo, pero no se detectó promo asociada en control.")
    else:
        promo_dates = promos["dias_para_vencer"].dropna()
        if not promo_dates.empty:
            min_days = int(promo_dates.min())
            if min_days < 0:
                insights.append("Tiene al menos una promo vencida que requiere revisión.")
            elif min_days == 0:
                insights.append("Tiene promo que vence hoy. Requiere decisión inmediata.")
            elif min_days == 1:
                insights.append("Tiene promo que vence mañana. Conviene definir continuidad hoy.")
            elif min_days == 2:
                insights.append("Tiene promo que vence pasado mañana. Ya está entrando en ventana de control.")
            elif min_days <= 7:
                insights.append("Tiene promo dentro de los próximos 7 días. Conviene anticiparse.")

        if promos["MLC"].nunique(dropna=True) > 1:
            insights.append("El SKU tiene múltiples publicaciones asociadas. Conviene revisar consistencia entre campañas.")
        if promos["Ads/Comentario"].fillna("").eq("").any():
            insights.append("Hay promos sin comentario Ads. Falta contexto operativo.")
        if "Precio promocional" in promos.columns and pd.notna(costo):
            under_cost = promos["Precio promocional"].dropna() <= float(costo)
            if under_cost.any():
                insights.append("Al menos una promo tiene precio promocional menor o igual al costo.")
        if pd.notna(precio_meli) and "Precio promocional" in promos.columns:
            cheaper = promos["Precio promocional"].dropna() < float(precio_meli)
            if cheaper.any():
                insights.append("Hay precio promocional más agresivo que el precio Meli base, revisar margen real.")

    if pd.notna(comentario) and str(comentario).strip():
        insights.append("El producto ya tiene comentario interno en la maestra. Úsalo como contexto de decisión.")

    if purchase_pack is not None:
        stats = purchase_pack.get("stats")
        last_row = purchase_pack.get("last_row")
        if stats and last_row:
            days_since = (pd.Timestamp(date.today()) - pd.Timestamp(stats["fecha_ultima_compra"])).days
            if days_since <= 30:
                insights.append("La última compra es reciente. Puedes contrastar rápido costo actual versus costo de reposición.")
            else:
                insights.append("La última compra no es reciente. Ojo con decidir usando un costo que podría estar desactualizado.")
            if pd.notna(product.get("ÚLTIMO COSTO")) and pd.notna(stats.get("precio_ultimo")):
                delta = float(product.get("ÚLTIMO COSTO")) - float(stats.get("precio_ultimo"))
                if abs(delta) > 0:
                    if delta > 0:
                        insights.append("El último costo en maestra está por encima del último precio de compra detectado.")
                    else:
                        insights.append("El último costo en maestra está por debajo del último precio de compra detectado. Conviene revisar criterio de actualización.")
            if stats.get("proveedores_distintos", 0) > 1:
                insights.append("Este SKU ha sido comprado a más de un proveedor. Vale la pena comparar dispersión de precio antes de decidir promo.")

    return insights



def urgency_label(promos: pd.DataFrame) -> tuple[str, str]:
    if promos.empty:
        return "Sin promo", "secondary"
    min_days = promos["dias_para_vencer"].dropna()
    if min_days.empty:
        return "Sin fecha", "secondary"
    return urgency_from_days(min_days.min())



def agenda_dataframe(data: dict) -> pd.DataFrame:
    promo = data["promo_expanded"].copy()
    master = data["master"][["SKU", "DESCRIPCIÓN", "ÚLTIMO COSTO", "PRECIO MELI REAL", "COMENTARIO"]].copy()

    df = promo.merge(master, on="SKU", how="left")
    df["urgencia"], df["urgencia_color"] = zip(*df["dias_para_vencer"].map(urgency_from_days))
    return df


# ----------------------------
# UI
# ----------------------------
st.title("📈 Ficha comercial y agenda de promociones")
st.caption("Vista consolidada desde MAESTRA de precios + MLC -SKU + CONTROL DE PROMOCIONES, con historial de compras opcional.")

with st.sidebar:
    st.header("Fuente de datos")
    uploaded = st.file_uploader("Sube tu Excel maestro", type=["xlsx"], key="master_uploader")
    default_path = find_default_file()

    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        file_label = uploaded.name
    elif default_path is not None:
        file_bytes = default_path.read_bytes()
        file_label = default_path.name
        st.info(f"Usando archivo local: {file_label}")
    else:
        st.warning("Sube el archivo Excel para comenzar.")
        st.stop()

    st.markdown("---")
    st.subheader("Archivo de compras")
    st.caption("Opcional. Úsalo para ver última compra, proveedor e historial de costo.")
    compras_upload = st.file_uploader(
        "Sube el archivo de compras actualizado",
        type=["xlsx"],
        key="compras_uploader",
    )
    if compras_upload is None:
        st.info("Cuando lo subas, la ficha mostrará historial de compras por SKU.")

try:
    data = load_workbook(file_bytes)
except Exception as e:
    st.error(f"No pude leer el archivo maestro: {e}")
    st.stop()

purchases_data = None
if compras_upload is not None:
    try:
        purchases_data = load_purchases(compras_upload.getvalue())
    except Exception as e:
        st.error(f"No pude leer el archivo de compras: {e}")

search_df = build_search_index(data["product_base"])
agenda_df = agenda_dataframe(data)

today_count = int((agenda_df["dias_para_vencer"] == 0).sum())
tomorrow_count = int((agenda_df["dias_para_vencer"] == 1).sum())
day2_count = int((agenda_df["dias_para_vencer"] == 2).sum())
week_count = int(agenda_df["dias_para_vencer"].between(0, 7, inclusive="both").sum())

c1, c2, c3, c4 = st.columns(4)
with c1:
    style_metric_card("Promos que vencen hoy", str(today_count), "prioridad máxima")
with c2:
    style_metric_card("Promos que vencen mañana", str(tomorrow_count), "ventana de reacción")
with c3:
    style_metric_card("Vencen pasado mañana", str(day2_count), "control preventivo")
with c4:
    if purchases_data is None:
        style_metric_card("Archivo de compras", "No cargado", "opcional para historial de costo")
    else:
        style_metric_card(
            "Archivo de compras",
            str(len(purchases_data["compras"])),
            f"líneas válidas desde {purchases_data['sheet_name']}",
        )

tab1, tab2, tab3 = st.tabs(["Consulta producto", "Agenda comercial", "Edición puntual"])

with tab1:
    st.subheader("Consulta producto")

    q1, q2 = st.columns([2, 1])
    with q1:
        query = st.text_input(
            "Busca por SKU, descripción o MLC",
            placeholder="Ej: 110203002020, abrazadera, MLC1789384668",
        )
    with q2:
        only_with_promos = st.toggle("Solo con promos asociadas", value=False)

    filtered = search_df.copy()
    if query:
        q = query.strip().lower()
        filtered = filtered[filtered["search_text"].str.contains(re.escape(q), regex=True, na=False)]
    if only_with_promos:
        filtered = filtered[filtered["promo_rows_by_sku"] > 0]

    filtered = filtered.sort_values(by=["promo_rows_by_sku", "SKU"], ascending=[False, True])

    if filtered.empty:
        st.warning("No encontré productos con ese criterio.")
    else:
        options = []
        option_map = {}
        for _, row in filtered.head(250).iterrows():
            desc = str(row.get("DESCRIPCIÓN", ""))[:90]
            urg_txt, _ = urgency_label(data["promo_expanded"][data["promo_expanded"]["SKU"] == row["SKU"]])
            label = f"{int(row['SKU'])} · {desc} · {urg_txt}"
            options.append(label)
            option_map[label] = int(row["SKU"])

        selected_label = st.selectbox("Selecciona un producto", options, index=0)
        selected_sku = option_map[selected_label]

        product, promos = get_product_view(selected_sku, data)
        purchase_pack = get_purchase_view(selected_sku, purchases_data)

        if product:
            urg_txt, urg_color = urgency_label(promos)
            title = str(product.get("DESCRIPCIÓN", "Sin descripción"))
            sku_str = str(int(product["SKU"]))
            mlcs = product.get("mlc_list", []) or []

            st.markdown("---")
            st.markdown(
                f"""
                <div style="border:1px solid #e5e7eb;border-radius:18px;padding:18px 20px;background:#fafafa;">
                    <div style="font-size:1.45rem;font-weight:700;">{title}</div>
                    <div style="margin-top:8px;font-size:0.95rem;color:#475569;">
                        SKU <b>{sku_str}</b> · UBIC <b>{product.get('UBIC', '—') if pd.notna(product.get('UBIC')) else '—'}</b>
                    </div>
                    <div style="margin-top:12px;">
                        {pill(urg_txt, urg_color)}
                        {pill(f"{len(mlcs)} publicaciones mapeadas", "secondary")}
                        {pill(f"{len(promos)} filas promo", "secondary")}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            a, b, c = st.columns(3)
            with a:
                st.markdown("#### Precio base")
                st.write(f"**Último costo:** {fmt_money(product.get('ÚLTIMO COSTO'))}")
                st.write(f"**Precio neto:** {fmt_money(product.get('PRECIO NETO'))}")
                st.write(f"**Precio bruto:** {fmt_money(product.get('PRECIO BRUTO'))}")
                st.write(f"**Precio Meli real:** {fmt_money(product.get('PRECIO MELI REAL'))}")
                st.write(f"**Precio B2C:** {fmt_money(product.get('PRECIO B2C'))}")
                st.write(f"**Cambio de precio:** {fmt_money(product.get('CAMBIO DE PRECIO'))}")
            with b:
                st.markdown("#### Márgenes")
                st.write(f"**Margen local:** {fmt_percent(product.get('MARGEN LOCAL'))}")
                st.write(f"**Margen Meli 1:** {fmt_percent(product.get('MARGEN MELI 1'))}")
                st.write(f"**Margen Meli 2:** {fmt_percent(product.get('MARGEN MELI 2'))}")
                st.write(f"**% dcto maestra:** {fmt_percent(product.get('% DCTO'))}")
                st.write(f"**Fecha venci maestra:** {fmt_date(product.get('FECHA VENCI'))}")
            with c:
                st.markdown("#### Relación con publicaciones")
                if mlcs:
                    st.write("**MLC asociadas:**")
                    st.code("\n".join(mlcs))
                else:
                    st.info("No encontré publicaciones en la hoja MLC -SKU para este SKU.")
                st.write(f"**Comentario maestra:** {product.get('COMENTARIO') if pd.notna(product.get('COMENTARIO')) else '—'}")

            st.markdown("#### Promociones asociadas")
            if promos.empty:
                st.info("No encontré promociones vinculadas en CONTROL DE PROMOCIONES.")
            else:
                promo_view = promos.copy()
                promo_view["Estado"] = promo_view["dias_para_vencer"].map(lambda x: urgency_from_days(x)[0])
                promo_view["Campaña próxima"] = promo_view["proxima_campana"].dt.date
                promo_view["Campaña 1"] = promo_view["Campaña 1"].dt.date
                promo_view["Campaña 2"] = promo_view["Campaña 2"].dt.date
                promo_view["Campaña 3"] = promo_view["Campaña 3"].dt.date
                promo_view["Campaña 4"] = promo_view["Campaña 4"].dt.date
                promo_view["% F"] = promo_view["% F"].map(fmt_percent)
                promo_view["Precio promocional"] = promo_view["Precio promocional"].map(fmt_money)

                show_cols = [
                    "MLC",
                    "Descripción",
                    "% F",
                    "Precio promocional",
                    "Motivo promoción",
                    "Ads/Comentario",
                    "Campaña próxima",
                    "Estado",
                    "Campaña 1",
                    "Campaña 2",
                    "Campaña 3",
                    "Campaña 4",
                ]
                show_cols = [c for c in show_cols if c in promo_view.columns]
                st.dataframe(promo_view[show_cols], use_container_width=True, hide_index=True)

            st.markdown("#### Historial de compras")
            if purchases_data is None:
                st.info("Sube el archivo de compras en la barra lateral para activar esta sección.")
            elif purchase_pack is None or purchase_pack.get("stats") is None:
                st.warning("No encontré compras para este SKU en el archivo cargado.")
            else:
                stats = purchase_pack["stats"]
                last_row = purchase_pack["last_row"]
                suppliers = purchase_pack["suppliers"].copy()
                history = purchase_pack["history"].copy()

                p1, p2, p3, p4 = st.columns(4)
                with p1:
                    style_metric_card("Última compra", fmt_date(stats.get("fecha_ultima_compra")), f"{int(stats.get('compras_registradas', 0))} compras registradas")
                with p2:
                    style_metric_card("Último precio compra", fmt_money(stats.get("precio_ultimo")), f"rango {fmt_money(stats.get('precio_min'))} a {fmt_money(stats.get('precio_max'))}")
                with p3:
                    style_metric_card("Proveedor última compra", str(last_row.get("Razón Social", "—")), f"folio {last_row.get('Folio', '—')}")
                with p4:
                    style_metric_card("Proveedores históricos", str(int(stats.get("proveedores_distintos", 0))), f"promedio {fmt_money(stats.get('precio_prom'))}")

                if pd.notna(product.get("ÚLTIMO COSTO")) and pd.notna(stats.get("precio_ultimo")):
                    delta_abs = float(product.get("ÚLTIMO COSTO")) - float(stats.get("precio_ultimo"))
                    delta_pct = delta_abs / float(stats.get("precio_ultimo")) if float(stats.get("precio_ultimo")) else np.nan
                    color = "green" if delta_abs <= 0 else "orange"
                    st.markdown(
                        pill(
                            f"Diferencia vs último precio de compra: {fmt_money(delta_abs)} ({fmt_percent(delta_pct)})",
                            color,
                        ),
                        unsafe_allow_html=True,
                    )

                chart_df = history[["Fecha", "Precio Un."]].copy().sort_values("Fecha")
                chart_df = chart_df.rename(columns={"Precio Un.": "Precio unitario"}).set_index("Fecha")
                st.line_chart(chart_df)

                left_hist, right_hist = st.columns([1.2, 1.8])
                with left_hist:
                    st.markdown("##### Proveedores del SKU")
                    sup_view = suppliers.copy()
                    sup_view["ultima_fecha"] = sup_view["ultima_fecha"].map(fmt_date)
                    sup_view["primer_fecha"] = sup_view["primer_fecha"].map(fmt_date)
                    sup_view["precio_min"] = sup_view["precio_min"].map(fmt_money)
                    sup_view["precio_max"] = sup_view["precio_max"].map(fmt_money)
                    sup_view["precio_ult"] = sup_view["precio_ult"].map(fmt_money)
                    cols_sup = ["Razón Social", "compras", "cantidad_total", "primer_fecha", "ultima_fecha", "precio_min", "precio_max", "precio_ult"]
                    st.dataframe(sup_view[cols_sup], use_container_width=True, hide_index=True)
                with right_hist:
                    st.markdown("##### Detalle histórico")
                    hist_view = history.sort_values(["Fecha", "#"], ascending=[False, False]).copy()
                    hist_view["Fecha"] = hist_view["Fecha"].map(fmt_date)
                    hist_view["Precio Un."] = hist_view["Precio Un."].map(fmt_money)
                    if "Total Línea" in hist_view.columns:
                        hist_view["Total Línea"] = hist_view["Total Línea"].map(fmt_money)
                    cols_hist = ["Fecha", "Razón Social", "Documento", "Folio", "Cantidad", "Precio Un.", "Total Línea", "Concepto / Artículo"]
                    cols_hist = [c for c in cols_hist if c in hist_view.columns]
                    st.dataframe(hist_view[cols_hist], use_container_width=True, hide_index=True, height=320)

            st.markdown("#### Lectura comercial")
            for insight in decision_rules(product, promos, purchase_pack):
                st.write(f"- {insight}")

with tab2:
    st.subheader("Agenda comercial")

    left, right = st.columns([1.2, 2.2])
    with left:
        status_filter = st.multiselect(
            "Filtrar por urgencia",
            ["Vence hoy", "Vence mañana", "Vence pasado mañana", "Próximos 7 días", "Vencida", "Sin fecha"],
            default=["Vence hoy", "Vence mañana", "Vence pasado mañana", "Próximos 7 días"],
        )
        solo_sin_ads = st.toggle("Solo sin comentario Ads", value=False)
        texto = st.text_input("Buscar SKU / descripción / MLC", placeholder="Ej: MLC1789 o abrazadera")

    ag = agenda_df.copy()
    ag["Estado"] = ag["dias_para_vencer"].map(lambda x: urgency_from_days(x)[0])

    if status_filter:
        buckets = []
        for s in status_filter:
            if s == "Próximos 7 días":
                buckets.append(ag["dias_para_vencer"].between(3, 7, inclusive="both"))
            elif s == "Vence hoy":
                buckets.append(ag["dias_para_vencer"] == 0)
            elif s == "Vence mañana":
                buckets.append(ag["dias_para_vencer"] == 1)
            elif s == "Vence pasado mañana":
                buckets.append(ag["dias_para_vencer"] == 2)
            elif s == "Vencida":
                buckets.append(ag["dias_para_vencer"] < 0)
            elif s == "Sin fecha":
                buckets.append(ag["dias_para_vencer"].isna())
        if buckets:
            mask = buckets[0].copy()
            for extra in buckets[1:]:
                mask = mask | extra
            ag = ag[mask]

    if solo_sin_ads:
        ag = ag[ag["Ads/Comentario"].fillna("").eq("")]

    if texto:
        text_l = texto.lower()
        mask = (
            ag["SKU"].astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
            | ag["MLC"].astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
            | ag["Descripción"].fillna("").astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
        )
        ag = ag[mask]

    ag = ag.sort_values(by=["dias_para_vencer", "SKU"], na_position="last")

    with right:
        st.caption("Ordenado por urgencia de vencimiento.")
        agenda_show = ag.copy()
        agenda_show["Precio promocional"] = agenda_show["Precio promocional"].map(fmt_money)
        agenda_show["ÚLTIMO COSTO"] = agenda_show["ÚLTIMO COSTO"].map(fmt_money)
        agenda_show["PRECIO MELI REAL"] = agenda_show["PRECIO MELI REAL"].map(fmt_money)
        agenda_show["% F"] = agenda_show["% F"].map(fmt_percent)
        agenda_show["Próxima campaña"] = agenda_show["proxima_campana"].dt.date

        cols = [
            "SKU",
            "MLC",
            "Descripción",
            "% F",
            "Precio promocional",
            "ÚLTIMO COSTO",
            "PRECIO MELI REAL",
            "Ads/Comentario",
            "Motivo promoción",
            "Próxima campaña",
            "Estado",
        ]
        cols = [c for c in cols if c in agenda_show.columns]
        st.dataframe(agenda_show[cols], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Edición puntual")
    st.caption("Este módulo no sobreescribe el Excel original. Guarda cambios en memoria y te deja descargar un respaldo de ediciones.")

    sku_input = st.text_input("SKU a editar", placeholder="Ej: 110203002020")
    if sku_input and sku_input.isdigit():
        sku_edit = int(sku_input)
        product, promos = get_product_view(sku_edit, data)

        if not product:
            st.warning("No encontré ese SKU.")
        else:
            st.write(f"**Producto:** {product.get('DESCRIPCIÓN', '—')}")
            st.write(f"**SKU:** {sku_edit}")

            if "edits" not in st.session_state:
                st.session_state["edits"] = []

            with st.form("edit_form"):
                nuevo_comentario = st.text_area(
                    "Comentario maestra",
                    value="" if pd.isna(product.get("COMENTARIO")) else str(product.get("COMENTARIO")),
                    height=120,
                )
                nuevo_estado = st.selectbox(
                    "Estado interno sugerido",
                    ["Mantener", "Revisar hoy", "Revisar mañana", "Sin promo", "Crítico"],
                )
                nota_ads = st.text_input("Nota Ads / operativa", value="")
                submitted = st.form_submit_button("Guardar edición en memoria")

            if submitted:
                st.session_state["edits"].append(
                    {
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "SKU": sku_edit,
                        "Descripción": product.get("DESCRIPCIÓN"),
                        "Comentario maestra nuevo": nuevo_comentario,
                        "Estado sugerido": nuevo_estado,
                        "Nota Ads": nota_ads,
                    }
                )
                st.success("Edición guardada en memoria.")

            if st.session_state["edits"]:
                edits_df = pd.DataFrame(st.session_state["edits"])
                st.markdown("#### Ediciones acumuladas")
                st.dataframe(edits_df, use_container_width=True, hide_index=True)

                output = BytesIO()
                edits_df.to_excel(output, index=False)
                output.seek(0)
                st.download_button(
                    "Descargar respaldo de ediciones",
                    data=output,
                    file_name="ediciones_promos_y_precios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        st.info("Ingresa un SKU numérico para editar comentarios o decisiones internas.")

st.markdown("---")
st.caption(
    "Modelo de unión: SKU (maestra) → MLC -SKU → CONTROL DE PROMOCIONES. "
    "Si cargas el archivo de compras, la ficha agrega última compra, proveedores e historial de precio unitario."
)
