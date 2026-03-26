import re
from copy import deepcopy
from datetime import date
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Aurora | Ficha Comercial", page_icon="📈", layout="wide")

DEFAULT_MASTER_FILE = "MAESTRA PRECIOS Y PROMOS (2).xlsx"
DEFAULT_PURCHASES_FILE = "compras_23-03-2026 19_17_02.xlsx"
MASTER_SHEET = "MAESTRA de precios"
MAP_SHEET = "MLC -SKU"
PROMO_SHEET = "CONTROL DE PROMOCIONES"
PURCHASES_SHEET = "Reporte"
TARGET_EDIT_SHEETS = [MASTER_SHEET, MAP_SHEET, PROMO_SHEET]


# ----------------------------
# Helpers
# ----------------------------
def clean_col(col: str) -> str:
    col = str(col).replace("\n", " ").strip()
    return re.sub(r"\s+", " ", col)


def clean_cell(v):
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, np.generic):
        return v.item()
    return v


def parse_publication_ids(value) -> list[str]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    ids = re.findall(r"\d{8,14}", text)
    return [f"MLC{m}" for m in ids]


def coerce_numeric(series: pd.Series) -> pd.Series:
    if series is None:
        return series
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan, "NaT": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def coerce_date(series: pd.Series, dayfirst: bool = False) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst).dt.normalize()


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
    <span style="background:{bg};color:{fg};padding:0.25rem 0.55rem;border-radius:999px;font-size:0.82rem;font-weight:600;display:inline-block;margin-right:0.35rem;">{text}</span>
    """


def style_metric_card(title: str, value: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div style="border:1px solid #e5e7eb;border-radius:18px;padding:16px 18px;background:white;min-height:110px;">
            <div style="font-size:0.82rem;color:#64748b;margin-bottom:6px;">{title}</div>
            <div style="font-size:1.6rem;font-weight:700;line-height:1.1;">{value}</div>
            <div style="font-size:0.84rem;color:#64748b;margin-top:8px;">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [clean_col(c) for c in out.columns]
    return out


def find_local_file(filename: str) -> Path | None:
    candidates = [Path.cwd() / filename, Path(__file__).resolve().parent / filename, Path("/mnt/data") / filename]
    for p in candidates:
        if p.exists():
            return p
    return None


def load_all_sheets(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(BytesIO(file_bytes))
    sheets = {}
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet)
        sheets[sheet] = normalize_headers(df)
    return sheets


def initialize_state(master_bytes: bytes, master_name: str, purchases_bytes: bytes | None, purchases_name: str | None):
    file_sig = (master_name, len(master_bytes), hash(master_bytes[:5000]))
    purchases_sig = None if purchases_bytes is None else (purchases_name, len(purchases_bytes), hash(purchases_bytes[:5000]))

    if st.session_state.get("master_sig") != file_sig:
        all_sheets = load_all_sheets(master_bytes)
        st.session_state["master_sig"] = file_sig
        st.session_state["master_file_name"] = master_name
        st.session_state["master_file_bytes"] = master_bytes
        st.session_state["all_sheets"] = all_sheets
        st.session_state["edited_sheets"] = deepcopy(all_sheets)
        st.session_state["edit_log"] = []

    if st.session_state.get("purchases_sig") != purchases_sig:
        st.session_state["purchases_sig"] = purchases_sig
        st.session_state["purchases_file_name"] = purchases_name
        st.session_state["purchases_file_bytes"] = purchases_bytes


def get_working_sheets() -> dict[str, pd.DataFrame]:
    return st.session_state["edited_sheets"]


def derive_model(sheets: dict[str, pd.DataFrame]) -> dict:
    master = sheets.get(MASTER_SHEET, pd.DataFrame()).copy()
    mlc_map = sheets.get(MAP_SHEET, pd.DataFrame()).copy()
    promo = sheets.get(PROMO_SHEET, pd.DataFrame()).copy()

    if master.empty or mlc_map.empty or promo.empty:
        raise ValueError("Faltan hojas requeridas: MAESTRA de precios, MLC -SKU o CONTROL DE PROMOCIONES.")

    if "SKU" not in master.columns:
        raise ValueError("La hoja MAESTRA de precios no tiene la columna SKU.")
    master["SKU"] = coerce_numeric(master["SKU"]).astype("Int64")
    master = master[master["SKU"].notna()].copy()

    numeric_master_cols = [
        "ÚLTIMO COSTO", "MARGEN LOCAL", "PRECIO NETO", "PRECIO BRUTO", "MARGEN MELI 1", "NETO MELI 1",
        "PRECIO MELI REAL", "PRECIO B2C", "% DCTO", "MARGEN MELI 2", "NETO MELI 2", "VENTA BRUTO MELI 2",
        "UBIC", "CAMBIO DE PRECIO",
    ]
    for col in numeric_master_cols:
        if col in master.columns:
            master[col] = coerce_numeric(master[col])
    if "FECHA VENCI" in master.columns:
        master["FECHA VENCI"] = coerce_date(master["FECHA VENCI"]) 

    if "SKU" not in mlc_map.columns:
        raise ValueError("La hoja MLC -SKU no tiene la columna SKU.")
    pub_col = "Número de publicación" if "Número de publicación" in mlc_map.columns else mlc_map.columns[0]
    mlc_map["SKU"] = coerce_numeric(mlc_map["SKU"]).astype("Int64")
    mlc_map["MLC"] = (
        mlc_map[pub_col].astype(str).str.extract(r"(\d{8,14})", expand=False).map(lambda x: f"MLC{x}" if pd.notna(x) else np.nan)
    )
    mlc_map = mlc_map[mlc_map["SKU"].notna()].copy()

    sku_col = next((c for c in promo.columns if c.strip() == ""), None)
    if sku_col is None:
        sku_col = promo.columns[0]
    promo = promo.rename(columns={sku_col: "SKU"}).copy()
    promo["SKU"] = coerce_numeric(promo["SKU"]).astype("Int64")
    if "% F" in promo.columns:
        promo["% F"] = coerce_numeric(promo["% F"])
    if "Precio promocional" in promo.columns:
        promo["Precio promocional"] = coerce_numeric(promo["Precio promocional"])
    for c in ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]:
        if c in promo.columns:
            promo[c] = coerce_date(promo[c])
    if "Ads/Comentario" not in promo.columns:
        promo["Ads/Comentario"] = np.nan
    if "Motivo promoción" not in promo.columns:
        promo["Motivo promoción"] = np.nan

    if "N° Publicación" in promo.columns:
        promo["mlc_candidates"] = promo["N° Publicación"].apply(parse_publication_ids)
    else:
        promo["mlc_candidates"] = [[] for _ in range(len(promo))]

    promo_expanded = promo.explode("mlc_candidates").rename(columns={"mlc_candidates": "MLC"})
    promo_expanded["MLC"] = promo_expanded["MLC"].fillna("")

    campaign_cols = [c for c in ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"] if c in promo_expanded.columns]

    def calc_next_date(row):
        vals = [row[c] for c in campaign_cols if pd.notna(row[c])]
        if not vals:
            return pd.NaT
        return min(vals)

    promo_expanded["proxima_campana"] = promo_expanded.apply(calc_next_date, axis=1)
    today = pd.Timestamp(date.today())
    promo_expanded["dias_para_vencer"] = (promo_expanded["proxima_campana"] - today).dt.days

    product_base = master.merge(
        mlc_map.groupby("SKU")["MLC"].agg(lambda x: sorted(set(v for v in x if pd.notna(v)))).reset_index(name="mlc_list"),
        on="SKU", how="left"
    )
    product_base["mlc_list"] = product_base["mlc_list"].apply(lambda x: x if isinstance(x, list) else [])
    promo_by_sku = promo_expanded.groupby("SKU").size().reset_index(name="promo_rows_by_sku")
    product_base = product_base.merge(promo_by_sku, on="SKU", how="left")
    product_base["promo_rows_by_sku"] = product_base["promo_rows_by_sku"].fillna(0).astype(int)
    product_base["search_text"] = (
        product_base["SKU"].astype(str) + " "
        + product_base.get("DESCRIPCIÓN", pd.Series("", index=product_base.index)).fillna("").astype(str)
        + " " + product_base["mlc_list"].apply(lambda x: " ".join(x) if isinstance(x, list) else "")
    ).str.lower()

    return {
        "master": master,
        "mlc_map": mlc_map,
        "promo": promo,
        "promo_expanded": promo_expanded,
        "product_base": product_base,
    }


def get_product_view(sku: int, model: dict):
    product = model["product_base"].loc[model["product_base"]["SKU"] == sku].copy()
    if product.empty:
        return None, pd.DataFrame()
    product_row = product.iloc[0].to_dict()
    promos = model["promo_expanded"][model["promo_expanded"]["SKU"] == sku].copy()
    if promos.empty:
        mapped_mlcs = set(product_row.get("mlc_list", []) or [])
        if mapped_mlcs:
            promos = model["promo_expanded"][model["promo_expanded"]["MLC"].isin(mapped_mlcs)].copy()
    if not promos.empty:
        promos = promos.sort_values(by=["dias_para_vencer", "Precio promocional"], ascending=[True, True], na_position="last")
    return product_row, promos


def urgency_label(promos: pd.DataFrame) -> tuple[str, str]:
    if promos.empty:
        return "Sin promo", "secondary"
    min_days = promos["dias_para_vencer"].dropna()
    if min_days.empty:
        return "Sin fecha", "secondary"
    return urgency_from_days(min_days.min())


def agenda_dataframe(model: dict) -> pd.DataFrame:
    promo = model["promo_expanded"].copy()
    master = model["master"][[c for c in ["SKU", "DESCRIPCIÓN", "ÚLTIMO COSTO", "PRECIO MELI REAL", "COMENTARIO"] if c in model["master"].columns]].copy()
    df = promo.merge(master, on="SKU", how="left")
    df["urgencia"], df["urgencia_color"] = zip(*df["dias_para_vencer"].map(urgency_from_days))
    return df


def decision_rules(product: dict, promos: pd.DataFrame) -> list[str]:
    insights = []
    precio_meli = product.get("PRECIO MELI REAL")
    costo = product.get("ÚLTIMO COSTO")
    if promos.empty:
        insights.append("Sin promociones vinculadas en el control.")
        if pd.notna(product.get("FECHA VENCI")):
            delta = (product["FECHA VENCI"] - pd.Timestamp(date.today())).days
            if delta <= 2:
                insights.append("La maestra indica vencimiento próximo, pero no se detectó promo asociada en control.")
        return insights
    min_days = promos["dias_para_vencer"].dropna()
    if not min_days.empty:
        md = int(min_days.min())
        if md == 0:
            insights.append("Tiene promo que vence hoy. Requiere decisión inmediata.")
        elif md == 1:
            insights.append("Tiene promo que vence mañana. Conviene definir continuidad hoy.")
        elif md == 2:
            insights.append("Tiene promo que vence pasado mañana.")
        elif md < 0:
            insights.append("Tiene al menos una promo vencida.")
        elif md <= 7:
            insights.append("Tiene promo dentro de los próximos 7 días.")
    if promos["MLC"].nunique(dropna=True) > 1:
        insights.append("El SKU tiene múltiples publicaciones asociadas.")
    if "Ads/Comentario" in promos.columns and promos["Ads/Comentario"].fillna("").eq("").any():
        insights.append("Hay promos sin comentario Ads.")
    if pd.notna(costo) and "Precio promocional" in promos.columns and (promos["Precio promocional"].dropna() <= float(costo)).any():
        insights.append("Al menos una promo tiene precio promocional menor o igual al costo.")
    if pd.notna(precio_meli) and "Precio promocional" in promos.columns and (promos["Precio promocional"].dropna() < float(precio_meli)).any():
        insights.append("Hay precio promocional más agresivo que el precio Meli base.")
    return insights


def normalize_purchases(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = normalize_headers(df)
    if "SKU" not in out.columns:
        return pd.DataFrame()
    out["SKU"] = coerce_numeric(out["SKU"]).astype("Int64")
    if "Fecha" in out.columns:
        out["Fecha"] = coerce_date(out["Fecha"], dayfirst=True)
    if "Precio Un." in out.columns:
        out["Precio Un."] = coerce_numeric(out["Precio Un."])
    if "Razón Social" not in out.columns:
        out["Razón Social"] = np.nan
    return out[out["SKU"].notna()].copy()


def get_purchase_view(purchases_df: pd.DataFrame, sku: int):
    if purchases_df is None or purchases_df.empty:
        return None, pd.DataFrame()
    p = purchases_df[purchases_df["SKU"] == sku].copy()
    if p.empty:
        return None, pd.DataFrame()
    sort_cols = [c for c in ["Fecha", "#"] if c in p.columns]
    if sort_cols:
        p = p.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position="last")
    latest = p.iloc[0].to_dict()
    return latest, p


def build_updated_workbook(original_bytes: bytes, edited_sheets: dict[str, pd.DataFrame]) -> bytes:
    wb = load_workbook(filename=BytesIO(original_bytes))
    for sheet_name in TARGET_EDIT_SHEETS:
        if sheet_name not in wb.sheetnames or sheet_name not in edited_sheets:
            continue
        ws = wb[sheet_name]
        df = edited_sheets[sheet_name].copy()
        df = df.replace({pd.NaT: None, np.nan: None})
        max_existing_row = ws.max_row
        max_existing_col = ws.max_column
        for r in range(1, max_existing_row + 1):
            for c in range(1, max_existing_col + 1):
                ws.cell(r, c).value = None
        for c_idx, col in enumerate(df.columns, start=1):
            ws.cell(1, c_idx).value = col
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(r_idx, c_idx).value = clean_cell(value)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def log_edit(msg: str):
    st.session_state.setdefault("edit_log", []).append({"timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"), "acción": msg})


# ----------------------------
# Sidebar / load files
# ----------------------------
st.title("📈 Ficha comercial y agenda de promociones")
st.caption("Vista consolidada desde MAESTRA de precios + MLC -SKU + CONTROL DE PROMOCIONES")

with st.sidebar:
    st.header("Archivos")
    uploaded_master = st.file_uploader("Sube la maestra actualizada", type=["xlsx"], key="master_uploader")
    default_master = find_local_file(DEFAULT_MASTER_FILE)

    if uploaded_master is not None:
        master_bytes = uploaded_master.getvalue()
        master_name = uploaded_master.name
    elif default_master is not None:
        master_bytes = default_master.read_bytes()
        master_name = default_master.name
        st.info(f"Usando maestra local: {master_name}")
    else:
        st.warning("Sube el archivo maestro para comenzar.")
        st.stop()

    uploaded_purchases = st.file_uploader("Sube compras actualizado (opcional)", type=["xlsx"], key="purchases_uploader")
    default_purchases = find_local_file(DEFAULT_PURCHASES_FILE)
    if uploaded_purchases is not None:
        purchases_bytes = uploaded_purchases.getvalue()
        purchases_name = uploaded_purchases.name
    elif default_purchases is not None:
        purchases_bytes = default_purchases.read_bytes()
        purchases_name = default_purchases.name
        st.caption(f"Compras local: {purchases_name}")
    else:
        purchases_bytes = None
        purchases_name = None

initialize_state(master_bytes, master_name, purchases_bytes, purchases_name)

try:
    sheets = get_working_sheets()
    model = derive_model(sheets)
except Exception as e:
    st.error(f"No pude leer o modelar la maestra: {e}")
    st.stop()

purchases_df = pd.DataFrame()
if purchases_bytes is not None:
    try:
        purchases_df = normalize_purchases(pd.read_excel(BytesIO(purchases_bytes), sheet_name=PURCHASES_SHEET))
    except Exception as e:
        st.warning(f"No pude leer el archivo de compras: {e}")

agenda_df = agenda_dataframe(model)
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
    style_metric_card("Vencen en 7 días", str(week_count), "agenda comercial")

updated_workbook = build_updated_workbook(st.session_state["master_file_bytes"], st.session_state["edited_sheets"])
with st.sidebar:
    st.download_button(
        "Descargar maestra actualizada",
        data=updated_workbook,
        file_name=f"editado_{st.session_state['master_file_name']}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    if st.session_state.get("edit_log"):
        st.caption(f"Cambios acumulados: {len(st.session_state['edit_log'])}")


tab1, tab2, tab3 = st.tabs(["Consulta producto", "Agenda comercial", "Edición total"])

with tab1:
    st.subheader("Consulta producto")
    q1, q2 = st.columns([2, 1])
    with q1:
        query = st.text_input("Busca por SKU, descripción o MLC", placeholder="Ej: 110203002020, abrazadera, MLC1789384668")
    with q2:
        only_with_promos = st.toggle("Solo con promos asociadas", value=False)

    filtered = model["product_base"].copy()
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
            urg_txt, _ = urgency_label(model["promo_expanded"][model["promo_expanded"]["SKU"] == row["SKU"]])
            label = f"{int(row['SKU'])} · {desc} · {urg_txt}"
            options.append(label)
            option_map[label] = int(row["SKU"])
        selected_label = st.selectbox("Selecciona un producto", options, index=0)
        selected_sku = option_map[selected_label]

        product, promos = get_product_view(selected_sku, model)
        latest_buy, buy_hist = get_purchase_view(purchases_df, selected_sku)

        urg_txt, urg_color = urgency_label(promos)
        st.markdown("---")
        title = str(product.get("DESCRIPCIÓN", "Sin descripción"))
        sku_str = str(int(product["SKU"]))
        mlcs = product.get("mlc_list", []) or []
        st.markdown(
            f"""
            <div style="border:1px solid #e5e7eb;border-radius:18px;padding:18px 20px;background:#fafafa;">
                <div style="font-size:1.45rem;font-weight:700;">{title}</div>
                <div style="margin-top:8px;font-size:0.95rem;color:#475569;">SKU <b>{sku_str}</b> · UBIC <b>{product.get('UBIC', '—') if pd.notna(product.get('UBIC')) else '—'}</b></div>
                <div style="margin-top:12px;">{pill(urg_txt, urg_color)}{pill(f'{len(mlcs)} publicaciones mapeadas', 'secondary')}{pill(f'{len(promos)} filas promo', 'secondary')}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        a, b, c, d = st.columns(4)
        with a:
            st.markdown("#### Precio base")
            for label, key in [("Último costo", "ÚLTIMO COSTO"), ("Precio neto", "PRECIO NETO"), ("Precio bruto", "PRECIO BRUTO"), ("Precio Meli real", "PRECIO MELI REAL"), ("Precio B2C", "PRECIO B2C")]:
                if key in product:
                    st.write(f"**{label}:** {fmt_money(product.get(key))}")
        with b:
            st.markdown("#### Márgenes")
            for label, key in [("Margen local", "MARGEN LOCAL"), ("Margen Meli 1", "MARGEN MELI 1"), ("Margen Meli 2", "MARGEN MELI 2"), ("% dcto maestra", "% DCTO")]:
                if key in product:
                    st.write(f"**{label}:** {fmt_percent(product.get(key))}")
            if "FECHA VENCI" in product:
                st.write(f"**Fecha venci maestra:** {fmt_date(product.get('FECHA VENCI'))}")
        with c:
            st.markdown("#### Publicaciones / control")
            if mlcs:
                st.code("\n".join(mlcs))
            else:
                st.info("Sin publicaciones mapeadas.")
            st.write(f"**Comentario maestra:** {product.get('COMENTARIO') if pd.notna(product.get('COMENTARIO')) else '—'}")
        with d:
            st.markdown("#### Historial compras")
            if latest_buy:
                st.write(f"**Última compra:** {fmt_date(latest_buy.get('Fecha'))}")
                st.write(f"**Último precio compra:** {fmt_money(latest_buy.get('Precio Un.'))}")
                st.write(f"**Proveedor:** {latest_buy.get('Razón Social') or '—'}")
                if pd.notna(product.get("ÚLTIMO COSTO")) and pd.notna(latest_buy.get("Precio Un.")):
                    diff = float(latest_buy.get("Precio Un.")) - float(product.get("ÚLTIMO COSTO"))
                    st.write(f"**Dif. vs último costo maestra:** {fmt_money(diff)}")
            else:
                st.info("Sin historial de compras cargado para este SKU.")

        st.markdown("#### Promociones asociadas")
        if promos.empty:
            st.info("No encontré promociones vinculadas en CONTROL DE PROMOCIONES.")
        else:
            promo_view = promos.copy()
            promo_view["Estado"] = promo_view["dias_para_vencer"].map(lambda x: urgency_from_days(x)[0])
            promo_view["Campaña próxima"] = promo_view["proxima_campana"].map(fmt_date)
            if "% F" in promo_view.columns:
                promo_view["% F"] = promo_view["% F"].map(fmt_percent)
            if "Precio promocional" in promo_view.columns:
                promo_view["Precio promocional"] = promo_view["Precio promocional"].map(fmt_money)
            for col in ["Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"]:
                if col in promo_view.columns:
                    promo_view[col] = promo_view[col].map(fmt_date)
            show_cols = [c for c in ["MLC", "Descripción", "% F", "Precio promocional", "Motivo promoción", "Ads/Comentario", "Campaña próxima", "Estado", "Campaña 1", "Campaña 2", "Campaña 3", "Campaña 4"] if c in promo_view.columns]
            st.dataframe(promo_view[show_cols], use_container_width=True, hide_index=True)

        if latest_buy and not buy_hist.empty:
            st.markdown("#### Variación histórica de compra")
            chart_df = buy_hist[[c for c in ["Fecha", "Precio Un."] if c in buy_hist.columns]].dropna()
            if not chart_df.empty:
                chart_df = chart_df.sort_values("Fecha")
                st.line_chart(chart_df.set_index("Fecha")["Precio Un."], use_container_width=True)
            view_cols = [c for c in ["Fecha", "Razón Social", "Documento", "Folio", "Cantidad", "Precio Un.", "Total Línea"] if c in buy_hist.columns]
            show_hist = buy_hist[view_cols].copy()
            if "Precio Un." in show_hist.columns:
                show_hist["Precio Un."] = show_hist["Precio Un."].map(fmt_money)
            if "Total Línea" in show_hist.columns:
                show_hist["Total Línea"] = coerce_numeric(show_hist["Total Línea"]).map(fmt_money)
            if "Fecha" in show_hist.columns:
                show_hist["Fecha"] = show_hist["Fecha"].map(fmt_date)
            st.dataframe(show_hist, use_container_width=True, hide_index=True)

        st.markdown("#### Lectura comercial")
        for insight in decision_rules(product, promos):
            st.write(f"- {insight}")

with tab2:
    st.subheader("Agenda comercial")
    left, right = st.columns([1.15, 2.25])
    with left:
        status_filter = st.multiselect(
            "Filtrar por urgencia",
            ["Vence hoy", "Vence mañana", "Vence pasado mañana", "Próximos 7 días", "Vencida", "Sin fecha"],
            default=["Vence hoy", "Vence mañana", "Vence pasado mañana", "Próximos 7 días"],
        )
        solo_sin_ads = st.toggle("Solo sin comentario Ads", value=False)
        texto = st.text_input("Buscar SKU / descripción / MLC", placeholder="Ej: MLC1789 o abrazadera", key="agenda_text")

    ag = agenda_df.copy()
    ag["Estado"] = ag["dias_para_vencer"].map(lambda x: urgency_from_days(x)[0])
    if status_filter:
        masks = []
        for s in status_filter:
            if s == "Próximos 7 días":
                masks.append(ag["dias_para_vencer"].between(3, 7, inclusive="both"))
            elif s == "Vence hoy":
                masks.append(ag["dias_para_vencer"] == 0)
            elif s == "Vence mañana":
                masks.append(ag["dias_para_vencer"] == 1)
            elif s == "Vence pasado mañana":
                masks.append(ag["dias_para_vencer"] == 2)
            elif s == "Vencida":
                masks.append(ag["dias_para_vencer"] < 0)
            elif s == "Sin fecha":
                masks.append(ag["dias_para_vencer"].isna())
        if masks:
            mask = masks[0]
            for extra in masks[1:]:
                mask = mask | extra
            ag = ag[mask]
    if solo_sin_ads and "Ads/Comentario" in ag.columns:
        ag = ag[ag["Ads/Comentario"].fillna("").eq("")]
    if texto:
        t = texto.lower()
        ag = ag[
            ag["SKU"].astype(str).str.contains(re.escape(t), case=False, regex=True, na=False)
            | ag["MLC"].astype(str).str.contains(re.escape(t), case=False, regex=True, na=False)
            | ag["Descripción"].fillna("").astype(str).str.contains(re.escape(t), case=False, regex=True, na=False)
        ]
    ag = ag.sort_values(by=["dias_para_vencer", "SKU"], na_position="last")

    with right:
        st.caption("Ordenado por urgencia de vencimiento.")
        agenda_show = ag.copy()
        for col in ["Precio promocional", "ÚLTIMO COSTO", "PRECIO MELI REAL"]:
            if col in agenda_show.columns:
                agenda_show[col] = agenda_show[col].map(fmt_money)
        if "% F" in agenda_show.columns:
            agenda_show["% F"] = agenda_show["% F"].map(fmt_percent)
        if "proxima_campana" in agenda_show.columns:
            agenda_show["Próxima campaña"] = agenda_show["proxima_campana"].map(fmt_date)
        cols = [c for c in ["SKU", "MLC", "Descripción", "% F", "Precio promocional", "ÚLTIMO COSTO", "PRECIO MELI REAL", "Ads/Comentario", "Motivo promoción", "Próxima campaña", "Estado"] if c in agenda_show.columns]
        st.dataframe(agenda_show[cols], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Edición total")
    st.caption("Aquí sí puedes editar y esos cambios quedan reflejados en la maestra descargable. El archivo original no se sobrescribe automáticamente desde el navegador.")

    sku_text = st.text_input("SKU a editar", placeholder="Ej: 110203002020")
    if sku_text and sku_text.isdigit():
        sku_edit = int(sku_text)
        product, promos = get_product_view(sku_edit, model)
        if not product:
            st.warning("No encontré ese SKU en la maestra.")
        else:
            st.write(f"**Producto:** {product.get('DESCRIPCIÓN', '—')}")
            st.write(f"**SKU:** {sku_edit}")

            master_df = st.session_state["edited_sheets"][MASTER_SHEET]
            map_df = st.session_state["edited_sheets"][MAP_SHEET]
            promo_df = st.session_state["edited_sheets"][PROMO_SHEET]

            master_mask = coerce_numeric(master_df["SKU"]).astype("Int64") == sku_edit
            map_mask = coerce_numeric(map_df["SKU"]).astype("Int64") == sku_edit

            promo_work = promo_df.copy()
            promo_sku_col = next((c for c in promo_work.columns if clean_col(c) == ""), None)
            if promo_sku_col is None:
                promo_sku_col = promo_work.columns[0]
            promo_mask = coerce_numeric(promo_work[promo_sku_col]).astype("Int64") == sku_edit

            st.markdown("#### 1) Fila de maestra")
            st.caption("Puedes editar cualquier campo de la fila del producto.")
            edited_master = st.data_editor(
                master_df.loc[master_mask].reset_index().rename(columns={"index": "_row_id_"}),
                use_container_width=True,
                num_rows="fixed",
                key=f"master_editor_{sku_edit}",
                hide_index=True,
            )
            if st.button("Guardar cambios en maestra", key=f"save_master_{sku_edit}"):
                if not edited_master.empty:
                    row_id = int(edited_master.iloc[0]["_row_id_"])
                    new_row = edited_master.drop(columns=["_row_id_"]).iloc[0]
                    for col in new_row.index:
                        st.session_state["edited_sheets"][MASTER_SHEET].at[row_id, col] = new_row[col]
                    log_edit(f"Actualizada fila maestra de SKU {sku_edit}")
                    st.success("Cambios de maestra guardados en memoria.")
                    st.rerun()

            st.markdown("#### 2) Relación SKU ↔ publicaciones")
            st.caption("También puedes corregir o agregar publicaciones manualmente.")
            map_view = map_df.loc[map_mask].reset_index().rename(columns={"index": "_row_id_"})
            if map_view.empty:
                base_pub_col = "Número de publicación" if "Número de publicación" in map_df.columns else map_df.columns[0]
                map_view = pd.DataFrame([{"_row_id_": -1, base_pub_col: None, "SKU": sku_edit}])
            edited_map = st.data_editor(map_view, use_container_width=True, num_rows="dynamic", key=f"map_editor_{sku_edit}", hide_index=True)
            colm1, colm2 = st.columns([1, 1])
            with colm1:
                if st.button("Guardar publicaciones", key=f"save_map_{sku_edit}"):
                    kept = edited_map.drop(columns=["_row_id_"]).copy()
                    kept = kept[~(kept.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1))]
                    st.session_state["edited_sheets"][MAP_SHEET] = pd.concat([
                        map_df.loc[~map_mask], kept
                    ], ignore_index=True)
                    log_edit(f"Actualizada relación MLC-SKU de SKU {sku_edit}")
                    st.success("Publicaciones guardadas.")
                    st.rerun()
            with colm2:
                if st.button("Eliminar todas las publicaciones del SKU", key=f"clear_map_{sku_edit}"):
                    st.session_state["edited_sheets"][MAP_SHEET] = map_df.loc[~map_mask].copy().reset_index(drop=True)
                    log_edit(f"Eliminadas publicaciones mapeadas de SKU {sku_edit}")
                    st.success("Publicaciones eliminadas.")
                    st.rerun()

            st.markdown("#### 3) Control de promociones")
            st.caption("Aquí puedes editar fechas, Ads, comentarios, precios, campañas y cualquier otro campo del control.")
            promo_view = promo_df.loc[promo_mask].reset_index().rename(columns={"index": "_row_id_"})
            if promo_view.empty:
                empty_row = {col: None for col in promo_df.columns}
                empty_row[promo_sku_col] = sku_edit
                promo_view = pd.DataFrame([{"_row_id_": -1, **empty_row}])
            edited_promo = st.data_editor(promo_view, use_container_width=True, num_rows="dynamic", key=f"promo_editor_{sku_edit}", hide_index=True)
            cp1, cp2 = st.columns([1, 1])
            with cp1:
                if st.button("Guardar promociones", key=f"save_promo_{sku_edit}"):
                    kept = edited_promo.drop(columns=["_row_id_"]).copy()
                    kept = kept[~(kept.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1))]
                    st.session_state["edited_sheets"][PROMO_SHEET] = pd.concat([
                        promo_df.loc[~promo_mask], kept
                    ], ignore_index=True)
                    log_edit(f"Actualizado control de promociones de SKU {sku_edit}")
                    st.success("Promociones guardadas.")
                    st.rerun()
            with cp2:
                if st.button("Eliminar promos del SKU", key=f"clear_promo_{sku_edit}"):
                    st.session_state["edited_sheets"][PROMO_SHEET] = promo_df.loc[~promo_mask].copy().reset_index(drop=True)
                    log_edit(f"Eliminadas promociones de SKU {sku_edit}")
                    st.success("Promociones eliminadas.")
                    st.rerun()

            st.markdown("#### 4) Respaldo de cambios")
            if st.session_state.get("edit_log"):
                st.dataframe(pd.DataFrame(st.session_state["edit_log"]), use_container_width=True, hide_index=True)
            st.info("Cuando termines, descarga la maestra actualizada desde la barra lateral.")
    else:
        st.info("Ingresa un SKU numérico para editar todo su bloque de información.")

st.markdown("---")
st.caption("Modelo de unión: SKU (maestra) → MLC -SKU → CONTROL DE PROMOCIONES. Cambios editados se integran en la descarga de la maestra actualizada.")
