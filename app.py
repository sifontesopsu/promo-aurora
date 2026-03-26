
import re
from datetime import date
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Aurora | Ficha Comercial", page_icon="📈", layout="wide")

DEFAULT_FILE = "MAESTRA PRECIOS Y PROMOS (3).xlsx"
MASTER_SHEET = "MAESTRA de precios"
MAP_SHEET = "MLC -SKU"

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

PROMO_SLOTS = [
    {"slot": 1, "label": "Promo 1", "mlc_col": "MLC", "price_col": "PRECIO B2C PUBLICADO ", "date_col": "FECHA VENCI", "comment_col": "COMENTARIO"},
    {"slot": 2, "label": "Promo 2", "mlc_col": "MLC.1", "price_col": "PRECIO B2C", "date_col": "FECHA VENCI.1", "comment_col": "COMENTARIO.1"},
]


def clean_col(col: str) -> str:
    col = str(col).replace("\n", " ").strip()
    return re.sub(r"\s+", " ", col)


def coerce_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "nan": np.nan, "None": np.nan, "NaT": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def coerce_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.normalize()


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
        return pd.Timestamp(value).strftime("%d/%m/%Y")
    except Exception:
        return str(value)


def normalize_sku(x):
    if pd.isna(x):
        return None
    s = str(x).strip().replace(".0", "")
    digits = re.sub(r"\D", "", s)
    return digits or None


def normalize_mlc(x):
    if pd.isna(x):
        return None
    s = str(x).upper().strip()
    if not s:
        return None
    digits = re.findall(r"\d{8,14}", s)
    if not digits:
        return None
    return f"MLC{digits[0]}"


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
    return f'<span style="background:{bg};color:{fg};padding:0.25rem 0.55rem;border-radius:999px;font-size:0.82rem;font-weight:600;display:inline-block;margin-right:0.35rem;">{text}</span>'


def find_default_file() -> Path | None:
    p = Path(f"/mnt/data/{DEFAULT_FILE}")
    return p if p.exists() else None


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


@st.cache_data(show_spinner=False)
def load_workbook(source_bytes: bytes):
    excel = pd.ExcelFile(BytesIO(source_bytes))
    sheets = {name: pd.read_excel(excel, sheet_name=name) for name in excel.sheet_names}

    master = sheets.get(MASTER_SHEET, pd.DataFrame()).copy()
    mlc_map = sheets.get(MAP_SHEET, pd.DataFrame()).copy()

    master.columns = [clean_col(c) for c in master.columns]
    mlc_map.columns = [clean_col(c) for c in mlc_map.columns]

    if "Unnamed: 12" in master.columns and "MLC_aux" not in master.columns:
        master = master.rename(columns={"Unnamed: 12": "MLC_aux"})

    if "SKU" not in master.columns:
        raise ValueError("La hoja 'MAESTRA de precios' no tiene la columna SKU.")

    master["SKU"] = coerce_numeric(master["SKU"]).astype("Int64")
    master = master[master["SKU"].notna()].copy()

    numeric_master_cols = [
        "ÚLTIMO COSTO", "MARGEN LOCAL", "PRECIO NETO", "PRECIO BRUTO",
        "MARGEN MELI 1", "NETO MELI 1", "MONTO EN SIMULACIÓN", "CAMBIO DE PRECIO",
        "PRECIO B2C PUBLICADO ", "PRECIO B2C", "% DCTO", "% DCTO.1", "UBIC"
    ]
    for col in numeric_master_cols:
        if col in master.columns:
            master[col] = coerce_numeric(master[col])

    for col in ["FECHA VENCI", "FECHA VENCI.1"]:
        if col in master.columns:
            master[col] = coerce_date(master[col])

    master["SKU_norm"] = master["SKU"].astype("Int64").astype(str)

    if "Número de publicación" not in mlc_map.columns or "SKU" not in mlc_map.columns:
        raise ValueError("La hoja 'MLC -SKU' no tiene las columnas esperadas.")

    mlc_map["SKU"] = coerce_numeric(mlc_map["SKU"]).astype("Int64")
    mlc_map["MLC"] = mlc_map["Número de publicación"].apply(normalize_mlc)
    mlc_map = mlc_map[mlc_map["SKU"].notna() & mlc_map["MLC"].notna()].copy()
    mlc_map["SKU_norm"] = mlc_map["SKU"].astype(str)
    mlc_map = mlc_map[["SKU", "SKU_norm", "MLC"]].drop_duplicates()

    promos = build_master_promos(master)

    product_base = master.merge(
        mlc_map.groupby("SKU")["MLC"].agg(lambda x: sorted(set(x))).reset_index(name="mlc_list"),
        on="SKU",
        how="left",
    )
    product_base["mlc_list"] = product_base["mlc_list"].apply(lambda x: x if isinstance(x, list) else [])

    promo_by_sku = promos.groupby("SKU").size().reset_index(name="promo_rows_by_sku") if not promos.empty else pd.DataFrame(columns=["SKU", "promo_rows_by_sku"])
    product_base = product_base.merge(promo_by_sku, on="SKU", how="left")
    product_base["promo_rows_by_sku"] = product_base["promo_rows_by_sku"].fillna(0).astype(int)

    return {"sheets": sheets, "master": master, "mlc_map": mlc_map, "promos": promos, "product_base": product_base}


def build_master_promos(master: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for idx, row in master.iterrows():
        sku = row["SKU"]
        sku_norm = str(int(sku)) if pd.notna(sku) else None
        desc = row.get("DESCRIPCIÓN")
        mlc_aux = normalize_mlc(row.get("MLC_aux")) if "MLC_aux" in master.columns else None

        for slot_cfg in PROMO_SLOTS:
            mlc = normalize_mlc(row.get(slot_cfg["mlc_col"]))
            if slot_cfg["slot"] == 1 and mlc is None:
                mlc = mlc_aux
            price = row.get(slot_cfg["price_col"], np.nan)
            price = pd.to_numeric(pd.Series([price]), errors="coerce").iloc[0]
            fveni = pd.to_datetime(row.get(slot_cfg["date_col"]), errors="coerce")
            if pd.notna(fveni):
                fveni = pd.Timestamp(fveni).normalize()
            comment = row.get(slot_cfg["comment_col"])
            has_data = bool(mlc) or pd.notna(price) or pd.notna(fveni) or (pd.notna(comment) and str(comment).strip() != "")
            if not has_data:
                continue
            rows.append({
                "master_index": idx,
                "promo_slot": slot_cfg["slot"],
                "slot_label": slot_cfg["label"],
                "SKU": sku,
                "SKU_norm": sku_norm,
                "MLC": mlc,
                "Descripción": desc,
                "Precio promocional": price,
                "FECHA VENCI": fveni,
                "Comentario": "" if pd.isna(comment) else str(comment),
            })
    promos = pd.DataFrame(rows)
    if promos.empty:
        return pd.DataFrame(columns=["master_index", "promo_slot", "slot_label", "SKU", "SKU_norm", "MLC", "Descripción", "Precio promocional", "FECHA VENCI", "Comentario", "dias_para_vencer"])
    today = pd.Timestamp(date.today())
    promos["dias_para_vencer"] = (promos["FECHA VENCI"] - today).dt.days
    return promos.sort_values(["dias_para_vencer", "SKU"], na_position="last")


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
    promos_df = data["promos"]

    product = base.loc[base["SKU"] == sku].copy()
    if product.empty:
        return None, pd.DataFrame()

    product_row = product.iloc[0].to_dict()
    promos = promos_df[promos_df["SKU"] == sku].copy()

    mapped_mlcs = set(product_row.get("mlc_list", []) or [])
    if promos.empty and mapped_mlcs:
        promos = promos_df[promos_df["MLC"].isin(mapped_mlcs)].copy()

    if not promos.empty:
        promos["dias_para_vencer"] = (promos["FECHA VENCI"] - pd.Timestamp(date.today())).dt.days
        promos = promos.sort_values(by=["dias_para_vencer", "Precio promocional"], ascending=[True, True], na_position="last")
    return product_row, promos


def decision_rules(product: dict, promos: pd.DataFrame) -> list[str]:
    insights = []
    precio_meli = product.get("MONTO EN SIMULACIÓN")
    costo = product.get("ÚLTIMO COSTO")
    comentario = product.get("COMENTARIO")

    if promos.empty:
        insights.append("Sin promos registradas en la maestra.")
        return insights

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
            insights.append("Tiene promo que vence pasado mañana. Ya entra en ventana de control.")
        elif min_days <= 7:
            insights.append("Tiene promo dentro de los próximos 7 días. Conviene anticiparse.")

    if promos["MLC"].nunique(dropna=True) > 1:
        insights.append("El SKU tiene múltiples publicaciones asociadas. Conviene revisar consistencia.")
    if promos["Comentario"].fillna("").eq("").any():
        insights.append("Hay promos sin comentario. Falta contexto operativo.")
    if "Precio promocional" in promos.columns and pd.notna(costo):
        under_cost = promos["Precio promocional"].dropna() <= float(costo)
        if under_cost.any():
            insights.append("Al menos una promo tiene precio menor o igual al costo.")
    if pd.notna(precio_meli) and "Precio promocional" in promos.columns:
        cheaper = promos["Precio promocional"].dropna() < float(precio_meli)
        if cheaper.any():
            insights.append("Hay precio promocional más agresivo que el monto en simulación.")
    if pd.notna(comentario) and str(comentario).strip():
        insights.append("El producto ya tiene comentario interno en la maestra. Úsalo como contexto.")
    return insights


def urgency_label(promos: pd.DataFrame) -> tuple[str, str]:
    if promos.empty:
        return "Sin promo", "secondary"
    min_days = promos["dias_para_vencer"].dropna()
    if min_days.empty:
        return "Sin fecha", "secondary"
    return urgency_from_days(min_days.min())


def agenda_dataframe(data: dict) -> pd.DataFrame:
    promos = data["promos"].copy()
    master = data["master"][["SKU", "DESCRIPCIÓN", "ÚLTIMO COSTO", "MONTO EN SIMULACIÓN", "COMENTARIO"]].copy()
    if promos.empty:
        return promos
    df = promos.merge(master, on="SKU", how="left", suffixes=("", "_master"))
    labels = df["dias_para_vencer"].map(urgency_from_days)
    df["Estado"] = labels.map(lambda x: x[0])
    df["urgencia_color"] = labels.map(lambda x: x[1])
    return df


def update_master_promo(master_df: pd.DataFrame, sku: int, slot: int, price, end_date, comment):
    master_df = master_df.copy()
    mask = master_df["SKU"] == sku
    if not mask.any():
        return master_df

    price_col = "PRECIO B2C PUBLICADO " if slot == 1 else "PRECIO B2C"
    date_col = "FECHA VENCI" if slot == 1 else "FECHA VENCI.1"
    comment_col = "COMENTARIO" if slot == 1 else "COMENTARIO.1"

    idx = master_df.index[mask][0]
    if price_col in master_df.columns:
        master_df.loc[idx, price_col] = price if price not in (None, "") else np.nan
    if date_col in master_df.columns:
        master_df.loc[idx, date_col] = pd.to_datetime(end_date) if end_date else pd.NaT
    if comment_col in master_df.columns:
        master_df.loc[idx, comment_col] = comment
    return master_df


def make_download_bytes(sheets: dict, master_df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if name == MASTER_SHEET:
                master_df.to_excel(writer, sheet_name=name, index=False)
            else:
                df.to_excel(writer, sheet_name=name[:31], index=False)
    out.seek(0)
    return out.getvalue()


# UI
st.title("📈 Ficha comercial y agenda de promociones")
st.caption("Vista consolidada solo desde MAESTRA de precios + MLC -SKU. No se lee nada desde hojas antiguas de promociones.")

with st.sidebar:
    st.header("Fuente de datos")
    uploaded = st.file_uploader("Sube tu Excel maestro", type=["xlsx"])
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

# Session state for edited master
source_signature = (file_label, len(file_bytes))
if st.session_state.get("source_signature") != source_signature:
    raw = load_workbook(file_bytes)
    st.session_state["source_signature"] = source_signature
    st.session_state["sheets"] = raw["sheets"]
    st.session_state["master_df"] = raw["master"].copy()
    st.session_state["mlc_map_df"] = raw["mlc_map"].copy()

# rebuild runtime views from in-memory master
runtime = {
    "sheets": st.session_state["sheets"],
    "master": st.session_state["master_df"].copy(),
    "mlc_map": st.session_state["mlc_map_df"].copy(),
}
runtime["promos"] = build_master_promos(runtime["master"])
runtime["product_base"] = runtime["master"].merge(
    runtime["mlc_map"].groupby("SKU")["MLC"].agg(lambda x: sorted(set(x))).reset_index(name="mlc_list"),
    on="SKU", how="left"
)
runtime["product_base"]["mlc_list"] = runtime["product_base"]["mlc_list"].apply(lambda x: x if isinstance(x, list) else [])
promo_by_sku = runtime["promos"].groupby("SKU").size().reset_index(name="promo_rows_by_sku") if not runtime["promos"].empty else pd.DataFrame(columns=["SKU", "promo_rows_by_sku"])
runtime["product_base"] = runtime["product_base"].merge(promo_by_sku, on="SKU", how="left")
runtime["product_base"]["promo_rows_by_sku"] = runtime["product_base"]["promo_rows_by_sku"].fillna(0).astype(int)

search_df = build_search_index(runtime["product_base"])
agenda_df = agenda_dataframe(runtime)

today_count = int((agenda_df["dias_para_vencer"] == 0).sum()) if not agenda_df.empty else 0
tomorrow_count = int((agenda_df["dias_para_vencer"] == 1).sum()) if not agenda_df.empty else 0
day2_count = int((agenda_df["dias_para_vencer"] == 2).sum()) if not agenda_df.empty else 0
week_count = int(agenda_df["dias_para_vencer"].between(0, 7, inclusive="both").sum()) if not agenda_df.empty else 0

c1, c2, c3, c4 = st.columns(4)
with c1:
    style_metric_card("Promos que vencen hoy", str(today_count), "prioridad máxima")
with c2:
    style_metric_card("Promos que vencen mañana", str(tomorrow_count), "ventana de reacción")
with c3:
    style_metric_card("Vencen pasado mañana", str(day2_count), "control preventivo")
with c4:
    style_metric_card("Vencen en 7 días", str(week_count), "agenda comercial")

tab1, tab2, tab3 = st.tabs(["Consulta producto", "Agenda comercial", "Edición puntual"])

with tab1:
    st.subheader("Consulta producto")

    q1, q2 = st.columns([2, 1])
    with q1:
        query = st.text_input("Busca por SKU, descripción o MLC", placeholder="Ej: 110203002020, abrazadera, MLC1789384668")
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
            urg_txt, _ = urgency_label(runtime["promos"][runtime["promos"]["SKU"] == row["SKU"]])
            label = f"{int(row['SKU'])} · {desc} · {urg_txt}"
            options.append(label)
            option_map[label] = int(row["SKU"])

        selected_label = st.selectbox("Selecciona un producto", options, index=0)
        selected_sku = option_map[selected_label]

        product, promos = get_product_view(selected_sku, runtime)

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
                        SKU <b>{sku_str}</b> · UBIC <b>{product.get("UBIC", "—") if pd.notna(product.get("UBIC")) else "—"}</b>
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
                st.write(f"**Monto en simulación:** {fmt_money(product.get('MONTO EN SIMULACIÓN'))}")
                st.write(f"**Cambio de precio:** {fmt_money(product.get('CAMBIO DE PRECIO'))}")
            with b:
                st.markdown("#### Márgenes")
                st.write(f"**Margen local:** {fmt_percent(product.get('MARGEN LOCAL'))}")
                st.write(f"**Margen Meli 1:** {fmt_percent(product.get('MARGEN MELI 1'))}")
                st.write(f"**Fecha venci maestra:** {fmt_date(product.get('FECHA VENCI'))}")
                st.write(f"**Fecha venci 2:** {fmt_date(product.get('FECHA VENCI.1'))}")
            with c:
                st.markdown("#### Relación con publicaciones")
                if mlcs:
                    st.write("**MLC asociadas:**")
                    st.code("\n".join(mlcs))
                else:
                    st.info("No encontré publicaciones en la hoja MLC -SKU para este SKU.")
                st.write(f"**Comentario maestra:** {product.get('COMENTARIO') if pd.notna(product.get('COMENTARIO')) else '—'}")

            st.markdown("#### Promos de la maestra")
            if promos.empty:
                st.info("No encontré promos registradas en la maestra para este SKU.")
            else:
                promo_view = promos.copy()
                promo_view["Estado"] = promo_view["dias_para_vencer"].map(lambda x: urgency_from_days(x)[0])
                promo_view["FECHA VENCI"] = promo_view["FECHA VENCI"].map(fmt_date)
                promo_view["Precio promocional"] = promo_view["Precio promocional"].map(fmt_money)
                show_cols = ["slot_label", "MLC", "Descripción", "Precio promocional", "Comentario", "FECHA VENCI", "Estado"]
                st.dataframe(promo_view[show_cols], use_container_width=True, hide_index=True)

            st.markdown("#### Lectura comercial")
            for insight in decision_rules(product, promos):
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
        solo_sin_comentario = st.toggle("Solo sin comentario", value=False)
        texto = st.text_input("Buscar SKU / descripción / MLC", placeholder="Ej: MLC1789 o abrazadera")

    ag = agenda_df.copy()
    if not ag.empty and status_filter:
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

    if solo_sin_comentario and not ag.empty:
        ag = ag[ag["Comentario"].fillna("").eq("")]

    if texto and not ag.empty:
        text_l = texto.lower()
        mask = (
            ag["SKU"].astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
            | ag["MLC"].astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
            | ag["Descripción"].fillna("").astype(str).str.contains(re.escape(text_l), case=False, regex=True, na=False)
        )
        ag = ag[mask]

    ag = ag.sort_values(by=["dias_para_vencer", "SKU"], na_position="last") if not ag.empty else ag

    with right:
        st.caption("Ordenado por urgencia de vencimiento.")
        if ag.empty:
            st.info("No hay promos que mostrar con esos filtros.")
        else:
            agenda_show = ag.copy()
            agenda_show["Precio promocional"] = agenda_show["Precio promocional"].map(fmt_money)
            agenda_show["ÚLTIMO COSTO"] = agenda_show["ÚLTIMO COSTO"].map(fmt_money)
            agenda_show["MONTO EN SIMULACIÓN"] = agenda_show["MONTO EN SIMULACIÓN"].map(fmt_money)
            agenda_show["FECHA VENCI"] = agenda_show["FECHA VENCI"].map(fmt_date)

            cols = ["SKU", "MLC", "Descripción", "Precio promocional", "ÚLTIMO COSTO", "MONTO EN SIMULACIÓN", "Comentario", "FECHA VENCI", "Estado"]
            st.dataframe(agenda_show[cols], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Edición puntual")
    st.caption("Edita directamente los campos de promo en la maestra cargada y descarga el Excel actualizado.")

    sku_input = st.text_input("SKU a editar", placeholder="Ej: 110203002020")
    if sku_input and sku_input.isdigit():
        sku_edit = int(sku_input)
        product, promos = get_product_view(sku_edit, runtime)

        if not product:
            st.warning("No encontré ese SKU.")
        else:
            st.write(f"**Producto:** {product.get('DESCRIPCIÓN', '—')}")
            st.write(f"**SKU:** {sku_edit}")

            with st.form("edit_form"):
                comentario_maestra = st.text_area(
                    "Comentario maestra",
                    value="" if pd.isna(product.get("COMENTARIO")) else str(product.get("COMENTARIO")),
                    height=100,
                )

                st.markdown("#### Promo 1")
                p1c1, p1c2 = st.columns(2)
                p1_price = p1c1.number_input("Precio B2C publicado", min_value=0.0, value=float(pd.to_numeric(product.get("PRECIO B2C PUBLICADO "), errors="coerce") if pd.notna(product.get("PRECIO B2C PUBLICADO ")) else 0.0), step=1.0)
                current_p1_date = pd.to_datetime(product.get("FECHA VENCI"), errors="coerce")
                p1_date = p1c2.date_input("FECHA VENCI", value=(current_p1_date.date() if pd.notna(current_p1_date) else None), format="DD/MM/YYYY")
                p1_comment = st.text_input("Comentario promo 1", value="" if pd.isna(product.get("COMENTARIO")) else str(product.get("COMENTARIO")))

                st.markdown("#### Promo 2")
                p2c1, p2c2 = st.columns(2)
                p2_price = p2c1.number_input("Precio B2C 2", min_value=0.0, value=float(pd.to_numeric(product.get("PRECIO B2C"), errors="coerce") if pd.notna(product.get("PRECIO B2C")) else 0.0), step=1.0)
                current_p2_date = pd.to_datetime(product.get("FECHA VENCI.1"), errors="coerce")
                p2_date = p2c2.date_input("FECHA VENCI.1", value=(current_p2_date.date() if pd.notna(current_p2_date) else None), format="DD/MM/YYYY")
                p2_comment = st.text_input("Comentario promo 2", value="" if pd.isna(product.get("COMENTARIO.1")) else str(product.get("COMENTARIO.1")))

                submitted = st.form_submit_button("Guardar cambios en memoria", type="primary", use_container_width=True)

            if submitted:
                master_df = st.session_state["master_df"].copy()
                master_df.loc[master_df["SKU"] == sku_edit, "COMENTARIO"] = comentario_maestra
                master_df = update_master_promo(master_df, sku_edit, 1, p1_price if p1_price > 0 else np.nan, p1_date, p1_comment)
                master_df = update_master_promo(master_df, sku_edit, 2, p2_price if p2_price > 0 else np.nan, p2_date, p2_comment)
                st.session_state["master_df"] = master_df
                st.success("Cambios guardados en memoria.")

    else:
        st.info("Ingresa un SKU numérico para editar la promo directamente desde la maestra.")

download_bytes = make_download_bytes(st.session_state["sheets"], st.session_state["master_df"])
st.markdown("---")
st.download_button(
    "Descargar Excel actualizado",
    data=download_bytes,
    file_name="maestra_precios_actualizada.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.caption("Modelo de unión: SKU (maestra) → MLC -SKU. Las promos se leen solo desde la hoja MAESTRA de precios.")
