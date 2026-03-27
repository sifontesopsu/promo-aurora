
import hashlib
import io
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Aurora Promos", layout="wide")


# ----------------------------- Helpers ----------------------------- #

SHEET_MASTER = "MAESTRA de precios"
SHEET_BRIDGE = "MLC -SKU"
SHEET_RELAMPAGO = "Relampago mi pagina"


def file_signature(uploaded_file) -> str:
    data = uploaded_file.getvalue()
    return hashlib.md5(data).hexdigest()


def normalize_text(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() == "nan":
        return ""
    return s


def normalize_sku(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if s.lower() == "nan":
        return ""
    return s


def normalize_mlc(val) -> str:
    s = normalize_text(val).upper().replace(" ", "")
    if not s:
        return ""
    if s.startswith("MLC"):
        return s
    digits = "".join(ch for ch in s if ch.isdigit())
    return f"MLC{digits}" if digits else s


def to_number(val):
    if pd.isna(val) or val == "":
        return None
    try:
        return float(val)
    except Exception:
        try:
            cleaned = str(val).replace("$", "").replace(".", "").replace(",", ".")
            return float(cleaned)
        except Exception:
            return None


def to_date_only(val):
    if pd.isna(val) or val == "":
        return None
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def format_date(val) -> str:
    d = to_date_only(val)
    return d.strftime("%d/%m/%Y") if d else ""


def format_money(val) -> str:
    num = to_number(val)
    if num is None:
        return "—"
    return f"${num:,.0f}".replace(",", ".")


def format_margin(val) -> str:
    num = to_number(val)
    if num is None:
        return "—"
    return f"{num*100:.1f}%"


def first_existing(columns, candidates):
    for c in candidates:
        if c in columns:
            return c
    return None


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {name: pd.read_excel(io.BytesIO(file_bytes), sheet_name=name) for name in xls.sheet_names}
    return sheets, xls.sheet_names


@st.cache_data(show_spinner=False)
def load_purchases(file_bytes: bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    # Normalize columns
    df.columns = [str(c).strip() for c in df.columns]

    date_col = first_existing(df.columns, ["Fecha"])
    sku_col = first_existing(df.columns, ["SKU"])
    desc_col = first_existing(df.columns, ["Concepto / Artículo", "Detalle Concepto Compra"])
    provider_col = first_existing(df.columns, ["Razón Social"])
    price_col = first_existing(df.columns, ["Precio Un."])
    qty_col = first_existing(df.columns, ["Cantidad"])

    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)

    if sku_col is None:
        df["SKU_norm"] = ""
    else:
        df["SKU_norm"] = df[sku_col].apply(normalize_sku)

    if desc_col is None:
        df["Descripcion_compra"] = ""
    else:
        df["Descripcion_compra"] = df[desc_col].fillna("").astype(str)

    if provider_col is None:
        df["Proveedor"] = ""
    else:
        df["Proveedor"] = df[provider_col].fillna("").astype(str)

    if price_col is None:
        df["Precio_compra"] = np.nan
    else:
        df["Precio_compra"] = pd.to_numeric(df[price_col], errors="coerce")

    if qty_col is None:
        df["Cantidad_compra"] = np.nan
    else:
        df["Cantidad_compra"] = pd.to_numeric(df[qty_col], errors="coerce")

    work = df[df["SKU_norm"] != ""].copy()

    if date_col is None or work.empty:
        summary = pd.DataFrame(columns=[
            "SKU_norm", "Ultima_fecha_compra", "Ultimo_precio_compra",
            "Proveedor_ultima_compra", "Cantidad_ultima_compra",
            "Precio_compra_anterior", "Variacion_vs_anterior"
        ])
        history = work
        return summary, history

    work = work.sort_values(date_col)
    last = work.groupby("SKU_norm", as_index=False).tail(1).copy()

    prev = work.groupby("SKU_norm").nth(-2).reset_index()
    prev = prev[["SKU_norm", "Precio_compra"]].rename(columns={"Precio_compra": "Precio_compra_anterior"})

    summary = last[["SKU_norm", date_col, "Precio_compra", "Proveedor", "Cantidad_compra"]].rename(columns={
        date_col: "Ultima_fecha_compra",
        "Precio_compra": "Ultimo_precio_compra",
        "Proveedor": "Proveedor_ultima_compra",
        "Cantidad_compra": "Cantidad_ultima_compra",
    })
    summary = summary.merge(prev, on="SKU_norm", how="left")
    summary["Variacion_vs_anterior"] = np.where(
        summary["Precio_compra_anterior"].notna() & (summary["Precio_compra_anterior"] != 0),
        (summary["Ultimo_precio_compra"] - summary["Precio_compra_anterior"]) / summary["Precio_compra_anterior"],
        np.nan,
    )
    return summary, work


def parse_relampago_sheet(df_raw: pd.DataFrame) -> pd.DataFrame:
    # The sheet has no real header; parse by position
    raw = pd.read_excel(io.BytesIO(st.session_state.master_file_bytes), sheet_name=SHEET_RELAMPAGO, header=None)
    raw = raw.iloc[:, :6].copy()
    raw.columns = ["SKU", "DESCRIPCION", "PRECIO_B2C", "MLC", "COMENTARIO", "ESTADO"]
    raw["SKU_norm"] = raw["SKU"].apply(normalize_sku)
    raw["DESCRIPCION"] = raw["DESCRIPCION"].fillna("").astype(str)
    raw["PRECIO_B2C"] = pd.to_numeric(raw["PRECIO_B2C"], errors="coerce")
    raw["MLC_norm"] = raw["MLC"].apply(normalize_mlc)
    raw["COMENTARIO"] = raw["COMENTARIO"].fillna("").astype(str)
    raw["ESTADO"] = raw["ESTADO"].fillna("").astype(str)
    raw = raw[(raw["SKU_norm"] != "") | (raw["DESCRIPCION"].str.strip() != "")]
    return raw


def ensure_master_types(df: pd.DataFrame) -> pd.DataFrame:
    master = df.copy()
    master.columns = [str(c) for c in master.columns]
    master["SKU_norm"] = master["SKU"].apply(normalize_sku)
    for c in ["Unnamed: 12", "MLC", "MLC.1"]:
        if c in master.columns:
            if c == "Unnamed: 12":
                master["MLC_BASE"] = master[c].apply(normalize_mlc)
            else:
                master[c] = master[c].apply(normalize_mlc)
    for c in ["FECHA VENCI", "FECHA VENCI.1"]:
        if c in master.columns:
            master[c] = pd.to_datetime(master[c], errors="coerce").dt.date
    return master


def ensure_bridge_types(df: pd.DataFrame) -> pd.DataFrame:
    bridge = df.copy()
    bridge.columns = [str(c) for c in bridge.columns]
    bridge["SKU_norm"] = bridge["SKU"].apply(normalize_sku)
    bridge["MLC_norm"] = bridge["Número de publicación"].apply(normalize_mlc)
    return bridge[bridge["SKU_norm"] != ""]


def margin_local(cost, precio_bruto):
    c = to_number(cost)
    pb = to_number(precio_bruto)
    if c is None or pb is None or pb == 0:
        return None, None
    neto = pb / 1.19
    return neto, (neto - c) / neto if neto else None


def margin_meli(cost, monto_simulacion):
    c = to_number(cost)
    ms = to_number(monto_simulacion)
    if c is None or ms is None or ms == 0:
        return None, None
    neto = ms / 1.19
    return neto, (neto - c) / neto if neto else None


def build_promos(master: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for idx, row in master.iterrows():
        sku = row["SKU_norm"]
        desc = normalize_text(row.get("DESCRIPCIÓN"))
        for slot in [1, 2]:
            mlc_col = "MLC" if slot == 1 else "MLC.1"
            price_col = "PRECIO B2C PUBLICADO " if slot == 1 else "PRECIO B2C"
            date_col = "FECHA VENCI" if slot == 1 else "FECHA VENCI.1"
            comment_col = "COMENTARIO" if slot == 1 else "COMENTARIO.1"
            mlc = normalize_mlc(row.get(mlc_col))
            price = to_number(row.get(price_col))
            fecha = to_date_only(row.get(date_col))
            comment = normalize_text(row.get(comment_col))
            if not any([mlc, price is not None, fecha, comment]):
                continue
            rows.append({
                "row_idx": idx,
                "slot": slot,
                "SKU_norm": sku,
                "DESCRIPCION": desc,
                "MLC_norm": mlc,
                "PRECIO_B2C": price,
                "FECHA_VENCI": fecha,
                "COMENTARIO": comment,
            })
    promos = pd.DataFrame(rows)
    if promos.empty:
        return pd.DataFrame(columns=["row_idx","slot","SKU_norm","DESCRIPCION","MLC_norm","PRECIO_B2C","FECHA_VENCI","COMENTARIO","STATUS","DIAS"])
    today = date.today()
    promos["DIAS"] = promos["FECHA_VENCI"].apply(lambda d: (d - today).days if d else None)
    def status(d):
        if d is None:
            return "Sin fecha"
        delta = (d - today).days
        if delta < 0:
            return "Vencida"
        if delta == 0:
            return "Vence hoy"
        if delta == 1:
            return "Vence mañana"
        if delta == 2:
            return "Vence en 2 días"
        if delta == 3:
            return "Vence en 3 días"
        return "Vigente"
    promos["STATUS"] = promos["FECHA_VENCI"].apply(status)
    return promos


def build_download_bytes(all_sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False, header=not (name == SHEET_RELAMPAGO and list(df.columns) == list(range(len(df.columns)))))
    buf.seek(0)
    return buf.getvalue()


def promo_card_label(row: pd.Series) -> str:
    sku = row["SKU_norm"]
    desc = normalize_text(row["DESCRIPCION"])[:55]
    mlc = row["MLC_norm"] or "Sin MLC"
    fecha = format_date(row["FECHA_VENCI"]) or "Sin fecha"
    return f"{sku}\n{desc}\n{mlc}\n{fecha}"


def filter_promos(promos, q, statuses):
    out = promos.copy()
    if q:
        qn = q.strip().lower()
        out = out[
            out["SKU_norm"].str.lower().str.contains(qn, na=False)
            | out["DESCRIPCION"].str.lower().str.contains(qn, na=False)
            | out["MLC_norm"].str.lower().str.contains(qn, na=False)
        ]
    if statuses:
        out = out[out["STATUS"].isin(statuses)]
    order = {"Vence hoy":0,"Vence mañana":1,"Vence en 2 días":2,"Vence en 3 días":3,"Vigente":4,"Sin fecha":5,"Vencida":6}
    out["_order"] = out["STATUS"].map(order).fillna(99)
    out = out.sort_values(["_order","FECHA_VENCI","SKU_norm","slot"])
    return out.drop(columns=["_order"])


# ----------------------------- State ----------------------------- #

def initialize_from_files(master_upload, compras_upload):
    master_sig = file_signature(master_upload)
    compras_sig = file_signature(compras_upload) if compras_upload else ""

    if st.session_state.get("master_sig") != master_sig:
        master_bytes = master_upload.getvalue()
        sheets, order = load_workbook(master_bytes)
        st.session_state.master_sig = master_sig
        st.session_state.master_file_bytes = master_bytes
        st.session_state.sheet_order = order

        st.session_state.master_df = ensure_master_types(sheets[SHEET_MASTER])
        st.session_state.bridge_df = ensure_bridge_types(sheets[SHEET_BRIDGE])

        if SHEET_RELAMPAGO in order:
            st.session_state.relampago_df = parse_relampago_sheet(sheets[SHEET_RELAMPAGO])
        else:
            st.session_state.relampago_df = pd.DataFrame(columns=["SKU","DESCRIPCION","PRECIO_B2C","MLC","COMENTARIO","ESTADO","SKU_norm","MLC_norm"])

        st.session_state.other_sheets = {
            name: df.copy() for name, df in sheets.items()
            if name not in [SHEET_MASTER, SHEET_BRIDGE, SHEET_RELAMPAGO]
        }

    if compras_upload:
        if st.session_state.get("compras_sig") != compras_sig:
            summary, history = load_purchases(compras_upload.getvalue())
            st.session_state.compras_sig = compras_sig
            st.session_state.purchase_summary = summary
            st.session_state.purchase_history = history
    else:
        st.session_state.compras_sig = ""
        st.session_state.purchase_summary = pd.DataFrame(columns=[
            "SKU_norm","Ultima_fecha_compra","Ultimo_precio_compra","Proveedor_ultima_compra",
            "Cantidad_ultima_compra","Precio_compra_anterior","Variacion_vs_anterior"
        ])
        st.session_state.purchase_history = pd.DataFrame()


# ----------------------------- Dialogs ----------------------------- #

@st.dialog("Editar promo", width="large")
def edit_promo_dialog(promo_index: int):
    promos = build_promos(st.session_state.master_df)
    if promo_index >= len(promos):
        st.warning("La promo ya no está disponible.")
        return
    row = promos.iloc[promo_index]
    row_idx = int(row["row_idx"])
    slot = int(row["slot"])
    master = st.session_state.master_df

    price_col = "PRECIO B2C PUBLICADO " if slot == 1 else "PRECIO B2C"
    date_col = "FECHA VENCI" if slot == 1 else "FECHA VENCI.1"
    comment_col = "COMENTARIO" if slot == 1 else "COMENTARIO.1"

    st.markdown(f"### {row['SKU_norm']}")
    st.caption(normalize_text(row["DESCRIPCION"]))

    c1, c2 = st.columns([1.3, 1])
    with c1:
        new_date = st.date_input(
            "Fecha de vencimiento",
            value=row["FECHA_VENCI"] or date.today(),
            format="DD/MM/YYYY",
            key=f"date_{promo_index}"
        )
    with c2:
        new_price = st.number_input(
            "Precio B2C publicado",
            min_value=0,
            value=int(row["PRECIO_B2C"] or 0),
            step=100,
            key=f"price_{promo_index}",
        )

    new_comment = st.text_area(
        "Comentario",
        value=row["COMENTARIO"],
        height=100,
        key=f"comment_{promo_index}",
    )

    if st.button("Guardar cambios", type="primary", key=f"savepromo_{promo_index}", use_container_width=True):
        st.session_state.master_df.at[row_idx, price_col] = new_price
        st.session_state.master_df.at[row_idx, date_col] = pd.to_datetime(new_date).date()
        st.session_state.master_df.at[row_idx, comment_col] = new_comment
        st.success("Promo actualizada.")
        st.rerun()


# ----------------------------- UI ----------------------------- #

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    .metric-card {
        background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; padding:14px 16px;
        box-shadow:0 1px 3px rgba(0,0,0,.06);
    }
    .soft-card {
        background:#ffffff; border:1px solid #e5e7eb; border-radius:18px; padding:14px 16px;
        box-shadow:0 1px 2px rgba(0,0,0,.04);
    }
    div[data-testid="stButton"] > button[kind="secondary"]{
        border-radius:18px; min-height:120px; white-space:pre-wrap; text-align:left; padding:14px 16px;
        border:1px solid #e5e7eb; background:#fff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Aurora • Pricing y promos")

with st.sidebar:
    st.header("Archivos")
    master_upload = st.file_uploader("Maestra saneada", type=["xlsx"], key="master_upload")
    compras_upload = st.file_uploader("Compras", type=["xlsx"], key="compras_upload")
    st.caption("La app trabaja solo con Maestra, MLC -SKU, Relámpago y Compras.")

if not master_upload:
    st.info("Sube la maestra saneada para comenzar.")
    st.stop()

initialize_from_files(master_upload, compras_upload)

master_df = st.session_state.master_df
bridge_df = st.session_state.bridge_df
relampago_df = st.session_state.relampago_df
purchase_summary = st.session_state.purchase_summary
purchase_history = st.session_state.purchase_history

promos_df = build_promos(master_df)

tab1, tab2, tab3, tab4 = st.tabs(["Cockpit por producto", "Operador de promos", "Relámpago mi página", "Alta de producto"])

with tab1:
    search = st.text_input("Buscar por SKU o descripción")
    options = master_df.copy()
    if search:
        s = search.strip().lower()
        options = options[
            options["SKU_norm"].str.lower().str.contains(s, na=False)
            | options["DESCRIPCIÓN"].fillna("").astype(str).str.lower().str.contains(s, na=False)
        ]
    if options.empty:
        st.warning("No hay productos para ese filtro.")
    else:
        labels = options.apply(lambda r: f"{r['SKU_norm']} · {normalize_text(r['DESCRIPCIÓN'])[:90]}", axis=1).tolist()
        idx = st.selectbox("Producto", range(len(options)), format_func=lambda i: labels[i])
        prod = options.iloc[idx]
        sku = prod["SKU_norm"]

        neto_local, margen_local_calc = margin_local(prod.get("ÚLTIMO COSTO"), prod.get("PRECIO BRUTO"))
        neto_meli, margen_meli_calc = margin_meli(prod.get("ÚLTIMO COSTO"), prod.get("MONTO EN SIMULACIÓN"))

        pcols = st.columns(4)
        pcols[0].metric("SKU", sku)
        pcols[1].metric("Último costo", format_money(prod.get("ÚLTIMO COSTO")))
        pcols[2].metric("Margen local", format_margin(prod.get("MARGEN LOCAL") if pd.notna(prod.get("MARGEN LOCAL")) else margen_local_calc))
        pcols[3].metric("Margen Meli 1", format_margin(prod.get("MARGEN MELI 1") if pd.notna(prod.get("MARGEN MELI 1")) else margen_meli_calc))

        c1, c2 = st.columns([1.2, 1])
        with c1:
            st.markdown("#### Datos base")
            base_view = pd.DataFrame([{
                "Descripción": normalize_text(prod.get("DESCRIPCIÓN")),
                "Ubicación": normalize_text(prod.get("UBIC")),
                "Precio bruto tienda": format_money(prod.get("PRECIO BRUTO")),
                "Precio neto": format_money(prod.get("PRECIO NETO") if pd.notna(prod.get("PRECIO NETO")) else neto_local),
                "Monto en simulación": format_money(prod.get("MONTO EN SIMULACIÓN")),
                "Neto Meli 1": format_money(prod.get(" NETO MELI 1") if pd.notna(prod.get(" NETO MELI 1")) else neto_meli),
            }])
            st.dataframe(base_view, use_container_width=True, hide_index=True)

            mlcs = set()
            base_mlc = normalize_mlc(prod.get("Unnamed: 12"))
            if base_mlc:
                mlcs.add(base_mlc)
            for col in ["MLC", "MLC.1"]:
                v = normalize_mlc(prod.get(col))
                if v:
                    mlcs.add(v)
            bridge_mlcs = bridge_df.loc[bridge_df["SKU_norm"] == sku, "MLC_norm"].dropna().tolist()
            mlcs.update([m for m in bridge_mlcs if m])
            st.markdown("#### MLC asociados")
            if mlcs:
                st.write(" · ".join(sorted(mlcs)))
            else:
                st.caption("Sin MLC asociados.")

        with c2:
            st.markdown("#### Promos desde la maestra")
            prod_promos = promos_df[promos_df["SKU_norm"] == sku].copy()
            if prod_promos.empty:
                st.info("No tiene promo activa registrada en la maestra.")
            else:
                show = prod_promos[["MLC_norm", "PRECIO_B2C", "FECHA_VENCI", "COMENTARIO", "STATUS"]].copy()
                show.columns = ["MLC", "Precio B2C", "Fecha", "Comentario", "Estado"]
                show["Precio B2C"] = show["Precio B2C"].apply(format_money)
                show["Fecha"] = show["Fecha"].apply(format_date)
                st.dataframe(show, use_container_width=True, hide_index=True)

            st.markdown("#### Relámpago")
            rel = relampago_df[relampago_df["SKU_norm"] == sku].copy()
            if rel.empty:
                st.caption("No está en Relámpago.")
            else:
                show_rel = rel[["PRECIO_B2C", "COMENTARIO", "ESTADO"]].copy()
                show_rel.columns = ["Precio B2C", "Comentario", "Estado"]
                show_rel["Precio B2C"] = show_rel["Precio B2C"].apply(format_money)
                st.dataframe(show_rel, use_container_width=True, hide_index=True)

        st.markdown("#### Compras")
        compra = purchase_summary[purchase_summary["SKU_norm"] == sku].copy()
        if compra.empty:
            st.caption("Sin historial de compras para este SKU.")
        else:
            crow = compra.iloc[0]
            cc = st.columns(4)
            cc[0].metric("Última compra", format_date(crow["Ultima_fecha_compra"]))
            cc[1].metric("Último precio compra", format_money(crow["Ultimo_precio_compra"]))
            cc[2].metric("Proveedor", normalize_text(crow["Proveedor_ultima_compra"]) or "—")
            cc[3].metric("Variación vs anterior", format_margin(crow["Variacion_vs_anterior"]))
            hist = purchase_history[purchase_history["SKU_norm"] == sku].copy()
            if not hist.empty:
                display_cols = []
                for col in ["Fecha", "Proveedor", "Descripcion_compra", "Cantidad_compra", "Precio_compra"]:
                    if col in hist.columns:
                        display_cols.append(col)
                h = hist[display_cols].copy()
                if "Fecha" in h.columns:
                    h["Fecha"] = pd.to_datetime(h["Fecha"], errors="coerce").dt.date.apply(format_date)
                if "Precio_compra" in h.columns:
                    h["Precio_compra"] = h["Precio_compra"].apply(format_money)
                st.dataframe(h.sort_values(by=display_cols[0], ascending=False) if display_cols else h, use_container_width=True, hide_index=True)

with tab2:
    st.markdown("### Operador de promos")
    q = st.text_input("Buscar promo", placeholder="SKU, descripción o MLC")
    statuses = st.multiselect(
        "Estados",
        ["Vence hoy", "Vence mañana", "Vence en 2 días", "Vence en 3 días", "Vigente", "Sin fecha", "Vencida"],
        default=["Vence hoy", "Vence mañana", "Vence en 2 días", "Vence en 3 días", "Vigente", "Sin fecha"]
    )
    filtered = filter_promos(promos_df, q, statuses)

    left, right = st.columns([1.2, 1])
    with left:
        st.caption(f"{len(filtered)} promos")
    with right:
        new_bulk_date = st.date_input("Cambio masivo de fecha", value=date.today(), format="DD/MM/YYYY", key="bulk_date")
        if st.button("Aplicar fecha a promos filtradas", use_container_width=True):
            for _, row in filtered.iterrows():
                row_idx = int(row["row_idx"])
                slot = int(row["slot"])
                col = "FECHA VENCI" if slot == 1 else "FECHA VENCI.1"
                st.session_state.master_df.at[row_idx, col] = pd.to_datetime(new_bulk_date).date()
            st.success("Fecha actualizada para las promos filtradas.")
            st.rerun()

    if filtered.empty:
        st.info("No hay promos para ese filtro.")
    else:
        per_row = 4
        items = filtered.reset_index(drop=True)
        for start in range(0, len(items), per_row):
            cols = st.columns(per_row)
            for col, i in zip(cols, range(start, min(start + per_row, len(items)))):
                row = items.iloc[i]
                with col:
                    if st.button(promo_card_label(row), key=f"card_{i}", use_container_width=True):
                        edit_promo_dialog(i)

with tab3:
    st.markdown("### Relámpago mi página")
    r1, r2 = st.columns([1.3, 1])
    with r1:
        st.dataframe(
            relampago_df[["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "COMENTARIO", "ESTADO"]]
            .rename(columns={"SKU_norm":"SKU", "PRECIO_B2C":"Precio B2C", "DESCRIPCION":"Descripción"}),
            use_container_width=True,
            hide_index=True
        )
    with r2:
        st.markdown("#### Agregar / editar")
        sku_r = st.text_input("SKU", key="rel_sku")
        existing_idx = None
        if sku_r:
            matches = relampago_df[relampago_df["SKU_norm"] == normalize_sku(sku_r)]
            if not matches.empty:
                existing_idx = matches.index[0]
                ex = matches.iloc[0]
                default_desc = ex["DESCRIPCION"]
                default_price = int(ex["PRECIO_B2C"] or 0)
                default_comment = ex["COMENTARIO"]
                default_state = ex["ESTADO"]
            else:
                prod_match = master_df[master_df["SKU_norm"] == normalize_sku(sku_r)]
                default_desc = normalize_text(prod_match.iloc[0]["DESCRIPCIÓN"]) if not prod_match.empty else ""
                default_price = 0
                default_comment = ""
                default_state = ""
        else:
            default_desc = ""
            default_price = 0
            default_comment = ""
            default_state = ""

        desc_r = st.text_input("Descripción", value=default_desc, key="rel_desc")
        price_r = st.number_input("Precio B2C", min_value=0, value=int(default_price), step=100, key="rel_price")
        comm_r = st.text_input("Comentario", value=default_comment, key="rel_comment")
        state_r = st.text_input("Estado", value=default_state, key="rel_state")

        csave, cdel = st.columns(2)
        with csave:
            if st.button("Guardar", use_container_width=True):
                row = {
                    "SKU": normalize_sku(sku_r),
                    "DESCRIPCION": desc_r,
                    "PRECIO_B2C": price_r,
                    "MLC": "",
                    "COMENTARIO": comm_r,
                    "ESTADO": state_r,
                    "SKU_norm": normalize_sku(sku_r),
                    "MLC_norm": "",
                }
                if existing_idx is None:
                    st.session_state.relampago_df = pd.concat([st.session_state.relampago_df, pd.DataFrame([row])], ignore_index=True)
                else:
                    for k, v in row.items():
                        st.session_state.relampago_df.at[existing_idx, k] = v
                st.success("Relámpago actualizado.")
                st.rerun()
        with cdel:
            if existing_idx is not None and st.button("Eliminar", use_container_width=True):
                st.session_state.relampago_df = st.session_state.relampago_df.drop(index=existing_idx).reset_index(drop=True)
                st.success("Eliminado.")
                st.rerun()

with tab4:
    st.markdown("### Alta de producto")
    c1, c2, c3 = st.columns(3)
    sku_new = c1.text_input("SKU nuevo", key="new_sku")
    desc_new = c2.text_input("Descripción", key="new_desc")
    ubic_new = c3.text_input("Ubicación", key="new_ubic")

    c4, c5, c6 = st.columns(3)
    costo_new = c4.number_input("Último costo", min_value=0, value=0, step=1, key="new_cost")
    precio_tienda_new = c5.number_input("Precio bruto en tienda", min_value=0, value=0, step=100, key="new_store")
    monto_sim_new = c6.number_input("Monto en simulación", min_value=0, value=0, step=100, key="new_sim")

    neto_loc, marg_loc = margin_local(costo_new, precio_tienda_new)
    neto_mel, marg_mel = margin_meli(costo_new, monto_sim_new)

    m1, m2 = st.columns(2)
    m1.metric("Margen local proyectado", format_margin(marg_loc))
    m2.metric("Margen Meli 1 proyectado", format_margin(marg_mel))

    st.markdown("#### Promo base")
    p1, p2, p3 = st.columns(3)
    mlc1_new = p1.text_input("MLC", key="new_mlc1")
    b2c1_new = p2.number_input("Precio B2C publicado", min_value=0, value=0, step=100, key="new_b2c1")
    fecha1_new = p3.date_input("Fecha venci", value=date.today(), format="DD/MM/YYYY", key="new_date1")
    comment1_new = st.text_input("Comentario", key="new_comment1")

    mlcs_extra = st.text_input("MLC adicionales (separados por coma)", key="new_mlcs_extra")
    add_rel = st.checkbox("Agregar también a Relámpago", key="new_add_rel")

    if st.button("Crear producto", type="primary"):
        sku_norm = normalize_sku(sku_new)
        if not sku_norm:
            st.error("Debes ingresar un SKU.")
        elif sku_norm in set(master_df["SKU_norm"]):
            st.error("Ese SKU ya existe.")
        else:
            new_row = {col: np.nan for col in master_df.columns}
            new_row["SKU"] = sku_norm
            new_row["SKU_norm"] = sku_norm
            new_row["DESCRIPCIÓN"] = desc_new
            new_row["UBIC"] = ubic_new
            new_row["ÚLTIMO COSTO"] = costo_new
            new_row["PRECIO BRUTO"] = precio_tienda_new
            new_row["PRECIO NETO"] = neto_loc
            new_row["MARGEN LOCAL"] = marg_loc
            new_row["MONTO EN SIMULACIÓN"] = monto_sim_new
            new_row[" NETO MELI 1"] = neto_mel
            new_row["MARGEN MELI 1"] = marg_mel
            new_row["MLC"] = normalize_mlc(mlc1_new)
            new_row["PRECIO B2C PUBLICADO "] = b2c1_new if b2c1_new > 0 else np.nan
            new_row["FECHA VENCI"] = pd.to_datetime(fecha1_new).date()
            new_row["COMENTARIO"] = comment1_new
            st.session_state.master_df = pd.concat([st.session_state.master_df, pd.DataFrame([new_row])], ignore_index=True)

            bridge_rows = []
            for mlc in [normalize_mlc(mlc1_new)] + [normalize_mlc(x) for x in mlcs_extra.split(",") if normalize_mlc(x)]:
                if mlc:
                    bridge_rows.append({"SKU": sku_norm, "Número de publicación": mlc, "SKU_norm": sku_norm, "MLC_norm": mlc})
            if bridge_rows:
                st.session_state.bridge_df = pd.concat([st.session_state.bridge_df, pd.DataFrame(bridge_rows)], ignore_index=True)

            if add_rel and b2c1_new > 0:
                rel_row = {
                    "SKU": sku_norm,
                    "DESCRIPCION": desc_new,
                    "PRECIO_B2C": b2c1_new,
                    "MLC": normalize_mlc(mlc1_new),
                    "COMENTARIO": comment1_new,
                    "ESTADO": "",
                    "SKU_norm": sku_norm,
                    "MLC_norm": normalize_mlc(mlc1_new),
                }
                st.session_state.relampago_df = pd.concat([st.session_state.relampago_df, pd.DataFrame([rel_row])], ignore_index=True)

            st.success("Producto creado.")
            st.rerun()

# ----------------------------- Download ----------------------------- #

st.divider()
master_export = st.session_state.master_df.drop(columns=["SKU_norm"], errors="ignore")
bridge_export = st.session_state.bridge_df.drop(columns=["SKU_norm", "MLC_norm"], errors="ignore")

rel = st.session_state.relampago_df.copy()
rel_export = rel[["SKU", "DESCRIPCION", "PRECIO_B2C", "MLC", "COMENTARIO", "ESTADO"]].copy()

all_sheets = {}
for name in st.session_state.sheet_order:
    if name == SHEET_MASTER:
        all_sheets[name] = master_export
    elif name == SHEET_BRIDGE:
        all_sheets[name] = bridge_export
    elif name == SHEET_RELAMPAGO:
        all_sheets[name] = rel_export
    else:
        all_sheets[name] = st.session_state.other_sheets[name]

download_bytes = build_download_bytes(all_sheets)
st.download_button(
    "Descargar Excel actualizado",
    data=download_bytes,
    file_name="MAESTRA_PRECIOS_Y_PROMOS_actualizada.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
