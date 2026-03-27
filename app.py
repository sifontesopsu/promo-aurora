
import io
import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Aurora Promos", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {background: #fafafa; border: 1px solid #ececec; padding: .6rem .8rem; border-radius: 14px;}
.promo-chip {display:inline-block; padding:.2rem .55rem; border-radius:999px; background:#f4f4f5; font-size:.85rem; margin-right:.35rem;}
.small-muted {color:#666; font-size:.88rem;}
.card-shell {border:1px solid #e8e8ea; border-radius:18px; padding:.55rem .7rem; background:white; min-height:124px;}
.section-shell {border:1px solid #ececec; border-radius:18px; padding: .9rem 1rem; background: #fff;}
</style>
""", unsafe_allow_html=True)

VAT = 1.19

# ---------- Helpers ----------
def clean_str(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    return s

def norm_sku(x):
    s = clean_str(x)
    if not s:
        return ""
    s = re.sub(r"\.0$", "", s)
    return s

def norm_mlc(x):
    s = clean_str(x).upper().replace(" ", "")
    return s

def coalesce(*values):
    for v in values:
        if pd.notna(v) and clean_str(v) != "":
            return v
    return np.nan

def to_float(x):
    try:
        if pd.isna(x) or clean_str(x) == "":
            return np.nan
        return float(x)
    except Exception:
        return np.nan

def to_date_value(x):
    if pd.isna(x) or clean_str(x) == "":
        return None
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def fmt_date(x):
    d = to_date_value(x)
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")

def fmt_money(x):
    v = to_float(x)
    if pd.isna(v):
        return "-"
    return f"${int(round(v)):,}".replace(",", ".")

def fmt_pct(x):
    v = to_float(x)
    if pd.isna(v):
        return "-"
    return f"{v:.2f}%"

def margin_local(cost, bruto):
    c = to_float(cost)
    b = to_float(bruto)
    if pd.isna(c) or pd.isna(b) or b == 0:
        return np.nan
    neto = b / VAT
    if neto == 0:
        return np.nan
    return ((neto - c) / neto) * 100

def margin_meli(cost, simulacion):
    c = to_float(cost)
    s = to_float(simulacion)
    if pd.isna(c) or pd.isna(s) or s == 0:
        return np.nan
    neto = s / VAT
    if neto == 0:
        return np.nan
    return ((neto - c) / neto) * 100

def signature_from_upload(upload):
    if upload is None:
        return None
    payload = upload.getvalue()
    return (upload.name, len(payload), hash(payload[:100000]))

@st.cache_data(show_spinner=False)
def read_excel_book(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return xls.sheet_names

@st.cache_data(show_spinner=False)
def load_master_book(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names
    required = ["MAESTRA de precios", "MLC -SKU"]
    missing = [s for s in required if s not in sheet_names]
    if missing:
        raise ValueError(f"Faltan hojas requeridas: {', '.join(missing)}")

    master = pd.read_excel(io.BytesIO(file_bytes), sheet_name="MAESTRA de precios")
    bridge = pd.read_excel(io.BytesIO(file_bytes), sheet_name="MLC -SKU")

    rel_sheet = None
    for s in sheet_names:
        if clean_str(s).lower() == "relampago mi pagina":
            rel_sheet = s
            break

    if rel_sheet:
        rel_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=rel_sheet, header=None)
        # force first 6 columns only, assign clean names
        rel_raw = rel_raw.iloc[:, :6].copy()
        while rel_raw.shape[1] < 6:
            rel_raw[rel_raw.shape[1]] = np.nan
        rel_raw.columns = ["SKU", "DESCRIPCION", "PRECIO_B2C", "AUX", "COMENTARIO", "ESTADO"]
        rel_raw = rel_raw.dropna(how="all")
    else:
        rel_raw = pd.DataFrame(columns=["SKU", "DESCRIPCION", "PRECIO_B2C", "AUX", "COMENTARIO", "ESTADO"])

    return master, bridge, rel_raw, sheet_names

@st.cache_data(show_spinner=False)
def load_purchases(file_bytes: bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    cols = {c: clean_str(c) for c in df.columns}
    # normalize purchase columns
    sku_col = next((c for c in df.columns if clean_str(c).lower() == "sku"), None)
    fecha_col = next((c for c in df.columns if clean_str(c).lower() == "fecha"), None)
    proveedor_col = next((c for c in df.columns if clean_str(c).lower() in ["razón social", "razon social"]), None)
    precio_col = next((c for c in df.columns if clean_str(c).lower().startswith("precio un")), None)
    cantidad_col = next((c for c in df.columns if clean_str(c).lower() == "cantidad"), None)
    articulo_col = next((c for c in df.columns if "concepto / artículo" in clean_str(c).lower() or "concepto / articulo" in clean_str(c).lower()), None)

    out = pd.DataFrame()
    out["SKU_norm"] = df[sku_col].map(norm_sku) if sku_col else ""
    out["Fecha"] = pd.to_datetime(df[fecha_col], dayfirst=True, errors="coerce") if fecha_col else pd.NaT
    out["Proveedor"] = df[proveedor_col].map(clean_str) if proveedor_col else ""
    out["Precio"] = pd.to_numeric(df[precio_col], errors="coerce") if precio_col else np.nan
    out["Cantidad"] = pd.to_numeric(df[cantidad_col], errors="coerce") if cantidad_col else np.nan
    out["Articulo"] = df[articulo_col].map(clean_str) if articulo_col else ""
    out = out[(out["SKU_norm"] != "") | (out["Articulo"] != "")]
    out = out.sort_values(["SKU_norm", "Fecha"], ascending=[True, False])

    if out.empty:
        summary = pd.DataFrame(columns=["SKU_norm", "ultima_fecha", "ultimo_precio", "ultimo_proveedor", "cantidad_ultima", "precio_anterior", "variacion_pct"])
    else:
        def summarize(g):
            g = g.sort_values("Fecha", ascending=False)
            first = g.iloc[0]
            second_price = g.iloc[1]["Precio"] if len(g) > 1 else np.nan
            var_pct = np.nan
            if pd.notna(first["Precio"]) and pd.notna(second_price) and second_price != 0:
                var_pct = ((first["Precio"] - second_price) / second_price) * 100
            return pd.Series({
                "SKU_norm": first["SKU_norm"],
                "ultima_fecha": first["Fecha"],
                "ultimo_precio": first["Precio"],
                "ultimo_proveedor": first["Proveedor"],
                "cantidad_ultima": first["Cantidad"],
                "precio_anterior": second_price,
                "variacion_pct": var_pct,
            })
        summary = out.groupby("SKU_norm", dropna=False).apply(summarize).reset_index(drop=True)
    return out, summary

def prepare_master(master_df, bridge_df, rel_raw):
    master = master_df.copy()
    master.columns = [clean_str(c) for c in master.columns]
    bridge = bridge_df.copy()
    bridge.columns = [clean_str(c) for c in bridge.columns]
    rel = rel_raw.copy()

    if "SKU" not in master.columns:
        raise ValueError("La hoja MAESTRA de precios no tiene columna SKU.")
    master["SKU_norm"] = master["SKU"].map(norm_sku)
    master["DESCRIPCIÓN"] = master.get("DESCRIPCIÓN", "").map(clean_str)
    master["MLC_BASE"] = master.get("Unnamed: 12", np.nan).map(norm_mlc) if "Unnamed: 12" in master.columns else ""

    bridge["SKU_norm"] = bridge["SKU"].map(norm_sku) if "SKU" in bridge.columns else ""
    pub_col = next((c for c in bridge.columns if clean_str(c).lower() == "número de publicación" or clean_str(c).lower() == "numero de publicacion"), None)
    bridge["MLC_norm"] = bridge[pub_col].map(norm_mlc) if pub_col else ""

    rel["SKU_norm"] = rel["SKU"].map(norm_sku) if "SKU" in rel.columns else ""
    rel["DESCRIPCION"] = rel.get("DESCRIPCION", "").map(clean_str)
    rel["COMENTARIO"] = rel.get("COMENTARIO", "").map(clean_str)
    rel["ESTADO"] = rel.get("ESTADO", "").map(clean_str)
    rel["PRECIO_B2C"] = pd.to_numeric(rel.get("PRECIO_B2C"), errors="coerce")
    rel = rel[(rel["SKU_norm"] != "") | (rel["DESCRIPCION"] != "")]

    # promos only from master
    promo_rows = []
    slot_specs = [
        {"slot": 1, "mlc_col": "MLC", "price_col": "PRECIO B2C PUBLICADO", "date_col": "FECHA VENCI", "comment_col": "COMENTARIO"},
        {"slot": 2, "mlc_col": "MLC.1", "price_col": "PRECIO B2C", "date_col": "FECHA VENCI.1", "comment_col": "COMENTARIO.1"},
    ]
    # account for exact column with trailing space
    for i, spec in enumerate(slot_specs):
        if spec["price_col"] not in master.columns:
            for c in master.columns:
                if clean_str(c).lower() == clean_str(spec["price_col"]).lower():
                    spec["price_col"] = c
        if spec["mlc_col"] not in master.columns:
            continue
        for idx, row in master.iterrows():
            mlc = norm_mlc(row.get(spec["mlc_col"]))
            price = to_float(row.get(spec["price_col"]))
            dt = to_date_value(row.get(spec["date_col"]))
            comm = clean_str(row.get(spec["comment_col"]))
            if mlc or pd.notna(price) or dt or comm:
                promo_rows.append({
                    "source_row": idx,
                    "slot": spec["slot"],
                    "SKU_norm": row["SKU_norm"],
                    "SKU": row.get("SKU"),
                    "DESCRIPCION": clean_str(row.get("DESCRIPCIÓN")),
                    "MLC_norm": mlc,
                    "PRECIO_B2C": price,
                    "FECHA_VENCI": dt,
                    "COMENTARIO": comm,
                })
    promos = pd.DataFrame(promo_rows)
    if promos.empty:
        promos = pd.DataFrame(columns=["source_row","slot","SKU_norm","SKU","DESCRIPCION","MLC_norm","PRECIO_B2C","FECHA_VENCI","COMENTARIO"])

    # bridge summary
    mlc_map = bridge.groupby("SKU_norm")["MLC_norm"].apply(lambda s: sorted([x for x in s if x])).to_dict()
    return master, bridge, rel, promos, mlc_map

def promo_status(dt):
    d = to_date_value(dt)
    if not d:
        return "Sin fecha"
    delta = (d - date.today()).days
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

def build_download(master_df, bridge_df, rel_df, original_bytes):
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            if sheet == "MAESTRA de precios":
                master_df.drop(columns=[c for c in master_df.columns if c.endswith("_norm") or c in ["MLC_BASE"]], errors="ignore").to_excel(writer, sheet_name=sheet, index=False)
            elif sheet == "MLC -SKU":
                bridge_df.drop(columns=[c for c in bridge_df.columns if c.endswith("_norm")], errors="ignore").to_excel(writer, sheet_name=sheet, index=False)
            elif clean_str(sheet).lower() == "relampago mi pagina":
                rel_out = rel_df.copy()
                rel_out = rel_out[["SKU","DESCRIPCION","PRECIO_B2C","AUX","COMENTARIO","ESTADO"]]
                rel_out.to_excel(writer, sheet_name=sheet, header=False, index=False)
            else:
                pd.read_excel(io.BytesIO(original_bytes), sheet_name=sheet).to_excel(writer, sheet_name=sheet, index=False)
    return output.getvalue()

def product_options(master):
    opts = []
    for _, r in master.iterrows():
        sku = r["SKU_norm"]
        desc = clean_str(r.get("DESCRIPCIÓN"))
        opts.append((f"{sku} · {desc[:80]}", sku))
    return opts

def ensure_state(master_bytes, purchases_bytes):
    sig = (signature_from_upload(master_bytes), signature_from_upload(purchases_bytes))
    if st.session_state.get("loaded_sig") == sig:
        return
    master_df, bridge_df, rel_raw, sheet_names = load_master_book(master_bytes.getvalue())
    master, bridge, rel, promos, mlc_map = prepare_master(master_df, bridge_df, rel_raw)
    st.session_state.master_df = master
    st.session_state.bridge_df = bridge
    st.session_state.rel_df = rel
    st.session_state.promos_df = promos
    st.session_state.mlc_map = mlc_map
    st.session_state.sheet_names = sheet_names
    if purchases_bytes is not None:
        purchases, summary = load_purchases(purchases_bytes.getvalue())
    else:
        purchases = pd.DataFrame(columns=["SKU_norm","Fecha","Proveedor","Precio","Cantidad","Articulo"])
        summary = pd.DataFrame(columns=["SKU_norm","ultima_fecha","ultimo_precio","ultimo_proveedor","cantidad_ultima","precio_anterior","variacion_pct"])
    st.session_state.purchases_df = purchases
    st.session_state.purchase_summary = summary
    st.session_state.loaded_sig = sig
    st.session_state.data_dirty = False

# ---------- Sidebar ----------
st.sidebar.header("Archivos")
master_upload = st.sidebar.file_uploader("Maestra saneada", type=["xlsx"], key="master_upload")
purchases_upload = st.sidebar.file_uploader("Compras", type=["xlsx"], key="purchases_upload")
if not master_upload:
    st.info("Sube la maestra para comenzar.")
    st.stop()

try:
    ensure_state(master_upload, purchases_upload)
except Exception as e:
    st.error(f"No pude leer o modelar la maestra: {e}")
    st.stop()

master = st.session_state.master_df
bridge = st.session_state.bridge_df
rel_df = st.session_state.rel_df
promos = st.session_state.promos_df
mlc_map = st.session_state.mlc_map
purchases_df = st.session_state.purchases_df
purchase_summary = st.session_state.purchase_summary

with st.sidebar:
    st.markdown("---")
    if st.button("Recargar desde archivos"):
        st.session_state.loaded_sig = None
        st.rerun()
    dl_name = f"MAESTRA_PRECIOS_Y_PROMOS_actualizada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    download_bytes = build_download(master, bridge, rel_df, master_upload.getvalue())
    st.download_button("Descargar Excel actualizado", data=download_bytes, file_name=dl_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

tabs = st.tabs(["Cockpit por producto", "Operador de promos", "Relámpago", "Alta de producto"])

# ---------- Cockpit ----------
with tabs[0]:
    st.subheader("Cockpit por producto")
    opts = product_options(master)
    label_map = {label: sku for label, sku in opts}
    choice = st.selectbox("Buscar producto", options=list(label_map.keys()), index=0 if opts else None)
    if not opts:
        st.warning("No hay productos en la maestra.")
    else:
        sku = label_map[choice]
        row = master[master["SKU_norm"] == sku].iloc[0]
        desc = clean_str(row.get("DESCRIPCIÓN"))
        mlcs = sorted(set([x for x in mlc_map.get(sku, []) if x] + [norm_mlc(row.get("MLC")), norm_mlc(row.get("MLC.1")), norm_mlc(row.get("MLC_BASE"))]))
        mlcs = [x for x in mlcs if x]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("SKU", sku)
        c2.metric("Precio bruto tienda", fmt_money(row.get("PRECIO BRUTO")))
        c3.metric("Monto en simulación", fmt_money(row.get("MONTO EN SIMULACIÓN")))
        c4.metric("Promos asociadas", int((promos["SKU_norm"] == sku).sum()))

        left, right = st.columns([1.2, 1])
        with left:
            st.markdown('<div class="section-shell">', unsafe_allow_html=True)
            st.markdown(f"### {desc}")
            chips = "".join([f'<span class="promo-chip">{m}</span>' for m in mlcs]) or '<span class="small-muted">Sin MLC asociado</span>'
            st.markdown(chips, unsafe_allow_html=True)
            p1, p2, p3 = st.columns(3)
            p1.metric("Último costo", fmt_money(row.get("ÚLTIMO COSTO")))
            p2.metric("Margen local", fmt_pct(row.get("MARGEN LOCAL")))
            p3.metric("Margen Meli 1", fmt_pct(row.get("MARGEN MELI 1")))
            st.markdown(f"<div class='small-muted'>Ubicación: {clean_str(row.get('UBIC'))}</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("#### Promos de maestra")
            product_promos = promos[promos["SKU_norm"] == sku].copy()
            if product_promos.empty:
                st.info("No tiene promo cargada en la maestra.")
            else:
                show = product_promos[["slot", "MLC_norm", "PRECIO_B2C", "FECHA_VENCI", "COMENTARIO"]].copy()
                show["FECHA_VENCI"] = show["FECHA_VENCI"].map(fmt_date)
                show["PRECIO_B2C"] = show["PRECIO_B2C"].map(fmt_money)
                show["Estado"] = product_promos["FECHA_VENCI"].map(promo_status)
                show = show.rename(columns={"slot":"Slot","MLC_norm":"MLC","PRECIO_B2C":"Precio B2C","FECHA_VENCI":"Fecha venci","COMENTARIO":"Comentario"})
                st.dataframe(show, use_container_width=True, hide_index=True)

        with right:
            st.markdown("#### Compras")
            purchase_row = purchase_summary[purchase_summary["SKU_norm"] == sku]
            if purchase_row.empty:
                st.info("Sin compras detectadas para este SKU.")
            else:
                pr = purchase_row.iloc[0]
                a1, a2 = st.columns(2)
                a1.metric("Última compra", fmt_date(pr["ultima_fecha"]))
                a2.metric("Último precio", fmt_money(pr["ultimo_precio"]))
                b1, b2 = st.columns(2)
                b1.metric("Proveedor", clean_str(pr["ultimo_proveedor"]) or "-")
                b2.metric("Variación vs anterior", fmt_pct(pr["variacion_pct"]))
                hist = purchases_df[purchases_df["SKU_norm"] == sku].copy().head(15)
                hist["Fecha"] = hist["Fecha"].dt.strftime("%d/%m/%Y")
                hist["Precio"] = hist["Precio"].map(fmt_money)
                st.dataframe(hist[["Fecha","Proveedor","Cantidad","Precio","Articulo"]], use_container_width=True, hide_index=True)

            st.markdown("#### Relámpago mi página")
            rel_row = rel_df[rel_df["SKU_norm"] == sku]
            if rel_row.empty:
                st.info("No está en relámpago.")
            else:
                rr = rel_row.iloc[0]
                r1, r2 = st.columns(2)
                r1.metric("Precio relámpago", fmt_money(rr.get("PRECIO_B2C")))
                r2.metric("Estado", clean_str(rr.get("ESTADO")) or "-")
                st.caption(clean_str(rr.get("COMENTARIO")) or "Sin comentario")

# ---------- Dialog ----------
@st.dialog("Editar promo")
def edit_promo_dialog(promo_idx):
    promo_row = st.session_state.promos_df.loc[promo_idx]
    st.write(f"**SKU:** {promo_row['SKU_norm']}")
    st.write(clean_str(promo_row["DESCRIPCION"]))
    new_price = st.number_input("Precio B2C", min_value=0.0, step=100.0, value=float(to_float(promo_row["PRECIO_B2C"]) or 0))
    current_date = to_date_value(promo_row["FECHA_VENCI"])
    new_date = st.date_input("Fecha venci", value=current_date or date.today(), format="DD/MM/YYYY")
    new_comment = st.text_area("Comentario", value=clean_str(promo_row["COMENTARIO"]), height=90)
    c1, c2 = st.columns(2)
    if c1.button("Guardar", use_container_width=True):
        slot = int(promo_row["slot"])
        row_idx = int(promo_row["source_row"])
        price_col = "PRECIO B2C PUBLICADO " if slot == 1 else "PRECIO B2C"
        date_col = "FECHA VENCI" if slot == 1 else "FECHA VENCI.1"
        comment_col = "COMENTARIO" if slot == 1 else "COMENTARIO.1"
        st.session_state.master_df.at[row_idx, price_col] = new_price
        st.session_state.master_df.at[row_idx, date_col] = pd.Timestamp(new_date)
        st.session_state.master_df.at[row_idx, comment_col] = new_comment
        st.session_state.promos_df.at[promo_idx, "PRECIO_B2C"] = new_price
        st.session_state.promos_df.at[promo_idx, "FECHA_VENCI"] = new_date
        st.session_state.promos_df.at[promo_idx, "COMENTARIO"] = new_comment
        st.session_state.data_dirty = True
        st.rerun()
    if c2.button("Cancelar", use_container_width=True):
        st.rerun()

# ---------- Operador ----------
with tabs[1]:
    st.subheader("Operador de promos")
    origin_filter = st.selectbox("Origen", ["Promos maestra", "Relámpago"], index=0)
    if origin_filter == "Promos maestra":
        band = promos.copy()
        q = st.text_input("Buscar por SKU, descripción o MLC")
        urgency = st.selectbox("Filtro fecha", ["Todas", "Vence hoy", "Vence mañana", "Vence en 3 días", "Sin fecha", "Vencida"])
        if q:
            qq = q.strip().lower()
            band = band[
                band["SKU_norm"].str.lower().str.contains(qq, na=False) |
                band["DESCRIPCION"].str.lower().str.contains(qq, na=False) |
                band["MLC_norm"].str.lower().str.contains(qq, na=False)
            ]
        if urgency != "Todas":
            band["status"] = band["FECHA_VENCI"].map(promo_status)
            band = band[band["status"] == urgency]
        band = band.sort_values(by=["FECHA_VENCI","SKU_norm"], ascending=[True, True])

        st.markdown("#### Cambio masivo de fecha")
        m1, m2, m3 = st.columns([1,1,1.4])
        mass_date = m1.date_input("Nueva fecha", value=date.today(), format="DD/MM/YYYY", key="mass_date")
        slot_choice = m2.selectbox("Aplicar a slot", ["Todos", "1", "2"], index=0)
        if m3.button("Aplicar a filas filtradas", use_container_width=True):
            target_idx = band.index.tolist()
            if slot_choice != "Todos":
                target_idx = [i for i in target_idx if int(st.session_state.promos_df.loc[i, "slot"]) == int(slot_choice)]
            for idx in target_idx:
                rowp = st.session_state.promos_df.loc[idx]
                row_idx = int(rowp["source_row"])
                date_col = "FECHA VENCI" if int(rowp["slot"]) == 1 else "FECHA VENCI.1"
                st.session_state.master_df.at[row_idx, date_col] = pd.Timestamp(mass_date)
                st.session_state.promos_df.at[idx, "FECHA_VENCI"] = mass_date
            st.session_state.data_dirty = True
            st.success(f"Fecha aplicada a {len(target_idx)} promo(s).")

        st.markdown("#### Bandeja")
        if band.empty:
            st.info("No hay promos para mostrar con esos filtros.")
        else:
            cols = st.columns(4)
            for i, (idx, r) in enumerate(band.iterrows()):
                with cols[i % 4]:
                    txt = f"{r['SKU_norm']}\n{clean_str(r['DESCRIPCION'])[:34]}\n{clean_str(r['MLC_norm']) or '-'}\n{fmt_date(r['FECHA_VENCI']) or 'Sin fecha'}"
                    if st.button(txt, key=f"card_{idx}", use_container_width=True):
                        edit_promo_dialog(idx)
    else:
        st.markdown("#### Lista Relámpago")
        q = st.text_input("Buscar relámpago", key="rel_search")
        show_rel = rel_df.copy()
        if q:
            qq = q.strip().lower()
            show_rel = show_rel[
                show_rel["SKU_norm"].str.lower().str.contains(qq, na=False) |
                show_rel["DESCRIPCION"].str.lower().str.contains(qq, na=False)
            ]
        st.dataframe(show_rel[["SKU","DESCRIPCION","PRECIO_B2C","COMENTARIO","ESTADO"]], use_container_width=True, hide_index=True)

# ---------- Relámpago ----------
with tabs[2]:
    st.subheader("Relámpago mi página")
    st.caption("Lista editable simple.")
    if rel_df.empty:
        st.info("No hay registros en Relámpago.")
    else:
        edited = st.data_editor(
            rel_df[["SKU","DESCRIPCION","PRECIO_B2C","COMENTARIO","ESTADO"]].reset_index(drop=True),
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            key="rel_editor"
        )
        if st.button("Guardar cambios relámpago"):
            keep_aux = st.session_state.rel_df.get("AUX", pd.Series([np.nan]*len(st.session_state.rel_df)))
            new_rel = edited.copy()
            new_rel["AUX"] = np.nan
            new_rel["SKU_norm"] = new_rel["SKU"].map(norm_sku)
            st.session_state.rel_df = new_rel[["SKU","DESCRIPCION","PRECIO_B2C","AUX","COMENTARIO","ESTADO","SKU_norm"]]
            st.session_state.data_dirty = True
            st.success("Relámpago actualizado en memoria.")

# ---------- Alta ----------
with tabs[3]:
    st.subheader("Alta de producto")
    c1, c2, c3 = st.columns(3)
    sku_new = c1.text_input("SKU nuevo")
    desc_new = c2.text_input("Descripción")
    ubic_new = c3.text_input("Ubicación")

    c4, c5, c6 = st.columns(3)
    cost_new = c4.number_input("Último costo", min_value=0.0, step=100.0, value=0.0)
    bruto_tienda = c5.number_input("Precio bruto en tienda", min_value=0.0, step=100.0, value=0.0)
    monto_sim = c6.number_input("Monto en simulación", min_value=0.0, step=100.0, value=0.0)

    c7, c8, c9 = st.columns(3)
    margin_loc = margin_local(cost_new, bruto_tienda)
    margin_me = margin_meli(cost_new, monto_sim)
    c7.metric("Margen local proyectado", fmt_pct(margin_loc))
    c8.metric("Margen Meli 1 proyectado", fmt_pct(margin_me))
    c9.metric("Precio neto tienda", fmt_money(bruto_tienda / VAT if bruto_tienda else np.nan))

    st.markdown("#### Promo base")
    p1, p2, p3, p4 = st.columns(4)
    mlc1 = p1.text_input("MLC", key="new_mlc1")
    b2c1 = p2.number_input("Precio B2C publicado", min_value=0.0, step=100.0, value=0.0, key="new_b2c1")
    fecha1 = p3.date_input("Fecha venci", value=date.today(), format="DD/MM/YYYY", key="new_date1")
    comm1 = p4.text_input("Comentario", key="new_comm1")

    add_rel = st.checkbox("Agregar a Relámpago mi página")
    rel_estado = st.text_input("Estado relámpago", value="LIQUIDACION") if add_rel else ""

    if st.button("Crear producto"):
        if not norm_sku(sku_new):
            st.error("Debes ingresar un SKU válido.")
        else:
            new_row = {c: np.nan for c in st.session_state.master_df.columns}
            new_row["SKU"] = sku_new
            new_row["SKU_norm"] = norm_sku(sku_new)
            new_row["DESCRIPCIÓN"] = desc_new
            new_row["UBIC"] = ubic_new
            new_row["ÚLTIMO COSTO"] = cost_new
            new_row["PRECIO BRUTO"] = bruto_tienda
            new_row["PRECIO NETO"] = bruto_tienda / VAT if bruto_tienda else np.nan
            new_row["MARGEN LOCAL"] = margin_loc
            new_row["MONTO EN SIMULACIÓN"] = monto_sim
            new_row[" NETO MELI 1"] = monto_sim / VAT if monto_sim else np.nan
            new_row["MARGEN MELI 1"] = margin_me
            new_row["MLC"] = norm_mlc(mlc1)
            # promo fields from current model
            price_col = "PRECIO B2C PUBLICADO " if "PRECIO B2C PUBLICADO " in st.session_state.master_df.columns else "PRECIO B2C PUBLICADO"
            new_row[price_col] = b2c1 if b2c1 else np.nan
            new_row["FECHA VENCI"] = pd.Timestamp(fecha1) if fecha1 else pd.NaT
            new_row["COMENTARIO"] = comm1
            st.session_state.master_df = pd.concat([st.session_state.master_df, pd.DataFrame([new_row])], ignore_index=True)

            if norm_mlc(mlc1):
                bridge_cols = list(st.session_state.bridge_df.columns)
                pub_col = next((c for c in bridge_cols if clean_str(c).lower() in ["número de publicación", "numero de publicacion"]), bridge_cols[1] if len(bridge_cols) > 1 else "Número de publicación")
                new_bridge = {c: np.nan for c in bridge_cols}
                new_bridge["SKU"] = sku_new
                new_bridge["SKU_norm"] = norm_sku(sku_new)
                new_bridge[pub_col] = norm_mlc(mlc1)
                new_bridge["MLC_norm"] = norm_mlc(mlc1)
                st.session_state.bridge_df = pd.concat([st.session_state.bridge_df, pd.DataFrame([new_bridge])], ignore_index=True)

            if add_rel:
                rel_new = {
                    "SKU": sku_new,
                    "DESCRIPCION": desc_new,
                    "PRECIO_B2C": b2c1 if b2c1 else np.nan,
                    "AUX": np.nan,
                    "COMENTARIO": comm1,
                    "ESTADO": rel_estado,
                    "SKU_norm": norm_sku(sku_new),
                }
                st.session_state.rel_df = pd.concat([st.session_state.rel_df, pd.DataFrame([rel_new])], ignore_index=True)

            # refresh derived data
            st.session_state.master_df, st.session_state.bridge_df, st.session_state.rel_df, st.session_state.promos_df, st.session_state.mlc_map = prepare_master(
                st.session_state.master_df.drop(columns=["SKU_norm","MLC_BASE"], errors="ignore"),
                st.session_state.bridge_df.drop(columns=["SKU_norm","MLC_norm"], errors="ignore"),
                st.session_state.rel_df.drop(columns=["SKU_norm"], errors="ignore"),
            )
            st.session_state.data_dirty = True
            st.success("Producto creado.")
