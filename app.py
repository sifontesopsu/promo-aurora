
import io
import hashlib
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Aurora Pricing", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
def file_signature(uploaded_file) -> str:
    data = uploaded_file.getvalue()
    return hashlib.md5(data).hexdigest()


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


def re_is_numberlike(s: str) -> bool:
    import re
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", s))


def norm_mlc(value) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    s = str(value).strip().upper()
    if not s or s == "NAN":
        return ""
    return s


def safe_float(value, default=np.nan):
    try:
        if value is None or (isinstance(value, str) and not value.strip()):
            return default
        return float(value)
    except Exception:
        return default


def to_date_only(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return pd.NaT
    try:
        return pd.to_datetime(value, errors="coerce").normalize()
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


def margin_display(value) -> str:
    x = safe_float(value, np.nan)
    if np.isnan(x):
        return "—"
    # normalize decimal fractions to percentage points
    if abs(x) <= 2:
        x = x * 100
    return f"{x:.2f}%"


def calc_margin_local(cost, price_bruto):
    cost = safe_float(cost, np.nan)
    bruto = safe_float(price_bruto, np.nan)
    if np.isnan(cost) or np.isnan(bruto) or bruto == 0:
        return np.nan
    neto = bruto / 1.19
    if neto == 0:
        return np.nan
    return ((neto - cost) / neto) * 100


def calc_margin_meli1(cost, monto_sim):
    cost = safe_float(cost, np.nan)
    monto = safe_float(monto_sim, np.nan)
    if np.isnan(cost) or np.isnan(monto) or monto == 0:
        return np.nan
    neto = monto / 1.19
    if neto == 0:
        return np.nan
    return ((neto - cost) / neto) * 100


def promo_status(dt):
    dt = to_date_only(dt)
    if pd.isna(dt):
        return "Vencen en 1 mes", 30
    today = pd.Timestamp(date.today())
    delta = (dt - today).days
    if delta < 0:
        return "Vencidas", -1
    if delta == 0:
        return "Vencen hoy", 0
    if delta == 1:
        return "Vencen mañana", 1
    if delta == 2:
        return "Vencen pasado mañana", 2
    if delta <= 7:
        return "Vencen en 7 días", 7
    if delta <= 15:
        return "Vencen en 15 días", 15
    return "Vencen en 1 mes", 30


def _find_sheet(sheet_names, wanted):
    lowers = {name.lower().strip(): name for name in sheet_names}
    for name in sheet_names:
        low = name.lower().strip()
        if low == wanted.lower().strip():
            return name
    for name in sheet_names:
        if wanted.lower().strip() in name.lower().strip():
            return name
    return None


@st.cache_data(show_spinner=False)
def load_workbook_cached(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xls.sheet_names
    maestra_name = _find_sheet(names, "MAESTRA de precios")
    bridge_name = _find_sheet(names, "MLC -SKU")
    rel_name = _find_sheet(names, "Relampago mi pagina")
    if not maestra_name or not bridge_name:
        raise ValueError(f"Hojas requeridas no encontradas. Disponibles: {', '.join(names)}")

    master_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=maestra_name)
    bridge_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=bridge_name)
    rel_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=rel_name, header=None) if rel_name else pd.DataFrame()

    return {
        "sheet_names": names,
        "maestra_name": maestra_name,
        "bridge_name": bridge_name,
        "rel_name": rel_name,
        "master_df": master_df,
        "bridge_df": bridge_df,
        "rel_df": rel_df,
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
    if "SKU" not in raw.columns:
        raw["SKU"] = np.nan
    if "Fecha" not in raw.columns:
        raw["Fecha"] = pd.NaT
    if "Razón Social" not in raw.columns:
        raw["Razón Social"] = np.nan
    if "Precio Un." not in raw.columns:
        raw["Precio Un."] = np.nan
    if "Cantidad" not in raw.columns:
        raw["Cantidad"] = np.nan

    raw["SKU_norm"] = raw["SKU"].map(norm_sku)
    raw["Fecha_dt"] = pd.to_datetime(raw["Fecha"], dayfirst=True, errors="coerce").dt.normalize()
    raw = raw[raw["SKU_norm"] != ""].copy()
    raw = raw.sort_values(["SKU_norm", "Fecha_dt"])

    if raw.empty:
        return {
            "raw": raw,
            "summary": pd.DataFrame(columns=["SKU_norm", "ultima_fecha", "ultimo_precio", "ultimo_proveedor", "ultima_cantidad", "compra_anterior", "variacion_pct"]),
            "by_sku": {}
        }

    by_sku = {sku: grp.copy() for sku, grp in raw.groupby("SKU_norm", sort=False)}

    summary_rows = []
    for sku, grp in by_sku.items():
        grp = grp.sort_values("Fecha_dt")
        last = grp.iloc[-1]
        prev_price = safe_float(grp.iloc[-2]["Precio Un."], np.nan) if len(grp) >= 2 else np.nan
        last_price = safe_float(last["Precio Un."], np.nan)
        variation = np.nan
        if not np.isnan(last_price) and not np.isnan(prev_price) and prev_price != 0:
            variation = ((last_price - prev_price) / prev_price) * 100
        summary_rows.append({
            "SKU_norm": sku,
            "ultima_fecha": last["Fecha_dt"],
            "ultimo_precio": last_price,
            "ultimo_proveedor": last.get("Razón Social", ""),
            "ultima_cantidad": safe_float(last.get("Cantidad"), np.nan),
            "compra_anterior": prev_price,
            "variacion_pct": variation,
            "compras_total": len(grp),
        })
    summary = pd.DataFrame(summary_rows)
    return {"raw": raw, "summary": summary, "by_sku": by_sku}


def normalize_master(master_df: pd.DataFrame) -> pd.DataFrame:
    df = master_df.copy()
    # normalize expected cols
    for col in ["SKU", "DESCRIPCIÓN", "UBIC", "ÚLTIMO COSTO", "MARGEN LOCAL", "PRECIO BRUTO",
                "MARGEN MELI 1", "MONTO EN SIMULACIÓN", "PRECIO B2C PUBLICADO ", "FECHA VENCI",
                "COMENTARIO", "MLC", "Unnamed: 12", "MLC.1", "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1"]:
        if col not in df.columns:
            df[col] = np.nan

    df["SKU_norm"] = df["SKU"].map(norm_sku)
    df["DESC_norm"] = df["DESCRIPCIÓN"].fillna("").astype(str)
    df["MLC_slot1"] = df["MLC"].map(norm_mlc)
    fallback = df["Unnamed: 12"].map(norm_mlc) if "Unnamed: 12" in df.columns else ""
    df.loc[df["MLC_slot1"] == "", "MLC_slot1"] = fallback[df["MLC_slot1"] == ""]
    df["MLC_slot2"] = df["MLC.1"].map(norm_mlc)

    for c in ["FECHA VENCI", "FECHA VENCI.1"]:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.normalize()

    return df


def normalize_bridge(bridge_df: pd.DataFrame) -> pd.DataFrame:
    df = bridge_df.copy()
    sku_col = "SKU" if "SKU" in df.columns else df.columns[0]
    mlc_col = "Número de publicación" if "Número de publicación" in df.columns else df.columns[-1]
    df["SKU_norm"] = df[sku_col].map(norm_sku)
    df["MLC_norm"] = df[mlc_col].map(norm_mlc)
    df = df[(df["SKU_norm"] != "") & (df["MLC_norm"] != "")].copy()
    return df[["SKU_norm", "MLC_norm"]].drop_duplicates()


def normalize_rel(rel_df: pd.DataFrame) -> pd.DataFrame:
    if rel_df is None or rel_df.empty:
        return pd.DataFrame(columns=["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"])
    df = rel_df.copy()
    # expected 6 columns no header
    while df.shape[1] < 6:
        df[df.shape[1]] = np.nan
    df = df.iloc[:, :6]
    df.columns = ["SKU_raw", "DESCRIPCION", "PRECIO_B2C", "EXTRA", "TIPO", "ESTADO"]
    df["SKU_norm"] = df["SKU_raw"].map(norm_sku)
    df = df[df["SKU_norm"] != ""].copy()
    return df[["SKU_norm", "DESCRIPCION", "PRECIO_B2C", "TIPO", "ESTADO"]]


def build_model(master_df, bridge_df, rel_df, purchases):
    master = normalize_master(master_df)
    bridge = normalize_bridge(bridge_df)
    rel = normalize_rel(rel_df)

    # map MLCs
    mlc_map = bridge.groupby("SKU_norm")["MLC_norm"].apply(list).to_dict()
    for idx, row in master.iterrows():
        sku = row["SKU_norm"]
        extra = []
        for mlc in [row["MLC_slot1"], row["MLC_slot2"]]:
            if mlc:
                extra.append(mlc)
        if sku in mlc_map:
            extra.extend(mlc_map[sku])
        mlc_map[sku] = sorted([m for m in set(extra) if m])

    promos_rows = []
    for idx, row in master.iterrows():
        sku = row["SKU_norm"]
        desc = row["DESCRIPCIÓN"]
        for slot, mlc_col, price_col, date_col, comment_col in [
            (1, "MLC_slot1", "PRECIO B2C PUBLICADO ", "FECHA VENCI", "COMENTARIO"),
            (2, "MLC_slot2", "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1"),
        ]:
            mlc = row[mlc_col]
            price = row[price_col]
            dt = row[date_col]
            comment = row[comment_col]
            if mlc or not pd.isna(price) or not pd.isna(dt) or (isinstance(comment, str) and comment.strip()):
                status, order = promo_status(dt)
                promos_rows.append({
                    "master_index": idx,
                    "SKU_norm": sku,
                    "DESCRIPCIÓN": desc,
                    "slot": slot,
                    "MLC": mlc,
                    "PRECIO_B2C": price,
                    "FECHA_VENCI": dt,
                    "COMENTARIO": comment,
                    "STATUS": status,
                    "STATUS_ORDER": order,
                })
    promos = pd.DataFrame(promos_rows)
    if promos.empty:
        promos = pd.DataFrame(columns=["master_index", "SKU_norm", "DESCRIPCIÓN", "slot", "MLC", "PRECIO_B2C", "FECHA_VENCI", "COMENTARIO", "STATUS", "STATUS_ORDER"])

    purchases_summary = purchases["summary"]
    purchase_map = purchases["by_sku"]

    sku_options = master[master["SKU_norm"] != ""]["SKU_norm"].tolist()
    sku_desc = {row["SKU_norm"]: row["DESCRIPCIÓN"] for _, row in master.iterrows() if row["SKU_norm"]}

    return {
        "master": master,
        "bridge": bridge,
        "rel": rel,
        "promos": promos,
        "mlc_map": mlc_map,
        "purchases_summary": purchases_summary,
        "purchase_map": purchase_map,
        "sku_options": sku_options,
        "sku_desc": sku_desc,
        "dirty": False,
    }


def init_state_from_upload(master_up, purchases_up):
    master_sig = file_signature(master_up)
    purchases_sig = file_signature(purchases_up) if purchases_up else ""
    combined = f"{master_sig}|{purchases_sig}"

    if st.session_state.get("model_sig") == combined:
        return

    wb = load_workbook_cached(master_up.getvalue())
    purchases = load_purchases_cached(purchases_up.getvalue() if purchases_up else b"")
    model = build_model(wb["master_df"], wb["bridge_df"], wb["rel_df"], purchases)
    st.session_state.model = model
    st.session_state.workbook_meta = wb
    st.session_state.model_sig = combined
    st.session_state.download_sig = ""


def get_product_row(model, sku):
    master = model["master"]
    rows = master[master["SKU_norm"] == sku]
    if rows.empty:
        return None
    return rows.iloc[0]


def refresh_promos_only():
    model = st.session_state.model
    master = model["master"]
    promos_rows = []
    for idx, row in master.iterrows():
        for slot, mlc_col, price_col, date_col, comment_col in [
            (1, "MLC_slot1", "PRECIO B2C PUBLICADO ", "FECHA VENCI", "COMENTARIO"),
            (2, "MLC_slot2", "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1"),
        ]:
            mlc = row[mlc_col]
            price = row[price_col]
            dt = row[date_col]
            comment = row[comment_col]
            if mlc or not pd.isna(price) or not pd.isna(dt) or (isinstance(comment, str) and comment.strip()):
                status, order = promo_status(dt)
                promos_rows.append({
                    "master_index": idx,
                    "SKU_norm": row["SKU_norm"],
                    "DESCRIPCIÓN": row["DESCRIPCIÓN"],
                    "slot": slot,
                    "MLC": mlc,
                    "PRECIO_B2C": price,
                    "FECHA_VENCI": dt,
                    "COMENTARIO": comment,
                    "STATUS": status,
                    "STATUS_ORDER": order,
                })
    model["promos"] = pd.DataFrame(promos_rows) if promos_rows else pd.DataFrame(columns=model["promos"].columns)
    model["dirty"] = True


def update_single_promo(master_index: int, slot: int, price, dt, comment):
    model = st.session_state.model
    master = model["master"]
    if slot == 1:
        price_col, date_col, comment_col = "PRECIO B2C PUBLICADO ", "FECHA VENCI", "COMENTARIO"
    else:
        price_col, date_col, comment_col = "PRECIO B2C", "FECHA VENCI.1", "COMENTARIO.1"
    master.at[master_index, price_col] = safe_float(price, np.nan)
    master.at[master_index, date_col] = pd.to_datetime(dt).normalize() if dt else pd.NaT
    master.at[master_index, comment_col] = comment
    refresh_promos_only()


@st.cache_data(show_spinner=False)
def build_download_bytes(master_df: pd.DataFrame, original_bytes: bytes, maestra_name: str):
    buffer_in = io.BytesIO(original_bytes)
    xls = pd.ExcelFile(buffer_in)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            if sheet == maestra_name:
                master_df.drop(columns=[c for c in ["SKU_norm", "DESC_norm", "MLC_slot1", "MLC_slot2"] if c in master_df.columns], errors="ignore").to_excel(writer, sheet_name=sheet, index=False)
            else:
                pd.read_excel(io.BytesIO(original_bytes), sheet_name=sheet, header=None if "relampago" in sheet.lower() else 0).to_excel(writer, sheet_name=sheet, index=False, header=not ("relampago" in sheet.lower()))
    return out.getvalue()


# -----------------------------
# UI
# -----------------------------
st.title("Aurora Pricing App")

with st.sidebar:
    master_up = st.file_uploader("Maestra saneada", type=["xlsx"], key="master")
    purchases_up = st.file_uploader("Compras", type=["xlsx"], key="purchases")
    st.caption("La app usa solo MAESTRA de precios, MLC -SKU, Relampago mi pagina y Compras.")

if not master_up:
    st.info("Sube la maestra saneada para comenzar.")
    st.stop()

try:
    init_state_from_upload(master_up, purchases_up)
except Exception as e:
    st.error(f"No pude leer o modelar la maestra: {e}")
    st.stop()

model = st.session_state.model
product_options = [f"{sku} — {model['sku_desc'].get(sku, '')}" for sku in model["sku_options"]]
selected_label = st.selectbox("Buscar producto", product_options, index=0 if product_options else None)
selected_sku = selected_label.split(" — ")[0] if selected_label else ""
row = get_product_row(model, selected_sku)

tabs = st.tabs(["Cockpit", "Operador de promos", "Relámpago", "Alta de producto", "Descargar"])

with tabs[0]:
    if row is None:
        st.warning("No encontré el SKU.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("SKU", selected_sku)
        c2.metric("Costo", fmt_money(row.get("ÚLTIMO COSTO")))
        c3.metric("Margen local", margin_display(calc_margin_local(row.get("ÚLTIMO COSTO"), row.get("PRECIO BRUTO"))))
        c4.metric("Margen Meli 1", margin_display(calc_margin_meli1(row.get("ÚLTIMO COSTO"), row.get("MONTO EN SIMULACIÓN"))))

        st.subheader("Ficha de decisión")
        a, b = st.columns([1.2, 1])
        with a:
            st.write(f"**Descripción:** {row.get('DESCRIPCIÓN', '')}")
            st.write(f"**Ubicación:** {row.get('UBIC', '—')}")
            st.write(f"**Precio bruto tienda:** {fmt_money(row.get('PRECIO BRUTO'))}")
            st.write(f"**Monto en simulación:** {fmt_money(row.get('MONTO EN SIMULACIÓN'))}")
            st.write(f"**Precio B2C publicado:** {fmt_money(row.get('PRECIO B2C PUBLICADO '))}")
            st.write(f"**Campaña Ads:** {row.get('CAMPAÑA PADS', '—') if pd.notna(row.get('CAMPAÑA PADS', np.nan)) else '—'}")
            mlcs = model["mlc_map"].get(selected_sku, [])
            st.write(f"**MLC asociados:** {', '.join(mlcs) if mlcs else '—'}")
        with b:
            promos_sku = model["promos"][model["promos"]["SKU_norm"] == selected_sku].sort_values(["STATUS_ORDER", "slot"])
            if promos_sku.empty:
                st.info("No tiene promo activa configurada en la maestra.")
            else:
                show = promos_sku[["slot", "MLC", "PRECIO_B2C", "FECHA_VENCI", "COMENTARIO", "STATUS"]].copy()
                show["PRECIO_B2C"] = show["PRECIO_B2C"].map(fmt_money)
                show["FECHA_VENCI"] = show["FECHA_VENCI"].map(fmt_date)
                show.columns = ["Slot", "MLC", "Precio B2C", "Fecha venci", "Comentario", "Estado"]
                st.dataframe(show, use_container_width=True, hide_index=True)

        st.subheader("Relámpago")
        rel_sku = model["rel"][model["rel"]["SKU_norm"] == selected_sku]
        if rel_sku.empty:
            st.write("No está en relámpago.")
        else:
            rel_show = rel_sku.copy()
            rel_show["PRECIO_B2C"] = rel_show["PRECIO_B2C"].map(fmt_money)
            rel_show.columns = ["SKU", "Descripción", "Precio B2C", "Tipo", "Estado"]
            st.dataframe(rel_show, use_container_width=True, hide_index=True)

        st.subheader("Compras")
        ps = model["purchases_summary"]
        purchase_row = ps[ps["SKU_norm"] == selected_sku] if ("SKU_norm" in ps.columns) else pd.DataFrame()
        if purchase_row.empty:
            st.write("No encontré compras para este SKU.")
        else:
            pr = purchase_row.iloc[0]
            x1, x2, x3, x4 = st.columns(4)
            x1.metric("Última compra", fmt_date(pr["ultima_fecha"]))
            x2.metric("Último precio", fmt_money(pr["ultimo_precio"]))
            x3.metric("Proveedor", str(pr["ultimo_proveedor"]))
            x4.metric("Variación", margin_display(pr["variacion_pct"]))
            hist = model["purchase_map"].get(selected_sku, pd.DataFrame()).copy()
            if not hist.empty:
                hist = hist[["Fecha_dt", "Razón Social", "Cantidad", "Precio Un."]].sort_values("Fecha_dt", ascending=False)
                hist["Fecha_dt"] = hist["Fecha_dt"].map(fmt_date)
                hist["Precio Un."] = hist["Precio Un."].map(fmt_money)
                hist.columns = ["Fecha", "Proveedor", "Cantidad", "Precio Unitario"]
                st.dataframe(hist, use_container_width=True, hide_index=True)

with tabs[1]:
    st.subheader("Operador de promos")
    promos = model["promos"].copy()
    if promos.empty:
        st.info("No hay promos configuradas en la maestra.")
    else:
        left, right = st.columns([1, 2])
        with left:
            status_options = [
                "Vencidas",
                "Vencen hoy",
                "Vencen mañana",
                "Vencen pasado mañana",
                "Vencen en 7 días",
                "Vencen en 15 días",
                "Vencen en 1 mes",
            ]
            status_filter = st.multiselect(
                "Estado",
                status_options,
                default=["Vencidas", "Vencen hoy"],
            )
            text_filter = st.text_input("Buscar por SKU / descripción / MLC")
            promos = promos[promos["STATUS"].isin(status_filter)] if status_filter else promos.iloc[0:0]
            if text_filter:
                q = text_filter.lower().strip()
                promos = promos[
                    promos["SKU_norm"].astype(str).str.lower().str.contains(q, na=False) |
                    promos["DESCRIPCIÓN"].astype(str).str.lower().str.contains(q, na=False) |
                    promos["MLC"].astype(str).str.lower().str.contains(q, na=False)
                ]
            mass_date = st.date_input("Cambio masivo de fecha", value=None, format="DD/MM/YYYY")
            if st.button("Aplicar fecha masiva a filtradas"):
                if mass_date and not promos.empty:
                    for _, p in promos.iterrows():
                        update_single_promo(int(p["master_index"]), int(p["slot"]), p["PRECIO_B2C"], mass_date, p["COMENTARIO"])
                    st.success("Fecha actualizada.")
                    st.rerun()

        with right:
            cols = st.columns(4)
            for i, (_, p) in enumerate(promos.sort_values(["STATUS_ORDER", "SKU_norm"]).iterrows()):
                with cols[i % 4]:
                    with st.container(border=True):
                        st.markdown(f"**{p['SKU_norm']}**")
                        st.caption(str(p["DESCRIPCIÓN"])[:55])
                        st.write(f"`{p['MLC'] or '—'}`")
                        st.write(fmt_date(p["FECHA_VENCI"]))
                        if st.button("Abrir", key=f"open_{p['master_index']}_{p['slot']}"):
                            st.session_state.edit_target = (int(p["master_index"]), int(p["slot"]))
                            st.rerun()

        if "edit_target" in st.session_state:
            master_index, slot = st.session_state.edit_target
            current = model["promos"][
                (model["promos"]["master_index"] == master_index) &
                (model["promos"]["slot"] == slot)
            ]
            if not current.empty:
                cp = current.iloc[0]
                @st.dialog("Editar promo")
                def edit_promo_dialog():
                    st.write(f"**SKU:** {cp['SKU_norm']}")
                    st.write(f"**Descripción:** {cp['DESCRIPCIÓN']}")
                    st.write(f"**MLC:** {cp['MLC'] or '—'}")
                    new_date = st.date_input("Fecha venci", value=cp["FECHA_VENCI"].date() if pd.notna(cp["FECHA_VENCI"]) else None, format="DD/MM/YYYY")
                    with st.expander("Campos secundarios"):
                        new_price = st.number_input("Precio B2C", value=float(cp["PRECIO_B2C"]) if pd.notna(cp["PRECIO_B2C"]) else 0.0, step=100.0)
                        new_comment = st.text_area("Comentario", value="" if pd.isna(cp["COMENTARIO"]) else str(cp["COMENTARIO"]))
                    c1, c2 = st.columns(2)
                    if c1.button("Guardar", use_container_width=True):
                        update_single_promo(master_index, slot, new_price, new_date, new_comment)
                        st.session_state.pop("edit_target", None)
                        st.rerun()
                    if c2.button("Cerrar", use_container_width=True):
                        st.session_state.pop("edit_target", None)
                        st.rerun()
                edit_promo_dialog()

with tabs[2]:
    st.subheader("Relámpago mi página")
    rel = model["rel"].copy()
    if rel.empty:
        st.info("No hay registros de relámpago.")
    else:
        rel_show = rel.copy()
        rel_show["PRECIO_B2C"] = rel_show["PRECIO_B2C"].map(fmt_money)
        rel_show.columns = ["SKU", "Descripción", "Precio B2C", "Tipo", "Estado"]
        st.dataframe(rel_show, use_container_width=True, hide_index=True)

with tabs[3]:
    st.subheader("Alta de producto")
    c1, c2, c3 = st.columns(3)
    new_sku = c1.text_input("SKU")
    new_desc = c2.text_input("Descripción")
    new_ubic = c3.text_input("Ubicación")
    d1, d2, d3 = st.columns(3)
    new_cost = d1.number_input("Último costo", min_value=0.0, step=100.0)
    new_bruto_tienda = d2.number_input("Precio bruto en tienda", min_value=0.0, step=100.0)
    new_monto_sim = d3.number_input("Monto en simulación", min_value=0.0, step=100.0)

    st.info(
        f"Margen local proyectado: **{margin_display(calc_margin_local(new_cost, new_bruto_tienda))}**  \n"
        f"Margen Meli 1 proyectado: **{margin_display(calc_margin_meli1(new_cost, new_monto_sim))}**"
    )

    e1, e2, e3 = st.columns(3)
    promo_mlc = e1.text_input("MLC promo base")
    promo_b2c = e2.number_input("Precio B2C publicado", min_value=0.0, step=100.0)
    promo_fecha = e3.date_input("Fecha venci promo base", value=None, format="DD/MM/YYYY")
    promo_comment = st.text_input("Comentario promo base")
    add_rel = st.checkbox("Agregar también a relámpago")

    if st.button("Crear producto"):
        if not new_sku or not new_desc:
            st.error("SKU y descripción son obligatorios.")
        else:
            master = model["master"]
            if (master["SKU_norm"] == norm_sku(new_sku)).any():
                st.error("Ese SKU ya existe.")
            else:
                new_row = {c: np.nan for c in master.columns}
                new_row["SKU"] = new_sku
                new_row["DESCRIPCIÓN"] = new_desc
                new_row["UBIC"] = new_ubic
                new_row["ÚLTIMO COSTO"] = new_cost
                new_row["PRECIO BRUTO"] = new_bruto_tienda
                new_row["MONTO EN SIMULACIÓN"] = new_monto_sim
                new_row["MARGEN LOCAL"] = calc_margin_local(new_cost, new_bruto_tienda)
                new_row["MARGEN MELI 1"] = calc_margin_meli1(new_cost, new_monto_sim)
                new_row["SKU_norm"] = norm_sku(new_sku)
                new_row["DESC_norm"] = new_desc
                new_row["MLC_slot1"] = norm_mlc(promo_mlc)
                new_row["MLC"] = promo_mlc
                new_row["PRECIO B2C PUBLICADO "] = promo_b2c if promo_b2c > 0 else np.nan
                new_row["FECHA VENCI"] = pd.to_datetime(promo_fecha).normalize() if promo_fecha else pd.NaT
                new_row["COMENTARIO"] = promo_comment
                model["master"] = pd.concat([master, pd.DataFrame([new_row])], ignore_index=True)
                model["sku_options"].append(norm_sku(new_sku))
                model["sku_desc"][norm_sku(new_sku)] = new_desc
                if promo_mlc:
                    model["mlc_map"][norm_sku(new_sku)] = [norm_mlc(promo_mlc)]
                if add_rel:
                    rel_new = pd.DataFrame([{
                        "SKU_norm": norm_sku(new_sku),
                        "DESCRIPCION": new_desc,
                        "PRECIO_B2C": promo_b2c if promo_b2c > 0 else np.nan,
                        "TIPO": "LIQUIDACION",
                        "ESTADO": np.nan,
                    }])
                    model["rel"] = pd.concat([model["rel"], rel_new], ignore_index=True)
                refresh_promos_only()
                st.success("Producto creado.")
                st.rerun()

with tabs[4]:
    st.subheader("Descargar maestra actualizada")
    if st.button("Preparar Excel"):
        wb = st.session_state.workbook_meta
        payload = build_download_bytes(model["master"], wb["file_bytes"], wb["maestra_name"])
        st.session_state.download_bytes = payload
    if st.session_state.get("download_bytes"):
        st.download_button(
            "Descargar Excel actualizado",
            data=st.session_state.download_bytes,
            file_name="MAESTRA_PRECIOS_ACTUALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
