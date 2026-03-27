
import io
import re
import hashlib
from datetime import date
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Aurora Pricing Cockpit", page_icon="📈", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem;}
.big-title {font-size: 1.8rem; font-weight: 800; margin-bottom: .2rem;}
.subtle {color:#6b7280; font-size:.95rem;}
.card {
    border:1px solid #e5e7eb; border-radius:16px; padding:14px 14px 10px 14px;
    background:#fff; min-height:132px; display:flex; flex-direction:column; justify-content:space-between;
}
.mini {font-size:.8rem; color:#6b7280;}
.card-title {font-weight:700; font-size:.92rem; line-height:1.2rem; margin-top:.2rem; max-height:2.4rem; overflow:hidden;}
.badge {display:inline-block; padding:.18rem .55rem; border-radius:999px; font-size:.74rem; font-weight:700; border:1px solid transparent;}
.badge-red {background:#fef2f2; color:#b91c1c; border-color:#fecaca;}
.badge-orange {background:#fff7ed; color:#c2410c; border-color:#fed7aa;}
.badge-yellow {background:#fefce8; color:#a16207; border-color:#fde68a;}
.badge-green {background:#f0fdf4; color:#166534; border-color:#bbf7d0;}
.badge-gray {background:#f9fafb; color:#4b5563; border-color:#e5e7eb;}
.section-box {border:1px solid #e5e7eb; border-radius:18px; padding:14px; background:#fff;}
.small-muted {font-size:.82rem; color:#6b7280;}
</style>
""", unsafe_allow_html=True)

MASTER_SHEET = "MAESTRA de precios"
MAP_SHEET = "MLC -SKU"
REL_SHEET_CANDIDATES = ["Relampago mi pagina", "Relámpago mi página"]

PROMO_SLOTS = [
    {"slot_key":"promo_1","label":"Promo 1","mlc_col":"MLC","price_col":"PRECIO B2C PUBLICADO ","date_col":"FECHA VENCI","comment_col":"COMENTARIO"},
    {"slot_key":"promo_2","label":"Promo 2","mlc_col":"MLC.1","price_col":"PRECIO B2C","date_col":"FECHA VENCI.1","comment_col":"COMENTARIO.1"},
]

REL_COLS = ["SKU","Descripción","Precio promocional","Extra","Motivo promoción","COMENTARIO"]


def safe_float(v, default=np.nan):
    try:
        if pd.isna(v):
            return default
        if isinstance(v, str):
            v = v.replace("$","").replace(".","").replace(",",".").strip()
        return float(v)
    except Exception:
        return default


def money(v):
    if pd.isna(v):
        return "—"
    try:
        return f"${int(round(float(v))):,}".replace(",", ".")
    except Exception:
        return "—"


def pct(v):
    if pd.isna(v):
        return "—"
    try:
        return f"{float(v)*100:.1f}%"
    except Exception:
        return "—"


def normalize_sku(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.replace(".0", "")
    digits = re.sub(r"\D", "", s)
    return digits or s.upper()


def normalize_mlc(v):
    if pd.isna(v):
        return None
    s = str(v).strip().upper().replace(" ", "")
    if not s or s == "NAN":
        return None
    m = re.search(r"(MLC\d+)", s)
    return m.group(1) if m else s


def extract_mlcs(v):
    if pd.isna(v):
        return []
    s = str(v).upper()
    mlcs = re.findall(r"MLC\d+", s)
    if mlcs:
        return list(dict.fromkeys(mlcs))
    parts = re.split(r"[,;/\s]+", s)
    parts = [p.strip() for p in parts if p.strip()]
    return [normalize_mlc(p) for p in parts if normalize_mlc(p)]


def first_nonempty(*vals):
    for v in vals:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s and s.lower() != "nan":
            return s
    return ""


def urgency_info(fecha):
    if pd.isna(fecha):
        return ("Sin fecha", "gray", 99)
    days = int((pd.Timestamp(fecha).normalize() - pd.Timestamp.today().normalize()).days)
    if days < 0:
        return ("Vencida", "red", -1)
    if days == 0:
        return ("Vence hoy", "red", 0)
    if days == 1:
        return ("Vence mañana", "orange", 1)
    if days == 2:
        return ("En 2 días", "yellow", 2)
    if days == 3:
        return ("En 3 días", "yellow", 3)
    return ("Vigente", "green", days)


def render_badge(text, cls):
    return f'<span class="badge badge-{cls}">{text}</span>'


def file_signature(uploaded):
    if uploaded is None:
        return None
    raw = uploaded.getvalue()
    return hashlib.md5(raw).hexdigest()


@st.cache_data(show_spinner=False)
def read_excel_sheets(raw_bytes):
    xl = pd.ExcelFile(BytesIO(raw_bytes))
    return {name: pd.read_excel(BytesIO(raw_bytes), sheet_name=name) for name in xl.sheet_names}


def pick_rel_sheet(sheets):
    for name in REL_SHEET_CANDIDATES:
        if name in sheets:
            return name
    return None


def prep_relampago(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=REL_COLS + ["SKU_norm"])
    rel = df.copy()
    if list(rel.columns) != REL_COLS and len(rel.columns) >= 6:
        rel = rel.iloc[:, :6].copy()
        rel.columns = REL_COLS
    for c in REL_COLS:
        if c not in rel.columns:
            rel[c] = np.nan
    rel["SKU_norm"] = rel["SKU"].apply(normalize_sku)
    rel["Precio promocional"] = pd.to_numeric(rel["Precio promocional"], errors="coerce")
    rel["COMENTARIO"] = rel["COMENTARIO"].astype(str).replace("nan","").str.strip()
    rel["Descripción"] = rel["Descripción"].astype(str).replace("nan","").str.strip()
    return rel[[*REL_COLS, "SKU_norm"]].copy()


def normalize_desc_for_match(v):
    s = str(v or "").upper()
    s = re.sub(r"\[UBC:.*?\]", "", s)
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def prepare_compras_dataframe(df):
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame()
    raw = df.copy()
    cols_up = {c: str(c).strip().upper() for c in raw.columns}
    sku_col = next((c for c,u in cols_up.items() if "SKU" in u or "CODIGO" in u), None)
    desc_col = next((c for c,u in cols_up.items() if "DESCRIP" in u or "ARTICULO" in u or "PRODUCTO" in u), None)
    prov_col = next((c for c,u in cols_up.items() if "PROVEEDOR" in u), None)
    price_col = next((c for c,u in cols_up.items() if "PRECIO" in u and ("COSTO" in u or "COMPRA" in u or "NETO" in u)), None)
    qty_col = next((c for c,u in cols_up.items() if "CANT" in u), None)
    date_col = next((c for c,u in cols_up.items() if "FECHA" in u), None)
    if price_col is None:
        numeric = raw.select_dtypes(include=[np.number]).columns.tolist()
        price_col = numeric[0] if numeric else None
    c = pd.DataFrame()
    c["SKU_norm"] = raw[sku_col].apply(normalize_sku) if sku_col else None
    c["Descripción"] = raw[desc_col] if desc_col else ""
    c["desc_norm"] = c["Descripción"].apply(normalize_desc_for_match)
    c["Proveedor"] = raw[prov_col] if prov_col else ""
    c["Precio compra"] = pd.to_numeric(raw[price_col], errors="coerce") if price_col else np.nan
    c["Cantidad"] = pd.to_numeric(raw[qty_col], errors="coerce") if qty_col else np.nan
    c["Fecha"] = pd.to_datetime(raw[date_col], errors="coerce") if date_col else pd.NaT
    c = c[(c["SKU_norm"].notna()) | (c["desc_norm"].str.len() > 5)].copy()
    c = c.sort_values("Fecha")
    summary_rows = []
    for sku, grp in c[c["SKU_norm"].notna()].groupby("SKU_norm"):
        grp = grp.sort_values("Fecha")
        last = grp.iloc[-1]
        prev = grp.iloc[-2] if len(grp) > 1 else None
        var = np.nan
        if prev is not None and pd.notna(prev["Precio compra"]) and prev["Precio compra"] != 0 and pd.notna(last["Precio compra"]):
            var = (last["Precio compra"] - prev["Precio compra"]) / prev["Precio compra"]
        summary_rows.append({
            "SKU_norm": sku,
            "purchase_match_method": "Compras por SKU exacto",
            "last_purchase_date": last["Fecha"],
            "last_purchase_price": last["Precio compra"],
            "last_supplier": first_nonempty(last["Proveedor"]),
            "last_purchase_qty": last["Cantidad"],
            "purchase_change_vs_prev": var,
            "purchase_min": grp["Precio compra"].min(),
            "purchase_max": grp["Precio compra"].max(),
            "supplier_history": " | ".join(pd.unique([str(x) for x in grp["Proveedor"] if str(x).strip() and str(x).lower() != "nan"])),
        })
    return c, pd.DataFrame(summary_rows)


def build_master_promos(master):
    cards = []
    if master.empty:
        return pd.DataFrame(columns=["source_kind","slot_key","source_index","SKU_norm","Descripción","MLC_norm","Precio promocional","FECHA VENCI","COMENTARIO"])
    for idx, row in master.iterrows():
        sku = normalize_sku(row.get("SKU"))
        desc = row.get("DESCRIPCIÓN")
        for slot in PROMO_SLOTS:
            mlcs = extract_mlcs(row.get(slot["mlc_col"])) if slot["mlc_col"] in master.columns else []
            price = pd.to_numeric(pd.Series([row.get(slot["price_col"])]), errors="coerce").iloc[0] if slot["price_col"] in master.columns else np.nan
            fecha = pd.to_datetime(row.get(slot["date_col"]), errors="coerce") if slot["date_col"] in master.columns else pd.NaT
            comment = first_nonempty(row.get(slot["comment_col"]))
            should_create = bool(mlcs) or pd.notna(price) or pd.notna(fecha) or bool(comment)
            if not should_create:
                continue
            for mlc in (mlcs or [None]):
                cards.append({
                    "source_kind": "maestra",
                    "slot_key": slot["slot_key"],
                    "slot_label": slot["label"],
                    "source_index": idx,
                    "SKU_norm": sku,
                    "Descripción": desc,
                    "MLC_norm": mlc,
                    "Precio promocional": price,
                    "FECHA VENCI": fecha,
                    "COMENTARIO": comment,
                })
    promo = pd.DataFrame(cards)
    if promo.empty:
        return pd.DataFrame(columns=["source_kind","slot_key","source_index","SKU_norm","Descripción","MLC_norm","Precio promocional","FECHA VENCI","COMENTARIO","days_to_next","urgency_text","urgency_cls","urgency_rank"])
    promo["days_to_next"] = (pd.to_datetime(promo["FECHA VENCI"]).dt.normalize() - pd.Timestamp.today().normalize()).dt.days
    urg = promo["FECHA VENCI"].apply(urgency_info)
    promo["urgency_text"] = urg.apply(lambda x: x[0])
    promo["urgency_cls"] = urg.apply(lambda x: x[1])
    promo["urgency_rank"] = urg.apply(lambda x: x[2])
    return promo


def aggregate_product(master, bridge, promo_cards, relampago):
    product = master.copy()
    product["SKU_norm"] = product["SKU"].apply(normalize_sku)
    if "Unnamed: 12" in product.columns and "MLC_aux" not in product.columns:
        product = product.rename(columns={"Unnamed: 12":"MLC_aux"})
    bridge2 = pd.DataFrame(columns=["SKU_norm","MLC_norm"])
    if not bridge.empty and len(bridge.columns)>=2:
        bridge2 = pd.DataFrame({
            "SKU_norm": bridge.iloc[:,0].apply(normalize_sku),
            "MLC_norm": bridge.iloc[:,1].apply(normalize_mlc),
        }).dropna().drop_duplicates()
    master_links = []
    for c in ["MLC","MLC.1","MLC_aux"]:
        if c in product.columns:
            tmp = pd.DataFrame({"SKU_norm": product["SKU_norm"], "MLC_norm": product[c].apply(extract_mlcs)})
            tmp = tmp.explode("MLC_norm").dropna()
            master_links.append(tmp)
    sku_mlc = pd.concat(master_links + [bridge2], ignore_index=True) if master_links or not bridge2.empty else pd.DataFrame(columns=["SKU_norm","MLC_norm"])
    sku_mlc = sku_mlc.drop_duplicates()
    links_agg = sku_mlc.groupby("SKU_norm")["MLC_norm"].agg(lambda s: list(pd.unique([x for x in s if pd.notna(x)]))).reset_index() if not sku_mlc.empty else pd.DataFrame(columns=["SKU_norm","MLC_norm"])

    promo_agg = pd.DataFrame(columns=["SKU_norm","total_promos","next_campaign_date","min_promo_price","max_promo_price","promo_comment"])
    if not promo_cards.empty:
        promo_agg = promo_cards.groupby("SKU_norm").agg(
            total_promos=("SKU_norm","count"),
            next_campaign_date=("FECHA VENCI","min"),
            min_promo_price=("Precio promocional","min"),
            max_promo_price=("Precio promocional","max"),
            promo_comment=("COMENTARIO", lambda s: " | ".join([str(x) for x in s if str(x).strip() and str(x).lower()!="nan"][:4])),
        ).reset_index()

    rel_agg = pd.DataFrame(columns=["SKU_norm","relampago_count","relampago_min_price","relampago_max_price","relampago_comment"])
    if not relampago.empty:
        rel_agg = relampago.groupby("SKU_norm").agg(
            relampago_count=("SKU_norm","count"),
            relampago_min_price=("Precio promocional","min"),
            relampago_max_price=("Precio promocional","max"),
            relampago_comment=("COMENTARIO", lambda s: " | ".join([str(x) for x in s if str(x).strip() and str(x).lower()!="nan"][:4])),
        ).reset_index()

    product = product.merge(links_agg, on="SKU_norm", how="left")
    product = product.merge(promo_agg, on="SKU_norm", how="left")
    product = product.merge(rel_agg, on="SKU_norm", how="left")
    product["search_text"] = (
        product["SKU_norm"].fillna("").astype(str) + " " +
        product.get("DESCRIPCIÓN", pd.Series("", index=product.index)).fillna("").astype(str) + " " +
        product["MLC_norm"].apply(lambda x: " ".join(x) if isinstance(x, list) else "")
    ).str.upper()
    return product, bridge2


@st.cache_data(show_spinner=False)
def build_model(master_bytes, compras_bytes=None):
    sheets = read_excel_sheets(master_bytes)
    master = sheets.get(MASTER_SHEET, pd.DataFrame()).copy()
    bridge = sheets.get(MAP_SHEET, pd.DataFrame()).copy()
    rel_sheet = sheets.get(pick_rel_sheet(sheets), pd.DataFrame()).copy() if pick_rel_sheet(sheets) else pd.DataFrame()
    relampago = prep_relampago(rel_sheet)
    promo_cards = build_master_promos(master)
    product, bridge2 = aggregate_product(master, bridge, promo_cards, relampago)
    compras = pd.DataFrame()
    compras_summary = pd.DataFrame()
    if compras_bytes is not None:
        c_sheets = read_excel_sheets(compras_bytes)
        best_name = max(c_sheets, key=lambda k: len(c_sheets[k]))
        compras, compras_summary = prepare_compras_dataframe(c_sheets[best_name])
        if not compras_summary.empty:
            product = product.merge(compras_summary, on="SKU_norm", how="left")
    return {
        "sheets": sheets,
        "master": master,
        "bridge": bridge2,
        "relampago": relampago,
        "promo_cards": promo_cards,
        "product": product,
        "compras": compras,
    }


def display_date(v):
    dt = pd.to_datetime(v, errors="coerce")
    return dt.strftime("%d/%m/%Y") if pd.notna(dt) else "—"


def promo_price_from_row(row, slot):
    return pd.to_numeric(pd.Series([row.get(slot["price_col"])]), errors="coerce").iloc[0]


def update_master_promo(master_df, source_index, slot_key, new_price, new_date, new_comment):
    slot = next(s for s in PROMO_SLOTS if s["slot_key"] == slot_key)
    if slot["price_col"] in master_df.columns:
        master_df.loc[source_index, slot["price_col"]] = new_price if new_price and new_price > 0 else np.nan
    if slot["date_col"] in master_df.columns:
        master_df.loc[source_index, slot["date_col"]] = pd.to_datetime(new_date).normalize() if new_date else pd.NaT
    if slot["comment_col"] in master_df.columns:
        master_df.loc[source_index, slot["comment_col"]] = new_comment
    return master_df


def compras_candidates_for_row(compras_df, row):
    if compras_df.empty:
        return pd.DataFrame(), "Sin archivo de compras"
    sku = row.get("SKU_norm")
    exact = compras_df[compras_df["SKU_norm"] == sku].copy()
    if not exact.empty:
        return exact.sort_values("Fecha"), "Compras por SKU exacto"
    desc = normalize_desc_for_match(row.get("DESCRIPCIÓN"))
    if not desc:
        return pd.DataFrame(), "Sin match"
    tokens = [t for t in desc.split() if len(t) >= 4]
    if not tokens:
        return pd.DataFrame(), "Sin match"
    mask = compras_df["desc_norm"].fillna("").apply(lambda s: sum(t in s for t in tokens) >= max(2, min(3, len(tokens))))
    approx = compras_df[mask].copy()
    if not approx.empty:
        return approx.sort_values("Fecha"), "Compras por descripción aproximada"
    return pd.DataFrame(), "Sin match"


def make_download_workbook(sheets, master_df, rel_df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if name == MASTER_SHEET:
                master_df.to_excel(writer, sheet_name=name, index=False)
            elif name in REL_SHEET_CANDIDATES:
                base = rel_df[REL_COLS].copy() if not rel_df.empty else pd.DataFrame(columns=REL_COLS)
                base.to_excel(writer, sheet_name=name, index=False, header=False)
            else:
                df.to_excel(writer, sheet_name=name, index=False)
    return out.getvalue()


@st.dialog("Editar promo", width="large")
def promo_dialog(card):
    st.write(f"**SKU:** {card.get('SKU_norm')}")
    st.write(f"**Descripción:** {card.get('Descripción')}")
    if card.get("MLC_norm"):
        st.write(f"**MLC:** {card.get('MLC_norm')}")
    master_df = st.session_state.master_df.copy()
    source_index = int(card["source_index"])
    slot = next(s for s in PROMO_SLOTS if s["slot_key"] == card["slot_key"])
    row = master_df.loc[source_index]
    current_date = pd.to_datetime(row.get(slot["date_col"]), errors="coerce")
    current_price = promo_price_from_row(row, slot)
    current_comment = first_nonempty(row.get(slot["comment_col"]))

    with st.form(f"promo_form_{source_index}_{slot['slot_key']}"):
        new_date = st.date_input("Fecha de vencimiento", value=current_date.date() if pd.notna(current_date) else date.today(), format="DD/MM/YYYY")
        with st.expander("Precio B2C publicado y comentario", expanded=False):
            c1, c2 = st.columns(2)
            new_price = c1.number_input("Precio B2C publicado", min_value=0.0, value=float(safe_float(current_price,0.0)), step=1.0)
            new_comment = c2.text_input("Comentario", value=str(current_comment))
        save = st.form_submit_button("Guardar", type="primary", use_container_width=True)
    if save:
        st.session_state.master_df = update_master_promo(master_df, source_index, card["slot_key"], new_price, new_date, new_comment)
        st.session_state.needs_rebuild = True
        st.rerun()


st.sidebar.title("Aurora Pricing Cockpit")
master_file = st.sidebar.file_uploader("Maestra integrada", type=["xlsx"], key="master")
compras_file = st.sidebar.file_uploader("Compras históricas", type=["xlsx"], key="compras")
st.sidebar.caption("Solo se usa MAESTRA de precios, MLC -SKU, Relámpago mi página y Compras. No se usa CONTROL DE PROMOCIONES.")
if master_file is None:
    st.info("Carga la maestra integrada para empezar.")
    st.stop()

master_sig = file_signature(master_file)
compras_sig = file_signature(compras_file) if compras_file else None
if st.session_state.get("loaded_master_sig") != master_sig or st.session_state.get("loaded_compras_sig") != compras_sig:
    model = build_model(master_file.getvalue(), compras_file.getvalue() if compras_file else None)
    st.session_state.loaded_master_sig = master_sig
    st.session_state.loaded_compras_sig = compras_sig
    st.session_state.sheets = model["sheets"]
    st.session_state.master_df = model["master"]
    st.session_state.bridge_df = model["bridge"]
    st.session_state.relampago_df = model["relampago"]
    st.session_state.compras_df = model["compras"]
    st.session_state.needs_rebuild = True

if st.session_state.get("needs_rebuild", True):
    tmp_sheets = dict(st.session_state.sheets)
    tmp_sheets[MASTER_SHEET] = st.session_state.master_df.copy()
    rel_name = pick_rel_sheet(tmp_sheets) or REL_SHEET_CANDIDATES[0]
    tmp_sheets[rel_name] = st.session_state.relampago_df[REL_COLS].copy() if not st.session_state.relampago_df.empty else pd.DataFrame(columns=REL_COLS)
    tmp_bytes = make_download_workbook(tmp_sheets, st.session_state.master_df, st.session_state.relampago_df)
    rebuilt = build_model(tmp_bytes, compras_file.getvalue() if compras_file else None)
    st.session_state.model = rebuilt
    st.session_state.download_bytes = tmp_bytes
    st.session_state.needs_rebuild = False

model = st.session_state.model
product_df = model["product"].copy()
promo_cards = model["promo_cards"].copy()
relampago_df = st.session_state.relampago_df.copy()
compras_df = st.session_state.compras_df.copy()

page = st.sidebar.radio("Módulos", ["Cockpit por producto", "Operador de promos", "Alta de producto"])
st.sidebar.download_button(
    "Descargar Excel actualizado",
    data=st.session_state.download_bytes,
    file_name="MAESTRA_PRECIOS_Y_PROMOS_actualizada.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

if page == "Cockpit por producto":
    st.markdown('<div class="big-title">Cockpit por producto</div>', unsafe_allow_html=True)
    q = st.text_input("Buscar por SKU, descripción o MLC")
    filt = product_df.copy()
    if q.strip():
        filt = filt[filt["search_text"].str.contains(q.upper(), na=False)]
    options = {f"{r['SKU_norm']} · {str(r.get('DESCRIPCIÓN',''))[:90]}": r["SKU_norm"] for _, r in filt.head(300).iterrows()}
    selected = st.selectbox("Producto", options=list(options.keys())) if options else None
    if not selected:
        st.info("No encontré productos con ese criterio.")
        st.stop()
    sku = options[selected]
    row = product_df[product_df["SKU_norm"] == sku].iloc[0]
    sku_promos = promo_cards[promo_cards["SKU_norm"] == sku].copy()
    sku_rel = relampago_df[relampago_df["SKU_norm"] == sku].copy()
    compras_match, compras_method = compras_candidates_for_row(compras_df, row)

    b1, b2 = st.columns(2)
    with b1:
        promo_badge = urgency_info(sku_promos["FECHA VENCI"].min())[0:2] if not sku_promos.empty else ("Sin promo", "gray")
        st.markdown(render_badge(*promo_badge), unsafe_allow_html=True)
    with b2:
        rel_badge = ("En relámpago","green") if not sku_rel.empty else ("Sin relámpago","gray")
        st.markdown(render_badge(*rel_badge), unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1.2, 1, 1])
    with c1:
        st.markdown("### Identidad")
        st.write(f"**SKU:** {sku}")
        st.write(f"**Descripción:** {row.get('DESCRIPCIÓN','—')}")
        st.write(f"**Ubicación:** {first_nonempty(row.get('UBIC'), row.get('T-L'),'—')}")
        mlcs = row.get("MLC_norm") if isinstance(row.get("MLC_norm"), list) else []
        st.write(f"**MLCs asociados:** {', '.join(mlcs) if mlcs else '—'}")
        st.write(f"**Comentario maestra:** {first_nonempty(row.get('COMENTARIO'),'—')}")
        st.markdown("### Compras")
        st.markdown(render_badge(compras_method, "green" if "exacto" in compras_method.lower() else "yellow" if "aproximada" in compras_method.lower() else "gray"), unsafe_allow_html=True)
        if compras_match.empty:
            st.info("No encontré compras para este producto.")
        else:
            grp = compras_match.sort_values("Fecha")
            last = grp.iloc[-1]
            prev = grp.iloc[-2] if len(grp)>1 else None
            var = np.nan
            if prev is not None and pd.notna(prev["Precio compra"]) and prev["Precio compra"] != 0 and pd.notna(last["Precio compra"]):
                var = (last["Precio compra"] - prev["Precio compra"]) / prev["Precio compra"]
            st.write(f"**Última compra:** {display_date(last['Fecha'])}")
            st.write(f"**Último precio compra:** {money(last['Precio compra'])}")
            st.write(f"**Proveedor último:** {first_nonempty(last['Proveedor'],'—')}")
            st.write(f"**Cantidad última compra:** {money(last['Cantidad']) if pd.notna(last['Cantidad']) else '—'}")
            st.write(f"**Variación vs compra anterior:** {pct(var)}")
            st.write(f"**Rango histórico:** {money(grp['Precio compra'].min())} a {money(grp['Precio compra'].max())}")
            provs = " | ".join(pd.unique([str(x) for x in grp["Proveedor"] if str(x).strip() and str(x).lower()!="nan"]))
            st.write(f"**Proveedores históricos:** {provs or '—'}")
    with c2:
        st.markdown("### Precio y rentabilidad")
        m1, m2 = st.columns(2)
        m1.metric("Precio neto", money(row.get("PRECIO NETO")))
        m2.metric("Cambio precio", money(row.get("CAMBIO DE PRECIO")))
        m3, m4 = st.columns(2)
        m3.metric("Margen local", pct(row.get("MARGEN LOCAL")))
        m4.metric("Monto en simulación", money(row.get("MONTO EN SIMULACIÓN")))
        st.write(f"**Neto Meli 1:** {money(row.get(' NETO MELI 1'))}")
        st.write(f"**Margen Meli 1:** {pct(row.get('MARGEN MELI 1'))}")
        st.write(f"**Precio promo mínimo:** {money(sku_promos['Precio promocional'].min()) if not sku_promos.empty else '—'}")
        st.write(f"**Precio promo máximo:** {money(sku_promos['Precio promocional'].max()) if not sku_promos.empty else '—'}")
        st.write(f"**Precio relámpago mínimo:** {money(sku_rel['Precio promocional'].min()) if not sku_rel.empty else '—'}")
        st.write(f"**Precio relámpago máximo:** {money(sku_rel['Precio promocional'].max()) if not sku_rel.empty else '—'}")
        st.write(f"**Comentario promo:** {first_nonempty(sku_promos['COMENTARIO'].iloc[0] if not sku_promos.empty else None,'—')}")
        st.write(f"**Comentario relámpago:** {first_nonempty(sku_rel['COMENTARIO'].iloc[0] if not sku_rel.empty else None,'—')}")
    with c3:
        st.markdown("### Promos maestra")
        st.metric("Filas promo", int(len(sku_promos)))
        if sku_promos.empty:
            st.info("No hay promos en maestra.")
        else:
            show = sku_promos[["slot_label","MLC_norm","Precio promocional","FECHA VENCI","COMENTARIO"]].copy()
            show["FECHA VENCI"] = show["FECHA VENCI"].apply(display_date)
            st.dataframe(show, use_container_width=True, hide_index=True, height=240)
        st.markdown("### Relámpago mi página")
        st.metric("Filas relámpago", int(len(sku_rel)))
        if sku_rel.empty:
            st.info("No está en relámpago mi página.")
        else:
            show2 = sku_rel[["SKU","Descripción","Precio promocional","Motivo promoción","COMENTARIO"]].copy()
            st.dataframe(show2, use_container_width=True, hide_index=True, height=210)

    st.markdown("### Lectura automática")
    notes = []
    if not sku_promos.empty:
        t,_,_ = urgency_info(sku_promos["FECHA VENCI"].min())
        notes.append(f"Producto con promo en maestra. Estado principal: {t.lower()}.")
    if not sku_rel.empty:
        notes.append("Producto presente en relámpago mi página.")
    if compras_match.empty:
        notes.append("No hay compras detectadas para este SKU.")
    if pd.notna(row.get("MARGEN LOCAL")) and float(row.get("MARGEN LOCAL")) < 0:
        notes.append("Margen local negativo.")
    if not notes:
        notes.append("Producto bajo control. No veo alertas críticas inmediatas.")
    for n in notes:
        st.write(f"- {n}")

elif page == "Operador de promos":
    st.markdown('<div class="big-title">Operador de promos</div>', unsafe_allow_html=True)
    st.caption("Solo maestra de precios y relámpago mi página. Sin CONTROL DE PROMOCIONES.")
    tab1, tab2 = st.tabs(["Bandeja de promos maestra", "Relámpago mi página"])

    with tab1:
        cards = promo_cards.copy()
        h1,h2,h3,h4 = st.columns(4)
        h1.metric("Vencen hoy", int((cards["urgency_text"]=="Vence hoy").sum()) if not cards.empty else 0)
        h2.metric("Vencen mañana", int((cards["urgency_text"]=="Vence mañana").sum()) if not cards.empty else 0)
        h3.metric("Vencen en 3 días", int((cards["urgency_text"]=="En 3 días").sum()) if not cards.empty else 0)
        h4.metric("Sin fecha", int((cards["urgency_text"]=="Sin fecha").sum()) if not cards.empty else 0)

        f1,f2,f3 = st.columns([1,1,2])
        urgency_filter = f1.selectbox("Prioridad", ["Todos","Vencidas","Hoy","Mañana","En 2 días","En 3 días","Sin fecha","Vigentes"])
        without_price = f2.checkbox("Sin precio B2C")
        search = f3.text_input("Buscar SKU / MLC / descripción")

        work = cards.copy()
        if urgency_filter == "Vencidas":
            work = work[work["urgency_text"]=="Vencida"]
        elif urgency_filter == "Hoy":
            work = work[work["urgency_text"]=="Vence hoy"]
        elif urgency_filter == "Mañana":
            work = work[work["urgency_text"]=="Vence mañana"]
        elif urgency_filter == "En 2 días":
            work = work[work["urgency_text"]=="En 2 días"]
        elif urgency_filter == "En 3 días":
            work = work[work["urgency_text"]=="En 3 días"]
        elif urgency_filter == "Sin fecha":
            work = work[work["urgency_text"]=="Sin fecha"]
        elif urgency_filter == "Vigentes":
            work = work[work["urgency_cls"]=="green"]
        if without_price:
            work = work[work["Precio promocional"].isna()]
        if search.strip():
            ss = search.upper()
            mask = work["SKU_norm"].fillna("").astype(str).str.contains(ss, na=False) | work["MLC_norm"].fillna("").astype(str).str.contains(ss, na=False) | work["Descripción"].fillna("").astype(str).str.upper().str.contains(ss, na=False)
            work = work[mask]
        work = work.sort_values(["urgency_rank","FECHA VENCI","SKU_norm"], ascending=[True,True,True]).reset_index(drop=True)

        st.markdown("### Bandeja visual")
        if work.empty:
            st.info("No hay promos con esos filtros.")
        else:
            cols = st.columns(4)
            for i, (_, card) in enumerate(work.iterrows()):
                col = cols[i % 4]
                with col:
                    st.markdown(
                        f"""
                        <div class="card">
                            <div>
                                {render_badge(card['urgency_text'], card['urgency_cls'])}
                                <div class="mini" style="margin-top:.5rem;">{card['SKU_norm']}</div>
                                <div class="card-title">{first_nonempty(card.get('Descripción'),'—')}</div>
                            </div>
                            <div class="small-muted">MLC: {first_nonempty(card.get('MLC_norm'),'—')}<br>Fecha: {display_date(card.get('FECHA VENCI'))}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button("Editar", key=f"edit_card_{card['source_index']}_{card['slot_key']}", use_container_width=True):
                        promo_dialog(card)
        st.markdown("---")
        st.markdown("#### Cambio masivo de fecha")
        if work.empty:
            st.info("No hay filas filtradas para actualizar.")
        else:
            with st.form("bulk_date_form"):
                bulk_date = st.date_input("Nueva fecha", value=date.today(), format="DD/MM/YYYY")
                submitted = st.form_submit_button("Aplicar a filas filtradas", type="primary")
            if submitted:
                base = st.session_state.master_df.copy()
                for _, card in work.iterrows():
                    base = update_master_promo(base, int(card["source_index"]), card["slot_key"], card["Precio promocional"], bulk_date, card.get("COMENTARIO",""))
                st.session_state.master_df = base
                st.session_state.needs_rebuild = True
                st.success(f"Fecha actualizada en {len(work)} promos.")
                st.rerun()

    with tab2:
        st.caption("Lista para agregar, sacar o modificar ofertas relámpago.")
        rel_edit = st.data_editor(
            relampago_df[REL_COLS].copy() if not relampago_df.empty else pd.DataFrame(columns=REL_COLS),
            use_container_width=True, hide_index=True, num_rows="dynamic", height=520, key="rel_editor")
        if st.button("Guardar lista relámpago", type="primary", use_container_width=True):
            st.session_state.relampago_df = prep_relampago(rel_edit)
            st.session_state.needs_rebuild = True
            st.success("Lista relámpago actualizada.")
            st.rerun()

else:
    st.markdown('<div class="big-title">Alta de producto</div>', unsafe_allow_html=True)
    st.caption("Crea un SKU nuevo directamente sobre la maestra actual.")
    with st.form("alta_producto"):
        c1,c2,c3 = st.columns(3)
        new_sku = c1.text_input("SKU")
        desc = c2.text_input("Descripción")
        ubic = c3.text_input("Ubicación")
        costo = c1.number_input("Último costo", min_value=0.0, value=0.0, step=1.0)
        precio_tienda = c2.number_input("Precio bruto en tienda", min_value=0.0, value=0.0, step=1.0)
        monto_sim = c3.number_input("MONTO EN SIMULACIÓN", min_value=0.0, value=0.0, step=1.0)
        st.markdown("### Promo base")
        p1,p2,p3 = st.columns(3)
        mlc = p1.text_input("MLC")
        precio_b2c = p2.number_input("Precio B2C publicado", min_value=0.0, value=0.0, step=1.0)
        fecha = p3.date_input("Fecha vencimiento", value=None, format="DD/MM/YYYY")
        comentario = st.text_input("Comentario")
        add_rel = st.checkbox("Agregar a relámpago")

        precio_neto = precio_tienda/1.19 if precio_tienda else np.nan
        margen_local = ((precio_neto-costo)/precio_neto) if pd.notna(precio_neto) and precio_neto else np.nan
        neto_meli = monto_sim/1.19 if monto_sim else np.nan
        margen_meli = ((neto_meli-costo)/neto_meli) if pd.notna(neto_meli) and neto_meli else np.nan
        r1,r2 = st.columns(2)
        r1.metric("Margen local proyectado", pct(margen_local))
        r2.metric("Margen Meli 1 proyectado", pct(margen_meli))
        submit = st.form_submit_button("Crear producto", type="primary", use_container_width=True)

    if submit:
        sku_norm = normalize_sku(new_sku)
        if not sku_norm:
            st.error("SKU inválido.")
        elif sku_norm in set(st.session_state.master_df["SKU"].apply(normalize_sku).dropna()):
            st.error("Ese SKU ya existe.")
        else:
            master = st.session_state.master_df.copy()
            new_row = {c: np.nan for c in master.columns}
            assignments = {
                "SKU": new_sku,
                "DESCRIPCIÓN": desc,
                "UBIC": ubic,
                "ÚLTIMO COSTO": costo,
                "PRECIO BRUTO": precio_tienda,
                "PRECIO NETO": precio_neto,
                "MARGEN LOCAL": margen_local,
                "MONTO EN SIMULACIÓN": monto_sim,
                " NETO MELI 1": neto_meli,
                "MARGEN MELI 1": margen_meli,
                "MLC": mlc,
                "PRECIO B2C PUBLICADO ": precio_b2c if precio_b2c > 0 else np.nan,
                "FECHA VENCI": pd.to_datetime(fecha).normalize() if fecha else pd.NaT,
                "COMENTARIO": comentario,
            }
            for k,v in assignments.items():
                if k in new_row:
                    new_row[k] = v
            st.session_state.master_df = pd.concat([master, pd.DataFrame([new_row])], ignore_index=True)

            if mlc:
                bridge = st.session_state.bridge_df.copy()
                bridge = pd.concat([bridge, pd.DataFrame([{"SKU_norm": sku_norm, "MLC_norm": normalize_mlc(mlc)}])], ignore_index=True).drop_duplicates()
                st.session_state.bridge_df = bridge

            if add_rel:
                rel = st.session_state.relampago_df.copy()
                rel_new = {
                    "SKU": new_sku,
                    "Descripción": desc,
                    "Precio promocional": precio_b2c if precio_b2c > 0 else np.nan,
                    "Extra": np.nan,
                    "Motivo promoción": "LIQUIDACION",
                    "COMENTARIO": comentario,
                }
                st.session_state.relampago_df = pd.concat([rel, pd.DataFrame([rel_new])], ignore_index=True)

            st.session_state.needs_rebuild = True
            st.success("Producto creado en memoria. Descarga el Excel actualizado.")
            st.rerun()
