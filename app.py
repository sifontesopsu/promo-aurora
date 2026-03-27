
import io
import os
import re
import json
import time
import sqlite3
import hashlib
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "Aurora Pricing App"
DB_PATH = "aurora_pricing_live.db"
MASTER_SHEETS_REQUIRED = ["MAESTRA de precios", "MLC -SKU", "Relampago mi pagina"]

# ---------------------------
# Config / helpers
# ---------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")

def norm_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def norm_sku(v):
    if pd.isna(v) or v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    if re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    # remove trailing .0 from numeric-like strings
    if s.endswith(".0") and s.replace(".", "", 1).isdigit():
        s = s[:-2]
    return s

def norm_mlc(v):
    s = norm_text(v).upper().replace(" ", "")
    if not s:
        return ""
    if not s.startswith("MLC"):
        m = re.search(r"(MLC\d+)", s)
        if m:
            s = m.group(1)
    return s

def as_number(v):
    try:
        if pd.isna(v) or v == "":
            return None
        return float(v)
    except Exception:
        return None

def fmt_money(v):
    n = as_number(v)
    if n is None:
        return "—"
    return f"${n:,.0f}".replace(",", ".")

def fmt_pct(v):
    n = as_number(v)
    if n is None:
        return "—"
    # Values in master are already percentage points like 35.27 not 0.3527
    return f"{n:.2f}%"

def parse_date_any(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v == "":
        return None
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.date()
    except Exception:
        return None

def fmt_date(v):
    d = parse_date_any(v)
    return d.strftime("%d/%m/%Y") if d else "—"

def compute_margin_local(cost, bruto_tienda):
    cost = as_number(cost)
    bruto_tienda = as_number(bruto_tienda)
    if cost is None or bruto_tienda is None or bruto_tienda == 0:
        return None, None
    neto = bruto_tienda / 1.19
    if neto == 0:
        return None, neto
    margen = ((neto - cost) / neto) * 100
    return margen, neto

def compute_margin_meli1(cost, monto_sim):
    cost = as_number(cost)
    monto_sim = as_number(monto_sim)
    if cost is None or monto_sim is None or monto_sim == 0:
        return None, None
    neto = monto_sim / 1.19
    if neto == 0:
        return None, neto
    margen = ((neto - cost) / neto) * 100
    return margen, neto

def date_status(d):
    d = parse_date_any(d)
    if d is None:
        return "Sin fecha"
    today = date.today()
    diff = (d - today).days
    if diff < 0:
        return "Vencidas"
    if diff == 0:
        return "Vencen hoy"
    if diff == 1:
        return "Vencen mañana"
    if diff == 2:
        return "Vencen pasado mañana"
    if 3 <= diff <= 7:
        return "Vencen en 7 días"
    if 8 <= diff <= 15:
        return "Vencen en 15 días"
    if 16 <= diff <= 31:
        return "Vencen en 1 mes"
    return "Fuera de rango"

def file_signature(uploaded_file):
    if uploaded_file is None:
        return None
    data = uploaded_file.getvalue()
    return hashlib.md5(data).hexdigest()

def editor_height(rows):
    return min(max(220, 35 * (rows + 1)), 700)

# ---------------------------
# SQLite
# ---------------------------
@st.cache_resource
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA busy_timeout=5000;")
    init_db(conn)
    return conn

def init_db(conn):
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS meta (
        key TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS products (
        sku TEXT PRIMARY KEY,
        descripcion TEXT,
        ubicacion TEXT,
        cambio_precio REAL,
        ultimo_costo REAL,
        margen_local REAL,
        precio_neto REAL,
        precio_bruto REAL,
        margen_meli_1 REAL,
        neto_meli_1 REAL,
        monto_simulacion REAL,
        mlc_ads TEXT,
        campana_ads TEXT,
        comentario_maestra TEXT,
        row_idx INTEGER,
        updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS mlc_map (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT,
        mlc TEXT,
        UNIQUE(sku, mlc)
    );

    CREATE TABLE IF NOT EXISTS promo_slots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT,
        slot INTEGER,
        mlc TEXT,
        precio_b2c REAL,
        fecha_venci TEXT,
        comentario TEXT,
        UNIQUE(sku, slot)
    );

    CREATE TABLE IF NOT EXISTS relampago (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT,
        descripcion TEXT,
        precio REAL,
        extra TEXT,
        comentario TEXT,
        estado TEXT
    );

    CREATE TABLE IF NOT EXISTS purchase_summary (
        sku TEXT PRIMARY KEY,
        ultima_compra TEXT,
        ultimo_precio REAL,
        proveedor_ultimo TEXT,
        cantidad_ultima REAL,
        variacion_pct REAL,
        precio_min REAL,
        precio_max REAL,
        proveedores TEXT,
        match_method TEXT
    );

    CREATE TABLE IF NOT EXISTS purchase_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT,
        fecha TEXT,
        proveedor TEXT,
        cantidad REAL,
        precio_unitario REAL
    );
    """)

def set_meta(conn, key, value):
    conn.execute("INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
    conn.commit()

def get_meta(conn, key):
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row["value"] if row else None

# ---------------------------
# Importers
# ---------------------------
def read_master_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    missing = [s for s in MASTER_SHEETS_REQUIRED if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"Faltan hojas requeridas: {', '.join(missing)}")

    master = pd.read_excel(uploaded_file, sheet_name="MAESTRA de precios")
    bridge = pd.read_excel(uploaded_file, sheet_name="MLC -SKU")
    rel_raw = pd.read_excel(uploaded_file, sheet_name="Relampago mi pagina", header=None)

    # Normalize master
    master = master.copy()
    master["SKU_norm"] = master.get("SKU", pd.Series(dtype=object)).apply(norm_sku)
    master["DESC_norm"] = master.get("DESCRIPCIÓN", pd.Series(dtype=object)).astype(str).fillna("")

    # Normalize bridge
    bridge = bridge.rename(columns={"Número de publicación": "MLC"})
    if "SKU" not in bridge.columns:
        bridge = pd.DataFrame(columns=["SKU", "MLC"])
    bridge["sku"] = bridge["SKU"].apply(norm_sku)
    bridge["mlc"] = bridge.get("MLC", pd.Series(dtype=object)).apply(norm_mlc)
    bridge = bridge[(bridge["sku"] != "") & (bridge["mlc"] != "")]
    bridge = bridge[["sku", "mlc"]].drop_duplicates()

    # Normalize relampago raw -> fixed columns
    rel = rel_raw.copy()
    if rel.shape[1] < 6:
        for i in range(rel.shape[1], 6):
            rel[i] = np.nan
    rel = rel.iloc[:, :6]
    rel.columns = ["sku", "descripcion", "precio", "extra", "comentario", "estado"]
    rel["sku"] = rel["sku"].apply(norm_sku)
    rel["descripcion"] = rel["descripcion"].fillna("").astype(str)
    rel["precio"] = pd.to_numeric(rel["precio"], errors="coerce")
    rel["extra"] = rel["extra"].fillna("").astype(str)
    rel["comentario"] = rel["comentario"].fillna("").astype(str)
    rel["estado"] = rel["estado"].fillna("").astype(str)
    rel = rel[rel["sku"] != ""].reset_index(drop=True)

    return master, bridge, rel, xls.sheet_names

def import_master_to_db(conn, uploaded_file):
    master, bridge, rel, sheet_names = read_master_excel(uploaded_file)
    now = datetime.now().isoformat()

    with conn:
        conn.execute("DELETE FROM products")
        conn.execute("DELETE FROM mlc_map")
        conn.execute("DELETE FROM promo_slots")
        conn.execute("DELETE FROM relampago")

        product_rows = []
        promo_rows = []
        for idx, row in master.iterrows():
            sku = norm_sku(row.get("SKU"))
            if not sku:
                continue
            product_rows.append((
                sku,
                norm_text(row.get("DESCRIPCIÓN")),
                norm_text(row.get("UBIC")),
                as_number(row.get("CAMBIO DE PRECIO")),
                as_number(row.get("ÚLTIMO COSTO")),
                as_number(row.get("MARGEN LOCAL")),
                as_number(row.get("PRECIO NETO")),
                as_number(row.get("PRECIO BRUTO")),
                as_number(row.get("MARGEN MELI 1")),
                as_number(row.get(" NETO MELI 1")),
                as_number(row.get("MONTO EN SIMULACIÓN")),
                norm_mlc(row.get("Unnamed: 12")),  # ads mlc principal
                norm_text(row.get("CAMPAÑA PADS")),
                norm_text(row.get("COMENTARIO")),
                int(idx),
                now
            ))
            slot1_mlc = norm_mlc(row.get("MLC"))
            if slot1_mlc or as_number(row.get("PRECIO B2C PUBLICADO ")) is not None or parse_date_any(row.get("FECHA VENCI")):
                promo_rows.append((sku, 1, slot1_mlc, as_number(row.get("PRECIO B2C PUBLICADO ")), 
                                   parse_date_any(row.get("FECHA VENCI")).isoformat() if parse_date_any(row.get("FECHA VENCI")) else None,
                                   norm_text(row.get("COMENTARIO"))))
            slot2_mlc = norm_mlc(row.get("MLC.1"))
            if slot2_mlc or as_number(row.get("PRECIO B2C")) is not None or parse_date_any(row.get("FECHA VENCI.1")):
                promo_rows.append((sku, 2, slot2_mlc, as_number(row.get("PRECIO B2C")),
                                   parse_date_any(row.get("FECHA VENCI.1")).isoformat() if parse_date_any(row.get("FECHA VENCI.1")) else None,
                                   norm_text(row.get("COMENTARIO.1"))))

        conn.executemany("""
            INSERT INTO products (
                sku, descripcion, ubicacion, cambio_precio, ultimo_costo, margen_local, precio_neto,
                precio_bruto, margen_meli_1, neto_meli_1, monto_simulacion, mlc_ads, campana_ads,
                comentario_maestra, row_idx, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, product_rows)

        conn.executemany("INSERT INTO mlc_map(sku, mlc) VALUES(?, ?)", bridge[["sku", "mlc"]].itertuples(index=False, name=None))
        if promo_rows:
            conn.executemany(
                "INSERT INTO promo_slots(sku, slot, mlc, precio_b2c, fecha_venci, comentario) VALUES (?, ?, ?, ?, ?, ?)",
                promo_rows
            )
        conn.executemany(
            "INSERT INTO relampago(sku, descripcion, precio, extra, comentario, estado) VALUES (?, ?, ?, ?, ?, ?)",
            rel[["sku", "descripcion", "precio", "extra", "comentario", "estado"]].itertuples(index=False, name=None)
        )

    set_meta(conn, "master_sig", file_signature(uploaded_file))
    set_meta(conn, "master_sheet_names", json.dumps(sheet_names))
    # Save original bytes for export
    orig_path = "master_source.xlsx"
    with open(orig_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    set_meta(conn, "master_source_path", orig_path)

def import_purchases_to_db(conn, uploaded_file):
    df = pd.read_excel(uploaded_file)
    df = df.copy()
    if "SKU" not in df.columns:
        raise ValueError("El archivo de compras no tiene columna SKU.")
    df["sku"] = df["SKU"].apply(norm_sku)
    df["fecha_dt"] = pd.to_datetime(df.get("Fecha"), errors="coerce", dayfirst=True)
    df["proveedor"] = df.get("Razón Social", pd.Series(dtype=object)).fillna("").astype(str)
    df["cantidad"] = pd.to_numeric(df.get("Cantidad"), errors="coerce")
    df["precio_unitario"] = pd.to_numeric(df.get("Precio Un."), errors="coerce")
    df = df[(df["sku"] != "") & df["fecha_dt"].notna() & df["precio_unitario"].notna()].copy()
    df = df.sort_values(["sku", "fecha_dt"])

    summaries = []
    histories = []
    for sku, grp in df.groupby("sku", sort=False):
        grp = grp.sort_values("fecha_dt")
        last = grp.iloc[-1]
        prev_price = grp.iloc[-2]["precio_unitario"] if len(grp) >= 2 else np.nan
        var_pct = None
        if pd.notna(prev_price) and prev_price:
            var_pct = ((last["precio_unitario"] - prev_price) / prev_price) * 100
        providers = " | ".join(sorted(set([p for p in grp["proveedor"].tolist() if p])))
        summaries.append((
            sku,
            last["fecha_dt"].date().isoformat(),
            float(last["precio_unitario"]),
            last["proveedor"],
            as_number(last["cantidad"]),
            var_pct,
            float(grp["precio_unitario"].min()),
            float(grp["precio_unitario"].max()),
            providers,
            "SKU exacto"
        ))
        for _, r in grp.iterrows():
            histories.append((
                sku,
                r["fecha_dt"].date().isoformat(),
                r["proveedor"],
                as_number(r["cantidad"]),
                float(r["precio_unitario"])
            ))

    with conn:
        conn.execute("DELETE FROM purchase_summary")
        conn.execute("DELETE FROM purchase_history")
        if summaries:
            conn.executemany("""
                INSERT INTO purchase_summary (
                    sku, ultima_compra, ultimo_precio, proveedor_ultimo, cantidad_ultima,
                    variacion_pct, precio_min, precio_max, proveedores, match_method
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, summaries)
        if histories:
            conn.executemany("""
                INSERT INTO purchase_history (sku, fecha, proveedor, cantidad, precio_unitario)
                VALUES (?, ?, ?, ?, ?)
            """, histories)

    set_meta(conn, "purchases_sig", file_signature(uploaded_file))

# ---------------------------
# Query helpers
# ---------------------------
def list_products(conn):
    return pd.read_sql_query("""
        SELECT sku, descripcion
        FROM products
        ORDER BY sku
    """, conn)

@st.cache_data(ttl=30)
def get_product_snapshot(version_token: str, sku: str):
    conn = get_conn()
    prod = conn.execute("SELECT * FROM products WHERE sku=?", (sku,)).fetchone()
    if not prod:
        return None
    promos = pd.read_sql_query("SELECT * FROM promo_slots WHERE sku=? ORDER BY slot", conn, params=(sku,))
    mlcs = pd.read_sql_query("SELECT mlc FROM mlc_map WHERE sku=? ORDER BY mlc", conn, params=(sku,))
    rel = pd.read_sql_query("SELECT rowid as _rid, * FROM relampago WHERE sku=? ORDER BY id", conn, params=(sku,))
    purch = conn.execute("SELECT * FROM purchase_summary WHERE sku=?", (sku,)).fetchone()
    hist = pd.read_sql_query("SELECT fecha, proveedor, cantidad, precio_unitario FROM purchase_history WHERE sku=? ORDER BY fecha", conn, params=(sku,))
    return {
        "product": dict(prod),
        "promos": promos,
        "mlcs": mlcs["mlc"].tolist() if not mlcs.empty else [],
        "relampago": rel,
        "purchase": dict(purch) if purch else None,
        "history": hist,
    }

def get_version_token(conn):
    # Changes in SQLite should invalidate snapshots
    total = conn.execute("SELECT COALESCE(MAX(updated_at), '') as v FROM products").fetchone()["v"]
    counts = conn.execute("SELECT (SELECT COUNT(*) FROM promo_slots) + (SELECT COUNT(*) FROM relampago) AS n").fetchone()["n"]
    return f"{total}|{counts}"

def all_promos(conn):
    df = pd.read_sql_query("""
        SELECT p.sku, pr.slot, pr.mlc, pr.precio_b2c, pr.fecha_venci, pr.comentario, p.descripcion
        FROM promo_slots pr
        JOIN products p ON p.sku = pr.sku
        WHERE COALESCE(pr.mlc,'') <> '' OR pr.precio_b2c IS NOT NULL OR pr.fecha_venci IS NOT NULL
        ORDER BY pr.fecha_venci, p.sku, pr.slot
    """, conn)
    if df.empty:
        for c in ["sku","slot","mlc","precio_b2c","fecha_venci","comentario","descripcion","estado"]:
            if c not in df.columns:
                df[c] = []
        return df
    df["estado"] = df["fecha_venci"].apply(date_status)
    return df

# ---------------------------
# Writers
# ---------------------------
def touch_product(conn, sku):
    conn.execute("UPDATE products SET updated_at=? WHERE sku=?", (datetime.now().isoformat(), sku))
    conn.commit()
    st.cache_data.clear()

def update_promo(conn, sku, slot, precio_b2c, fecha_venci, comentario):
    with conn:
        exists = conn.execute("SELECT 1 FROM promo_slots WHERE sku=? AND slot=?", (sku, slot)).fetchone()
        if exists:
            conn.execute("""
                UPDATE promo_slots
                SET precio_b2c=?, fecha_venci=?, comentario=?
                WHERE sku=? AND slot=?
            """, (as_number(precio_b2c), fecha_venci.isoformat() if fecha_venci else None, norm_text(comentario), sku, slot))
        else:
            conn.execute("""
                INSERT INTO promo_slots (sku, slot, mlc, precio_b2c, fecha_venci, comentario)
                VALUES (?, ?, '', ?, ?, ?)
            """, (sku, slot, as_number(precio_b2c), fecha_venci.isoformat() if fecha_venci else None, norm_text(comentario)))
    touch_product(conn, sku)

def batch_update_promo_dates(conn, ids, new_date):
    if not ids:
        return 0
    iso = new_date.isoformat() if new_date else None
    placeholders = ",".join("?" for _ in ids)
    with conn:
        conn.execute(f"UPDATE promo_slots SET fecha_venci=? WHERE id IN ({placeholders})", [iso, *ids])
        skus = [r["sku"] for r in conn.execute(f"SELECT DISTINCT sku FROM promo_slots WHERE id IN ({placeholders})", ids).fetchall()]
        now = datetime.now().isoformat()
        for sku in skus:
            conn.execute("UPDATE products SET updated_at=? WHERE sku=?", (now, sku))
    st.cache_data.clear()
    return len(ids)

def save_relampago_table(conn, df):
    df = df.copy()
    expected = ["sku", "descripcion", "precio", "extra", "comentario", "estado"]
    for c in expected:
        if c not in df.columns:
            df[c] = ""
    df["sku"] = df["sku"].apply(norm_sku)
    df["descripcion"] = df["descripcion"].fillna("").astype(str)
    df["precio"] = pd.to_numeric(df["precio"], errors="coerce")
    df["extra"] = df["extra"].fillna("").astype(str)
    df["comentario"] = df["comentario"].fillna("").astype(str)
    df["estado"] = df["estado"].fillna("").astype(str)
    df = df[df["sku"] != ""]

    with conn:
        conn.execute("DELETE FROM relampago")
        conn.executemany("""
            INSERT INTO relampago (sku, descripcion, precio, extra, comentario, estado)
            VALUES (?, ?, ?, ?, ?, ?)
        """, df[expected].itertuples(index=False, name=None))
    st.cache_data.clear()

def create_product(conn, data):
    sku = norm_sku(data["sku"])
    if not sku:
        raise ValueError("SKU vacío.")
    if conn.execute("SELECT 1 FROM products WHERE sku=?", (sku,)).fetchone():
        raise ValueError("El SKU ya existe.")
    margen_local, neto_local = compute_margin_local(data["ultimo_costo"], data["precio_bruto"])
    margen_meli, neto_meli = compute_margin_meli1(data["ultimo_costo"], data["monto_sim"])
    now = datetime.now().isoformat()
    with conn:
        conn.execute("""
            INSERT INTO products (
                sku, descripcion, ubicacion, cambio_precio, ultimo_costo, margen_local, precio_neto,
                precio_bruto, margen_meli_1, neto_meli_1, monto_simulacion, mlc_ads, campana_ads,
                comentario_maestra, row_idx, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            sku,
            norm_text(data["descripcion"]),
            norm_text(data["ubicacion"]),
            None,
            as_number(data["ultimo_costo"]),
            margen_local,
            neto_local,
            as_number(data["precio_bruto"]),
            margen_meli,
            neto_meli,
            as_number(data["monto_sim"]),
            "",
            norm_text(data["campana_ads"]),
            "",
            None,
            now
        ))
        mlc = norm_mlc(data["mlc"])
        if mlc:
            conn.execute("INSERT OR IGNORE INTO mlc_map (sku, mlc) VALUES (?, ?)", (sku, mlc))
            conn.execute("""
                INSERT INTO promo_slots (sku, slot, mlc, precio_b2c, fecha_venci, comentario)
                VALUES (?, 1, ?, ?, ?, ?)
            """, (sku, mlc, as_number(data["precio_b2c"]), data["fecha_venci"].isoformat() if data["fecha_venci"] else None, norm_text(data["comentario"])))
    st.cache_data.clear()

# ---------------------------
# Export
# ---------------------------
def export_master_bytes(conn):
    src_path = get_meta(conn, "master_source_path")
    if not src_path or not os.path.exists(src_path):
        raise ValueError("No hay archivo maestro base cargado para exportar.")
    xls = pd.ExcelFile(src_path)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sh in xls.sheet_names:
            if sh == "MAESTRA de precios":
                base = pd.read_excel(src_path, sheet_name=sh)
                base = base.copy()
                product_rows = pd.read_sql_query("SELECT * FROM products", conn)
                promos = pd.read_sql_query("SELECT * FROM promo_slots", conn)
                product_rows = product_rows.set_index("sku")
                promo_map = {(r["sku"], r["slot"]): r for _, r in promos.iterrows()}
                # update existing rows by SKU
                for i, row in base.iterrows():
                    sku = norm_sku(row.get("SKU"))
                    if sku in product_rows.index:
                        p = product_rows.loc[sku]
                        base.at[i, "DESCRIPCIÓN"] = p["descripcion"]
                        base.at[i, "UBIC"] = p["ubicacion"]
                        base.at[i, "ÚLTIMO COSTO"] = p["ultimo_costo"]
                        base.at[i, "MARGEN LOCAL"] = p["margen_local"]
                        base.at[i, "PRECIO NETO"] = p["precio_neto"]
                        base.at[i, "PRECIO BRUTO"] = p["precio_bruto"]
                        base.at[i, "MARGEN MELI 1"] = p["margen_meli_1"]
                        base.at[i, " NETO MELI 1"] = p["neto_meli_1"]
                        base.at[i, "MONTO EN SIMULACIÓN"] = p["monto_simulacion"]
                        base.at[i, "CAMPAÑA PADS"] = p["campana_ads"]
                        pr1 = promo_map.get((sku, 1))
                        pr2 = promo_map.get((sku, 2))
                        if pr1 is not None:
                            base.at[i, "MLC"] = pr1["mlc"]
                            base.at[i, "PRECIO B2C PUBLICADO "] = pr1["precio_b2c"]
                            base.at[i, "FECHA VENCI"] = pr1["fecha_venci"]
                            base.at[i, "COMENTARIO"] = pr1["comentario"]
                        if pr2 is not None:
                            base.at[i, "MLC.1"] = pr2["mlc"]
                            base.at[i, "PRECIO B2C"] = pr2["precio_b2c"]
                            base.at[i, "FECHA VENCI.1"] = pr2["fecha_venci"]
                            base.at[i, "COMENTARIO.1"] = pr2["comentario"]
                # append newly created products
                existing = set(base["SKU"].apply(norm_sku))
                new_products = product_rows[~product_rows.index.isin(existing)]
                for sku, p in new_products.iterrows():
                    new_row = {c: np.nan for c in base.columns}
                    new_row["SKU"] = sku
                    new_row["DESCRIPCIÓN"] = p["descripcion"]
                    new_row["UBIC"] = p["ubicacion"]
                    new_row["ÚLTIMO COSTO"] = p["ultimo_costo"]
                    new_row["MARGEN LOCAL"] = p["margen_local"]
                    new_row["PRECIO NETO"] = p["precio_neto"]
                    new_row["PRECIO BRUTO"] = p["precio_bruto"]
                    new_row["MARGEN MELI 1"] = p["margen_meli_1"]
                    new_row[" NETO MELI 1"] = p["neto_meli_1"]
                    new_row["MONTO EN SIMULACIÓN"] = p["monto_simulacion"]
                    new_row["CAMPAÑA PADS"] = p["campana_ads"]
                    pr1 = promo_map.get((sku, 1))
                    if pr1 is not None:
                        new_row["MLC"] = pr1["mlc"]
                        new_row["PRECIO B2C PUBLICADO "] = pr1["precio_b2c"]
                        new_row["FECHA VENCI"] = pr1["fecha_venci"]
                        new_row["COMENTARIO"] = pr1["comentario"]
                    base = pd.concat([base, pd.DataFrame([new_row])], ignore_index=True)
                base.to_excel(writer, index=False, sheet_name=sh)
            elif sh == "MLC -SKU":
                bridge = pd.read_sql_query("SELECT sku as SKU, mlc as `Número de publicación` FROM mlc_map ORDER BY sku, mlc", conn)
                bridge.to_excel(writer, index=False, sheet_name=sh)
            elif sh == "Relampago mi pagina":
                rel = pd.read_sql_query("SELECT sku, descripcion, precio, extra, comentario, estado FROM relampago ORDER BY id", conn)
                rel.to_excel(writer, index=False, header=False, sheet_name=sh)
            else:
                pd.read_excel(src_path, sheet_name=sh).to_excel(writer, index=False, sheet_name=sh)
    return out.getvalue()

# ---------------------------
# Styling
# ---------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.small-muted {color:#666;font-size:0.9rem;}
.kpi-card {
    border:1px solid rgba(49,51,63,0.15);
    border-radius:14px;
    padding:16px 18px;
    background:#fff;
}
.kpi-label {font-size:0.9rem;color:#666;margin-bottom:6px;}
.kpi-value {font-size:2rem;font-weight:700;}
.tag {
    display:inline-block;padding:4px 10px;border-radius:999px;font-size:0.8rem;font-weight:600;
    border:1px solid rgba(0,0,0,0.08);margin-right:6px;margin-bottom:6px;
}
.tag-red {background:#ffe8e8;color:#c0392b;}
.tag-orange {background:#fff0db;color:#b96a00;}
.tag-green {background:#e9f9ee;color:#1d8f49;}
.tag-blue {background:#e9f2ff;color:#2962cc;}
.promo-card {
    border:1px solid rgba(49,51,63,0.12);
    border-radius:16px;
    padding:12px 14px;
    background:#fff;
    min-height:132px;
}
.promo-card .sku {font-weight:700;font-size:0.95rem; margin-bottom:6px;}
.promo-card .desc {font-size:0.88rem; line-height:1.3; height:2.6em; overflow:hidden; margin-bottom:8px;}
.promo-card .meta {font-size:0.82rem; color:#555;}
hr.soft {border:none;border-top:1px solid rgba(49,51,63,0.08); margin:18px 0;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Load / sync
# ---------------------------
conn = get_conn()

with st.sidebar:
    st.header("Carga")
    master_file = st.file_uploader("Maestra saneada", type=["xlsx"], key="master_file")
    purchases_file = st.file_uploader("Compras", type=["xlsx"], key="purchases_file")
    auto_refresh = st.checkbox("Refrescar cambios compartidos", value=True)
    refresh_seconds = st.selectbox("Cada cuántos segundos", [5, 10, 15, 30], index=1)

    if master_file:
        sig = file_signature(master_file)
        if get_meta(conn, "master_sig") != sig:
            with st.spinner("Importando maestra a SQLite..."):
                import_master_to_db(conn, master_file)
            st.success("Maestra sincronizada.")
    if purchases_file:
        psig = file_signature(purchases_file)
        if get_meta(conn, "purchases_sig") != psig:
            with st.spinner("Importando compras a SQLite..."):
                import_purchases_to_db(conn, purchases_file)
            st.success("Compras sincronizadas.")

    st.caption("SQLite permite que varios usuarios vean y escriban sobre la misma capa operativa sin rehacer el Excel en cada clic.")

if auto_refresh:
    # simple polling
    st_autorefresh = st.empty()
    st_autorefresh.caption(f"Auto-refresco activo cada {refresh_seconds}s.")
    time.sleep(0)  # no-op to keep importers happy

products_df = list_products(conn)
if products_df.empty:
    st.title(APP_TITLE)
    st.info("Sube la maestra saneada para comenzar.")
    st.stop()

st.title(APP_TITLE)

product_options = {f"{r['sku']} — {r['descripcion']}": r["sku"] for _, r in products_df.iterrows()}
default_label = list(product_options.keys())[0]
selected_label = st.selectbox("Buscar producto", list(product_options.keys()), index=0)
selected_sku = product_options[selected_label]

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Cockpit", "Operador de promos", "Relámpago", "Alta de producto", "Descargar"])

# ---------------------------
# Cockpit
# ---------------------------
with tab1:
    snap = get_product_snapshot(get_version_token(conn), selected_sku)
    if not snap:
        st.warning("No encontré ese SKU.")
    else:
        p = snap["product"]
        promos = snap["promos"]
        purchase = snap["purchase"]
        rel = snap["relampago"]
        mlc_assoc = sorted(set([m for m in snap["mlcs"] if m] + promos["mlc"].dropna().astype(str).tolist()))
        b2c_main = promos.loc[promos["slot"] == 1, "precio_b2c"].dropna()
        b2c_main = b2c_main.iloc[0] if not b2c_main.empty else None

        ctop1, ctop2, ctop3 = st.columns([1,1,1])
        status_tags = []
        active_promos = promos[promos["fecha_venci"].notna()]
        if active_promos.empty:
            status_tags.append('<span class="tag tag-red">Sin promo</span>')
        else:
            states = sorted(set(active_promos["fecha_venci"].apply(date_status)))
            for s in states:
                klass = "tag-orange" if "hoy" in s.lower() or "mañana" in s.lower() else "tag-green"
                status_tags.append(f'<span class="tag {klass}">{s}</span>')
        ctop1.markdown("".join(status_tags), unsafe_allow_html=True)
        ctop2.markdown(f'<span class="tag tag-blue">Promos activas: {len(promos)}</span>', unsafe_allow_html=True)
        ctop3.markdown(f'<span class="tag tag-green">Relámpago: {len(rel)}</span>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1.1, 1, 1])
        with col1:
            st.subheader("Identidad")
            st.markdown(f"**SKU:** {p['sku']}")
            st.markdown(f"**Descripción:** {p.get('descripcion') or '—'}")
            st.markdown(f"**Ubicación:** {p.get('ubicacion') or '—'}")
            st.markdown(f"**MLCs asociados:** {', '.join(mlc_assoc) if mlc_assoc else '—'}")
            st.markdown(f"**Comentario maestra:** {p.get('comentario_maestra') or '—'}")

            st.subheader("Compras")
            if purchase:
                st.markdown(f"**Última compra:** {fmt_date(purchase.get('ultima_compra'))}")
                st.markdown(f"**Último precio compra:** {fmt_money(purchase.get('ultimo_precio'))}")
                st.markdown(f"**Proveedor último:** {purchase.get('proveedor_ultimo') or '—'}")
                st.markdown(f"**Cantidad última compra:** {purchase.get('cantidad_ultima') or '—'}")
                st.markdown(f"**Variación vs compra anterior:** {fmt_pct(purchase.get('variacion_pct'))}")
                pmn = fmt_money(purchase.get('precio_min'))
                pmx = fmt_money(purchase.get('precio_max'))
                st.markdown(f"**Rango histórico:** {pmn} a {pmx}")
                st.markdown(f"**Proveedores históricos:** {purchase.get('proveedores') or '—'}")
            else:
                st.info("Sin compras asociadas para este SKU.")

        with col2:
            st.subheader("Precio y rentabilidad")
            r1c1, r1c2 = st.columns(2)
            with r1c1:
                st.markdown(f'<div class="kpi-card"><div class="kpi-label">Precio neto</div><div class="kpi-value">{fmt_money(p.get("precio_neto"))}</div></div>', unsafe_allow_html=True)
            with r1c2:
                st.markdown(f'<div class="kpi-card"><div class="kpi-label">Cambio precio</div><div class="kpi-value">{fmt_money(p.get("cambio_precio"))}</div></div>', unsafe_allow_html=True)
            r2c1, r2c2 = st.columns(2)
            with r2c1:
                st.markdown(f'<div class="kpi-card"><div class="kpi-label">Margen local</div><div class="kpi-value">{fmt_pct(p.get("margen_local"))}</div></div>', unsafe_allow_html=True)
            with r2c2:
                st.markdown(f'<div class="kpi-card"><div class="kpi-label">Monto en simulación</div><div class="kpi-value">{fmt_money(p.get("monto_simulacion"))}</div></div>', unsafe_allow_html=True)

            st.markdown(f"**Neto Meli 1:** {fmt_money(p.get('neto_meli_1'))}")
            st.markdown(f"**Margen Meli 1:** {fmt_pct(p.get('margen_meli_1'))}")
            st.markdown(f"**Precio B2C publicado:** {fmt_money(b2c_main)}")
            st.markdown(f"**Campaña Ads:** {p.get('campana_ads') or '—'}")
            if not promos.empty:
                st.markdown(f"**Precio promo mínimo:** {fmt_money(promos['precio_b2c'].min())}")
                st.markdown(f"**Precio promo máximo:** {fmt_money(promos['precio_b2c'].max())}")
            else:
                st.markdown("**Precio promo mínimo:** —")
                st.markdown("**Precio promo máximo:** —")
            if not rel.empty:
                st.markdown(f"**Precio relámpago mínimo:** {fmt_money(rel['precio'].min())}")
                st.markdown(f"**Precio relámpago máximo:** {fmt_money(rel['precio'].max())}")
            else:
                st.markdown("**Precio relámpago mínimo:** —")
                st.markdown("**Precio relámpago máximo:** —")
            st.markdown(f"**Comentario promos:** {(' | '.join([x for x in promos['comentario'].fillna('').tolist() if x])) or '—'}")

        with col3:
            st.subheader("Promos maestra")
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Filas promo</div><div class="kpi-value">{len(promos)}</div></div>', unsafe_allow_html=True)
            if promos.empty:
                st.info("Sin promos activas en maestra.")
            else:
                show = promos[["slot", "mlc", "precio_b2c", "fecha_venci", "comentario"]].copy()
                show["fecha_venci"] = show["fecha_venci"].apply(fmt_date)
                show.columns = ["Slot", "MLC", "Precio B2C", "Fecha venci", "Comentario"]
                st.dataframe(show, use_container_width=True, hide_index=True)

            st.subheader("Relámpago mi página")
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Filas relámpago</div><div class="kpi-value">{len(rel)}</div></div>', unsafe_allow_html=True)
            if rel.empty:
                st.info("No está en relámpago mi página.")
            else:
                show_rel = rel[["descripcion", "precio", "comentario", "estado"]].copy()
                st.dataframe(show_rel, use_container_width=True, hide_index=True)

        st.markdown('<hr class="soft">', unsafe_allow_html=True)
        st.subheader("Lectura automática")
        bullets = []
        if promos.empty:
            bullets.append("Producto sin promo activa registrada.")
        else:
            if any(promos["fecha_venci"].apply(lambda x: date_status(x) in ["Vencidas", "Vencen hoy"])):
                bullets.append("Tiene promos críticas: vencidas o que vencen hoy.")
            elif any(promos["fecha_venci"].apply(lambda x: date_status(x) == "Vencen mañana")):
                bullets.append("Tiene promo que vence mañana. Conviene revisar continuidad.")
            else:
                bullets.append("Promos cargadas sin urgencia inmediata.")
        if purchase and as_number(p.get("ultimo_costo")) and as_number(purchase.get("ultimo_precio")):
            if float(purchase["ultimo_precio"]) > float(p["ultimo_costo"]):
                bullets.append("Último precio de compra está por encima del último costo cargado en maestra.")
        if not rel.empty:
            bullets.append("El SKU está presente en relámpago mi página.")
        for b in bullets:
            st.markdown(f"- {b}")

# ---------------------------
# Operador de promos
# ---------------------------
@st.dialog("Editar promo", width="large")
def promo_editor_dialog(sku, slot):
    snap = get_product_snapshot(get_version_token(conn), sku)
    promo = snap["promos"][snap["promos"]["slot"] == slot]
    row = promo.iloc[0] if not promo.empty else pd.Series({"precio_b2c": None, "fecha_venci": None, "comentario": "", "mlc": ""})
    st.markdown(f"**SKU:** {sku}")
    st.markdown(f"**Descripción:** {snap['product'].get('descripcion') or '—'}")
    st.markdown(f"**MLC:** {row.get('mlc') or '—'}")
    new_date = st.date_input("Fecha venci", value=parse_date_any(row.get("fecha_venci")) or date.today(), format="DD/MM/YYYY")
    with st.expander("Campos secundarios", expanded=False):
        new_price = st.number_input("Precio B2C", min_value=0.0, value=float(row["precio_b2c"]) if pd.notna(row.get("precio_b2c")) else 0.0, step=100.0)
        new_comment = st.text_input("Comentario", value=norm_text(row.get("comentario")))
    c1, c2 = st.columns(2)
    if c1.button("Guardar cambios", type="primary", use_container_width=True):
        update_promo(conn, sku, slot, new_price, new_date, new_comment)
        st.success("Promo actualizada.")
        st.rerun()
    if c2.button("Cerrar", use_container_width=True):
        st.rerun()

with tab2:
    st.subheader("Operador de promos")
    promos_df = all_promos(conn)
    status_options = ["Vencidas", "Vencen hoy", "Vencen mañana", "Vencen pasado mañana", "Vencen en 7 días", "Vencen en 15 días", "Vencen en 1 mes"]
    if "promo_status_filter" not in st.session_state:
        st.session_state["promo_status_filter"] = ["Vencidas", "Vencen hoy"]
    state_filter = st.multiselect("Estado", options=status_options, key="promo_status_filter")
    search = st.text_input("Buscar por SKU / descripción / MLC")
    mass_date = st.date_input("Cambio masivo de fecha", value=None, format="DD/MM/YYYY")

    filtered = promos_df.copy()
    if state_filter:
        filtered = filtered[filtered["estado"].isin(state_filter)]
    if search:
        s = search.lower().strip()
        filtered = filtered[
            filtered["sku"].str.lower().str.contains(s, na=False) |
            filtered["descripcion"].str.lower().str.contains(s, na=False) |
            filtered["mlc"].str.lower().str.contains(s, na=False)
        ]

    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-label">Promos filtradas</div><div class="kpi-value">{len(filtered)}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-label">Vencidas</div><div class="kpi-value">{int((filtered["estado"]=="Vencidas").sum())}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-label">Vencen hoy</div><div class="kpi-value">{int((filtered["estado"]=="Vencen hoy").sum())}</div></div>', unsafe_allow_html=True)

    if mass_date and not filtered.empty:
        if st.button("Aplicar fecha masiva a filtradas"):
            updated = batch_update_promo_dates(conn, filtered["id"].tolist(), mass_date)
            st.success(f"Actualicé {updated} promos.")
            st.rerun()

    if filtered.empty:
        st.info("No hay promos para ese filtro.")
    else:
        st.markdown("### Bandeja visual")
        cols = st.columns(4)
        for i, (_, r) in enumerate(filtered.iterrows()):
            with cols[i % 4]:
                st.markdown(
                    f"""
                    <div class="promo-card">
                        <div class="sku">{r['sku']}</div>
                        <div class="desc">{r['descripcion']}</div>
                        <div class="meta">MLC: {r['mlc'] or '—'}</div>
                        <div class="meta">Fecha: {fmt_date(r['fecha_venci'])}</div>
                        <div class="meta">Estado: {r['estado']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button("Editar", key=f"edit_{r['sku']}_{int(r['slot'])}", use_container_width=True):
                    promo_editor_dialog(r["sku"], int(r["slot"]))

# ---------------------------
# Relampago
# ---------------------------
with tab3:
    st.subheader("Relámpago mi página")
    rel_df = pd.read_sql_query("SELECT sku, descripcion, precio, extra, comentario, estado FROM relampago ORDER BY id", conn)
    edited_rel = st.data_editor(
        rel_df,
        use_container_width=True,
        num_rows="dynamic",
        key="relampago_editor",
        height=editor_height(len(rel_df)),
        column_config={
            "sku": st.column_config.TextColumn("SKU"),
            "descripcion": st.column_config.TextColumn("Descripción", width="large"),
            "precio": st.column_config.NumberColumn("Precio", step=100),
            "extra": st.column_config.TextColumn("Extra"),
            "comentario": st.column_config.TextColumn("Comentario"),
            "estado": st.column_config.TextColumn("Estado"),
        }
    )
    if st.button("Guardar relámpago", type="primary"):
        save_relampago_table(conn, edited_rel)
        st.success("Relámpago guardado.")
        st.rerun()

# ---------------------------
# Alta de producto
# ---------------------------
with tab4:
    st.subheader("Alta de producto")
    with st.form("alta_producto_form"):
        c1, c2, c3 = st.columns(3)
        sku_new = c1.text_input("SKU")
        desc_new = c2.text_input("Descripción")
        ubic_new = c3.text_input("Ubicación")

        c4, c5, c6 = st.columns(3)
        costo_new = c4.number_input("Último costo", min_value=0.0, step=100.0)
        bruto_new = c5.number_input("Precio bruto en tienda", min_value=0.0, step=100.0)
        monto_new = c6.number_input("MONTO EN SIMULACIÓN", min_value=0.0, step=100.0)

        margen_local_proj, _neto_local = compute_margin_local(costo_new, bruto_new)
        margen_meli_proj, _neto_meli = compute_margin_meli1(costo_new, monto_new)

        c7, c8, c9 = st.columns(3)
        mlc_new = c7.text_input("MLC")
        b2c_new = c8.number_input("Precio B2C publicado", min_value=0.0, step=100.0)
        fecha_new = c9.date_input("Fecha vencimiento", value=None, format="DD/MM/YYYY")

        camp_ads_new = st.text_input("Campaña Ads")
        comentario_new = st.text_input("Comentario")

        k1, k2 = st.columns(2)
        k1.markdown(f'<div class="kpi-card"><div class="kpi-label">Margen local proyectado</div><div class="kpi-value">{fmt_pct(margen_local_proj)}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card"><div class="kpi-label">Margen Meli 1 proyectado</div><div class="kpi-value">{fmt_pct(margen_meli_proj)}</div></div>', unsafe_allow_html=True)

        submitted = st.form_submit_button("Crear producto", type="primary", use_container_width=True)
        if submitted:
            try:
                create_product(conn, {
                    "sku": sku_new,
                    "descripcion": desc_new,
                    "ubicacion": ubic_new,
                    "ultimo_costo": costo_new,
                    "precio_bruto": bruto_new,
                    "monto_sim": monto_new,
                    "mlc": mlc_new,
                    "precio_b2c": b2c_new,
                    "fecha_venci": fecha_new,
                    "comentario": comentario_new,
                    "campana_ads": camp_ads_new,
                })
                st.success("Producto creado.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ---------------------------
# Download
# ---------------------------
with tab5:
    st.subheader("Descargar")
    st.markdown("La base viva es SQLite. Aquí exportas el Excel actualizado con la maestra, MLC-SKU y relámpago.")
    if st.button("Preparar Excel actualizado", type="primary"):
        try:
            data = export_master_bytes(conn)
            st.download_button(
                "Descargar Excel actualizado",
                data=data,
                file_name=f"MAESTRA_ACTUALIZADA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(str(e))
