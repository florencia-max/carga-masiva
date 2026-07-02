"""
AllRide – Cuadratura de viajes
Lógica:
  1. Por cada viaje del cliente (fecha + ruta + empresa) calcular hora_allride = postura + 15 min
  2. Si ya existe en AllRide con esa hora exacta → OK, conservar
  3. Si NO existe con esa hora pero hay otro horario de la misma ruta ese día → editar el más cercano
  4. Si hay viajes en AllRide de esa ruta ese día que NO coinciden con ningún viaje del cliente → cancelar
  5. Si no existe ningún viaje de esa ruta ese día en AllRide → crear (one-time service)
"""

import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import timedelta, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from copy import copy

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="AllRide – Cuadratura",
    page_icon="🚌",
    layout="wide",
)

EMP_MAP = {
    "KAUFMANN": "KAUFMAN",
    "FALABELLA - LA SERENA": "FALABELLA - SERENA",
    "FALABELLA - VALPARAÍSO": "FALABELLA - VALPARAISO",
    "TRADIS LOGÍSTICA FALABELLA": "TRADIS LOGISTICA FALABELLA",
    "TÍO TOMATE": "TIO TOMATE",
}

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def norm_str(s: str) -> str:
    s = str(s).strip().upper()
    for a, b in [("Á","A"),("É","E"),("Í","I"),("Ó","O"),("Ú","U"),("Ñ","N")]:
        s = s.replace(a, b)
    return s

def norm_empresa(s: str) -> str:
    n = norm_str(s)
    return EMP_MAP.get(n, n)

def norm_ruta(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"^RDD\s*-\s*", "", s)
    return norm_str(s)

def is_spot(ruta: str) -> bool:
    return "SPOT" in norm_str(ruta)

def parse_hora_str(val) -> str:
    """Devuelve 'HH:MM' desde cualquier formato."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "00:00"
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    if isinstance(val, timedelta):
        total = int(val.total_seconds())
        h, m = divmod(total // 60, 60)
        return f"{h:02d}:{m:02d}"
    if isinstance(val, (int, float)):
        frac = val - int(val)
        total_min = round(frac * 1440)
        h, m = divmod(total_min, 60)
        return f"{h % 24:02d}:{m:02d}"
    s = str(val).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return "00:00"

def sumar_15(hora_str: str) -> str:
    m = re.match(r"^(\d{1,2}):(\d{2})$", hora_str)
    if not m:
        return hora_str
    total = int(m.group(1)) * 60 + int(m.group(2)) + 15
    total %= 1440
    return f"{total // 60:02d}:{total % 60:02d}"

def hora_a_min(h: str) -> int:
    m = re.match(r"^(\d{1,2}):(\d{2})$", h)
    if not m:
        return 0
    return int(m.group(1)) * 60 + int(m.group(2))

def parse_fecha_allride(val) -> str:
    """Extrae fecha 'DD/MM/YYYY' de la celda de AllRide (puede ser 'DD/MM/YYYY, HH:MM')."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).split(",")[0].strip()
    m = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", s)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    return s

def parse_hora_allride(val) -> str:
    """Extrae hora 'HH:MM' de la celda de AllRide."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "00:00"
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    s = str(val)
    parts = s.split(",")
    if len(parts) > 1:
        return parse_hora_str(parts[1].strip())
    return parse_hora_str(val)

def detect_col(cols, candidates):
    """Encuentra la primera columna que coincida (case-insensitive)."""
    cols_norm = {norm_str(c): c for c in cols}
    for cand in candidates:
        n = norm_str(cand)
        if n in cols_norm:
            return cols_norm[n]
    return None

# ══════════════════════════════════════════════════════════════════════════════
# CARGA DE ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════════
def leer_consolidado(f) -> pd.DataFrame:
    xl = pd.ExcelFile(f)
    hoja = None
    for h in xl.sheet_names:
        if norm_str(h) in ["RESUMEN", "PROGRAMACION RM", "PROGRAMACION", "PROGRAMA"]:
            hoja = h
            break
    if hoja is None:
        hoja = xl.sheet_names[0]
    df = pd.read_excel(f, sheet_name=hoja)
    st.caption(f"📋 Consolidado: hoja **{hoja}** — {len(df)} filas")
    return df

def leer_allride(f) -> pd.DataFrame:
    df = pd.read_excel(f)
    st.caption(f"🚌 AllRide: {len(df)} filas — estados: {df['Estado'].unique().tolist() if 'Estado' in df.columns else '?'}")
    return df

# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZAR CONSOLIDADO
# ══════════════════════════════════════════════════════════════════════════════
def procesar_consolidado(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)

    col_ruta = detect_col(cols, ["NUEVO NOMBRE RUTA FINAL", "NOMBRE RUTA FINAL", "NOMBRE RUTA"])
    col_postura = detect_col(cols, ["HORA DE POSTURA"])
    col_hora = detect_col(cols, [
        "HORA DE LLEGADA A BODEGA", "HORA DE LLEGADA\n HORA DE SALIDA",
        "HORA DE SALIDA", "HORA DE LLEGADA", "HORA"
    ])
    col_fecha = detect_col(cols, ["FECHA"])
    col_empresa = detect_col(cols, ["EMPRESA"])
    col_tipo = detect_col(cols, ["TIPO DE PEDIDO", "TIPO"])

    if not col_ruta:
        st.error("❌ No se encontró columna de ruta en el consolidado.")
        st.stop()
    if not col_fecha:
        st.error("❌ No se encontró columna de fecha en el consolidado.")
        st.stop()

    out = pd.DataFrame()
    out["ruta_orig"]    = df[col_ruta].astype(str).str.strip()
    out["ruta_norm"]    = out["ruta_orig"].apply(norm_ruta)
    out["empresa_orig"] = df[col_empresa].astype(str).str.strip() if col_empresa else ""
    out["empresa_norm"] = out["empresa_orig"].apply(norm_empresa)
    out["fecha"]        = pd.to_datetime(df[col_fecha], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
    out["tipo_norm"]    = df[col_tipo].apply(lambda x: norm_str(str(x))) if col_tipo else "REGULAR"

    # Hora AllRide = postura + 15 min (si existe) o hora directa
    if col_postura:
        out["hora_postura"]   = df[col_postura].apply(parse_hora_str)
        out["hora_allride"]   = out["hora_postura"].apply(sumar_15)
    else:
        out["hora_postura"]   = df[col_hora].apply(parse_hora_str) if col_hora else "00:00"
        out["hora_allride"]   = out["hora_postura"]  # sin postura, se usa directa

    out["es_spot"] = out["ruta_norm"].apply(is_spot)
    out = out[out["ruta_orig"].notna() & (out["ruta_orig"] != "") & (out["ruta_orig"] != "NAN")]
    out = out[out["fecha"].notna() & (out["fecha"] != "NaT") & (out["fecha"] != "")]
    return out.reset_index(drop=True)

# ══════════════════════════════════════════════════════════════════════════════
# NORMALIZAR ALLRIDE
# ══════════════════════════════════════════════════════════════════════════════
TIPO_MAP_AR = {
    "SALIDA CALENDARIZADA": "REGULAR",
    "SALIDA REALIZADA": "REGULAR",
    "SERVICIO REGULAR": "REGULAR",
    "SERVICIO ESPECIAL": "SPOT",
}

def procesar_allride(df: pd.DataFrame) -> pd.DataFrame:
    col_fecha_raw = detect_col(list(df.columns), ["Fecha estimada de inicio", "Fecha de inicio"])
    col_ruta = detect_col(list(df.columns), ["Ruta"])
    col_empresa = detect_col(list(df.columns), ["Comunidades"])
    col_tipo = detect_col(list(df.columns), ["Tipo"])
    col_id = detect_col(list(df.columns), ["ID de servicio", "Servicio", "ID"])
    col_estado = detect_col(list(df.columns), ["Estado"])

    if not col_fecha_raw or not col_ruta:
        st.error("❌ No se encontraron columnas clave en el archivo AllRide.")
        st.stop()

    fecha_raw = df[col_fecha_raw] if col_fecha_raw else pd.Series([""] * len(df))

    out = pd.DataFrame()
    out["id"]          = df[col_id].astype(str).str.strip() if col_id else ""
    out["ruta_orig"]   = df[col_ruta].astype(str).str.strip()
    out["ruta_norm"]   = out["ruta_orig"].apply(norm_ruta)
    out["empresa_orig"] = df[col_empresa].astype(str).str.strip() if col_empresa else ""
    out["empresa_norm"] = out["empresa_orig"].apply(norm_empresa)
    out["tipo_orig"]   = df[col_tipo].astype(str).str.strip() if col_tipo else ""
    out["tipo_norm"]   = out["tipo_orig"].apply(lambda x: TIPO_MAP_AR.get(norm_str(x), norm_str(x)))
    out["estado"]      = df[col_estado].astype(str).str.strip() if col_estado else ""
    out["fecha"]       = fecha_raw.apply(parse_fecha_allride)
    out["hora"]        = fecha_raw.apply(parse_hora_allride)
    out["es_spot"]     = out["ruta_norm"].apply(is_spot)

    # Preservar todas las columnas originales para reconstruir archivo cancelación
    out["_orig_idx"] = df.index
    return out.reset_index(drop=True)

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA PRINCIPAL DE CUADRATURA
# ══════════════════════════════════════════════════════════════════════════════
def cuadrar(cli: pd.DataFrame, ar: pd.DataFrame):
    """
    Devuelve:
      - ok:       filas AllRide que ya están correctas (hora exacta match)
      - editar:   list of dicts {ar_idx, ar_row, cli_row, hora_nueva}
      - cancelar: filas AllRide a cancelar
      - crear:    filas cliente sin ningún AllRide ese día para esa ruta
    """
    resultados = {
        "ok": [],
        "editar": [],
        "cancelar": [],
        "crear": [],
    }

    # Agrupar AllRide por (fecha, ruta_norm, empresa_norm)
    ar_grupos = {}
    for idx, row in ar.iterrows():
        k = (row["fecha"], row["ruta_norm"], row["empresa_norm"])
        ar_grupos.setdefault(k, []).append((idx, row))

    # Set de índices AllRide ya "usados" (matched)
    ar_usados = set()

    # Para cada viaje del cliente
    for _, crow in cli.iterrows():
        k = (crow["fecha"], crow["ruta_norm"], crow["empresa_norm"])
        hora_objetivo = crow["hora_allride"]
        candidatos = ar_grupos.get(k, [])

        if not candidatos:
            # No hay ningún viaje de esa ruta ese día → CREAR
            resultados["crear"].append(crow)
            continue

        # ¿Alguno ya tiene la hora correcta?
        exactos = [(idx, r) for idx, r in candidatos if r["hora"] == hora_objetivo]
        if exactos:
            # Usar el primero exacto
            idx_match, row_match = exactos[0]
            ar_usados.add(idx_match)
            resultados["ok"].append(row_match)
            continue

        # No hay exacto → buscar el más cercano en hora para EDITAR
        min_objetivo = hora_a_min(hora_objetivo)
        no_usados = [(idx, r) for idx, r in candidatos if idx not in ar_usados]
        if not no_usados:
            resultados["crear"].append(crow)
            continue

        no_usados_sorted = sorted(no_usados, key=lambda x: abs(hora_a_min(x[1]["hora"]) - min_objetivo))
        idx_edit, row_edit = no_usados_sorted[0]
        ar_usados.add(idx_edit)
        resultados["editar"].append({
            "ar_idx": idx_edit,
            "ar_row": row_edit,
            "cli_row": crow,
            "hora_nueva": hora_objetivo,
            "hora_actual": row_edit["hora"],
        })

    # Todos los AllRide no usados → CANCELAR
    for k, grupo in ar_grupos.items():
        for idx, row in grupo:
            if idx not in ar_usados:
                resultados["cancelar"].append(row)

    return resultados

# ══════════════════════════════════════════════════════════════════════════════
# GENERADORES DE EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def gen_cancelacion(ar_orig_df: pd.DataFrame, cancel_filas: list, cancel_template_bytes) -> bytes:
    """Toma la plantilla de cancelación y marca con X las filas a cancelar."""
    ids_cancelar = set(str(r["id"]) for r in cancel_filas)
    wb = openpyxl.load_workbook(BytesIO(cancel_template_bytes))
    ws = wb.active

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 2)]
    try:
        col_id = header.index("ID de servicio") + 1
    except ValueError:
        col_id = 1
    try:
        col_cancel = header.index("Cancelar") + 1
    except ValueError:
        col_cancel = ws.max_column

    font_x = Font(name="Calibri", size=12, bold=True, color="C0392B")
    fill_x = PatternFill("solid", fgColor="FDEDEC")

    for r in range(2, ws.max_row + 1):
        id_val = str(ws.cell(r, col_id).value or "").strip()
        if id_val in ids_cancelar:
            ws.cell(r, col_cancel).value = "X"
            ws.cell(r, col_cancel).font = font_x
            ws.cell(r, col_cancel).fill = fill_x
            ws.cell(r, col_cancel).alignment = Alignment(horizontal="center")
        else:
            ws.cell(r, col_cancel).value = None

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def gen_edicion_horarios(editar_filas: list, edit_template_bytes) -> bytes:
    """
    Genera el archivo de edición de horarios a partir de la plantilla.
    Rellena: ID servicio, nueva hora.
    """
    wb = openpyxl.load_workbook(BytesIO(edit_template_bytes))
    ws = wb.active

    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 2)]
    # Buscar columnas relevantes
    def find_col(names):
        for n in names:
            for i, h in enumerate(header):
                if h and norm_str(str(h)) in [norm_str(n), norm_str(n).replace(" ", "")]:
                    return i + 1
        return None

    col_id    = find_col(["ID de servicio", "ID"])
    col_hora  = find_col(["Hora del servicio", "Hora de inicio", "Hora"])
    col_ruta  = find_col(["Ruta", "Nombre ruta"])

    if not col_id or not col_hora:
        # Si no encontramos columnas, construir desde cero
        ws.delete_rows(2, ws.max_row)
        for i, e in enumerate(editar_filas, 2):
            ws.cell(i, 1, e["ar_row"]["id"])
            ws.cell(i, 2, e["hora_nueva"])

    else:
        # Limpiar filas existentes
        ws.delete_rows(2, ws.max_row)
        for i, e in enumerate(editar_filas, 2):
            ar = e["ar_row"]
            if col_id:
                ws.cell(i, col_id, ar["id"]).font = Font(name="Calibri", size=12)
            if col_hora:
                ws.cell(i, col_hora, e["hora_nueva"]).font = Font(name="Calibri", size=12)
            if col_ruta:
                ws.cell(i, col_ruta, ar["ruta_orig"]).font = Font(name="Calibri", size=12)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def gen_resumen_excel(resultados: dict) -> bytes:
    """Genera un Excel con 4 hojas: OK, Editar, Cancelar, Crear."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#2C3E50", "font_color": "white", "border": 1})
        ok_fmt  = wb.add_format({"bg_color": "#D5F5E3", "border": 1})
        ed_fmt  = wb.add_format({"bg_color": "#FFF3CD", "border": 1})
        can_fmt = wb.add_format({"bg_color": "#FDEDEC", "border": 1})
        cre_fmt = wb.add_format({"bg_color": "#D6EAF8", "border": 1})

        def write_sheet(name, rows, cols, fmt, col_widths=None):
            ws = writer.sheets.get(name)
            if name not in writer.sheets:
                writer.book.add_worksheet(name)
                ws = writer.book.worksheets()[-1]
            for ci, c in enumerate(cols):
                ws.write(0, ci, c, hdr_fmt)
            for ri, row in enumerate(rows, 1):
                for ci, val in enumerate(row):
                    ws.write(ri, ci, str(val) if val is not None else "", fmt)
            if col_widths:
                for ci, w in enumerate(col_widths):
                    ws.set_column(ci, ci, w)

        # OK
        ok_rows = [(r["fecha"], r["hora"], r["ruta_orig"], r["empresa_orig"], r["tipo_orig"], r["id"]) for r in resultados["ok"]]
        ok_sheet = wb.add_worksheet("OK – Sin cambio")
        ok_sheet.write_row(0, 0, ["Fecha","Hora","Ruta","Empresa","Tipo","ID AllRide"], hdr_fmt)
        for ri, row in enumerate(ok_rows, 1):
            ok_sheet.write_row(ri, 0, [str(v) for v in row], ok_fmt)
        for ci, w in enumerate([12,8,50,28,22,32]):
            ok_sheet.set_column(ci, ci, w)

        # EDITAR
        ed_sheet = wb.add_worksheet("Editar – Cambio de hora")
        ed_sheet.write_row(0, 0, ["Fecha","Ruta","Empresa","ID AllRide","Hora actual AllRide","Hora nueva","Hora postura cliente"], hdr_fmt)
        for ri, e in enumerate(resultados["editar"], 1):
            ar = e["ar_row"]; cli = e["cli_row"]
            ed_sheet.write_row(ri, 0, [
                str(ar["fecha"]), str(ar["ruta_orig"]), str(ar["empresa_orig"]),
                str(ar["id"]), str(ar["hora"]), str(e["hora_nueva"]),
                str(cli.get("hora_postura",""))
            ], ed_fmt)
        for ci, w in enumerate([12,50,28,32,18,12,18]):
            ed_sheet.set_column(ci, ci, w)

        # CANCELAR
        can_sheet = wb.add_worksheet("Cancelar")
        can_sheet.write_row(0, 0, ["Fecha","Hora","Ruta","Empresa","Tipo","ID AllRide","Estado"], hdr_fmt)
        for ri, r in enumerate(resultados["cancelar"], 1):
            can_sheet.write_row(ri, 0, [
                str(r["fecha"]), str(r["hora"]), str(r["ruta_orig"]),
                str(r["empresa_orig"]), str(r["tipo_orig"]), str(r["id"]), str(r["estado"])
            ], can_fmt)
        for ci, w in enumerate([12,8,50,28,22,32,22]):
            can_sheet.set_column(ci, ci, w)

        # CREAR
        cre_sheet = wb.add_worksheet("Crear – Nuevos viajes")
        cre_sheet.write_row(0, 0, ["Fecha","Hora postura","Hora AllRide","Ruta","Empresa","Tipo"], hdr_fmt)
        for ri, r in enumerate(resultados["crear"], 1):
            cre_sheet.write_row(ri, 0, [
                str(r.get("fecha","")), str(r.get("hora_postura","")),
                str(r.get("hora_allride","")), str(r.get("ruta_orig","")),
                str(r.get("empresa_orig","")), str(r.get("tipo_norm",""))
            ], cre_fmt)
        for ci, w in enumerate([12,14,12,50,28,12]):
            cre_sheet.set_column(ci, ci, w)

    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.title("🚌 AllRide – Cuadratura de viajes")
st.markdown(
    "Sube los archivos para generar automáticamente: ediciones de horario, cancelaciones y viajes nuevos."
)

with st.sidebar:
    st.header("📁 Archivos")
    f_consolidado = st.file_uploader("Consolidado del cliente (.xlsx)", type=["xlsx","xls"], key="cli")
    f_allride     = st.file_uploader("Exportación AllRide (.xlsx)", type=["xlsx","xls"], key="ar")
    f_cancel_tmpl = st.file_uploader("Plantilla cancelación masiva (.xlsx)", type=["xlsx","xls"], key="can")
    f_edit_tmpl   = st.file_uploader("Plantilla edición horarios (.xlsx) — opcional", type=["xlsx","xls"], key="edit")

    st.divider()
    st.markdown("**Lógica aplicada:**")
    st.markdown("""
- ✅ **OK**: hora ya correcta en AllRide
- ✏️ **Editar**: mismo día y ruta, pero hora distinta → se edita el más cercano
- ❌ **Cancelar**: viaje en AllRide sin match en cliente
- ➕ **Crear**: viaje en cliente sin ningún AllRide ese día para esa ruta
    """)

if not f_consolidado or not f_allride or not f_cancel_tmpl:
    st.info("⬅️ Sube los tres archivos obligatorios para comenzar.")
    st.stop()

# ── Procesar ──────────────────────────────────────────────────────────────────
with st.spinner("Procesando..."):
    cli_raw = leer_consolidado(f_consolidado)
    ar_raw  = leer_allride(f_allride)

    cli = procesar_consolidado(cli_raw)
    ar  = procesar_allride(ar_raw)

    resultados = cuadrar(cli, ar)

n_ok     = len(resultados["ok"])
n_editar = len(resultados["editar"])
n_cancel = len(resultados["cancelar"])
n_crear  = len(resultados["crear"])

# ── Resumen visual ────────────────────────────────────────────────────────────
st.divider()
c1, c2, c3, c4 = st.columns(4)
c1.metric("✅ Sin cambio", n_ok)
c2.metric("✏️ Editar hora", n_editar)
c3.metric("❌ Cancelar", n_cancel)
c4.metric("➕ Crear nuevos", n_crear)

# ── Tabs detalle ──────────────────────────────────────────────────────────────
tab_ok, tab_edit, tab_cancel, tab_crear = st.tabs([
    f"✅ OK ({n_ok})",
    f"✏️ Editar ({n_editar})",
    f"❌ Cancelar ({n_cancel})",
    f"➕ Crear ({n_crear})",
])

with tab_ok:
    if resultados["ok"]:
        df_ok = pd.DataFrame([{
            "Fecha": r["fecha"], "Hora": r["hora"],
            "Ruta": r["ruta_orig"], "Empresa": r["empresa_orig"],
            "ID AllRide": r["id"]
        } for r in resultados["ok"]])
        st.dataframe(df_ok, use_container_width=True)
    else:
        st.info("No hay viajes que ya estén correctos.")

with tab_edit:
    if resultados["editar"]:
        df_edit = pd.DataFrame([{
            "Fecha": e["ar_row"]["fecha"],
            "Ruta": e["ar_row"]["ruta_orig"],
            "Empresa": e["ar_row"]["empresa_orig"],
            "ID AllRide": e["ar_row"]["id"],
            "Hora actual": e["hora_actual"],
            "→ Hora nueva": e["hora_nueva"],
            "Hora postura cliente": e["cli_row"].get("hora_postura",""),
        } for e in resultados["editar"]])
        st.dataframe(df_edit, use_container_width=True)
    else:
        st.info("No hay viajes que editar.")

with tab_cancel:
    if resultados["cancelar"]:
        df_cancel = pd.DataFrame([{
            "Fecha": r["fecha"], "Hora": r["hora"],
            "Ruta": r["ruta_orig"], "Empresa": r["empresa_orig"],
            "Tipo": r["tipo_orig"], "ID AllRide": r["id"], "Estado": r["estado"]
        } for r in resultados["cancelar"]])
        st.dataframe(df_cancel, use_container_width=True)
    else:
        st.info("No hay viajes que cancelar.")

with tab_crear:
    if resultados["crear"]:
        df_crear = pd.DataFrame([{
            "Fecha": r.get("fecha",""), "Hora postura": r.get("hora_postura",""),
            "Hora AllRide": r.get("hora_allride",""),
            "Ruta": r.get("ruta_orig",""), "Empresa": r.get("empresa_orig",""),
            "Tipo": r.get("tipo_norm","")
        } for r in resultados["crear"]])
        st.dataframe(df_crear, use_container_width=True)
    else:
        st.info("No hay viajes nuevos que crear.")

# ── Descargas ─────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📥 Descargar archivos")

cancel_tmpl_bytes = f_cancel_tmpl.read()
edit_tmpl_bytes   = f_edit_tmpl.read() if f_edit_tmpl else None

col_d1, col_d2, col_d3, col_d4 = st.columns(4)

# 1. Resumen
resumen_bytes = gen_resumen_excel(resultados)
col_d1.download_button(
    "📊 Resumen completo",
    data=resumen_bytes,
    file_name="Cuadratura_resumen.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# 2. Cancelación
if n_cancel > 0:
    cancel_bytes = gen_cancelacion(ar_raw, resultados["cancelar"], cancel_tmpl_bytes)
    col_d2.download_button(
        f"❌ Cancelación masiva ({n_cancel})",
        data=cancel_bytes,
        file_name="Cancelacion_masiva.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    col_d2.button("❌ Cancelación (0)", disabled=True, use_container_width=True)

# 3. Edición horarios
if n_editar > 0:
    if edit_tmpl_bytes:
        edit_bytes = gen_edicion_horarios(resultados["editar"], edit_tmpl_bytes)
        col_d3.download_button(
            f"✏️ Edición horarios ({n_editar})",
            data=edit_bytes,
            file_name="Edicion_horarios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        # Sin plantilla: generar tabla simple
        df_edit_dl = pd.DataFrame([{
            "ID de servicio": e["ar_row"]["id"],
            "Nueva hora": e["hora_nueva"],
            "Hora actual": e["hora_actual"],
            "Ruta": e["ar_row"]["ruta_orig"],
            "Fecha": e["ar_row"]["fecha"],
        } for e in resultados["editar"]])
        buf = BytesIO()
        df_edit_dl.to_excel(buf, index=False)
        col_d3.download_button(
            f"✏️ Edición horarios ({n_editar})",
            data=buf.getvalue(),
            file_name="Edicion_horarios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    col_d3.button("✏️ Edición (0)", disabled=True, use_container_width=True)

# 4. Crear (info solamente)
if n_crear > 0:
    df_crear_dl = pd.DataFrame([{
        "Fecha": r.get("fecha",""), "Hora postura": r.get("hora_postura",""),
        "Hora inicio AllRide": r.get("hora_allride",""),
        "Ruta": r.get("ruta_orig",""), "Empresa": r.get("empresa_orig",""),
        "Tipo": r.get("tipo_norm","")
    } for r in resultados["crear"]])
    buf = BytesIO()
    df_crear_dl.to_excel(buf, index=False)
    col_d4.download_button(
        f"➕ Viajes a crear ({n_crear})",
        data=buf.getvalue(),
        file_name="Viajes_a_crear.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    col_d4.button("➕ Crear (0)", disabled=True, use_container_width=True)

st.divider()
st.caption("AllRide Cuadratura v1.0 · Lógica: postura+15min · editar antes de crear · cancelar sobrantes")
