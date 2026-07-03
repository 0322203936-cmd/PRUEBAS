from flask import Flask, render_template, request, redirect, jsonify, make_response, send_file
import json
import os
import base64
from datetime import datetime
import requests as req_lib
import msal
import threading
from supabase import create_client, Client
import cv2
import numpy as np
import re
import google.generativeai as genai

# Configurar Gemini API
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))

# Intentamos importar pytesseract
try:
    import pytesseract
except ImportError:
    print("pytesseract no está instalado.")
except Exception as e:
    print(f"Error inicializando pytesseract: {e}")

# ============================================================
# MAPEO DE NOMBRES: Acuse Walmart → Nombre en Base de Datos
# Columna izquierda = como aparece en Supabase
# Columna derecha = como aparece en la hoja acuse de Walmart
# ============================================================
ACUSE_A_DB = {
    # Clave: Nombre en la hoja de Acuse (en minúsculas) -> Valor: Nombre en Supabase
    "bqt lili asiatic 6t":       "BQT LILI ASIATIC 6 T",
    "bqt 18 rosas":              "BQT ROSA 18 TALLOS",
    "bqt mday premium":          "BQT MIXTO PREMIUM",
    "rosas 12 mday":             "ROSA 12 TALLOS",
    "jarron mday":               "JARRON MDAY",
    "bqt mday m":                "BQT MIXTO M",
    "bqt rosas 12t":             "BOUQUET ROSAS 12 T",
    "bqt rosas 6t":              "BOUQUET ROSAS 6 T",
    "bqt snapdragon 8t":         "BQT SNAPDRAGON 8 T",
    "bqt rosas 12t baja":        "BOUQUETS DOCENA DE ROSAS",
    "bqt mixto 9t":              "BOUQUET MIXTO 9 T",
    "bqt mixto 12t":             "BOUQUET MIXTO 12 T",
    "bqt mixto 15t":             "BOUQUET MIXTO 15 T",
}

# Mapa inverso: DB name → lista de posibles nombres en acuse (para búsqueda)
DB_A_ACUSE = {}
for acuse_name, db_name in ACUSE_A_DB.items():
    DB_A_ACUSE.setdefault(db_name.upper(), []).append(acuse_name)
# ============================================================

def _corregir_rotacion_exif(img_bytes_raw):
    """Corrige rotación usando metadatos EXIF de la foto."""
    try:
        from PIL import Image as PILImage
        import io
        pil_img = PILImage.open(io.BytesIO(img_bytes_raw))
        exif = pil_img._getexif() if hasattr(pil_img, '_getexif') else None
        orientacion = None
        if exif:
            for tag, val in exif.items():
                import PIL.ExifTags
                if PIL.ExifTags.TAGS.get(tag) == 'Orientation':
                    orientacion = val
                    break
        rotaciones = {3: 180, 6: 270, 8: 90}
        if orientacion in rotaciones:
            pil_img = pil_img.rotate(rotaciones[orientacion], expand=True)
        buf = io.BytesIO()
        pil_img.save(buf, format='JPEG')
        return np.frombuffer(buf.getvalue(), np.uint8), True
    except Exception:
        return None, False

def mejorar_imagen_opencv(img_bytes, img_bytes_raw=None):
    """Preprocesa la imagen usando OpenCV para mejorar el OCR.
    Incluye corrección automática de rotación (EXIF + auto-detección)."""

    # 1. Intentar corrección EXIF primero
    if img_bytes_raw is not None:
        corregido, ok = _corregir_rotacion_exif(img_bytes_raw)
        if ok and corregido is not None:
            img_bytes = corregido

    # Solo decodificar y devolver, sin filtros destructivos que realzan la tinta de atrás
    img = cv2.imdecode(img_bytes, cv2.IMREAD_COLOR)
    return img




app = Flask(__name__)

DELETE_PASSWORD = "CFBCWALMEX"

# Supabase Config
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
supabase_client: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# ── SharePoint / Excel config ─────────────────────────────────────────────────
SP_TENANT_ID     = os.environ.get("SP_TENANT_ID",     "")
SP_CLIENT_ID     = os.environ.get("SP_CLIENT_ID",     "")
SP_CLIENT_SECRET = os.environ.get("SP_CLIENT_SECRET", "")
SP_SITE_URL      = os.environ.get("SP_SITE_URL",      "https://pacificafarms.sharepoint.com/sites/requerimientovsproyeccion")
SP_FILE_PATH     = os.environ.get("SP_FILE_PATH",     "/requerimiento vs proyeccion/WALMEX/Analisis Walmart.xlsx")
SP_SHEET_DETALLE  = os.environ.get("SP_SHEET_NAME", "Detalle")
SP_SHEET_GASTOS   = "REPORTE-GASTOSAPP"

HEADERS_DETALLE = [
    "Fecha de registro", "Tienda", "Fecha", "Usuario",
    "Producto", "Inventario", "Merma", "Razon de merma"
]

HEADERS_CF = [
    "Fecha de registro", "Tienda", "Fecha", "Usuario",
    "Producto", "Existencia"
]

HEADERS_GASTOS = [
    "Fecha de registro", "Tienda", "Fecha del Gasto", "Usuario",
    "Categoria", "Monto", "Fotos", "Viaticos", "Comentarios"
]

# Tabla Detalle: columnas A-H (col 1-8)
# Separador:    columna I (9) vacia
# Tabla CF:     columnas J-O (col 10-15)
COL_DETALLE_START = 1   # A
COL_CF_START      = 10  # J

# Colores en hex para Graph API (sin #)
COLOR_HEADER_DETALLE = "1A73E8"
COLOR_TEXT_HEADER    = "FFFFFF"
COLOR_HEADER_CF      = "0D9488"
COLOR_HEADER_GASTOS  = "E67E22"
COLOR_ROW_ALT        = "EBF3FD"
COLOR_ROW_ALT_CF     = "F0FDFA"


import time

_TOKEN_CACHE = None
_TOKEN_EXPIRY = 0
_SITE_ID_CACHE = None
_DOWNLOAD_URL_CACHE = {}

def _get_sp_token():
    global _TOKEN_CACHE, _TOKEN_EXPIRY
    if _TOKEN_CACHE and time.time() < _TOKEN_EXPIRY:
        return _TOKEN_CACHE
        
    msal_app = msal.ConfidentialClientApplication(
        SP_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{SP_TENANT_ID}",
        client_credential=SP_CLIENT_SECRET,
    )
    result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    
    token = result.get("access_token")
    if token:
        _TOKEN_CACHE = token
        _TOKEN_EXPIRY = time.time() + 3000 # Cache for 50 mins
    return token


def _get_site_id(headers):
    global _SITE_ID_CACHE
    if _SITE_ID_CACHE:
        return _SITE_ID_CACHE
        
    parts     = SP_SITE_URL.rstrip("/").split("/")
    hostname  = parts[2]
    site_path = "/".join(parts[3:])
    r = req_lib.get(
        f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}",
        headers=headers, timeout=30
    )
    r.raise_for_status()
    _SITE_ID_CACHE = r.json()["id"]
    return _SITE_ID_CACHE


def _get_base_url(site_id):
    return (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:{SP_FILE_PATH}:"
    )


def _fmt_fecha_excel(fecha_str):
    """Convierte fecha de DD/MM/YYYY a MM/DD/YY para el Excel."""
    try:
        return datetime.strptime(fecha_str, "%d/%m/%Y").strftime("%m/%d/%y")
    except Exception:
        return fecha_str  # si falla, deja el valor original


def _col_letter(n):
    """Convierte número de columna (1-based) a letra(s). Ej: 1->A, 27->AA"""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _ensure_sheet_exists(headers, base_url, sheet_name):
    """Crea la hoja si no existe."""
    r = req_lib.get(f"{base_url}/workbook/worksheets", headers=headers, timeout=30)
    if r.ok:
        names = [s.get("name", "") for s in r.json().get("value", [])]
        if sheet_name not in names:
            req_lib.post(
                f"{base_url}/workbook/worksheets",
                headers={**headers, "Content-Type": "application/json"},
                json={"name": sheet_name}, timeout=30
            )


def _format_range(headers_auth, base_url, address, bg_color, bold=False,
                  font_color="000000", font_size=10, sheet_name=SP_SHEET_DETALLE):
    """Aplica formato de relleno y fuente a un rango dado."""
    fmt_url = (
        f"{base_url}/workbook/worksheets/{sheet_name}"
        f"/range(address='{address}')/format"
    )
    req_lib.patch(fmt_url + "/fill",
        headers={**headers_auth, "Content-Type": "application/json"},
        json={"color": bg_color}, timeout=30)
    req_lib.patch(fmt_url + "/font",
        headers={**headers_auth, "Content-Type": "application/json"},
        json={"bold": bold, "color": font_color, "size": font_size}, timeout=30)


def _ensure_table_headers(headers_auth, base_url, col_start, col_headers, bg_color, sheet_name=SP_SHEET_DETALLE):
    """
    Verifica y escribe encabezados en la fila 1 a partir de col_start.
    También aplica formato a esa fila de encabezados.
    """
    col_end      = col_start + len(col_headers) - 1
    start_letter = _col_letter(col_start)
    end_letter   = _col_letter(col_end)
    address      = f"{start_letter}1:{end_letter}1"
    range_url    = (
        f"{base_url}/workbook/worksheets/{sheet_name}"
        f"/range(address='{address}')"
    )

    r = req_lib.get(range_url, headers=headers_auth, timeout=30)
    needs_write = True
    if r.ok:
        values = r.json().get("values", [[]])
        row = values[0] if values else []
        if row and all(str(row[i]).strip() == col_headers[i]
                       for i in range(len(col_headers)) if i < len(row)):
            needs_write = False

    if needs_write:
        req_lib.patch(range_url,
            headers={**headers_auth, "Content-Type": "application/json"},
            json={"values": [col_headers]}, timeout=30)
        _format_range(headers_auth, base_url, address,
                      bg_color=bg_color, bold=True,
                      font_color=COLOR_TEXT_HEADER, font_size=11, sheet_name=sheet_name)


def _find_next_empty_row_col(headers_auth, base_url, col_start, sheet_name=SP_SHEET_DETALLE):
    """
    Busca la primera fila vacia en la columna col_start (1-based),
    leyendo celda a celda para ignorar filas borradas.
    """
    used_url = f"{base_url}/workbook/worksheets/{sheet_name}/usedRange"
    r = req_lib.get(used_url, headers=headers_auth, timeout=30)
    if not r.ok:
        return 2

    row_count = r.json().get("rowCount", 1)
    if row_count <= 1:
        return 2

    col_letter = _col_letter(col_start)
    col_url = (
        f"{base_url}/workbook/worksheets/{sheet_name}"
        f"/range(address='{col_letter}1:{col_letter}{row_count}')"
    )
    r2 = req_lib.get(col_url, headers=headers_auth, timeout=30)
    if not r2.ok:
        return row_count + 1

    values = r2.json().get("values", [])
    last_row_with_data = 1
    for i, cell in enumerate(values):
        if cell and str(cell[0]).strip():
            last_row_with_data = i + 1

    return last_row_with_data + 1


def escribir_en_excel(filas_detalle, filas_cf):
    """
    Escribe ambas tablas en la hoja Detalle:
      - Tabla Merma/Inventario: columnas A-H (col 1-8)  encabezado azul
      - Separador:              columna I (9) vacia
      - Tabla Cuarto Frio:      columnas J-O (col 10-15) encabezado teal
    """
    try:
        token = _get_sp_token()
        if not token:
            print("[SP] No se pudo obtener token.")
            return

        auth_headers = {"Authorization": f"Bearer {token}"}
        site_id      = _get_site_id(auth_headers)
        base_url     = _get_base_url(site_id)

        _ensure_sheet_exists(auth_headers, base_url, SP_SHEET_DETALLE)

        # ── Tabla Merma / Inventario (columnas A-H) ───────────────────────
        if filas_detalle:
            _ensure_table_headers(auth_headers, base_url,
                                  COL_DETALLE_START, HEADERS_DETALLE,
                                  COLOR_HEADER_DETALLE)
            next_row = _find_next_empty_row_col(auth_headers, base_url, COL_DETALLE_START)
            n_cols   = len(HEADERS_DETALLE)
            s_col    = _col_letter(COL_DETALLE_START)
            e_col    = _col_letter(COL_DETALLE_START + n_cols - 1)
            end_row  = next_row + len(filas_detalle) - 1
            address  = f"{s_col}{next_row}:{e_col}{end_row}"

            resp = req_lib.patch(
                f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/range(address='{address}')",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"values": filas_detalle}, timeout=30
            )
            if resp.ok:
                for i in range(len(filas_detalle)):
                    row_idx = next_row + i
                    if row_idx % 2 == 0:
                        _format_range(auth_headers, base_url,
                                      f"{s_col}{row_idx}:{e_col}{row_idx}",
                                      bg_color=COLOR_ROW_ALT, font_size=10)
            else:
                print(f"[SP] Error Detalle: {resp.status_code} {resp.text[:200]}")

        # ── Tabla Cuarto Frio (columnas J-O) ─────────────────────────────
        if filas_cf:
            _ensure_table_headers(auth_headers, base_url,
                                  COL_CF_START, HEADERS_CF,
                                  COLOR_HEADER_CF)
            next_row = _find_next_empty_row_col(auth_headers, base_url, COL_CF_START)
            n_cols   = len(HEADERS_CF)
            s_col    = _col_letter(COL_CF_START)
            e_col    = _col_letter(COL_CF_START + n_cols - 1)
            end_row  = next_row + len(filas_cf) - 1
            address  = f"{s_col}{next_row}:{e_col}{end_row}"

            resp = req_lib.patch(
                f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/range(address='{address}')",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"values": filas_cf}, timeout=30
            )
            if resp.ok:
                for i in range(len(filas_cf)):
                    row_idx = next_row + i
                    if row_idx % 2 == 0:
                        _format_range(auth_headers, base_url,
                                      f"{s_col}{row_idx}:{e_col}{row_idx}",
                                      bg_color=COLOR_ROW_ALT_CF, font_size=10)
            else:
                print(f"[SP] Error CuartoFrio: {resp.status_code} {resp.text[:200]}")

    except Exception as e:
        print(f"[SP] Excepcion: {e}")

def escribir_gasto_en_excel(filas_gastos):
    """Escribe registros en la hoja Gastos."""
    if not filas_gastos:
        return
    try:
        token = _get_sp_token()
        if not token:
            print("[SP Gastos] No se pudo obtener token.")
            return

        auth_headers = {"Authorization": f"Bearer {token}"}
        site_id      = _get_site_id(auth_headers)
        base_url     = _get_base_url(site_id)

        _ensure_sheet_exists(auth_headers, base_url, SP_SHEET_GASTOS)
        
        _ensure_table_headers(auth_headers, base_url, 1, HEADERS_GASTOS, COLOR_HEADER_GASTOS, sheet_name=SP_SHEET_GASTOS)
        
        next_row = _find_next_empty_row_col(auth_headers, base_url, 1, sheet_name=SP_SHEET_GASTOS)
        n_cols   = len(HEADERS_GASTOS)
        s_col    = "A"
        e_col    = _col_letter(n_cols)
        end_row  = next_row + len(filas_gastos) - 1
        address  = f"{s_col}{next_row}:{e_col}{end_row}"

        resp = req_lib.patch(
            f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/range(address='{address}')",
            headers={**auth_headers, "Content-Type": "application/json"},
            json={"values": filas_gastos}, timeout=30
        )
        if resp.ok:
            for i in range(len(filas_gastos)):
                row_idx = next_row + i
                if row_idx % 2 == 0:
                    _format_range(auth_headers, base_url,
                                  f"{s_col}{row_idx}:{e_col}{row_idx}",
                                  bg_color=COLOR_ROW_ALT, font_size=10, sheet_name=SP_SHEET_GASTOS)
        else:
            print(f"[SP Gastos] Error: {resp.status_code} {resp.text[:200]}")

    except Exception as e:
        print(f"[SP Gastos] Excepcion: {e}")


# ── Base de datos eliminada ───────────────────────────────────────────────────

TIENDAS = [
    "SC MEXICALI NOVENA","SC NUEVO MEXICALI","SC PLAZA SAN PEDRO",
    "SC MEXICALI","SC PLAYAS DE TIJUANA","SC LOMAS DE SANTA FE",
    "SC GALERIAS DEL VALLE","SC TIJUANA 2000","SC TECATE GARITA",
    "SC ROSARITO","SC ENSENADA CENTRO","SC MACROPLAZA INSURGENTES",
    "SC ENSENADA","SC TIJUANA HIPODROMO","SC PACIFICO",
    "SC DIAZ ORDAZ"
]


# ── Rutas ─────────────────────────────────────────────

@app.route("/sw.js")
def service_worker_root():
    """
    Sirve el Service Worker desde la raiz (/) para que su scope cubra
    toda la aplicacion y pueda interceptar todas las rutas sin internet.
    Sin este header, el navegador limita el scope al directorio /static/.
    """
    sw_path = os.path.join(app.root_path, 'static', 'service-worker.js')
    with open(sw_path, 'r', encoding='utf-8') as f:
        content = f.read()
    resp = make_response(content)
    resp.headers['Content-Type']           = 'application/javascript; charset=utf-8'
    resp.headers['Service-Worker-Allowed'] = '/'
    resp.headers['Cache-Control']          = 'no-cache, no-store, must-revalidate'
    return resp


@app.route("/")
def home():
    """Pantalla de inicio con los dos modulos: Inventario y Gastos."""
    return render_template("home.html")

@app.route("/inventario", methods=["GET", "POST"])
def index():
    try:
        if request.method == "POST":
            tienda      = request.form.get("tienda")
            fecha       = request.form.get("fecha")
            usuario     = request.form.get("usuario")
            productos   = request.form.getlist("producto[]")
            inventarios = request.form.getlist("inventario[]")
            mermas      = request.form.getlist("merma[]")
            razones     = request.form.getlist("razon[]")
            fecha_reg   = datetime.now().strftime("%d/%m/%Y %H:%M")

            filas_detalle = []
            filas_cf      = []

            # Merma / Inventario
            for i in range(len(productos)):
                if not productos[i].strip():
                    continue
                try:
                    inv = int(inventarios[i]) if inventarios[i] else 0
                except ValueError:
                    inv = 0
                try:
                    mer = int(mermas[i]) if mermas[i] else 0
                except ValueError:
                    mer = 0

                if inv > 0 or mer > 0:
                    razon = razones[i] if i < len(razones) else ""
                    filas_detalle.append([
                        fecha_reg, tienda, _fmt_fecha_excel(fecha), usuario,
                        productos[i], inv, mer, razon
                    ])

            # Cuarto Frio
            cf_productos   = request.form.getlist("cf_producto[]")
            cf_existencias = request.form.getlist("cf_existencia[]")

            for i in range(len(cf_productos)):
                try:
                    existencia = int(cf_existencias[i]) if cf_existencias[i] else 0
                except ValueError:
                    existencia = 0

                if existencia > 0:
                    filas_cf.append([
                        fecha_reg, tienda, _fmt_fecha_excel(fecha), usuario,
                        cf_productos[i], existencia
                    ])

            if filas_detalle or filas_cf:
                t = threading.Thread(
                    target=escribir_en_excel,
                    args=(filas_detalle, filas_cf),
                    daemon=True
                )
                t.start()

            return redirect("/inventario?success=1")

        today = datetime.now().strftime("%d/%m/%Y")
        resp = make_response(render_template("index.html", tiendas=TIENDAS, today=today))
        resp.headers['Cache-Control'] = 'no-cache'
        return resp

    except Exception as e:
        return f"<h2>Error en la aplicacion:</h2><pre>{e}</pre>"


@app.route("/gastos")
def gastos():
    """Pantalla de captura de gastos (tickets por categoria)."""
    resp = make_response(render_template("gastos.html"))
    resp.headers['Cache-Control'] = 'no-cache'
    return resp

@app.route("/facturacion")
def facturacion():
    """Pantalla de facturación."""
    resp = make_response(render_template("facturacion.html"))
    resp.headers['Cache-Control'] = 'no-cache'
    return resp


def leer_desde_excel():
    token = _get_sp_token()
    if not token:
        return [], []
    auth_headers = {"Authorization": f"Bearer {token}"}
    site_id = _get_site_id(auth_headers)
    base_url = _get_base_url(site_id)
    
    used_url = f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/usedRange"
    r = req_lib.get(used_url, headers=auth_headers, timeout=30)
    if not r.ok:
        return [], []
    
    values = r.json().get("values", [])
    if not values:
        return [], []
    
    merma_rows = []
    cf_rows = []
    
    for i, row in enumerate(values):
        if i == 0: continue
        row_id = i + 1
        
        # Merma (Cols A-H -> index 0-7)
        if len(row) > 4 and str(row[1]).strip() and str(row[4]).strip():
            tienda = row[1] if len(row) > 1 else ""
            fecha = row[2] if len(row) > 2 else ""
            usuario = row[3] if len(row) > 3 else ""
            producto = row[4] if len(row) > 4 else ""
            inv = row[5] if len(row) > 5 else 0
            merma = row[6] if len(row) > 6 else 0
            razon = row[7] if len(row) > 7 else ""
            merma_rows.append([row_id, tienda, fecha, usuario, producto, inv, merma, razon, ""])
            
        # CF (Cols J-O -> index 9-14)
        if len(row) > 13 and str(row[10]).strip() and str(row[13]).strip():
            tienda_cf = row[10] if len(row) > 10 else ""
            fecha_cf = row[11] if len(row) > 11 else ""
            usuario_cf = row[12] if len(row) > 12 else ""
            producto_cf = row[13] if len(row) > 13 else ""
            existencia = row[14] if len(row) > 14 else 0
            cf_rows.append([row_id, tienda_cf, fecha_cf, usuario_cf, producto_cf, existencia, ""])

    merma_rows.reverse()
    cf_rows.reverse()
    return merma_rows, cf_rows


@app.route("/registros")
def registros():
    try:
        merma_rows, cf_rows = leer_desde_excel()
        return render_template("registros.html", registros=merma_rows, cf_registros=cf_rows)
    except Exception as e:
        return f"<h2>Error:</h2><pre>{e}</pre>"


@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):
    try:
        merma_rows, _ = leer_desde_excel()
        reg = next((r for r in merma_rows if r[0] == id), None)
        
        if request.method == "POST":
            tienda     = request.form.get("tienda")
            fecha      = request.form.get("fecha")
            usuario    = request.form.get("usuario")
            producto   = request.form.get("producto")
            inventario = request.form.get("inventario") or 0
            merma      = request.form.get("merma") or 0
            razon      = request.form.get("razon") or ""
            
            token = _get_sp_token()
            if token:
                auth_headers = {"Authorization": f"Bearer {token}"}
                site_id = _get_site_id(auth_headers)
                base_url = _get_base_url(site_id)
                address = f"B{id}:H{id}"
                valores = [[tienda, _fmt_fecha_excel(fecha), usuario, producto, inventario, merma, razon]]
                req_lib.patch(
                    f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/range(address='{address}')",
                    headers={**auth_headers, "Content-Type": "application/json"},
                    json={"values": valores}, timeout=30
                )
            return redirect("/registros")

        if not reg:
            return redirect("/registros")
        return render_template("editar.html", reg=reg, tiendas=TIENDAS)
    except Exception as e:
        return f"<h2>Error:</h2><pre>{e}</pre>"


@app.route("/borrar/<int:id>", methods=["POST"])
def borrar(id):
    password = request.form.get("password")
    if password != DELETE_PASSWORD:
        return jsonify({"ok": False, "msg": "Contrasena incorrecta"}), 403
    try:
        token = _get_sp_token()
        if token:
            auth_headers = {"Authorization": f"Bearer {token}"}
            site_id = _get_site_id(auth_headers)
            base_url = _get_base_url(site_id)
            req_lib.post(
                f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/range(address='A{id}:H{id}')/delete",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"shift": "Up"}, timeout=30
            )
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/borrar_cf/<int:id>", methods=["POST"])
def borrar_cf(id):
    password = request.form.get("password")
    if password != DELETE_PASSWORD:
        return jsonify({"ok": False, "msg": "Contrasena incorrecta"}), 403
    try:
        token = _get_sp_token()
        if token:
            auth_headers = {"Authorization": f"Bearer {token}"}
            site_id = _get_site_id(auth_headers)
            base_url = _get_base_url(site_id)
            req_lib.post(
                f"{base_url}/workbook/worksheets/{SP_SHEET_DETALLE}/range(address='J{id}:O{id}')/delete",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"shift": "Up"}, timeout=30
            )
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


# ── Endpoints para soporte offline ──────────────────────────────────────────

@app.route("/ping")
def ping():
    """Endpoint liviano para verificar conectividad desde el cliente."""
    resp = make_response(jsonify({"ok": True}), 200)
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp


# ── SharePoint: subida de fotos ───────────────────────────────────────────────

SP_GASTOS_FOLDER = "/requerimiento vs proyeccion/WALMEX/Gastos"

def subir_foto_supabase(imagen_base64, ruta_destino):
    """
    Sube una imagen (base64) a Supabase Storage y retorna la URL publica.
    """
    try:
        if ',' in imagen_base64:
            imagen_base64 = imagen_base64.split(',', 1)[1]
        img_bytes = base64.b64decode(imagen_base64)
        
        # Subir a supabase (reemplaza si existe)
        supabase_client.storage.from_("gastos-fotos").upload(
            ruta_destino, 
            img_bytes, 
            file_options={"content-type": "image/jpeg", "upsert": "true"}
        )
        
        # Obtener URL publica
        res = supabase_client.storage.from_("gastos-fotos").get_public_url(ruta_destino)
        return res
    except Exception as e:
        print(f"[SUPABASE] Error subiendo foto: {e}")
        return None

def subir_foto_sharepoint(imagen_base64, ruta_destino, auth_headers, base_url):
    """
    Sube una imagen (base64) a SharePoint via Graph API.
    ruta_destino: ej. 'Gastos/2025-06/CASETAS/Mizael_20250617_083045.jpg'
    """
    # Decodificar base64 (puede venir como data:image/jpeg;base64,...)
    if ',' in imagen_base64:
        imagen_base64 = imagen_base64.split(',', 1)[1]
    img_bytes = base64.b64decode(imagen_base64)

    # Construir URL de subida en el drive del sitio
    site_parts    = SP_SITE_URL.rstrip("/").split("/")
    sp_hostname   = site_parts[2]
    sp_site_path  = "/".join(site_parts[3:])

    token    = auth_headers["Authorization"].replace("Bearer ", "")
    site_url = f"https://graph.microsoft.com/v1.0/sites/{sp_hostname}:/{sp_site_path}"
    r = req_lib.get(site_url, headers=auth_headers, timeout=30)
    r.raise_for_status()
    site_id = r.json()["id"]

    upload_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{ruta_destino}:/content"
    )
    resp = req_lib.put(
        upload_url,
        headers={**auth_headers, "Content-Type": "image/jpeg"},
        data=img_bytes,
        timeout=60
    )
    return resp.ok


def procesar_gastos(pendiente):
    """
    Sube las fotos de un registro de gastos a SharePoint.
    Se ejecuta en un hilo separado.
    """
    try:
        token = _get_sp_token()
        if not token:
            print("[GASTOS] No se pudo obtener token.")
            return

        auth_headers = {"Authorization": f"Bearer {token}"}
        site_id  = _get_site_id(auth_headers)
        base_url = _get_base_url(site_id)

        tienda   = pendiente.get("tienda", "SinTienda").replace(" ", "_")
        usuario  = pendiente.get("usuario", "SinUsuario")
        fecha    = pendiente.get("fecha", "").replace("/", "-")  # DD-MM-YYYY
        fecha_reg = datetime.now()
        timestamp = fecha_reg.strftime("%Y%m%d_%H%M%S")
        mes_folder = fecha_reg.strftime("%Y-%m")

        categorias = ["casetas", "comida", "otros"]
        filas_gastos = []
        for cat in categorias:
            cat_data = pendiente.get(cat, {})
            fotos = cat_data.get("fotos", [])
            monto = cat_data.get("monto", 0)
            comentario = cat_data.get("comentario", "")
            if not fotos and monto == 0 and not comentario:
                continue
                
            rutas_fotos = []
            for i, foto_b64 in enumerate(fotos):
                nombre_archivo = f"{tienda}_{usuario}_{fecha}_{timestamp}_{i+1}.jpg"
                ruta_supa = f"{mes_folder}/{cat.upper()}/{nombre_archivo}"
                
                # Subir a Supabase
                url_publica = subir_foto_supabase(foto_b64, ruta_supa)
                if url_publica:
                    rutas_fotos.append(url_publica)
                    print(f"[GASTOS] Subida a Supabase: {url_publica}")
                else:
                    print(f"[GASTOS] Error al subir a Supabase: {ruta_supa}")
                    
            filas_gastos.append([
                fecha_reg.strftime("%d/%m/%Y %H:%M"),
                tienda.replace("_", " "),
                pendiente.get("fecha", ""),
                usuario,
                cat.upper(),
                monto,
                ",".join(rutas_fotos),
                "",
                comentario
            ])
            
        if filas_gastos:
            escribir_gasto_en_excel(filas_gastos)

    except Exception as e:
        print(f"[GASTOS] Excepcion: {e}")


@app.route("/gastos/sync", methods=["POST"])
def gastos_sync():
    """
    Recibe registros de gastos (fotos en base64) y los sube a SharePoint.
    Body JSON: { "pendientes": [ { tipo, tienda, usuario, fecha, casetas, comida, otros }, ... ] }
    """
    try:
        data = request.get_json(force=True)
        if not data or "pendientes" not in data:
            return jsonify({"ok": False, "msg": "Formato invalido"}), 400

        pendientes = data["pendientes"]
        if not pendientes:
            return jsonify({"ok": True, "sincronizados": 0})

        for p in pendientes:
            t = threading.Thread(target=procesar_gastos, args=(p,), daemon=True)
            t.start()

        return jsonify({"ok": True, "sincronizados": len(pendientes)})

    except Exception as e:
        print(f"[GASTOS SYNC] Error: {e}")
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/sync", methods=["POST"])
def sync():
    """
    Recibe registros capturados offline (almacenados en IndexedDB del navegador)
    y los escribe en SharePoint igual que la ruta principal.

    Formato esperado del body JSON:
    {
        "pendientes": [
            {
                "tipo": "form",          // Un envio de formulario completo
                "tienda": "SC MEXICALI",
                "fecha": "06/17/25",
                "usuario": "Mizael",
                "fecha_reg": "17/06/2025 08:30",
                "filas_detalle": [[fecha_reg, tienda, fecha, usuario, producto, inv, merma, razon], ...],
                "filas_cf": [[fecha_reg, tienda, fecha, usuario, producto, existencia], ...]
            },
            ...
        ]
    }
    """
    try:
        data = request.get_json(force=True)
        if not data or "pendientes" not in data:
            return jsonify({"ok": False, "msg": "Formato invalido"}), 400

        pendientes = data["pendientes"]
        if not pendientes:
            return jsonify({"ok": True, "sincronizados": 0})

        all_detalle = []
        all_cf      = []

        for p in pendientes:
            filas_d = p.get("filas_detalle", [])
            filas_c = p.get("filas_cf", [])
            if filas_d:
                all_detalle.extend(filas_d)
            if filas_c:
                all_cf.extend(filas_c)

        if all_detalle or all_cf:
            # Escribir en segundo plano igual que el envio normal
            t = threading.Thread(
                target=escribir_en_excel,
                args=(all_detalle, all_cf),
                daemon=True
            )
            t.start()

        return jsonify({"ok": True, "sincronizados": len(pendientes)})

    except Exception as e:
        print(f"[SYNC] Error: {e}")
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/foto")
def api_foto():
    """
    Descarga la imagen desde SharePoint (server-side) y la devuelve al browser.
    Evita problemas de CORS/expiración con las URLs pre-autenticadas de Microsoft Graph.
    """
    ruta = request.args.get("path")
    if not ruta:
        return "Ruta no proporcionada", 400

    token = _get_sp_token()
    if not token:
        return "No autorizado", 401

    auth_headers = {"Authorization": f"Bearer {token}"}
    try:
        site_id   = _get_site_id(auth_headers)
        ruta_limpia = ruta.lstrip("/")
        
        import time
        global _DOWNLOAD_URL_CACHE
        cached = _DOWNLOAD_URL_CACHE.get(ruta_limpia)
        
        if cached and cached['expiry'] > time.time():
            download_url = cached['url']
        else:
            meta_url  = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_limpia}"
            
            # Retry mechanism para SharePoint eventual consistency o throttling
            r = None
            for attempt in range(3):
                r = req_lib.get(meta_url, headers=auth_headers, timeout=15)
                if r.ok:
                    break
                if r.status_code in (404, 429, 503) and attempt < 2:
                    time.sleep(1.5)
                    continue
                    
            if not r or not r.ok:
                print(f"[FOTO] Metadata error {r.status_code if r else 'NA'}: {ruta_limpia}")
                return "Imagen no encontrada o error en SharePoint", 404

            download_url = r.json().get("@microsoft.graph.downloadUrl")
            if not download_url:
                return "Download URL no encontrada", 404
                
            _DOWNLOAD_URL_CACHE[ruta_limpia] = {
                'url': download_url,
                'expiry': time.time() + 3000
            }

        # Descargar la imagen en el servidor y enviarla directamente al browser
        img = None
        for attempt in range(3):
            img = req_lib.get(download_url, timeout=30)
            if img.ok:
                break
            if attempt < 2:
                time.sleep(1.5)
                
        if not img or not img.ok:
            print(f"[FOTO] Download error {img.status_code if img else 'NA'}: {ruta_limpia}")
            return "Error al descargar imagen desde Microsoft", 502

        content_type = img.headers.get("Content-Type", "image/jpeg")
        resp = make_response(img.content)
        resp.headers["Content-Type"]  = content_type
        resp.headers["Cache-Control"] = "public, max-age=3600"
        return resp

    except Exception as e:
        print(f"[FOTO] Excepcion: {e}")
        return str(e), 500


@app.route("/reporte")
def reporte():
    """Muestra el reporte de gastos leyendo la hoja Gastos."""
    try:
        token = _get_sp_token()
        gastos = []
        if token:
            auth_headers = {"Authorization": f"Bearer {token}"}
            site_id = _get_site_id(auth_headers)
            base_url = _get_base_url(site_id)
            
            used_url = f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/usedRange"
            r = req_lib.get(used_url, headers=auth_headers, timeout=30)
            if r.ok:
                values = r.json().get("values", [])
                if len(values) > 1:
                    # El índice de valores asume:
                    # 0: Fecha reg, 1: Tienda, 2: Fecha gasto, 3: Usuario, 4: Categoria, 5: Monto, 6: Fotos
                    grouped = {}
                    for idx, row in enumerate(values[1:]):
                        if len(row) >= 6:
                            row_num = idx + 2
                            tienda = row[1]
                            fecha = row[2]
                            usuario = row[3]
                            categoria = row[4]
                            
                            try:
                                monto = float(str(row[5]).replace('$', '').replace(',', '').strip()) if row[5] else 0.0
                            except ValueError:
                                monto = 0.0
                                
                            try:
                                viaticos_val = float(str(row[7]).replace('$', '').replace(',', '').strip()) if len(row) > 7 and row[7] else None
                            except ValueError:
                                viaticos_val = None
                                
                            comentario_str = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                            
                            fotos_str = row[6] if len(row) > 6 and row[6] else ""
                            fotos_list = [f.strip() for f in fotos_str.split(",") if f.strip()]
                            
                            key = (tienda, fecha, usuario)
                            
                            if key not in grouped:
                                grouped[key] = {
                                    "fecha_reg": row[0],
                                    "tienda": tienda,
                                    "fecha": fecha,
                                    "usuario": usuario,
                                    "categoria": [categoria] if categoria else [],
                                    "monto": monto,
                                    "fotos": list(fotos_list),
                                    "viaticos": viaticos_val,
                                    "comentarios": [comentario_str] if comentario_str else [],
                                    "detalles": {},
                                    "row_nums": [row_num]
                                }
                                if categoria:
                                    grouped[key]["detalles"][categoria] = {
                                        "monto": monto, 
                                        "fotos": list(fotos_list), 
                                        "comentarios": [comentario_str] if comentario_str else [],
                                        "row_nums": [row_num]
                                    }
                            else:
                                grouped[key]["monto"] += monto
                                grouped[key]["fotos"].extend(fotos_list)
                                grouped[key]["row_nums"].append(row_num)
                                if comentario_str:
                                    grouped[key]["comentarios"].append(comentario_str)
                                if viaticos_val is not None:
                                    grouped[key]["viaticos"] = viaticos_val
                                if categoria and categoria not in grouped[key]["categoria"]:
                                    grouped[key]["categoria"].append(categoria)
                                
                                if categoria:
                                    if categoria not in grouped[key]["detalles"]:
                                        grouped[key]["detalles"][categoria] = {"monto": 0.0, "fotos": [], "comentarios": [], "row_nums": []}
                                    grouped[key]["detalles"][categoria]["monto"] += monto
                                    grouped[key]["detalles"][categoria]["fotos"].extend(fotos_list)
                                    grouped[key]["detalles"][categoria]["row_nums"].append(row_num)
                                    if comentario_str:
                                        grouped[key]["detalles"][categoria]["comentarios"].append(comentario_str)

                                grouped[key]["fecha_reg"] = row[0] # Mostrar última fecha de actualización
                                
                    gastos = list(grouped.values())
                    gastos.reverse() # Mostrar los grupos más recientes primero
                            
        resp = make_response(render_template("reporte.html", gastos=gastos, tiendas=TIENDAS))
        resp.headers['Cache-Control'] = 'no-cache'
        return resp
    except Exception as e:
        return f"<h2>Error cargando reporte:</h2><pre>{e}</pre>"


@app.route("/api/editar_gasto", methods=["POST"])
def api_editar_gasto():
    """Actualiza Monto por categoría y Viáticos globales de un grupo de gastos."""
    data = request.json
    pwd = data.get("password")
    if pwd != "cfbc2026":
        return jsonify({"ok": False, "msg": "Contraseña incorrecta."}), 403
        
    categorias = data.get("categorias", {})
    nuevo_viatico = data.get("viaticos", 0)
    
    if not categorias:
        return jsonify({"ok": False, "msg": "No hay datos para editar."}), 400
        
    try:
        token = _get_sp_token()
        if not token:
            return jsonify({"ok": False, "msg": "Error de token SP."}), 500
            
        auth_headers = {"Authorization": f"Bearer {token}"}
        site_id = _get_site_id(auth_headers)
        base_url = _get_base_url(site_id)
        
        # Encontrar la primera fila global para guardar los viáticos
        todas_filas = []
        for cat_data in categorias.values():
            todas_filas.extend(cat_data.get("row_nums", []))
            
        if not todas_filas:
            return jsonify({"ok": False, "msg": "No hay filas para editar."}), 400
            
        primera_fila = min(todas_filas)
        
        # Actualizar viáticos (Columna H) solo si viene en el payload
        if "viaticos" in data:
            nuevo_viatico = data["viaticos"]
            resp_v = req_lib.patch(
                f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/range(address='H{primera_fila}')",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"values": [[nuevo_viatico]]}, timeout=30
            )
            if not resp_v.ok:
                return jsonify({"ok": False, "msg": f"Error editando viáticos: {resp_v.text}"}), 500
        
        # Actualizar los montos por categoría
        for cat, cat_data in categorias.items():
            cat_monto = cat_data.get("monto", 0)
            cat_rows = cat_data.get("row_nums", [])
            
            if not cat_rows: continue
            cat_rows.sort()
            r1 = cat_rows[0]
            
            # Actualizar monto en F de la primera fila de la categoría
            resp_m = req_lib.patch(
                f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/range(address='F{r1}')",
                headers={**auth_headers, "Content-Type": "application/json"},
                json={"values": [[cat_monto]]}, timeout=30
            )
            if not resp_m.ok:
                return jsonify({"ok": False, "msg": f"Error editando monto de {cat}: {resp_m.text}"}), 500
            
            # Poner en 0 las demás filas de esta categoría
            for rn in cat_rows[1:]:
                req_lib.patch(
                    f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/range(address='F{rn}')",
                    headers={**auth_headers, "Content-Type": "application/json"},
                    json={"values": [[0]]}, timeout=30
                )
                
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500

@app.route("/api/eliminar_foto", methods=["POST"])
def api_eliminar_foto():
    """Elimina una foto específica de la base de datos (Excel)."""
    data = request.json
    pwd = data.get("password")
    if pwd != "cfbc2026":
        return jsonify({"ok": False, "msg": "Contraseña incorrecta."}), 403
        
    foto_path = data.get("foto_path")
    if not foto_path:
        return jsonify({"ok": False, "msg": "No se proporcionó la ruta de la foto."}), 400
        
    try:
        token = _get_sp_token()
        if not token:
            return jsonify({"ok": False, "msg": "Error de token SP."}), 500
            
        auth_headers = {"Authorization": f"Bearer {token}"}
        site_id = _get_site_id(auth_headers)
        base_url = _get_base_url(site_id)
        
        # Obtener todas las filas para buscar la foto
        used_url = f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/usedRange"
        r = req_lib.get(used_url, headers=auth_headers, timeout=30)
        if not r.ok:
            return jsonify({"ok": False, "msg": "Error obteniendo datos."}), 500
            
        values = r.json().get("values", [])
        for idx, row in enumerate(values):
            if len(row) > 6 and row[6]:
                fotos = [f.strip() for f in str(row[6]).split(",")]
                if foto_path in fotos:
                    fotos.remove(foto_path)
                    new_fotos_str = ",".join(fotos)
                    row_num = idx + 1
                    
                    # Actualizar celda G{row_num}
                    resp_p = req_lib.patch(
                        f"{base_url}/workbook/worksheets/{SP_SHEET_GASTOS}/range(address='G{row_num}')",
                        headers={**auth_headers, "Content-Type": "application/json"},
                        json={"values": [[new_fotos_str]]}, timeout=30
                    )
                    if resp_p.ok:
                        return jsonify({"ok": True, "msg": "Foto eliminada correctamente."})
                    else:
                        return jsonify({"ok": False, "msg": f"Error al actualizar celda: {resp_p.text}"}), 500
                        
        return jsonify({"ok": False, "msg": "Foto no encontrada en los registros."}), 404
        
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500

@app.route("/api/analizar_factura", methods=["POST"])
def api_analizar_factura():
    if 'imagen' not in request.files:
        return jsonify({"ok": False, "error": "No image provided"}), 400
        
    file = request.files['imagen']
    if file.filename == '':
        return jsonify({"ok": False, "error": "No selected file"}), 400
        
    try:
        import numpy as np
        import cv2
        import base64, time

        raw_bytes = file.read()
        
        # El folio siempre viene manual desde el modal del usuario
        folio_encontrado = request.form.get('folio_manual', '').strip()
        
        if not folio_encontrado:
            return jsonify({"ok": False, "error": "No se proporcionó un folio"}), 400
        
        print(f"Buscando folio manual: {folio_encontrado}", flush=True)
        
        # Buscar el folio en Supabase
        db_res = supabase_client.table("facturas_folios").select("*").eq("folio", folio_encontrado).execute()
        db_productos = db_res.data
        db_status = "FOUND" if db_productos else "NOT_FOUND"
        
        if db_status == "NOT_FOUND":
            return jsonify({
                "ok": True,
                "factura": {"serie": "N/A", "folio": folio_encontrado, "fecha": "N/A", "total": "N/A", "url_factura": ""},
                "db_status": "NOT_FOUND",
                "comparacion": [],
                "ocr_raw_text": ""
            })
        
        # Calcular comparación desde la DB
        import unicodedata
        comparacion = []
        total_calc = 0
        for db_p in db_productos:
            prod_name = str(db_p.get("producto", "")).strip()
            cant_db = float(db_p.get("unidades", 0))
            precio_db = float(db_p.get("precio_unidad", 0))
            total_calc += cant_db * precio_db
            comparacion.append({
                "producto_db": prod_name,
                "cantidad_db": cant_db,
                "precio_db": precio_db,
                "estado": "OK"
            })
        
        # Subir foto a Supabase
        url_factura_temp = ""
        try:
            file_bytes = np.frombuffer(raw_bytes, np.uint8)
            b64_str = base64.b64encode(file_bytes).decode('utf-8')
            ruta_supa = f"Facturas/factura_{folio_encontrado}_{int(time.time())}.jpg"
            url_factura_temp = subir_foto_supabase(b64_str, ruta_supa)
        except Exception as ex:
            print(f"Error subiendo foto factura: {ex}")
            
        return jsonify({
            "ok": True,
            "factura": {
                "serie": db_productos[0].get("salida", "N/A") if db_productos else "N/A",
                "folio": folio_encontrado,
                "fecha": db_productos[0].get("diario", "N/A") if db_productos else "N/A",
                "total": f"${total_calc:,.2f}",
                "url_factura": url_factura_temp
            },
            "db_status": db_status,
            "comparacion": comparacion,
            "productos_gemini": [],
            "ocr_raw_text": ""
        })
        
    except Exception as e:
        print(f"Error en analizar_factura: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500



@app.route('/api/analizar_recibo', methods=['POST'])
def analizar_recibo():
    if 'imagen' not in request.files:
        return jsonify({"ok": False, "error": "No image provided"}), 400
        
    file = request.files['imagen']
    esperados_json = request.form.get('esperados', '[]')
    
    try:
        import json
        esperados = json.loads(esperados_json)
    except Exception as e:
        esperados = []

    try:
        # Read the image
        img_bytes = file.read()
        
        import numpy as np
        import cv2
        import unicodedata
        import re
        
        # Iniciar PaddleOCR y procesar imagen
        file_bytes = np.frombuffer(img_bytes, np.uint8)
        img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        
        # --- PREPROCESAMIENTO PARA TESSERACT ---
        h, w = img.shape[:2]
        if w > 1000:
            scale = 1000 / w
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        elif w < 600:
            scale = 600 / w
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
            
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        # -----------------------------------------------
        
        import pytesseract
        from pytesseract import Output
        
        # Extraer datos con Tesseract
        data = pytesseract.image_to_data(img, lang='spa', output_type=Output.DICT)
        
        boxes = []
        n_boxes = len(data['text'])
        for i in range(n_boxes):
            conf = int(data['conf'][i])
            text = data['text'][i].strip()
            # Ignorar palabras vacías o con muy baja confianza
            if text and conf > 20:
                x = data['left'][i]
                y = data['top'][i]
                width = data['width'][i]
                height = data['height'][i]
                
                # Coordenadas centrales
                x_center = x + (width / 2)
                y_center = y + (height / 2)
                
                boxes.append({
                    "text": text,
                    "raw": text,
                    "clean": unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8').lower(),
                    "x": x_center,
                    "y": y_center
                })
                
        # Ordenar cajas de arriba a abajo, luego de izquierda a derecha (usando un margen de tolerancia en Y de 15 px)
        boxes.sort(key=lambda b: (round(b["y"] / 15.0), b["x"]))
        
        ocr_raw = "\n".join([b["text"] for b in boxes])
        print(f"--- OCR RECIBO (Tesseract OCR) --- Extrajo {len(boxes)} cajas.", flush=True)

        # 1. Encontrar en qué orden aparecen los productos esperados
        # Unimos todo el texto para poder buscar nombres completos (ej: "bouquet rosas 12 t")
        full_text_clean = " ".join([b["clean"] for b in boxes])
        
        productos_encontrados_con_indice = []
        
        for prod in esperados:
            prod_name = prod.get("producto", "")
            
            # 1a. Buscar si alguno de los nombres exactos de Walmart está en el ticket
            nombres_validos = DB_A_ACUSE.get(prod_name.upper(), [])
            encontrado = False
            indice_aparicion = -1
            
            for nombre_acuse in nombres_validos:
                idx = full_text_clean.find(nombre_acuse)
                if idx != -1:
                    encontrado = True
                    indice_aparicion = idx
                    break
                    
            # 1b. Fallback: matching difuso por si el producto no está mapeado
            if not encontrado:
                prod_clean = unicodedata.normalize('NFKD', prod_name).encode('ASCII', 'ignore').decode('utf-8').lower()
                words = [w for w in prod_clean.split() if len(w) > 3]
                
                # Buscar palabra por palabra en la lista de cajas para sacar un índice aproximado
                match_count = 0
                primer_indice_palabra = -1
                
                for i, b in enumerate(boxes):
                    if b["clean"] in words:
                        match_count += 1
                        if primer_indice_palabra == -1:
                            primer_indice_palabra = i
                            
                if len(words) > 0 and (match_count / len(words)) >= 0.5:
                    encontrado = True
                    # Aproximamos el índice en el string completo
                    indice_aparicion = primer_indice_palabra * 5 
                    
            if encontrado:
                productos_encontrados_con_indice.append((indice_aparicion, prod))

        # Ordenar los productos detectados según aparecieron en el texto (de arriba a abajo)
        productos_encontrados_con_indice.sort(key=lambda x: x[0])
        productos_encontrados = [p[1] for p in productos_encontrados_con_indice]

        # Si faltó alguno que no se detectó bien, lo agregamos al final para no perderlo
        for prod in esperados:
            if prod not in productos_encontrados:
                productos_encontrados.append(prod)

        # 2. Extraer Cantidad Recibida para cada producto
        # En el acuse, los números vienen después del nombre del producto (ej: "bouquet rosas 12 t 15.000 15.000")
        # El segundo número (o el último antes del siguiente producto) es la "Cant Recibida".
        
        full_text_para_numeros = full_text_clean.replace(",", ".")
        conciliacion = []
        
        for i, (idx_actual, prod) in enumerate(productos_encontrados_con_indice):
            prod_name = prod.get("producto", "")
            cant_esperada = float(prod.get("cantidad", 0))
            
            # Definir el segmento de texto que le pertenece a este producto
            # Va desde donde empieza este producto hasta donde empieza el siguiente (o fin del texto)
            idx_siguiente = productos_encontrados_con_indice[i+1][0] if i + 1 < len(productos_encontrados_con_indice) else len(full_text_para_numeros)
            
            segmento = full_text_para_numeros[idx_actual:idx_siguiente]
            
            # Buscar todos los números con formato decimal EXACTO de 3 dígitos (ej. 15.000)
            # Esto evita confundir con precios (ej. 113.00) o porcentajes (0.00%) que tienen 2.
            matches = re.findall(r'\b(\d+\.\d{3})\b', segmento)
            
            cantidades_validas = []
            for m in matches:
                val = float(m)
                if 0 < val < 500:
                    cantidades_validas.append(val)
                    
            cant_recibida = 0
            if cantidades_validas:
                # Tomamos el ÚLTIMO número válido de este segmento, que corresponde a la columna "Cant Recibida"
                val = cantidades_validas[-1]
                
                # Mitigar error de OCR común donde '10' se lee como '18'
                if val == 18 and cant_esperada == 10:
                    val = 10.0
                elif val == 18.000 and cant_esperada == 10:
                    val = 10.0
                    
                cant_recibida = val
            else:
                cant_recibida = 0 # No se encontró cantidad para este producto
                
            estado = "OK" if cant_recibida == cant_esperada else "DIFF"
            
            conciliacion.append({
                "producto": prod_name,
                "esperado": cant_esperada,
                "recibido": cant_recibida,
                "diferencia": cant_recibida - cant_esperada,
                "estado": estado
            })

        # Para los productos que no se detectaron en absoluto (los que se agregaron al final)
        productos_ya_procesados = [p[1].get("producto") for p in productos_encontrados_con_indice]
        for prod in esperados:
            prod_name = prod.get("producto", "")
            if prod_name not in productos_ya_procesados:
                cant_esperada = float(prod.get("cantidad", 0))
                conciliacion.append({
                    "producto": prod_name,
                    "esperado": cant_esperada,
                    "recibido": 0,
                    "diferencia": 0 - cant_esperada,
                    "estado": "DIFF"
                })
            
        # Re-ordenar la conciliación para que coincida con el orden original de `esperados`
        conciliacion_ordenada = []
        for prod in esperados:
            prod_name = prod.get("producto", "")
            # Buscar en conciliacion
            item_c = next((item for item in conciliacion if item["producto"] == prod_name), None)
            if item_c:
                conciliacion_ordenada.append(item_c)
            
        # Subir foto a Supabase si tenemos folio
        folio = request.form.get("folio")
        if folio:
            import base64, time
            try:
                b64_str = base64.b64encode(img_bytes).decode('utf-8')
                ruta_supa = f"Acuses/acuse_{folio}_{int(time.time())}.jpg"
                url_publica = subir_foto_supabase(b64_str, ruta_supa)
                if url_publica:
                      update_data = {"url_acuse": url_publica}
                      url_factura = request.form.get("url_factura")
                      if url_factura:
                          update_data["url_factura"] = url_factura
                      supabase_client.table("facturas_folios").update(update_data).eq("folio", folio).execute()
            except Exception as ex:
                print(f"Error subiendo foto acuse: {ex}")

        return jsonify({
            "ok": True,
            "conciliacion": conciliacion_ordenada,
            "ocr_raw": ocr_raw
        })
        
    except Exception as e:
        print(f"Error procesando recibo: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route('/api/actualizar_recibo', methods=['POST'])
def actualizar_recibo():
    try:
        data = request.json
        folio = data.get('folio')
        productos = data.get('productos', [])
        
        if not folio or not productos:
            return jsonify({"ok": False, "error": "Faltan datos de folio o productos."}), 400
            
        for p in productos:
            producto_nombre = p.get('producto')
            nueva_cantidad = p.get('nueva_cantidad')
            
            if producto_nombre and nueva_cantidad is not None:
                # Update the database
                supabase_client.table('facturas_folios').update({'unidades': nueva_cantidad}).eq('folio', folio).eq('producto', producto_nombre).execute()
                
        return jsonify({"ok": True})
    except Exception as e:
        print(f"Error actualizando recibo: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/sin_acuse", methods=["POST"])
def sin_acuse():
    data = request.json
    folio = data.get("folio")
    razon = data.get("razon")
    url_factura = data.get("url_factura")
    
    if not folio or not razon:
        return jsonify({"success": False, "error": "Folio y razón requeridos"}), 400
        
    try:
        # Actualizar la base de datos
        update_data = {"razon_sin_acuse": razon}
        if url_factura:
            update_data["url_factura"] = url_factura
        supabase_client.table("facturas_folios").update(update_data).eq("folio", folio).execute()
        return jsonify({"success": True})
    except Exception as e:
        print(f"[ERROR] Al reportar sin acuse: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/cancelar_factura", methods=["POST"])
def cancelar_factura():
    data = request.json
    folio = data.get("folio")
    if not folio:
        return jsonify({"success": False, "error": "Folio requerido"}), 400
        
    try:
        # Obtener los datos actuales de la factura
        res = supabase_client.table("facturas_folios").select("*").eq("folio", folio).execute()
        if not res.data:
            return jsonify({"success": False, "error": "Factura no encontrada en base de datos"}), 404
            
        registros_a_mover = res.data
        
        # Eliminar el id para evitar conflictos de llave primaria si es identity
        for reg in registros_a_mover:
            if "id" in reg:
                del reg["id"]
                
        # Insertar en facturas_canceladas
        supabase_client.table("facturas_canceladas").insert(registros_a_mover).execute()
        
        # Eliminar de facturas_folios
        supabase_client.table("facturas_folios").delete().eq("folio", folio).execute()
        
        return jsonify({"success": True})
    except Exception as e:
        print(f"[ERROR] Al cancelar factura: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == "__main__":

    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
