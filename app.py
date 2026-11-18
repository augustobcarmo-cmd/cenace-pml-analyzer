"""
PML CENACE — Streamlit
======================
Recurrent Energy México · Inteligencia de Mercado

Fork de «CENACE PML Analyzer» v9 de Sebastian Roldan (SRF)
https://github.com/sebasrf44/cenace-pml-analyzer
reducido a lo que trata de Precios Marginales Locales (PML).

Cambios respecto al original:
- Sin BESS: fuera el scoring por caso de uso, el «Excel de Análisis» BESS y las
  opciones BESS del Excel a la medida.
- Sin modo «Mapa» ni KMZ: el mapa queda solo en el dashboard, coloreado por PML.
- Un solo Excel de análisis de PML, con secciones a elegir (resumen, rankings,
  multi-año y spread por CCR).
- Registros ordenados por fecha y hora tras la descarga paralela.
- Aviso de publicación según el proceso (MTR: día de operación +7; MDA: día -1).
- Compatible con Plotly 7 (Scattermap) y Streamlit 1.67 (width="stretch").
- Catálogo de NodosP v2026-07-16, con la versión visible en la app.
"""

import streamlit as st
import pandas as pd
import requests
import json
import urllib3
import time
import io
import os
import re
import gc
from datetime import datetime, timedelta, date
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from difflib import SequenceMatcher

import plotly.express as px
import plotly.graph_objects as go

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ═══════════════════════════════════════════════════════════════════════
# CONFIG STREAMLIT
# ═══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="PML CENACE · Recurrent Energy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Workers fijo (no más slider)
MAX_WORKERS = 8

# ═══════════════════════════════════════════════════════════════════════
# COLORES
# ═══════════════════════════════════════════════════════════════════════
RE_NAVY     = "#0e346b"
RE_RED      = "#a0090c"
RE_BLUE     = "#2777bd"
RE_ALT      = "#EBF3FB"
RE_INFO     = "#D9E2F3"
TEXT_DARK   = "#1a1a1a"
TEXT_TITLE  = "#0a2347"
GRID_LIGHT  = "#D8D8D8"
AXIS_LINE   = "#444444"

PALETTE = [
    "#0e346b", "#a0090c", "#d4a017", "#1a8a3a", "#7e57c2",
    "#ff5722", "#00897b", "#c2185b", "#3949ab", "#5d4037",
    "#1976d2", "#558b2f", "#f57c00", "#8e24aa", "#455a64",
]


# ═══════════════════════════════════════════════════════════════════════
# MONEDA — helpers para el selector MXN / USD
# ═══════════════════════════════════════════════════════════════════════
def get_moneda():
    """Retorna 'MXN' o 'USD' según session_state."""
    return st.session_state.get("moneda", "MXN")


def simbolo_moneda():
    """Retorna '$' (MXN) o 'USD$' para diferenciarlos visualmente."""
    return "USD$" if get_moneda() == "USD" else "$"


def fmt_moneda():
    """Retorna formato Streamlit dataframe column compatible: '$%.2f' o 'USD$%.2f'."""
    if get_moneda() == "USD":
        return "USD$%.2f"
    return "$%.2f"


def label_moneda():
    """Retorna label completo: '($/MWh)' o '(USD$/MWh)'."""
    if get_moneda() == "USD":
        return "(USD$/MWh)"
    return "($/MWh)"


def label_moneda_short():
    """Retorna label corto sin /MWh."""
    if get_moneda() == "USD":
        return "USD$"
    return "$"

C_HEADER = "0e346b"
C_SUB    = "2777bd"
C_RED    = "a0090c"
C_WHITE  = "FFFFFF"
C_ALT    = "EBF3FB"
C_INFO   = "D9E2F3"
C_GREEN  = "1a8a3a"
C_GOLD   = "d4a017"

st.markdown(f"""
<style>
    .main-header {{
        background: linear-gradient(135deg, {RE_NAVY} 0%, {RE_BLUE} 100%);
        padding: 2rem; border-radius: 8px; margin-bottom: 2rem; color: white;
    }}
    .main-header h1 {{ color: white !important; margin: 0; font-size: 2rem; }}
    .main-header p {{ color: rgba(255,255,255,0.85); margin: 0.5rem 0 0 0; }}
    .pml-badge {{
        background: {RE_RED}; color: white; padding: 0.3rem 0.8rem;
        border-radius: 4px; font-weight: bold; font-size: 0.85rem;
        display: inline-block; margin-left: 1rem;
    }}
    .stButton > button[kind="primary"] {{
        background: {RE_NAVY}; color: white; border: none;
    }}
    .stButton > button[kind="primary"]:hover {{ background: {RE_RED}; }}
    .mode-badge {{
        background: #f0f4f8; border-left: 4px solid {RE_NAVY};
        padding: 0.6rem 1rem; border-radius: 4px; margin-bottom: 1rem;
        color: {TEXT_DARK}; font-size: 0.9rem;
    }}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# CONSTANTES API
# ═══════════════════════════════════════════════════════════════════════
BASE_URL    = "https://ws01.cenace.gob.mx:8082/SWPML/SIM"
BLOQUE_MAX  = 7
FORMATO     = "JSON"
TIMEOUT     = 60

OSM_HEADERS = {
    'Accept': 'application/json',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'PML-CENACE/1.0 (Recurrent Energy MX; github.com/augustobcarmo-cmd/cenace-pml-analyzer)',
    'Referer': 'https://recurrent-energy.com/',
}


# ═══════════════════════════════════════════════════════════════════════
# CARGAR CATÁLOGO (cached)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data
def cargar_catalogo_default():
    """Carga el catálogo del repo (default). Retorna (catalogo, version)."""
    catalog_path = os.path.join(os.path.dirname(__file__), 'data', 'catalogo_nodos.xlsx')
    if not os.path.exists(catalog_path):
        return {}, ""
    return _parse_catalogo_workbook(catalog_path)


def cargar_catalogo_uploaded(file_bytes):
    """Carga catálogo desde bytes subidos por el usuario. Retorna (catalogo, version)."""
    try:
        return _parse_catalogo_workbook(io.BytesIO(file_bytes))
    except Exception:
        return {}, ""


def _version_catalogo(nombre_hoja):
    """'Catalogo de NodosP V20260716' → 'v2026-07-16' ('' si no se reconoce)."""
    m = re.search(r"V(\d{4})(\d{2})(\d{2})", nombre_hoja or "", flags=re.IGNORECASE)
    return f"v{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""


def _parse_catalogo_workbook(path_or_buffer):
    """Convierte el catálogo oficial de NodosP (xlsx) a dict. Retorna (catalogo, version)."""
    import openpyxl
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True, read_only=True)
    hoja = wb.sheetnames[0]
    ws = wb[hoja]
    catalogo = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 6 or not r[3]:
            continue
        clave = str(r[3]).strip()
        catalogo[clave] = {
            'sistema':   str(r[0]).strip() if r[0] else '',
            'ccr':       str(r[1]).strip() if r[1] else '',
            'zona':      str(r[2]).strip() if r[2] else '',
            'nombre':    str(r[4]).strip() if r[4] else '',
            'kv':        int(r[5]) if r[5] else 0,
            'estado':    str(r[15]).strip() if len(r) > 15 and r[15] else '',
            'municipio': str(r[17]).strip() if len(r) > 17 and r[17] else '',
        }
    wb.close()
    return catalogo, _version_catalogo(hoja)


def get_catalogo_activo():
    """Retorna (catalogo, version, subido) — catálogo subido por el usuario o el del repo."""
    if st.session_state.get("catalogo_uploaded"):
        return (st.session_state["catalogo_uploaded"],
                st.session_state.get("catalogo_uploaded_version", ""), True)
    cat, version = cargar_catalogo_default()
    return cat, version, False


# ═══════════════════════════════════════════════════════════════════════
# OSM CACHE
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, ttl=86400 * 7)
def cargar_osm_subestaciones():
    osm_local = os.path.join(os.path.dirname(__file__), 'data', 'osm_subestaciones_mx.json')
    if os.path.exists(osm_local):
        with open(osm_local, 'r', encoding='utf-8') as f:
            return json.load(f)

    query = """
[out:json][timeout:180];
area["ISO3166-1"="MX"]->.mx;
(
  node["power"="substation"](area.mx);
  way["power"="substation"](area.mx);
);
out center tags;
"""
    try:
        r = requests.post('https://overpass-api.de/api/interpreter',
                         data={'data': query},
                         headers=OSM_HEADERS, timeout=200)
        if r.status_code != 200:
            return []
        data = r.json()
        elementos = data.get('elements', [])
    except Exception:
        return []

    def _parse_voltaje_max(v_raw):
        if not v_raw: return None
        v_clean = str(v_raw).replace('kV', '').replace('KV', '').replace('kv', '').strip()
        partes = [p.strip() for p in v_clean.split(';') if p.strip()]
        nums = []
        for p in partes:
            num_str = ''
            for ch in p:
                if ch.isdigit() or ch == '.':
                    num_str += ch
                else: break
            if num_str:
                try:
                    n = float(num_str)
                    nums.append(n / 1000 if n >= 1000 else n)
                except: pass
        return max(nums) if nums else None

    subs = []
    for el in elementos:
        tags = el.get('tags', {})
        sub_type = tags.get('substation', '')
        if sub_type in ('minor_distribution', 'traction', 'industrial',
                       'household', 'measurement'):
            continue
        if el['type'] == 'node':
            lat_e = el.get('lat', 0); lon_e = el.get('lon', 0)
        elif 'center' in el:
            lat_e = el['center'].get('lat', 0); lon_e = el['center'].get('lon', 0)
        else: continue
        if lat_e == 0 and lon_e == 0: continue
        v_kv = _parse_voltaje_max(tags.get('voltage', ''))
        subs.append({
            'osm_id': el['id'], 'osm_type': el['type'],
            'lat': round(lat_e, 6), 'lon': round(lon_e, 6),
            'name': tags.get('name', ''),
            'operator': tags.get('operator', ''),
            'voltage_kv': v_kv,
            'voltage_raw': tags.get('voltage', ''),
            'substation': sub_type,
            'rating': tags.get('rating', ''),
        })
    return subs


# ═══════════════════════════════════════════════════════════════════════
# FX — TIPO DE CAMBIO BANXICO (FIX serie SF43718)
# ═══════════════════════════════════════════════════════════════════════
BANXICO_API_BASE = "https://www.banxico.org.mx/SieAPIRest/service/v1/series"
SERIE_FIX_USD = "SF43718"  # Tipo de cambio FIX peso-dólar


@st.cache_data(show_spinner=False, ttl=86400)  # 24h cache
def cargar_fx_banxico(fecha_ini_yyyymmdd, fecha_fin_yyyymmdd, _token):
    """Descarga serie histórica del FIX entre dos fechas.

    Retorna dict {fecha_str: tipo_cambio_float} o {} si falla.
    Las fechas son en formato YYYY-MM-DD.
    """
    if not _token:
        return {}

    url = (f"{BANXICO_API_BASE}/{SERIE_FIX_USD}/datos/"
           f"{fecha_ini_yyyymmdd}/{fecha_fin_yyyymmdd}"
           f"?mediaType=json")
    headers = {'Bmx-Token': _token, 'Accept': 'application/json'}

    try:
        r = requests.get(url, headers=headers, timeout=30)
        if r.status_code != 200:
            return {}
        data = r.json()
        series = data.get('bmx', {}).get('series', [])
        if not series:
            return {}
        datos = series[0].get('datos', [])
    except Exception:
        return {}

    fx_dict = {}
    for d in datos:
        fecha_raw = d.get('fecha', '')  # formato dd/mm/yyyy
        valor_raw = d.get('dato', '')
        if not fecha_raw or not valor_raw or valor_raw == 'N/E':
            continue
        try:
            # convertir dd/mm/yyyy -> yyyy-mm-dd
            day, month, year = fecha_raw.split('/')
            fecha_norm = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            fx_dict[fecha_norm] = float(valor_raw)
        except (ValueError, AttributeError):
            continue
    return fx_dict


def construir_fx_lookup(fx_dict, fechas_necesarias):
    """Para fechas sin TC (fines de semana, feriados), usa el último día hábil disponible.

    Retorna dict {fecha_str: tc_aplicable}.
    """
    if not fx_dict:
        return {}

    # Ordenar fechas disponibles
    fechas_disponibles = sorted(fx_dict.keys())
    if not fechas_disponibles:
        return {}

    lookup = {}
    fechas_solicitadas = sorted(set(fechas_necesarias))

    idx = 0  # puntero a fechas disponibles
    last_known = fx_dict[fechas_disponibles[0]]

    for fecha_req in fechas_solicitadas:
        # Avanzar puntero hasta encontrar fecha <= fecha_req
        while idx < len(fechas_disponibles) and fechas_disponibles[idx] <= fecha_req:
            last_known = fx_dict[fechas_disponibles[idx]]
            idx += 1
        lookup[fecha_req] = last_known
    return lookup


def aplicar_conversion_usd(acumulado, fx_lookup):
    """Aplica conversión MXN→USD a todos los registros de acumulado.

    Modifica `acumulado` in-place dividiendo PML, PML_ENE, PML_PER, PML_CNG
    entre el tipo de cambio del día. Retorna el acumulado modificado.
    """
    if not fx_lookup:
        return acumulado

    for nodo, filas in acumulado.items():
        for f in filas:
            fecha = f.get("fecha", "")
            tc = fx_lookup.get(fecha)
            if not tc or tc <= 0:
                continue
            for k in ("pml", "pml_ene", "pml_per", "pml_cng"):
                v = f.get(k)
                if isinstance(v, (int, float)):
                    f[k] = round(v / tc, 4)
    return acumulado


def obtener_token_banxico():
    """Lee el token de Streamlit secrets. Retorna '' si no existe."""
    try:
        return st.secrets.get("BANXICO_TOKEN", "")
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════════════════════
# MATCHING GEOGRÁFICO
# ═══════════════════════════════════════════════════════════════════════
BBOX_ESTADO = {
    'AGUASCALIENTES':       (21.62, -102.87, 22.45, -101.84),
    'BAJA CALIFORNIA':      (28.00, -117.30, 32.72, -112.75),
    'BAJA CALIFORNIA SUR':  (22.87, -115.30, 28.00, -109.40),
    'CAMPECHE':             (17.81, -92.47,  20.85, -89.10),
    'CHIAPAS':              (14.53, -94.14,  17.99, -90.37),
    'CHIHUAHUA':            (25.55, -109.07, 31.78, -103.30),
    'COAHUILA':             (24.55, -103.96, 29.88, -99.84),
    'COLIMA':               (18.69, -104.74, 19.52, -103.49),
    'DURANGO':              (22.36, -107.20, 26.83, -102.46),
    'GUANAJUATO':           (19.91, -102.10, 21.85, -99.69),
    'GUERRERO':             (16.31, -102.18, 18.92, -98.00),
    'HIDALGO':              (19.63, -99.91,  21.40, -97.97),
    'JALISCO':              (18.92, -105.69, 22.75, -101.50),
    'MEXICO':               (18.36, -100.59, 20.29, -98.62),
    'CIUDAD DE MEXICO':     (19.05, -99.36,  19.59, -98.94),
    'MICHOACAN':            (17.92, -103.74, 20.40, -100.07),
    'MORELOS':              (18.32, -99.50,  19.13, -98.62),
    'NAYARIT':              (20.60, -105.78, 23.09, -103.71),
    'NUEVO LEON':           (23.18, -101.21, 27.81, -98.40),
    'OAXACA':               (15.65, -98.51,  18.66, -93.86),
    'PUEBLA':               (17.86, -99.06,  20.84, -96.71),
    'QUERETARO':            (20.02, -100.59, 21.67, -99.04),
    'QUINTANA ROO':         (17.88, -89.30,  21.62, -86.69),
    'SAN LUIS POTOSI':      (21.16, -102.30, 24.55, -98.34),
    'SINALOA':              (22.50, -109.45, 27.04, -105.41),
    'SONORA':               (26.32, -115.07, 32.51, -108.40),
    'TABASCO':              (17.25, -94.13,  18.65, -91.00),
    'TAMAULIPAS':           (22.21, -100.16, 27.68, -97.13),
    'TLAXCALA':             (19.06, -98.71,  19.72, -97.62),
    'VERACRUZ':             (17.15, -98.65,  22.46, -93.61),
    'YUCATAN':              (19.55, -90.42,  21.62, -87.53),
    'ZACATECAS':            (21.04, -104.34, 25.13, -100.81),
}

PALABRAS_GENERICAS = {
    'maniobras', 'planta', 'entronque', 'subestacion', 'subestación',
    'central', 'switching', 'rectificadora', 'tap', 'industrial',
    'cementos', 'pemex', 'cuadro', 'parque', 'potencia', 'aeropuerto',
    'norte', 'sur', 'este', 'oeste', 'poniente', 'oriente',
    'centro', 'cerro', 'valle', 'laguna', 'puerto', 'villa',
    'cabo', 'monte', 'sierra', 'mesa', 'isla', 'rio',
    'nuevo', 'nueva', 'viejo', 'vieja', 'gran', 'grande',
    'antiguo', 'antigua', 'segundo', 'segunda', 'primero', 'primera',
    'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x',
    '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
    'uno', 'dos', 'tres', 'cuatro', 'cinco',
    'de', 'la', 'el', 'las', 'los', 'del', 'al', 'con', 'por',
    'mexico', 'santa', 'santo', 'san',
}

PALABRAS_SEMIGENERICAS = {
    'hermosillo', 'guadalajara', 'monterrey', 'merida', 'chihuahua',
    'culiacan', 'puebla', 'queretaro', 'leon', 'morelia', 'tepic',
    'durango', 'mazatlan', 'oaxaca', 'campeche', 'cancun',
    'chetumal', 'colima', 'manzanillo', 'veracruz', 'tampico',
    'reynosa', 'matamoros', 'juarez', 'mexicali', 'tijuana',
    'ensenada', 'cardenas', 'guadalupe', 'progreso', 'lazaro',
}


def en_estado(lat, lon, nombre_estado):
    if not nombre_estado: return None
    clave = nombre_estado.upper().strip()
    for variant in [clave, clave.split(' ')[0],
                    ' '.join(clave.split(' ')[:2]),
                    ' '.join(clave.split(' ')[:3])]:
        bb = BBOX_ESTADO.get(variant)
        if bb:
            return bb[0] <= lat <= bb[2] and bb[1] <= lon <= bb[3]
    return None


def normalizar_nombre(s):
    if not s: return ''
    s = s.lower().strip()
    repl = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n','ü':'u'}
    for k, v in repl.items():
        s = s.replace(k, v)
    for prefix in ['s/e ', 's/e. ', 'se ', 'sub ', 'subestacion ',
                   'central ', 'planta ', 'c.t. ', 'ct ',
                   'c.h. ', 'ch ', 'c.f.e. ', 'cfe ']:
        if s.startswith(prefix):
            s = s[len(prefix):]
    for art in ['los ', 'las ', 'el ', 'la ']:
        if s.startswith(art):
            resto = s[len(art):]
            if len(resto.split(' ', 1)[0]) >= 4:
                s = resto
                break
    for sufijo in [' maniobras', ' switching', ' tap',
                   ' planta', ' central', ' cfe']:
        if s.endswith(sufijo):
            s = s[:-len(sufijo)]
    roman_map = {' i': ' 1', ' ii': ' 2', ' iii': ' 3', ' iv': ' 4',
                 ' v': ' 5', ' vi': ' 6', ' vii': ' 7', ' viii': ' 8',
                 ' ix': ' 9', ' x': ' 10'}
    for r, a in roman_map.items():
        if s.endswith(r):
            s = s[:-len(r)] + a
            break
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def similitud(a, b):
    if not a or not b: return 0.0
    return SequenceMatcher(None, a, b).ratio()


def palabras_clave(nombre_norm):
    if not nombre_norm: return []
    return [p for p in nombre_norm.split()
            if len(p) >= 5 and p not in PALABRAS_GENERICAS and not p.isdigit()]


def buscar_match(nombre_cenace, voltaje_cenace_kv, estado_cat, candidatos):
    nombre_norm = normalizar_nombre(nombre_cenace)
    if not nombre_norm:
        return None, 0.0, 'sin nombre'

    exactos = [s for s in candidatos if s['name_norm'] == nombre_norm]
    if exactos:
        if estado_cat:
            exactos_en_est = [s for s in exactos
                              if en_estado(s['lat'], s['lon'], estado_cat) is not False]
            if exactos_en_est:
                exactos = exactos_en_est
            elif len(nombre_norm) < 18:
                exactos = []
        if exactos:
            if voltaje_cenace_kv > 0 and len(exactos) > 1:
                con_v = [s for s in exactos if s['voltage_kv'] and
                         abs(s['voltage_kv'] - voltaje_cenace_kv) < voltaje_cenace_kv * 0.15]
                if con_v:
                    return con_v[0], 1.0, 'match exacto + voltaje'
            return exactos[0], 1.0, 'match exacto'

    substring_matches = []
    for s in candidatos:
        n_osm = s['name_norm']
        if nombre_norm in n_osm and len(nombre_norm) >= 4:
            if n_osm.startswith(nombre_norm + ' ') or n_osm == nombre_norm:
                cob = len(nombre_norm) / len(n_osm)
                if cob >= 0.40:
                    substring_matches.append((0.92, s, 'OSM extiende CENACE'))
            elif (' ' + nombre_norm) in n_osm or n_osm.endswith(' ' + nombre_norm):
                cob = len(nombre_norm) / len(n_osm)
                if cob >= 0.40:
                    substring_matches.append((0.88, s, 'CENACE dentro de OSM'))
        elif n_osm in nombre_norm and len(n_osm) >= 4:
            if nombre_norm.startswith(n_osm + ' ') or nombre_norm == n_osm:
                cob = len(n_osm) / len(nombre_norm)
                if cob >= 0.40:
                    substring_matches.append((0.90, s, 'CENACE extiende OSM'))

    difusos = []
    for s in candidatos:
        sim = similitud(nombre_norm, s['name_norm'])
        if sim >= 0.80:
            difusos.append((sim, s))

    kw_cenace = palabras_clave(nombre_norm)
    keyword_matches = []
    if kw_cenace:
        for s in candidatos:
            kw_osm = palabras_clave(s.get('name_norm', ''))
            if not kw_osm: continue
            compartidas = set(kw_cenace) & set(kw_osm)
            if not compartidas: continue
            todas_semi = all(p in PALABRAS_SEMIGENERICAS for p in compartidas)
            largas_no_semi = sum(1 for p in compartidas
                                 if len(p) >= 6 and p not in PALABRAS_SEMIGENERICAS)
            if todas_semi: base_score = 0.62
            elif largas_no_semi >= 1: base_score = 0.78
            else: base_score = 0.70
            ratio = len(compartidas) / max((len(kw_cenace) + len(kw_osm)) / 2, 1)
            score = base_score + (ratio * 0.12)
            keyword_matches.append((score, s, list(compartidas)))

    scored = []; seen = set()
    for sc, s, t in substring_matches:
        if s['osm_id'] not in seen:
            scored.append((sc, s, 'substring', t))
            seen.add(s['osm_id'])
    for sim, s in difusos:
        if s['osm_id'] not in seen:
            scored.append((sim, s, 'difuso', ''))
            seen.add(s['osm_id'])
    for sc, s, kws in keyword_matches:
        if s['osm_id'] not in seen:
            scored.append((sc, s, 'keyword', f'comparte "{",".join(kws[:2])}"'))
            seen.add(s['osm_id'])

    if not scored:
        return None, 0.0, 'no match'

    if estado_cat:
        scored = [(sc, s, t, st) for sc, s, t, st in scored
                  if en_estado(s['lat'], s['lon'], estado_cat) is not False]

    if not scored:
        return None, 0.0, 'sin match (estado)'

    if voltaje_cenace_kv > 0:
        filt = []
        for sim, s, t, st in scored:
            v = s['voltage_kv']
            ok = (v is not None and abs(v - voltaje_cenace_kv) < voltaje_cenace_kv * 0.15)
            unk = v is None
            if t == 'difuso' and sim < 0.92 and not (ok or unk):
                continue
            if t == 'keyword' and not (ok or unk):
                continue
            if ok: filt.append((sim + 0.02, s, t, st))
            elif unk: filt.append((sim, s, t, st))
            else: filt.append((sim - 0.10, s, t, st))
        scored = filt

    if not scored:
        return None, 0.0, 'sin match (voltaje)'

    scored.sort(key=lambda x: -x[0])
    sim_f, mejor, tipo, sub_tipo = scored[0]
    if tipo == 'substring': razon = sub_tipo
    elif tipo == 'keyword': razon = f'palabra-clave: {sub_tipo}'
    elif sim_f >= 0.95: razon = 'match difuso alto'
    elif sim_f >= 0.85: razon = 'match difuso medio'
    else: razon = f'match difuso ({sim_f:.2f})'
    return mejor, sim_f, razon


def matchear_nodos(nodos, catalogo, osm_subs):
    for s in osm_subs:
        if 'name_norm' not in s:
            s['name_norm'] = normalizar_nombre(s.get('name', ''))
    candidatos = [s for s in osm_subs if s.get('name_norm')]
    resultados = []
    for clave in nodos:
        info = catalogo.get(clave, {})
        nombre = info.get('nombre', clave)
        kv = info.get('kv', 0)
        if kv == 0:
            try: kv = int(clave.split('-')[-1])
            except: kv = 0
        ccr = info.get('ccr', '')
        zona = info.get('zona', '')
        estado = info.get('estado', '')
        municipio = info.get('municipio', '')

        match, sim, razon = buscar_match(nombre, kv, estado, candidatos)

        if match:
            if 'palabra-clave' in razon: calidad = '🥉 Aceptable'
            elif sim >= 0.95: calidad = '🥇 Excelente'
            elif sim >= 0.90: calidad = '🥈 Bueno'
            else: calidad = '🥉 Aceptable'
        else:
            calidad = '❌ Sin match'

        resultados.append({
            'clave': clave, 'nombre_cenace': nombre, 'ccr': ccr, 'zona': zona,
            'kv_cenace': kv, 'estado': estado, 'municipio': municipio,
            'calidad': calidad,
            'similitud': round(sim, 3) if match else 0,
            'razon': razon,
            'nombre_osm': match['name'] if match else '',
            'kv_osm': match['voltage_kv'] if match and match['voltage_kv'] else None,
            'voltage_raw': match['voltage_raw'] if match else '',
            'operator_osm': match['operator'] if match else '',
            'subtipo_osm': match['substation'] if match else '',
            'lat': match['lat'] if match else None,
            'lon': match['lon'] if match else None,
            'osm_id': match['osm_id'] if match else None,
            'osm_type': match['osm_type'] if match else None,
        })
    return resultados


# ═══════════════════════════════════════════════════════════════════════
# CENACE — descarga
# ═══════════════════════════════════════════════════════════════════════
def parse_fecha(s): return datetime.strptime(s, "%Y/%m/%d")
def fmt(d): return d.strftime("%Y/%m/%d")

def generar_bloques(ini_str, fin_str, max_dias=BLOQUE_MAX):
    ini = parse_fecha(ini_str); fin = parse_fecha(fin_str)
    bloques = []; cursor = ini
    while cursor <= fin:
        bloque_fin = min(cursor + timedelta(days=max_dias - 1), fin)
        bloques.append((fmt(cursor), fmt(bloque_fin)))
        cursor = bloque_fin + timedelta(days=1)
    return bloques

def construir_url(nodos_lista, fecha_ini, fecha_fin, sistema, proceso):
    nodos_str = ",".join(nodos_lista)
    ai, mi, di = fecha_ini.split("/")
    af, mf, df_ = fecha_fin.split("/")
    return f"{BASE_URL}/{sistema}/{proceso}/{nodos_str}/{ai}/{mi}/{di}/{af}/{mf}/{df_}/{FORMATO}"

def consultar(nodos_lista, fecha_ini, fecha_fin, sistema, proceso,
              errores_lista, lock, max_reintentos=4):
    url = construir_url(nodos_lista, fecha_ini, fecha_fin, sistema, proceso)
    if len(url) > 2000:
        with lock:
            errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                   "nodos": len(nodos_lista),
                                   "error": "URL muy larga"})
        return None
    session = requests.Session(); session.verify = False
    for intento in range(max_reintentos):
        try:
            r = session.get(url, timeout=TIMEOUT)
            if r.status_code == 200: return r.text
            elif r.status_code == 204: return None
            elif r.status_code == 429:
                time.sleep(2 ** (intento + 2)); continue
            else:
                if intento < max_reintentos - 1:
                    time.sleep(2 ** (intento + 1)); continue
                with lock:
                    errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                           "nodos": len(nodos_lista),
                                           "error": f"HTTP {r.status_code}"})
                return None
        except requests.exceptions.Timeout:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento); continue
            with lock:
                errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                       "nodos": len(nodos_lista),
                                       "error": "Timeout"})
            return None
        except Exception as e:
            if intento < max_reintentos - 1:
                time.sleep(2 ** intento); continue
            with lock:
                errores_lista.append({"fecha": f"{fecha_ini} → {fecha_fin}",
                                       "nodos": len(nodos_lista),
                                       "error": f"{type(e).__name__}"})
            return None
    return None

def _num(val):
    try: return float(val)
    except (TypeError, ValueError): return val

def parsear_json(texto):
    if not texto: return {}
    try: obj = json.loads(texto)
    except Exception: return {}
    datos = {}
    for nd in obj.get("Resultados", []):
        clv = nd.get("clv_nodo", "?")
        valores = nd.get("Valores", [])
        if not valores: continue
        registros = []
        for v in valores:
            registros.append({
                "fecha":   v.get("fecha", ""),
                "hora":    int(v.get("hora", 0)) if v.get("hora") else 0,
                "pml":     _num(v.get("pml", 0)),
                "pml_ene": _num(v.get("pml_ene", 0)),
                "pml_per": _num(v.get("pml_per", 0)),
                "pml_cng": _num(v.get("pml_cng", 0)),
            })
        datos[clv] = registros
    return datos

def descargar_pml(nodos, fecha_ini, fecha_fin, sistema, proceso, progress_cb=None):
    bloques = generar_bloques(fecha_ini, fecha_fin)
    LOTE = 10
    lotes_nodos = [nodos[i:i+LOTE] for i in range(0, len(nodos), LOTE)]
    consultas = []
    for lote in lotes_nodos:
        for bi, bf in bloques:
            consultas.append((lote, bi, bf))
    total = len(consultas)
    acumulado = {n: [] for n in nodos}
    errores_consulta = []; lock = Lock(); completed = 0
    def _job(c):
        lote, bi, bf = c
        return parsear_json(consultar(lote, bi, bf, sistema, proceso,
                                       errores_consulta, lock))
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_job, c): c for c in consultas}
        for fut in as_completed(futures):
            data = fut.result() or {}
            for nodo, regs in data.items():
                if nodo in acumulado:
                    acumulado[nodo].extend(regs)
            completed += 1
            if progress_cb:
                progress_cb(completed, total)
    # Las consultas paralelas terminan en cualquier orden: ordenar por fecha y hora
    acumulado = {n: sorted(regs, key=lambda f: (f["fecha"], f["hora"]))
                 for n, regs in acumulado.items() if regs}
    return acumulado, errores_consulta


def descargar_pml_auto(nodos, fecha_ini, fecha_fin, proceso, catalogo, progress_cb=None):
    """Auto-detecta el Sistema de cada nodo desde el catálogo y descarga.

    Si hay nodos de varios sistemas (raro: BCA/BCS son redes aisladas), hace
    consultas paralelas y combina resultados.

    Retorna (acumulado, errores, info_sistemas) donde info_sistemas es un dict
    {sistema: count_nodos}.
    """
    # Agrupar nodos por sistema según catálogo
    grupos = {"SIN": [], "BCA": [], "BCS": []}
    nodos_sin_sistema = []

    for nodo in nodos:
        info = catalogo.get(nodo, {}) if catalogo else {}
        sistema = (info.get("sistema") or "").strip().upper()
        if sistema in grupos:
            grupos[sistema].append(nodo)
        else:
            # Fallback: asumir SIN (95% de los nodos están ahí)
            nodos_sin_sistema.append(nodo)
            grupos["SIN"].append(nodo)

    info_sistemas = {s: len(ns) for s, ns in grupos.items() if ns}

    # Si todo es un solo sistema, llamada directa (sin overhead)
    sistemas_activos = [s for s, ns in grupos.items() if ns]
    if len(sistemas_activos) == 1:
        sistema = sistemas_activos[0]
        acumulado, errores = descargar_pml(
            grupos[sistema], fecha_ini, fecha_fin, sistema, proceso, progress_cb=progress_cb)
        return acumulado, errores, info_sistemas, nodos_sin_sistema

    # Multi-sistema: descargar cada uno en serie (no paralelo entre sistemas
    # para no saturar el progress bar; cada sistema YA usa workers paralelos internos)
    acumulado_total = {}
    errores_total = []

    # Progress bar combinado
    completed_jobs = [0]
    total_jobs_estimate = sum(
        len(generar_bloques(fecha_ini, fecha_fin)) * ((len(grupos[s]) + 9) // 10)
        for s in sistemas_activos
    )

    def cb_combined(done, total_per_system):
        if progress_cb and total_jobs_estimate > 0:
            current = completed_jobs[0] + done
            progress_cb(current, total_jobs_estimate)

    for sistema in sistemas_activos:
        nodos_grupo = grupos[sistema]
        if not nodos_grupo:
            continue
        n_jobs_grupo = len(generar_bloques(fecha_ini, fecha_fin)) * ((len(nodos_grupo) + 9) // 10)
        ac_g, err_g = descargar_pml(nodos_grupo, fecha_ini, fecha_fin, sistema, proceso,
                                      progress_cb=cb_combined)
        acumulado_total.update(ac_g)
        errores_total.extend(err_g)
        completed_jobs[0] += n_jobs_grupo

    return acumulado_total, errores_total, info_sistemas, nodos_sin_sistema


# ═══════════════════════════════════════════════════════════════════════
# EXCEL DATOS (datos crudos)
# ═══════════════════════════════════════════════════════════════════════
def generar_excel_datos(acumulado, sistema, proceso, fecha_ini, fecha_fin,
                          moneda="MXN", tc_info=None):
    _NOW = datetime.now().strftime("%Y-%m-%d %H:%M")
    _side = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=_side, right=_side, top=_side, bottom=_side)

    _sym = "USD$" if moneda == "USD" else "$"

    def hdr(cell, bg=C_HEADER, fg=C_WHITE, size=11):
        cell.font = Font(bold=True, color=fg, size=size, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE

    _FONT_DATO = Font(name="Arial", size=10)
    _ALIGN_NUM = Alignment(horizontal="center", vertical="center")
    _ALIGN_TXT = Alignment(horizontal="left", vertical="center")
    _FILL_ALT  = PatternFill("solid", start_color=C_ALT)

    COLS   = ["Fecha", "Hora", f"PML ({_sym}/MWh)", "Energía", "Pérdidas", "Congestión"]
    ANCHOS = [14, 8, 18, 16, 16, 18]
    ES_NUM = [False, True, True, True, True, True]

    wb = Workbook()
    ws = wb.active
    ws.title = "■ Portada"
    ws.sheet_properties.tabColor = C_HEADER
    ws.sheet_view.showGridLines = False

    ws.merge_cells("B2:H4")
    c = ws["B2"]
    c.value = "Precios Marginales Locales\nMercado Eléctrico Mayorista — CENACE"
    c.font = Font(bold=True, color=C_WHITE, size=22, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 5):
        ws.row_dimensions[r].height = 32

    ws.merge_cells("B5:H5")
    c = ws["B5"]; c.value = "Reporte de Datos"
    c.font = Font(bold=True, italic=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_RED)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("B7:H7")
    c = ws["B7"]; c.value = "PARÁMETROS DE LA CONSULTA"
    c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 22

    info_rows = [
        ("Sistema", sistema), ("Proceso", proceso),
        ("Fecha inicial", fecha_ini), ("Fecha final", fecha_fin),
        ("Nodos en reporte", f"{len(acumulado):,} nodos"),
        ("Total registros", f"{sum(len(f) for f in acumulado.values()):,} filas"),
        ("Moneda", f"{moneda}" + (f" · TC FIX promedio: {tc_info['tc_promedio']:.4f}" if tc_info and tc_info.get('tc_promedio') else "")),
        ("Fuente", "CENACE · servicio web SW-PML"),
        ("Fecha generación", _NOW),
    ]
    for ri, (lbl, val) in enumerate(info_rows, start=8):
        ws.merge_cells(f"B{ri}:D{ri}")
        c1 = ws[f"B{ri}"]; c1.value = lbl
        c1.font = Font(bold=True, color=C_HEADER, size=11, name="Arial")
        c1.fill = PatternFill("solid", start_color=C_INFO)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c1.border = BORDE
        ws.merge_cells(f"E{ri}:H{ri}")
        c2 = ws[f"E{ri}"]; c2.value = val
        c2.font = Font(size=11, name="Arial")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2.border = BORDE
        ws.row_dimensions[ri].height = 22

    # Pie de portada
    ws.merge_cells("B17:H17")
    c = ws["B17"]; c.value = "Recurrent Energy México · Inteligencia de Mercado"
    c.font = Font(italic=True, bold=True, color=C_HEADER, size=12, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=C_INFO)
    ws.row_dimensions[17].height = 24

    ws.merge_cells("B18:H18")
    c = ws["B18"]; c.value = "A subsidiary of Canadian Solar"
    c.font = Font(italic=True, color="555555", size=10, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[18].height = 18

    for col, w in [("A", 3), ("B", 14), ("C", 14), ("D", 14),
                   ("E", 14), ("F", 14), ("G", 14), ("H", 14), ("I", 3)]:
        ws.column_dimensions[col].width = w

    ws_res = wb.create_sheet("Resumen")
    ws_res.merge_cells("A1:F1")
    c = ws_res["A1"]; c.value = f"Resumen — {len(acumulado)} nodos · {sistema} · {proceso}"
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_res.row_dimensions[1].height = 28

    res_cols = ["Nodo", "# Registros", "PML promedio", "PML máximo", "PML mínimo", "Período"]
    for ci, col in enumerate(res_cols, 1):
        hdr(ws_res.cell(row=2, column=ci, value=col), bg=C_SUB)
    ws_res.row_dimensions[2].height = 22

    for ri, (nodo, filas) in enumerate(acumulado.items(), start=3):
        if not filas: continue
        pmls = [f["pml"] for f in filas if isinstance(f["pml"], (int, float))]
        bg = C_ALT if ri % 2 == 0 else None
        valores = [
            nodo, len(filas),
            round(sum(pmls)/len(pmls), 2) if pmls else "—",
            round(max(pmls), 2) if pmls else "—",
            round(min(pmls), 2) if pmls else "—",
            f"{filas[0]['fecha']} → {filas[-1]['fecha']}",
        ]
        for ci, v in enumerate(valores, 1):
            cc = ws_res.cell(row=ri, column=ci, value=v)
            cc.font = _FONT_DATO
            cc.alignment = _ALIGN_NUM if ci > 1 else _ALIGN_TXT
            cc.border = BORDE
            if bg: cc.fill = PatternFill("solid", start_color=bg)
            if ci in (3, 4, 5) and isinstance(v, (int, float)):
                cc.number_format = "#,##0.00"

    for ci, w in enumerate([14, 14, 14, 14, 14, 24], 1):
        ws_res.column_dimensions[get_column_letter(ci)].width = w
    ws_res.freeze_panes = "A3"

    for nodo, filas in acumulado.items():
        ws_n = wb.create_sheet(title=nodo[:31])
        ws_n.merge_cells("A1:F1")
        c = ws_n["A1"]; c.value = f"PML — Nodo: {nodo}"
        c.font = Font(bold=True, color=C_WHITE, size=13, name="Arial")
        c.fill = PatternFill("solid", start_color=C_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_n.row_dimensions[1].height = 28

        ws_n.merge_cells("A2:F2")
        c = ws_n["A2"]
        c.value = (f"Sistema: {sistema} | Proceso: {proceso} | "
                   f"Período: {fecha_ini} → {fecha_fin} | Total: {len(filas):,} | "
                   f"Fuente: CENACE SW-PML")
        c.font = Font(italic=True, size=9, name="Arial", color="555555")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", start_color=C_INFO)
        ws_n.row_dimensions[2].height = 16

        for ci, (enc, ancho) in enumerate(zip(COLS, ANCHOS), 1):
            hdr(ws_n.cell(row=3, column=ci, value=enc), bg=C_SUB)
            ws_n.column_dimensions[get_column_letter(ci)].width = ancho
        ws_n.row_dimensions[3].height = 22

        for fila in filas:
            ws_n.append([fila["fecha"], fila["hora"], fila["pml"],
                         fila["pml_ene"], fila["pml_per"], fila["pml_cng"]])
        for i in range(len(filas)):
            row_idx = i + 4
            es_alt = (i % 2 == 0)
            for ci in range(1, 7):
                cc = ws_n.cell(row=row_idx, column=ci)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ES_NUM[ci-1] else _ALIGN_TXT
                if es_alt: cc.fill = _FILL_ALT
                if ES_NUM[ci-1]: cc.number_format = "#,##0.00"
        ws_n.freeze_panes = "A4"

    buffer = io.BytesIO()
    wb.save(buffer); buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# EXCEL ANÁLISIS — secciones de PML que el usuario elige
# ═══════════════════════════════════════════════════════════════════════
SECCIONES_ANALISIS = {
    "resumen":         "Resumen estadístico",
    "top_pml":         "Top PML promedio",
    "top_volatilidad": "Top volatilidad",
    "top_negativos":   "Top horas negativas",
    "multiano":        "Multi-año por nodo",
    "spread_ccr":      "Spread mensual por CCR",
}


def generar_excel_analisis(df, df_resumen, opciones, sistema, proceso,
                           fecha_ini, fecha_fin, moneda="MXN", tc_info=None,
                           df_multianos_dict=None):
    """Excel de análisis de PML con las secciones elegidas.

    opciones: dict {clave de SECCIONES_ANALISIS: bool}.
    df_multianos_dict: {nodo: DataFrame con mes_num + una columna por año}.
    """
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.marker import Marker

    _NOW = datetime.now().strftime("%Y-%m-%d %H:%M")
    _side = Side(style="thin", color="BFBFBF")
    BORDE = Border(left=_side, right=_side, top=_side, bottom=_side)
    _sym = "USD$" if moneda == "USD" else "$"

    def hdr(cell, bg=C_HEADER, fg=C_WHITE, size=11):
        cell.font = Font(bold=True, color=fg, size=size, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE

    def titulo(ws_t, rango, texto, size=14, height=26, bg=C_HEADER):
        ws_t.merge_cells(rango)
        c = ws_t[rango.split(":")[0]]
        c.value = texto
        c.font = Font(bold=True, color=C_WHITE, size=size, name="Arial")
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws_t.row_dimensions[c.row].height = height

    _FONT_DATO = Font(name="Arial", size=10)
    _ALIGN_NUM = Alignment(horizontal="center", vertical="center")
    _ALIGN_TXT = Alignment(horizontal="left", vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = "■ Portada"
    ws.sheet_properties.tabColor = C_HEADER
    ws.sheet_view.showGridLines = False

    # ─── Portada ───
    ws.merge_cells("B2:H4")
    c = ws["B2"]
    c.value = "Precios Marginales Locales\nAnálisis de PML — CENACE"
    c.font = Font(bold=True, color=C_WHITE, size=22, name="Arial")
    c.fill = PatternFill("solid", start_color=C_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 5):
        ws.row_dimensions[r].height = 32

    ws.merge_cells("B5:H5")
    c = ws["B5"]; c.value = "Reporte de Análisis"
    c.font = Font(bold=True, italic=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_RED)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws.merge_cells("B7:H7")
    c = ws["B7"]; c.value = "PARÁMETROS DE LA CONSULTA"
    c.font = Font(bold=True, color=C_WHITE, size=12, name="Arial")
    c.fill = PatternFill("solid", start_color=C_SUB)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[7].height = 22

    secciones_inc = [SECCIONES_ANALISIS.get(k, k) for k, v in opciones.items() if v]
    info_rows = [
        ("Sistema", sistema), ("Proceso", proceso),
        ("Período", f"{fecha_ini} → {fecha_fin}"),
        ("Nodos", f"{df['nodo'].nunique() if not df.empty else 0}"),
        ("Moneda", f"{moneda}" + (f" · TC FIX promedio: {tc_info['tc_promedio']:.4f}"
                                     if tc_info and tc_info.get('tc_promedio') else "")),
        ("Fuente", "CENACE · servicio web SW-PML"),
        ("Secciones", ", ".join(secciones_inc) if secciones_inc else "—"),
        ("Fecha generación", _NOW),
    ]
    for ri, (lbl, val) in enumerate(info_rows, start=8):
        ws.merge_cells(f"B{ri}:D{ri}")
        c1 = ws[f"B{ri}"]; c1.value = lbl
        c1.font = Font(bold=True, color=C_HEADER, size=11, name="Arial")
        c1.fill = PatternFill("solid", start_color=C_INFO)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c1.border = BORDE
        ws.merge_cells(f"E{ri}:H{ri}")
        c2 = ws[f"E{ri}"]; c2.value = val
        c2.font = Font(size=11, name="Arial")
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        c2.border = BORDE
        ws.row_dimensions[ri].height = 22

    ws.merge_cells("B17:H17")
    c = ws["B17"]; c.value = "Recurrent Energy México · Inteligencia de Mercado"
    c.font = Font(italic=True, bold=True, color=C_HEADER, size=12, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", start_color=C_INFO)
    ws.row_dimensions[17].height = 24

    ws.merge_cells("B18:H18")
    c = ws["B18"]; c.value = "A subsidiary of Canadian Solar"
    c.font = Font(italic=True, color="555555", size=10, name="Arial")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[18].height = 18

    for col, w in [("A", 3), ("B", 18), ("C", 18), ("D", 18),
                   ("E", 18), ("F", 18), ("G", 18), ("H", 18), ("I", 3)]:
        ws.column_dimensions[col].width = w

    # ─── SECCIÓN: Resumen estadístico ───
    if opciones.get("resumen") and df_resumen is not None and not df_resumen.empty:
        ws_r = wb.create_sheet("📊 Resumen")
        titulo(ws_r, "A1:L1", f"Resumen estadístico por nodo ({_sym}/MWh)")

        cols = ["Nodo", "Nombre", "CCR", "# Reg", "Promedio", "Mediana",
                "Máximo", "Mínimo", "Volatilidad", "P95", "P5", "% horas neg"]
        for ci, col in enumerate(cols, 1):
            hdr(ws_r.cell(row=2, column=ci, value=col), bg=C_SUB)
        ws_r.row_dimensions[2].height = 22

        for i, (_, r) in enumerate(df_resumen.iterrows(), 1):
            bg = C_ALT if i % 2 == 0 else None
            valores = [str(r['nodo']), str(r['nombre']), str(r['ccr']), int(r['registros']),
                       r['promedio'], r['mediana'], r['maximo'], r['minimo'],
                       r['std'], r['p95'], r['p5'], r['% horas neg']]
            for ci, v in enumerate(valores, 1):
                if ci > 4:
                    v = float(v) if pd.notna(v) else None
                cc = ws_r.cell(row=i+2, column=ci, value=v)
                cc.font = _FONT_DATO
                cc.border = BORDE
                cc.alignment = _ALIGN_NUM if ci > 3 else _ALIGN_TXT
                if bg: cc.fill = PatternFill("solid", start_color=bg)
                if 5 <= ci <= 11: cc.number_format = "#,##0.00"
                elif ci == 12: cc.number_format = "0.0\"%\""

        for ci, w in enumerate([14, 22, 14, 10, 12, 12, 12, 12, 12, 12, 12, 12], 1):
            ws_r.column_dimensions[get_column_letter(ci)].width = w
        ws_r.freeze_panes = "A3"

    # ─── SECCIÓN: Rankings ───
    rankings = []
    if not df.empty:
        if opciones.get("top_pml"):
            rankings.append(("🏆 Top PML promedio", f"PML promedio ({_sym}/MWh)",
                             df.groupby("nodo")["pml"].mean(), "#,##0.00"))
        if opciones.get("top_volatilidad"):
            rankings.append(("📊 Top volatilidad", f"Volatilidad ({_sym}/MWh)",
                             df.groupby("nodo")["pml"].std(), "#,##0.00"))
        if opciones.get("top_negativos"):
            rankings.append(("🔻 Top horas negativas", "% de horas con PML < 0",
                             df.groupby("nodo")["pml"].apply(lambda x: (x < 0).mean() * 100),
                             "0.0\"%\""))

    if rankings:
        ws_t = wb.create_sheet("🏆 Rankings")
        titulo(ws_t, "A1:F1", "Rankings comparativos")
        cur_row = 3
        for nombre_rk, col_valor, serie, formato in rankings:
            titulo(ws_t, f"A{cur_row}:F{cur_row}", nombre_rk, size=12, height=22, bg=C_RED)
            cur_row += 1
            for ci, h in enumerate(["Rank", "Nodo", col_valor], 1):
                hdr(ws_t.cell(row=cur_row, column=ci, value=h), bg=C_SUB)
            cur_row += 1
            top = serie.dropna().sort_values(ascending=False).head(10).round(2)
            for i, (nodo_n, val) in enumerate(top.items(), 1):
                ws_t.cell(row=cur_row, column=1, value=i)
                ws_t.cell(row=cur_row, column=2, value=str(nodo_n))
                ws_t.cell(row=cur_row, column=3, value=float(val))
                for ci in range(1, 4):
                    cc = ws_t.cell(row=cur_row, column=ci)
                    cc.font = _FONT_DATO
                    cc.border = BORDE
                    cc.alignment = _ALIGN_NUM
                ws_t.cell(row=cur_row, column=3).number_format = formato
                cur_row += 1
            cur_row += 1

        for ci, w in enumerate([8, 18, 24, 14, 14, 14], 1):
            ws_t.column_dimensions[get_column_letter(ci)].width = w

    # ─── SECCIÓN: Multi-año (1 hoja por nodo) ───
    if opciones.get("multiano") and df_multianos_dict:
        meses_lbl = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
                     'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
        for nodo_n, df_n in df_multianos_dict.items():
            if df_n is None or df_n.empty:
                continue
            ws_g = wb.create_sheet(f"📈 {nodo_n[:24]}"[:31])
            titulo(ws_g, "A1:N1", f"PML promedio mensual por año · {nodo_n}")

            hdr(ws_g.cell(row=3, column=1, value="Mes"), bg=C_SUB)
            años_cols = sorted([c for c in df_n.columns if c != 'mes_num'])
            for ci, año in enumerate(años_cols, start=2):
                hdr(ws_g.cell(row=3, column=ci, value=str(año)), bg=C_SUB)

            for ri, mes_idx in enumerate(range(1, 13), start=4):
                ws_g.cell(row=ri, column=1, value=meses_lbl[mes_idx-1])
                ws_g.cell(row=ri, column=1).font = _FONT_DATO
                ws_g.cell(row=ri, column=1).border = BORDE
                for ci, año in enumerate(años_cols, start=2):
                    v = df_n[df_n['mes_num'] == mes_idx][año].values
                    val = float(v[0]) if len(v) > 0 and pd.notna(v[0]) else None
                    cc = ws_g.cell(row=ri, column=ci, value=val)
                    cc.font = _FONT_DATO
                    cc.border = BORDE
                    cc.alignment = _ALIGN_NUM
                    cc.number_format = "#,##0.00"

            chart = LineChart()
            chart.title = f"PML promedio mensual ({_sym}/MWh) · {nodo_n}"
            chart.style = 12
            chart.height = 12
            chart.width = 22
            chart.y_axis.title = f"PML promedio ({_sym}/MWh)"
            chart.x_axis.title = "Mes"
            data_ref = Reference(ws_g, min_col=2, max_col=1+len(años_cols),
                                  min_row=3, max_row=15)
            cats_ref = Reference(ws_g, min_col=1, min_row=4, max_row=15)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            for s in chart.series:
                s.smooth = False
                s.marker = Marker(symbol="circle", size=8)
            ws_g.add_chart(chart, "A18")

            for ci, w in enumerate([10] + [12] * len(años_cols), 1):
                ws_g.column_dimensions[get_column_letter(ci)].width = w

    # ─── SECCIÓN: Spread por CCR ───
    if opciones.get("spread_ccr") and not df.empty:
        ccrs_unicos = sorted([c for c in df['ccr'].unique() if c and c != "?"])
        if len(ccrs_unicos) >= 2:
            ws_c = wb.create_sheet("📊 Spread CCR")
            titulo(ws_c, "A1:N1", f"PML promedio mensual por CCR · {len(ccrs_unicos)} zonas")

            df_w = df[['ccr', 'fecha_dt', 'pml']].copy()
            df_w['periodo'] = df_w['fecha_dt'].dt.to_period('M').dt.to_timestamp()
            monthly = df_w.groupby(['ccr', 'periodo'], observed=True)['pml'].mean().reset_index()
            pivot = monthly.pivot(index='periodo', columns='ccr', values='pml').round(2)

            hdr(ws_c.cell(row=3, column=1, value="Periodo"), bg=C_SUB)
            ccr_cols = list(pivot.columns)
            for ci, ccr in enumerate(ccr_cols, start=2):
                hdr(ws_c.cell(row=3, column=ci, value=str(ccr)), bg=C_SUB)

            for ri, (per, row_vals) in enumerate(pivot.iterrows(), start=4):
                ws_c.cell(row=ri, column=1, value=per.strftime("%b %Y"))
                ws_c.cell(row=ri, column=1).font = _FONT_DATO
                ws_c.cell(row=ri, column=1).border = BORDE
                for ci, ccr in enumerate(ccr_cols, start=2):
                    val = row_vals[ccr]
                    cc = ws_c.cell(row=ri, column=ci, value=float(val) if pd.notna(val) else None)
                    cc.font = _FONT_DATO
                    cc.border = BORDE
                    cc.alignment = _ALIGN_NUM
                    cc.number_format = "#,##0.00"

            chart_ccr = LineChart()
            chart_ccr.title = f"PML promedio mensual por CCR ({_sym}/MWh)"
            chart_ccr.style = 12
            chart_ccr.height = 14
            chart_ccr.width = 26
            chart_ccr.y_axis.title = f"PML promedio ({_sym}/MWh)"
            chart_ccr.x_axis.title = "Periodo"
            n_rows = len(pivot)
            data_ref = Reference(ws_c, min_col=2, max_col=1+len(ccr_cols),
                                  min_row=3, max_row=3+n_rows)
            cats_ref = Reference(ws_c, min_col=1, min_row=4, max_row=3+n_rows)
            chart_ccr.add_data(data_ref, titles_from_data=True)
            chart_ccr.set_categories(cats_ref)
            for s in chart_ccr.series:
                s.smooth = False
            ws_c.add_chart(chart_ccr, f"A{4+n_rows+2}")

            for ci, w in enumerate([14] + [14] * len(ccr_cols), 1):
                ws_c.column_dimensions[get_column_letter(ci)].width = w

    buffer = io.BytesIO()
    wb.save(buffer); buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════════════
# DATAFRAMES (cached para performance)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, max_entries=4)
def acumulado_a_dataframe_cached(acumulado_id, _acumulado, _catalogo, moneda, _fx_lookup):
    """Convierte acumulado (en MXN base) a DF, aplicando conversión USD si aplica.

    Args:
        acumulado_id: hash key para invalidar cache
        _acumulado: dict {nodo: [registros_en_MXN]}  ← SIEMPRE EN MXN
        _catalogo: catálogo de nodos
        moneda: "MXN" o "USD" — incluido en cache key
        _fx_lookup: dict {fecha: TC} para conversión a USD (None si MXN)
    """
    rows = []
    for nodo, filas in _acumulado.items():
        info = _catalogo.get(nodo, {}) if _catalogo else {}
        ccr = info.get("ccr", "?")
        nombre = info.get("nombre", nodo)
        estado = info.get("estado", "")
        for f in filas:
            try: pml_val = float(f["pml"])
            except (TypeError, ValueError): continue

            # Conversión USD si aplica — se hace AQUÍ, no in-place al acumulado
            if moneda == "USD" and _fx_lookup:
                fecha_str = f.get("fecha", "")
                tc = _fx_lookup.get(fecha_str)
                if tc and tc > 0:
                    pml_val = pml_val / tc
                else:
                    continue  # skip filas sin TC válido

            rows.append({
                "nodo":   nodo,
                "ccr":    ccr,
                "nombre": nombre,
                "estado": estado,
                "fecha":  f["fecha"],
                "hora":   int(f["hora"]) if f["hora"] else 0,
                "pml":    pml_val,
            })
    df = pd.DataFrame(rows)
    rows = None  # liberar memoria de la lista
    gc.collect()

    if not df.empty:
        # Convertir a categorical para ahorrar memoria (string columnas repetitivas)
        df["nodo"] = df["nodo"].astype('category')
        df["ccr"] = df["ccr"].astype('category')
        df["nombre"] = df["nombre"].astype('category')
        df["estado"] = df["estado"].astype('category')
        df["hora"] = df["hora"].astype('int8')
        df["pml"] = df["pml"].astype('float32')  # vs float64 por defecto = 50% menos memoria

        df["fecha_dt"] = pd.to_datetime(df["fecha"], errors="coerce")
        df["mes"] = df["fecha_dt"].dt.month.astype('int8')
        df["mes_nombre"] = df["fecha_dt"].dt.strftime("%b %Y").astype('category')

        # Drop la columna fecha string (redundante con fecha_dt) — ahorra mucha memoria
        df = df.drop(columns=["fecha"])
        gc.collect()
    return df


def acumulado_a_dataframe(acumulado, catalogo, moneda=None, fx_lookup=None):
    """Wrapper que toma moneda + fx_lookup desde session_state si no se pasan.

    El acumulado debe estar SIEMPRE en MXN crudo. La conversión a USD se hace
    al construir el DataFrame.
    """
    if not acumulado: return pd.DataFrame()

    # Default: leer de session_state si no se pasaron
    if moneda is None:
        moneda = st.session_state.get("moneda", "MXN")
    if fx_lookup is None:
        fx_info_state = st.session_state.get("fx_info", {})
        fx_lookup = fx_info_state.get("tc_lookup", {}) if isinstance(fx_info_state, dict) else {}

    # Hash key: # nodos + total filas + primer nodo + moneda (clave!)
    # La moneda DEBE estar en el cache key para que MXN y USD tengan caches separados
    key = (len(acumulado),
           sum(len(v) for v in acumulado.values()),
           list(acumulado.keys())[0] if acumulado else "",
           moneda)
    return acumulado_a_dataframe_cached(str(key), acumulado, catalogo, moneda, fx_lookup)


@st.cache_data(show_spinner=False, max_entries=3)
def calcular_resumen_cached(df_hash, _df):
    if _df.empty: return pd.DataFrame()
    summary = _df.groupby("nodo").agg(
        nombre=("nombre", "first"),
        ccr=("ccr", "first"),
        registros=("pml", "count"),
        promedio=("pml", "mean"),
        mediana=("pml", "median"),
        maximo=("pml", "max"),
        minimo=("pml", "min"),
        std=("pml", "std"),
    ).round(2)
    pct_neg = _df.groupby("nodo")["pml"].apply(
        lambda x: (x < 0).sum() / len(x) * 100 if len(x) > 0 else 0
    ).round(1)
    summary["% horas neg"] = pct_neg
    summary["p95"] = _df.groupby("nodo")["pml"].quantile(0.95).round(2)
    summary["p5"] = _df.groupby("nodo")["pml"].quantile(0.05).round(2)
    return summary.reset_index().sort_values("promedio", ascending=False).reset_index(drop=True)


def calcular_resumen(df):
    if df.empty: return pd.DataFrame()
    key = f"{len(df)}-{df['pml'].sum():.0f}"
    return calcular_resumen_cached(key, df)


# Layout estándar para gráficas
def _layout_estandar(titulo, height=480):
    return dict(
        title=dict(
            text=titulo,
            font=dict(family="Arial Black", size=16, color=TEXT_TITLE),
            x=0.02, xanchor="left",
        ),
        height=height,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Arial", size=12, color=TEXT_DARK),
        margin=dict(l=70, r=40, t=70, b=60),
        hoverlabel=dict(bgcolor="white", font_size=12, font_color=TEXT_DARK,
                        bordercolor=AXIS_LINE),
        uirevision='constant',  # mantiene zoom/pan al actualizar
    )


def _ejes_estandar(fig, x_title=None, y_title=None):
    fig.update_xaxes(
        showgrid=True, gridcolor=GRID_LIGHT,
        linecolor=AXIS_LINE, linewidth=1.5,
        tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
        title=dict(text=x_title or "", font=dict(family="Arial", size=13, color=TEXT_DARK)),
    )
    fig.update_yaxes(
        showgrid=True, gridcolor=GRID_LIGHT,
        linecolor=AXIS_LINE, linewidth=1.5,
        tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
        title=dict(text=y_title or "", font=dict(family="Arial", size=13, color=TEXT_DARK)),
        zeroline=True, zerolinecolor=RE_RED, zerolinewidth=1.5,
    )
    return fig


def grafica_lineas_tiempo(df, max_nodos=15):
    if df.empty: return None
    promedios = df.groupby("nodo")["pml"].mean().sort_values(ascending=False)
    nodos_mostrar = promedios.head(max_nodos).index.tolist()
    df_plot = df[df["nodo"].isin(nodos_mostrar)].copy()
    daily = df_plot.groupby(["nodo", "fecha_dt"])["pml"].mean().reset_index()
    fig = px.line(
        daily, x="fecha_dt", y="pml", color="nodo",
        color_discrete_sequence=PALETTE,
    )
    fig.update_traces(line=dict(width=2.5), mode='lines')
    fig.update_layout(**_layout_estandar(
        f"Evolución temporal · Top {len(nodos_mostrar)} nodos por PML promedio"))
    fig.update_layout(
        hovermode="x unified",
        legend=dict(
            orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor=AXIS_LINE, borderwidth=1,
            font=dict(family="Arial", size=11, color=TEXT_DARK),
            title=dict(text="<b>Nodo</b>", font=dict(color=TEXT_DARK)),
        ),
        margin=dict(l=70, r=200, t=70, b=60),
    )
    return _ejes_estandar(fig, "Fecha", f"PML promedio diario {label_moneda()}")


def grafica_heatmap_horario(df, nodo_seleccionado):
    if df.empty: return None
    sub = df[df["nodo"] == nodo_seleccionado].copy()
    if sub.empty: return None
    pivot = sub.pivot_table(values="pml", index="hora", columns="mes_nombre",
                              aggfunc="mean").round(1)
    if not pivot.empty:
        meses_orden = sub.sort_values("fecha_dt")["mes_nombre"].drop_duplicates().tolist()
        pivot = pivot.reindex(columns=meses_orden)

    colorscale = [
        [0.0, "#1a5490"], [0.25, "#5ba0d4"], [0.5, "#FFFFFF"],
        [0.75, "#f5a653"], [1.0, "#a0090c"],
    ]
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values, x=pivot.columns, y=pivot.index,
        colorscale=colorscale,
        zmid=pivot.values.mean() if pivot.size > 0 else 0,
        colorbar=dict(
            title=dict(text=f"{label_moneda_short()}/MWh",
                       font=dict(color=TEXT_DARK, size=12)),
            thickness=15, len=0.85,
            tickfont=dict(color=TEXT_DARK, size=11),
        ),
        hovertemplate=(f"<b>Hora:</b> %{{y}}h<br><b>Mes:</b> %{{x}}<br>"
                       f"<b>PML:</b> {label_moneda_short()}%{{z}}<extra></extra>"),
        text=pivot.values,
        texttemplate="%{text:.0f}",
        textfont=dict(size=10, color=TEXT_DARK, family="Arial Black"),
    ))
    fig.update_layout(**_layout_estandar(
        f"Heatmap PML hora × mes · {nodo_seleccionado}", height=540))
    fig.update_yaxes(autorange="reversed", dtick=2,
                      linecolor=AXIS_LINE, linewidth=1.5,
                      tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
                      title=dict(text="Hora del día", font=dict(family="Arial", size=13, color=TEXT_DARK)))
    fig.update_xaxes(linecolor=AXIS_LINE, linewidth=1.5,
                      tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
                      title=dict(text="Mes", font=dict(family="Arial", size=13, color=TEXT_DARK)))
    return fig


def grafica_pml_multiano(df, nodo_seleccionado):
    """Gráfica de PML promedio mensual por año.

    Para el nodo seleccionado, calcula el PML promedio de cada mes en cada año
    presente en los datos. Genera puntos conectados con colores corporativos.
    """
    if df.empty: return None, []
    sub = df[df["nodo"] == nodo_seleccionado].copy()
    if sub.empty: return None, []

    sub["año"] = sub["fecha_dt"].dt.year
    sub["mes_num"] = sub["fecha_dt"].dt.month

    años_disponibles = sorted(sub["año"].unique())
    if len(años_disponibles) < 1:
        return None, []

    # Promedio mensual por año
    pivot = sub.pivot_table(values="pml", index="mes_num", columns="año",
                              aggfunc="mean").round(1)

    # Reindexar para tener todos los 12 meses
    pivot = pivot.reindex(range(1, 13))

    meses_labels = ['ene', 'feb', 'mar', 'abr', 'may', 'jun',
                    'jul', 'ago', 'sep', 'oct', 'nov', 'dic']

    # Colores corporativos por año (rotando paleta)
    color_anos = {
        años_disponibles[0]: RE_RED,
    }
    if len(años_disponibles) >= 2:
        color_anos[años_disponibles[1]] = "#5ba0d4"  # azul claro RE
    if len(años_disponibles) >= 3:
        color_anos[años_disponibles[2]] = RE_NAVY
    if len(años_disponibles) >= 4:
        color_anos[años_disponibles[3]] = "#d4a017"
    # Más años: rotar paleta
    for i, year in enumerate(años_disponibles[4:]):
        color_anos[year] = PALETTE[(i + 4) % len(PALETTE)]

    fig = go.Figure()
    promedios_anuales = {}

    for año in años_disponibles:
        if año not in pivot.columns:
            continue
        valores = pivot[año]
        promedio_anual = valores.mean()
        promedios_anuales[año] = promedio_anual

        sym = label_moneda_short()
        nombre_traza = f"PML Avg {año}: {sym}{promedio_anual:.0f}/MWh"

        fig.add_trace(go.Scatter(
            x=meses_labels,
            y=valores.values,
            mode='lines+markers',
            name=nombre_traza,
            line=dict(color=color_anos[año], width=2, dash='dot'),
            marker=dict(color=color_anos[año], size=10,
                        line=dict(color='white', width=1.5)),
            hovertemplate=f"<b>{año}</b><br>%{{x}}: {sym}%{{y:.1f}}<extra></extra>",
        ))

    fig.update_layout(**_layout_estandar(
        f"📊 PML promedio mensual · {nodo_seleccionado}", height=520))

    fig.update_layout(
        legend=dict(
            orientation="v", yanchor="top", y=0.98, xanchor="right", x=0.98,
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor=AXIS_LINE, borderwidth=1.5,
            font=dict(family="Arial", size=12, color=TEXT_DARK),
            itemsizing='constant',  # mantiene el dot del color visible
        ),
        margin=dict(l=70, r=40, t=70, b=60),
    )
    return _ejes_estandar(fig, "Mes", f"PML promedio {label_moneda()}"), años_disponibles


def grafica_spread_ccr(df, granularidad="mensual"):
    """Gráfica de spread por CCR — una línea continua por zona en el tiempo.

    Útil cuando hay nodos de varios CCRs en la consulta.
    Promedia todos los nodos de cada CCR por mes.

    Retorna (fig, ccrs_disponibles).
    """
    if df.empty:
        return None, []

    ccrs_unicos = sorted([c for c in df['ccr'].unique() if c and c != "?"])
    if len(ccrs_unicos) < 2:
        return None, ccrs_unicos

    # Pre-agregar a mensual para reducir memoria y mejorar legibilidad
    df_work = df[['ccr', 'fecha_dt', 'pml']].copy()
    df_work['periodo'] = df_work['fecha_dt'].dt.to_period('M').dt.to_timestamp()
    monthly = df_work.groupby(['ccr', 'periodo'])['pml'].mean().reset_index()

    # Asignar colores corporativos por CCR (consistente con paleta)
    color_ccr = {ccr: PALETTE[i % len(PALETTE)] for i, ccr in enumerate(ccrs_unicos)}

    fig = go.Figure()
    sym = label_moneda_short()

    for ccr in ccrs_unicos:
        sub = monthly[monthly['ccr'] == ccr].sort_values('periodo')
        if sub.empty:
            continue
        promedio = sub['pml'].mean()
        fig.add_trace(go.Scatter(
            x=sub['periodo'],
            y=sub['pml'],
            mode='lines+markers',
            name=f"{ccr} (avg {sym}{promedio:.0f})",
            line=dict(color=color_ccr[ccr], width=2.5),
            marker=dict(color=color_ccr[ccr], size=7,
                        line=dict(color='white', width=1)),
            hovertemplate=f"<b>{ccr}</b><br>%{{x|%b %Y}}: {sym}%{{y:.1f}}<extra></extra>",
        ))

    fig.update_layout(**_layout_estandar(
        f"📊 Spread mensual por CCR · {len(ccrs_unicos)} zonas comparadas", height=520))
    fig.update_layout(
        legend=dict(
            orientation="v", yanchor="top", y=0.98, xanchor="right", x=0.98,
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor=AXIS_LINE, borderwidth=1.5,
            font=dict(family="Arial", size=11, color=TEXT_DARK),
        ),
        hovermode='x unified',
    )
    return _ejes_estandar(fig, "Mes",
                          f"PML promedio mensual {label_moneda()}"), ccrs_unicos


def grafica_barras_top(df, metrica="promedio", top_n=10):
    if df.empty: return None
    summary = df.groupby("nodo")["pml"].agg(["mean", "std", "max", "min"])
    summary["pct_neg"] = df.groupby("nodo")["pml"].apply(lambda x: (x < 0).sum() / len(x) * 100)
    summary = summary.reset_index()
    if metrica == "promedio":
        col, titulo, color_bar = "mean", f"Top {top_n} · Mayor PML promedio", RE_NAVY
    elif metrica == "volatilidad":
        col, titulo, color_bar = "std", f"Top {top_n} · Mayor volatilidad", RE_RED
    elif metrica == "negativos":
        col, titulo, color_bar = "pct_neg", f"Top {top_n} · Más horas negativas", RE_BLUE
    else: return None
    top = summary.nlargest(top_n, col).round(2)

    text_template = "%{x:.2f}" if metrica != "negativos" else "%{x:.1f}%"
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=top[col], y=top["nodo"],
        orientation='h',
        marker=dict(color=color_bar, line=dict(color=TEXT_DARK, width=1)),
        text=top[col],
        texttemplate=text_template,
        textposition='outside',
        textfont=dict(family="Arial Black", size=11, color=TEXT_DARK),
        hovertemplate=f"<b>%{{y}}</b><br>{metrica.capitalize()}: %{{x:.2f}}<extra></extra>",
    ))
    fig.update_layout(**_layout_estandar(titulo, height=440))
    fig.update_layout(yaxis=dict(autorange="reversed"))
    x_title = metrica.capitalize() + f" {label_moneda()}" if metrica != "negativos" else "% de horas"
    return _ejes_estandar(fig, x_title, "")


# ═══════════════════════════════════════════════════════════════════════
# MAPA DE PML (nodos geolocalizados con OpenStreetMap) Y PML POR CCR
# ═══════════════════════════════════════════════════════════════════════
def grafica_mapa(matches_df, df_pml):
    """Mapa de los nodos geolocalizados, coloreados por su PML promedio."""
    if matches_df is None or matches_df.empty or df_pml is None or df_pml.empty:
        return None
    matches_ok = matches_df[matches_df['lat'].notna()].copy()
    if matches_ok.empty:
        return None

    pml_avg = (df_pml.groupby('nodo', observed=True)['pml']
               .agg(['mean', 'count', 'max', 'min']).reset_index())
    pml_avg.columns = ['clave', 'pml_avg', 'pml_count', 'pml_max', 'pml_min']
    pml_avg['clave'] = pml_avg['clave'].astype(str)
    map_df = matches_ok.merge(pml_avg, on='clave', how='left')
    map_df = map_df[map_df['pml_avg'].notna()]
    if map_df.empty:
        return None

    sym = label_moneda_short()
    map_df['hover'] = map_df.apply(
        lambda r: (
            f"<b>{r['clave']}</b><br>"
            f"<b>Nombre:</b> {r['nombre_cenace']}<br>"
            f"<b>CCR:</b> {r['ccr']} · {r['zona']}<br>"
            f"<b>Voltaje:</b> {r['kv_cenace']} kV<br>"
            f"<b>Estado:</b> {r['estado']}<br>"
            f"<b>Municipio:</b> {r['municipio']}<br>"
            f"━━━━━━━━━━━<br>"
            f"<b>PML promedio:</b> {sym}{r['pml_avg']:.2f}<br>"
            f"<b>PML máx:</b> {sym}{r['pml_max']:.2f}<br>"
            f"<b>PML mín:</b> {sym}{r['pml_min']:.2f}<br>"
            f"━━━━━━━━━━━<br>"
            f"<b>Ubicación OSM:</b> {r['calidad']} (sim {r['similitud']})<br>"
            f"<b>Subestación OSM:</b> {r['nombre_osm']}"
        ), axis=1
    )
    colorscale = [
        [0.0, "#1a8a3a"], [0.25, "#9bc24f"], [0.5, "#f5d017"],
        [0.75, "#f57f17"], [1.0, "#a0090c"],
    ]
    colorbar = dict(
        title=dict(text=f"<b>PML promedio<br>{label_moneda()}</b>",
                   font=dict(family="Arial", size=12, color=TEXT_DARK)),
        thickness=18, len=0.7, x=1.02,
        tickfont=dict(family="Arial", size=11, color=TEXT_DARK),
    )

    # Auto-zoom
    lat_min, lat_max = map_df['lat'].min(), map_df['lat'].max()
    lon_min, lon_max = map_df['lon'].min(), map_df['lon'].max()
    spread = max(lat_max - lat_min, lon_max - lon_min)
    if spread > 15: zoom = 4
    elif spread > 8: zoom = 5
    elif spread > 4: zoom = 6
    elif spread > 2: zoom = 7
    elif spread > 1: zoom = 8
    else: zoom = 9

    fig = go.Figure()
    # Capa 1: halo blanco
    fig.add_trace(go.Scattermap(
        lat=map_df['lat'], lon=map_df['lon'], mode='markers',
        marker=dict(size=28, color='white', opacity=0.95),
        hoverinfo='skip', showlegend=False,
    ))
    # Capa 2: marcador coloreado por PML promedio
    fig.add_trace(go.Scattermap(
        lat=map_df['lat'], lon=map_df['lon'], mode='markers+text',
        marker=dict(
            size=20, color=map_df['pml_avg'],
            colorscale=colorscale,
            cmin=map_df['pml_avg'].min(), cmax=map_df['pml_avg'].max(),
            colorbar=colorbar, showscale=True, opacity=0.95,
        ),
        text=map_df['clave'], textposition="top right",
        textfont=dict(size=11, color=TEXT_TITLE, family="Open Sans Bold"),
        hovertext=map_df['hover'], hoverinfo='text',
        name='Nodos', showlegend=False,
    ))
    # Capa 3: punto central
    fig.add_trace(go.Scattermap(
        lat=map_df['lat'], lon=map_df['lon'], mode='markers',
        marker=dict(size=4, color=TEXT_DARK, opacity=1.0),
        hoverinfo='skip', showlegend=False,
    ))

    fig.update_layout(
        map=dict(
            style='open-street-map',
            center=dict(lat=(lat_min + lat_max) / 2, lon=(lon_min + lon_max) / 2),
            zoom=zoom,
        ),
        height=680,
        margin=dict(l=0, r=0, t=50, b=0),
        title=dict(
            text=f"📍 PML promedio por nodo · {len(map_df)} nodos geolocalizados",
            font=dict(family="Arial Black", size=16, color=TEXT_TITLE),
            x=0.02, xanchor="left",
        ),
        showlegend=False,
        paper_bgcolor="white",
        hoverlabel=dict(bgcolor="white", font_size=12, font_color=TEXT_DARK,
                        bordercolor=AXIS_LINE),
        uirevision='constant',
    )
    return fig


def render_panel_ccr(df):
    """Tabla de PML por CCR, a partir del promedio de cada nodo con datos."""
    if df is None or df.empty:
        return
    prom_nodo = df.groupby(['ccr', 'nodo'], observed=True)['pml'].mean().reset_index()
    ccr_stats = prom_nodo.groupby('ccr', observed=True).agg(
        nodos=('nodo', 'count'),
        pml_promedio=('pml', 'mean'),
        pml_max=('pml', 'max'),
        pml_min=('pml', 'min'),
    ).round(2).reset_index().sort_values('pml_promedio', ascending=False)

    st.markdown("### 📊 PML promedio por CCR")
    st.dataframe(
        ccr_stats, width="stretch", hide_index=True,
        column_config={
            "ccr":          st.column_config.TextColumn("CCR", width="medium"),
            "nodos":        st.column_config.NumberColumn("# Nodos", format="%d"),
            "pml_promedio": st.column_config.NumberColumn("PML promedio", format=fmt_moneda()),
            "pml_max":      st.column_config.NumberColumn("Nodo más caro", format=fmt_moneda()),
            "pml_min":      st.column_config.NumberColumn("Nodo más barato", format=fmt_moneda()),
        },
    )


# ═══════════════════════════════════════════════════════════════════════
# RENDER DASHBOARD COMPLETO (sin Excel - solo análisis)
# ═══════════════════════════════════════════════════════════════════════
def render_dashboard(acumulado, catalogo, matches_df=None):
    st.divider()
    st.markdown("## 📊 Dashboard analítico")
    st.caption("Pasa el mouse sobre los gráficos para ver valores específicos.")

    df = acumulado_a_dataframe(acumulado, catalogo)
    if df.empty:
        st.warning("No hay datos numéricos válidos.")
        return

    col1, col2, col3, col4 = st.columns(4)
    pml_global = df["pml"].mean()
    pml_max = df["pml"].max()
    pct_neg = (df["pml"] < 0).sum() / len(df) * 100
    nodo_top = df.groupby("nodo")["pml"].mean().idxmax()
    sym = label_moneda_short()
    col1.metric("PML promedio global", f"{sym}{pml_global:.2f}")
    col2.metric("PML máximo", f"{sym}{pml_max:.2f}")
    col3.metric("% horas negativas", f"{pct_neg:.1f}%")
    col4.metric("Nodo top", nodo_top)

    # MAPA DE PML
    if matches_df is not None and not matches_df.empty:
        st.markdown("### 🗺️ Mapa de PML por nodo")
        n_mapeados = matches_df['lat'].notna().sum()
        n_consultados = len(matches_df)
        if n_mapeados == 0:
            st.warning(f"❌ Ninguno de los {n_consultados} nodos pudo ubicarse en OpenStreetMap.")
        else:
            cm1, cm2, cm3 = st.columns(3)
            cm1.metric("Ubicados", f"{n_mapeados}/{n_consultados}")
            cc = matches_df['calidad'].value_counts()
            n_exc = cc.get('🥇 Excelente', 0)
            n_bue = cc.get('🥈 Bueno', 0)
            n_ace = cc.get('🥉 Aceptable', 0)
            cm2.metric("🥇 Excelente / 🥈 Bueno", f"{n_exc} / {n_bue}")
            cm3.metric("🥉 Aceptable", n_ace)
            st.caption("Ubicación aproximada: nombre del nodo CENACE emparejado con "
                       "subestaciones de OpenStreetMap (no es la georreferencia oficial).")

            fig_mapa = grafica_mapa(matches_df, df)
            if fig_mapa:
                st.plotly_chart(fig_mapa, width="stretch",
                                config={'scrollZoom': True, 'displayModeBar': True})

            sin_match = matches_df[matches_df['lat'].isna()]
            if not sin_match.empty:
                with st.expander(f"❌ {len(sin_match)} nodos sin ubicar"):
                    st.dataframe(
                        sin_match[['clave', 'nombre_cenace', 'ccr', 'estado', 'razon']],
                        width="stretch", hide_index=True,
                    )

    # PML POR CCR
    render_panel_ccr(df)

    # TABLA RESUMEN
    st.markdown("### 📋 Tabla resumen por nodo")
    summary = calcular_resumen(df)
    if not summary.empty:
        st.dataframe(
            summary, width="stretch", hide_index=True,
            column_config={
                "nodo":        st.column_config.TextColumn("Clave", width="small"),
                "nombre":      st.column_config.TextColumn("Nombre", width="medium"),
                "ccr":         st.column_config.TextColumn("CCR", width="small"),
                "registros":   st.column_config.NumberColumn("# Reg", format="%d"),
                "promedio":    st.column_config.NumberColumn("Promedio", format=fmt_moneda()),
                "mediana":     st.column_config.NumberColumn("Mediana", format=fmt_moneda()),
                "maximo":      st.column_config.NumberColumn("Máximo", format=fmt_moneda()),
                "minimo":      st.column_config.NumberColumn("Mínimo", format=fmt_moneda()),
                "std":         st.column_config.NumberColumn("Volatilidad", format=fmt_moneda()),
                "p95":         st.column_config.NumberColumn("P95", format=fmt_moneda()),
                "p5":          st.column_config.NumberColumn("P5", format=fmt_moneda()),
                "% horas neg": st.column_config.NumberColumn("% Neg", format="%.1f%%"),
            },
        )

    st.markdown("### 📈 Evolución temporal del PML")
    fig_lin = grafica_lineas_tiempo(df, max_nodos=min(15, len(acumulado)))
    if fig_lin: st.plotly_chart(fig_lin, width="stretch")

    st.markdown("### 🔥 Heatmap horario × mensual")
    nodos_disp = sorted(df["nodo"].unique())
    nodo_h = st.selectbox("Selecciona nodo:", nodos_disp, index=0, key="sel_heatmap")
    fig_heat = grafica_heatmap_horario(df, nodo_h)
    if fig_heat: st.plotly_chart(fig_heat, width="stretch")

    # ─── PML PROMEDIO MULTI-AÑO ───
    años_total = df['fecha_dt'].dt.year.nunique() if 'fecha_dt' in df.columns else 1
    st.markdown("### 📊 PML promedio mensual por año")
    if años_total < 1:
        st.caption("Sin datos suficientes para gráfica multi-año.")
    else:
        nodo_multi = st.selectbox(
            "Selecciona nodo para ver evolución mensual por año:",
            nodos_disp,
            index=nodos_disp.index(nodo_h) if nodo_h in nodos_disp else 0,
            key="sel_multiano",
            help="Muestra promedio mensual del PML del nodo, comparando años disponibles."
        )
        fig_multi, años_disp = grafica_pml_multiano(df, nodo_multi)
        if fig_multi:
            st.plotly_chart(fig_multi, width="stretch")
            if len(años_disp) == 1:
                st.caption(f"ℹ️ Solo hay datos del año {años_disp[0]}. "
                           f"Para comparar años, consulta un período que cubra más de un año natural.")
        else:
            st.warning("No se pudo generar la gráfica multi-año.")

    # ─── SPREAD POR CCR (solo si ≥2 CCRs distintos) ───
    ccrs_unicos_dash = sorted([c for c in df['ccr'].unique() if c and c != "?"])
    if len(ccrs_unicos_dash) >= 2:
        st.markdown("### 📊 Spread mensual por CCR (zonas)")
        st.caption(
            f"Comparativa continua del PML promedio mensual entre **{len(ccrs_unicos_dash)} CCRs** "
            f"presentes en tu consulta. Útil para identificar diferenciales entre zonas."
        )
        fig_spread, _ = grafica_spread_ccr(df)
        if fig_spread:
            st.plotly_chart(fig_spread, width="stretch")

    st.markdown("### 🏆 Rankings comparativos")
    tab1, tab2, tab3 = st.tabs(["Mayor PML", "Mayor volatilidad", "Más horas negativas"])
    with tab1:
        fig = grafica_barras_top(df, "promedio", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, width="stretch")
    with tab2:
        fig = grafica_barras_top(df, "volatilidad", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, width="stretch")
    with tab3:
        fig = grafica_barras_top(df, "negativos", top_n=min(10, len(acumulado)))
        if fig: st.plotly_chart(fig, width="stretch")


# ═══════════════════════════════════════════════════════════════════════
# RENDER CENTRO DE DESCARGAS (modo Datos)
# ═══════════════════════════════════════════════════════════════════════
def _pivot_multianos(df):
    """{nodo: DataFrame mes_num × año} con el PML promedio mensual de cada nodo."""
    out = {}
    for nodo_n in sorted(df["nodo"].unique()):
        sub = df[df["nodo"] == nodo_n]
        if sub.empty:
            continue
        sub = sub.assign(año=sub["fecha_dt"].dt.year, mes_num=sub["fecha_dt"].dt.month)
        pivot = sub.pivot_table(values="pml", index="mes_num", columns="año",
                                aggfunc="mean").round(2)
        out[str(nodo_n)] = pivot.reset_index()
    return out


def render_centro_descargas(acumulado, catalogo, sistema, proceso, fecha_ini, fecha_fin,
                            moneda="MXN", tc_info=None):
    st.divider()
    st.markdown("## 📥 Centro de descargas")
    moneda_label = f"💱 Análisis en **{moneda}**"
    if moneda == "USD" and tc_info and tc_info.get("tc_promedio"):
        moneda_label += f" · TC FIX promedio: **{tc_info['tc_promedio']:.4f} MXN/USD**"
    st.caption(f"Elige los archivos que necesites. {moneda_label}")
    st.caption("⚡ Los archivos se generan solo cuando das click en cada botón.")

    df = acumulado_a_dataframe(acumulado, catalogo)
    df_resumen = calcular_resumen(df) if not df.empty else pd.DataFrame()
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    sufijo = f"_{moneda}"
    n_nodos = len(acumulado)

    if n_nodos < 20:
        desc_datos = f"~{n_nodos*150} KB · ~5s · 1 hoja por nodo + portada + resumen"
    elif n_nodos <= 50:
        desc_datos = f"~{n_nodos*200} KB · ~30s · 1 hoja por nodo"
    elif n_nodos <= 100:
        desc_datos = f"~{n_nodos*250} KB · 1-2 min · puede ser pesado al abrir"
    else:
        desc_datos = f"~{n_nodos*300} KB · 2-5 min · ⚠️ archivo muy pesado"

    n_nodos_df = df["nodo"].nunique() if not df.empty else 0
    ccrs = sorted([c for c in df['ccr'].unique() if c and c != "?"]) if not df.empty else []
    multiano_ok = 0 < n_nodos_df <= 20
    spread_ok = len(ccrs) >= 2

    st.markdown("##### 📦 Archivos disponibles")
    col1, col2 = st.columns(2, gap="medium")

    with col1:
        st.markdown("**📊 Excel de Datos**")
        st.caption("PML horario y sus componentes (energía, pérdidas, congestión).")
        st.caption(desc_datos)
        if st.button("⚡ Generar Excel de Datos", key="btn_gen_datos", type="primary",
                     width="stretch"):
            with st.spinner(f"Generando Excel de datos ({n_nodos} nodos)..."):
                st.session_state["excel_datos_bytes"] = generar_excel_datos(
                    acumulado, sistema, proceso, fecha_ini, fecha_fin,
                    moneda=moneda, tc_info=tc_info)
        if st.session_state.get("excel_datos_bytes"):
            st.download_button(
                label="📥 Descargar Excel de Datos",
                data=st.session_state["excel_datos_bytes"],
                file_name=f"PML_CENACE_Datos{sufijo}_{sistema}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_datos",
                width="stretch",
            )

    with col2:
        st.markdown("**📈 Excel de Análisis**")
        st.caption("Elige las secciones; cada una es una hoja del archivo.")
        opciones = {
            "resumen": st.toggle("📋 Resumen estadístico", value=True, key="xa_resumen"),
            "top_pml": st.toggle("🏆 Top PML promedio", value=True, key="xa_top_pml"),
            "top_volatilidad": st.toggle("📊 Top volatilidad", value=True, key="xa_top_vol"),
            "top_negativos": st.toggle("🔻 Top horas negativas", value=True, key="xa_top_neg"),
            "multiano": st.toggle(
                f"📈 Multi-año ({n_nodos_df} hojas, una por nodo)" if multiano_ok
                else "📈 Multi-año (solo con ≤20 nodos)",
                value=multiano_ok, disabled=not multiano_ok, key="xa_multi"),
            "spread_ccr": st.toggle(
                f"📊 Spread por CCR ({len(ccrs)} zonas)" if spread_ok
                else "📊 Spread por CCR (necesita ≥2 CCRs)",
                value=spread_ok, disabled=not spread_ok, key="xa_spread"),
        }
        opciones["multiano"] = opciones["multiano"] and multiano_ok
        opciones["spread_ccr"] = opciones["spread_ccr"] and spread_ok
        n_sec = sum(1 for v in opciones.values() if v)

        if st.button("⚡ Generar Excel de Análisis", key="btn_gen_anal", type="primary",
                     width="stretch", disabled=(n_sec == 0 or df.empty)):
            with st.spinner(f"Generando Excel de análisis ({n_sec} secciones)..."):
                st.session_state["excel_analisis_bytes"] = generar_excel_analisis(
                    df, df_resumen, opciones, sistema, proceso, fecha_ini, fecha_fin,
                    moneda=moneda, tc_info=tc_info,
                    df_multianos_dict=_pivot_multianos(df) if opciones["multiano"] else None,
                )
        if n_sec == 0:
            st.caption("ℹ️ Activa al menos una sección.")
        if st.session_state.get("excel_analisis_bytes"):
            st.download_button(
                label="📥 Descargar Excel de Análisis",
                data=st.session_state["excel_analisis_bytes"],
                file_name=f"PML_CENACE_Analisis{sufijo}_{sistema}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_anal",
                width="stretch",
            )


# ═══════════════════════════════════════════════════════════════════════
# UI PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <h1>⚡ PML CENACE
        <span class="pml-badge">MEM</span>
    </h1>
    <p>Descarga y análisis de Precios Marginales Locales · MTR y MDA · MXN o USD (FIX Banxico) — Recurrent Energy / Canadian Solar</p>
</div>
""", unsafe_allow_html=True)

catalogo, catalogo_version, catalogo_subido = get_catalogo_activo()
if catalogo:
    origen = "📤 subido en esta sesión" if catalogo_subido else "📦 del repositorio"
    version_txt = f" · {catalogo_version}" if catalogo_version else ""
    st.info(f"📚 Catálogo de NodosP del CENACE: **{len(catalogo):,} nodos**{version_txt} · {origen}")
else:
    st.warning("⚠️ Catálogo no cargado.")

# Inicializar session_state
if "consulta_ejecutada" not in st.session_state:
    st.session_state["consulta_ejecutada"] = False
if "acumulado" not in st.session_state:
    st.session_state["acumulado"] = {}
if "matches_df" not in st.session_state:
    st.session_state["matches_df"] = None
if "consulta_params" not in st.session_state:
    st.session_state["consulta_params"] = {}

with st.sidebar:
    st.markdown(
        f"<h2 style='color:{RE_BLUE};margin-bottom:0;font-weight:700;'>"
        f"⚡ Configuración</h2>",
        unsafe_allow_html=True,
    )
    st.caption("Consulta de PML")

    st.markdown("**🚀 Modo de uso**")
    modo = st.radio(
        "Selecciona qué hacer:",
        options=["🔬 Análisis", "📊 Datos"],
        index=0,
        help=(
            "**Análisis**: descarga + dashboard (KPIs, mapa, tablas y gráficas)\n\n"
            "**Datos**: descarga + Centro de descargas (Excel de datos y de análisis)"
        ),
        label_visibility="collapsed",
    )

    es_solo_datos = (modo == "📊 Datos")
    es_completo = (modo == "🔬 Análisis")
    necesita_mapa = es_completo   # el mapa de PML solo se arma en modo Análisis

    # FX Selector
    st.markdown("---")
    st.markdown("**💱 FX Selector**")
    moneda_sel = st.radio(
        "Moneda:",
        options=["🇲🇽 MXN (pesos)", "🇺🇸 USD (dólares)"],
        index=0,
        label_visibility="collapsed",
        help=(
            "**MXN**: PML directo del CENACE en pesos.\n\n"
            "**USD**: convertido día por día con el tipo de cambio "
            "FIX de Banxico (serie SF43718)."
        ),
    )
    st.session_state["moneda"] = "USD" if "USD" in moneda_sel else "MXN"

    st.markdown("---")
    st.markdown("**📍 Nodos CENACE**")

    # Procesar pending del botón "Agregar al textarea" (debe ir ANTES del text_area)
    if "nodos_pending" in st.session_state and st.session_state["nodos_pending"]:
        # Mover de pending al input
        st.session_state["nodos_input"] = st.session_state["nodos_pending"]
        st.session_state["nodos_pending"] = ""

    nodos_input = st.text_area(
        "Lista de claves",
        height=120,
        placeholder="01VAJ-230\n01XAL-230\n06FUN-115",
        key="nodos_input",
    )

    # Badge si los nodos vinieron del filtro
    if st.session_state.get("nodos_from_filter"):
        st.caption(f"📍 **{st.session_state['nodos_from_filter']} nodos del filtro**")

    # ─── SELECTOR JERÁRQUICO ───
    with st.expander("🔍 Buscar nodos por filtros (Estado / Zona / CCR / kV)"):
        if catalogo:
            cat_df = pd.DataFrame.from_dict(catalogo, orient='index').reset_index()
            cat_df = cat_df.rename(columns={'index': 'clave'})

            sistemas_disp = sorted(cat_df['sistema'].dropna().unique())
            sistemas_disp = [s for s in sistemas_disp if s]
            sis_sel = st.selectbox("🌎 Sistema", ["Todos"] + sistemas_disp, key="filt_sis")
            df_f = cat_df.copy() if sis_sel == "Todos" else cat_df[cat_df['sistema'] == sis_sel]

            estados_disp = sorted([e for e in df_f['estado'].dropna().unique() if e])
            est_sel = st.multiselect("🏛️ Estado(s)", estados_disp, key="filt_est")
            if est_sel:
                df_f = df_f[df_f['estado'].isin(est_sel)]

            municipios_disp = sorted([m for m in df_f['municipio'].dropna().unique() if m])
            mun_sel = st.multiselect("🏘️ Municipio(s)", municipios_disp, key="filt_mun")
            if mun_sel:
                df_f = df_f[df_f['municipio'].isin(mun_sel)]

            zonas_disp = sorted([z for z in df_f['zona'].dropna().unique() if z])
            zona_sel = st.multiselect("🗺️ Zona CENACE", zonas_disp, key="filt_zona")
            if zona_sel:
                df_f = df_f[df_f['zona'].isin(zona_sel)]

            ccr_disp = sorted([c for c in df_f['ccr'].dropna().unique() if c])
            ccr_sel = st.multiselect("📍 CCR (Centro Control Regional)", ccr_disp, key="filt_ccr")
            if ccr_sel:
                df_f = df_f[df_f['ccr'].isin(ccr_sel)]

            kvs_disp = sorted([int(k) for k in df_f['kv'].dropna().unique() if k > 0])
            kv_sel = st.multiselect("⚡ Voltaje (kV)", kvs_disp, key="filt_kv")
            if kv_sel:
                df_f = df_f[df_f['kv'].isin(kv_sel)]

            st.markdown(f"**📋 {len(df_f):,} nodos coinciden**")

            if len(df_f) > 0:
                # Buscador opcional
                buscar = st.text_input("🔎 Buscar por nombre/clave (opcional)", key="filt_buscar")
                if buscar:
                    mask = (df_f['clave'].str.contains(buscar, case=False, na=False) |
                            df_f['nombre'].str.contains(buscar, case=False, na=False))
                    df_f = df_f[mask]
                    st.caption(f"Tras filtro: **{len(df_f):,} nodos**")

                # Mostrar tabla compacta
                if len(df_f) > 0:
                    st.dataframe(
                        df_f[['clave', 'nombre', 'ccr', 'estado', 'kv']].head(100),
                        width="stretch", hide_index=True,
                        column_config={
                            "clave":  st.column_config.TextColumn("Clave", width="small"),
                            "nombre": st.column_config.TextColumn("Nombre", width="medium"),
                            "ccr":    st.column_config.TextColumn("CCR", width="small"),
                            "estado": st.column_config.TextColumn("Estado", width="small"),
                            "kv":     st.column_config.NumberColumn("kV", format="%d"),
                        },
                        height=200,
                    )
                    if len(df_f) > 100:
                        st.caption(f"⚠️ Mostrando solo los primeros 100 (refina filtros para ver todos)")

                    # Botón para agregar al textarea — limpia consulta previa
                    if st.button(f"✅ Agregar {len(df_f)} nodos (reemplaza textarea)",
                                 key="btn_add_nodos"):
                        nuevas_claves = "\n".join(df_f['clave'].tolist())
                        # Reemplazar (no concatenar) para que sea predecible
                        st.session_state["nodos_pending"] = nuevas_claves
                        st.session_state["nodos_from_filter"] = len(df_f)
                        # Limpiar resultado de consulta anterior (forzar re-ejecutar)
                        for k in ("consulta_ejecutada", "acumulado", "matches_df",
                                  "fx_info", "consulta_params",
                                  "excel_datos_bytes", "excel_analisis_bytes"):
                            if k in st.session_state:
                                if k == "consulta_ejecutada":
                                    st.session_state[k] = False
                                else:
                                    st.session_state[k] = None
                        st.rerun()
        else:
            st.warning("No hay catálogo cargado para filtrar.")

    # Proceso MTR/MDA — el sistema (SIN/BCA/BCS) se detecta de cada nodo con el catálogo
    proceso = st.selectbox(
        "⚙️ Proceso", ["MTR", "MDA"], index=0, key="proceso_sel",
        help="**MTR** = Mercado de Tiempo Real (ex post; se publica 7 días después del día "
             "de operación). **MDA** = Mercado del Día en Adelanto (se publica un día antes "
             "del día de operación).")
    st.caption("ℹ️ El sistema (SIN/BCA/BCS) se detecta de cada nodo con el catálogo.")

    st.markdown("**📅 Período**")
    col_f1, col_f2 = st.columns(2)
    today = date.today()
    fecha_max = today + timedelta(days=1)   # el MDA de mañana ya está publicado
    with col_f1:
        f_ini = st.date_input("Desde", value=today - timedelta(days=90),
                              max_value=fecha_max, key="f_ini")
    with col_f2:
        f_fin = st.date_input("Hasta", value=today - timedelta(days=14),
                              max_value=fecha_max, key="f_fin",
                              help="MTR: hasta el día de operación −7. MDA: hasta mañana.")

    if proceso == "MTR" and f_fin > today - timedelta(days=7):
        st.caption(
            "⚠️ El MTR se publica 7 días después del día de operación: "
            "los últimos días del período pueden venir vacíos."
        )

    st.divider()

    # Botón limpiar todo
    if st.button("🧹 Limpiar todo", key="btn_clear_all", width="stretch",
                  help="Reset completo: borra textarea, filtros, datos descargados y descargas pendientes."):
        # Limpiar todo el session_state
        keys_to_clear = [
            "consulta_ejecutada", "acumulado", "matches_df", "fx_info",
            "consulta_params", "nodos_input", "nodos_pending", "nodos_from_filter",
            "excel_datos_bytes", "excel_analisis_bytes", "moneda",
            # Filtros
            "filt_sis", "filt_est", "filt_mun", "filt_zona", "filt_ccr",
            "filt_kv", "filt_buscar",
            # Selectores y opciones del Excel de análisis
            "sel_heatmap", "sel_multiano", "conf_grande",
            "xa_resumen", "xa_top_pml", "xa_top_vol", "xa_top_neg", "xa_multi", "xa_spread",
        ]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    with st.expander("📤 Actualizar catálogo de nodos"):
        st.caption(
            "El catálogo se actualiza ~1 vez al mes. Si tienes una versión más nueva, "
            "súbelo aquí (.xlsx) y reemplazará el catálogo durante esta sesión."
        )
        uploaded_cat = st.file_uploader(
            "Catálogo CENACE (.xlsx)",
            type=["xlsx"],
            key="cat_upload",
            help="Archivo `Catálogo_NodosP_Sistema_Eléctrico_Nacional` del portal CENACE",
        )

        # Flag para evitar reprocesar el mismo archivo en cada rerun
        if uploaded_cat is not None:
            file_id = f"{uploaded_cat.name}_{uploaded_cat.size}"
            last_processed = st.session_state.get("cat_last_processed", "")
            if file_id != last_processed:
                # Solo procesar si es un archivo distinto al ya procesado
                file_bytes = uploaded_cat.read()
                new_cat, new_version = cargar_catalogo_uploaded(file_bytes)
                if new_cat:
                    st.session_state["catalogo_uploaded"] = new_cat
                    st.session_state["catalogo_uploaded_version"] = new_version
                    st.session_state["cat_last_processed"] = file_id
                    st.success(
                        f"✅ Catálogo actualizado: **{len(new_cat):,} nodos**"
                        f"{' · ' + new_version if new_version else ''}. "
                        f"Aplica desde la próxima interacción."
                    )
                else:
                    st.error("❌ No se pudo leer el archivo. Verifica que sea el catálogo CENACE oficial.")
                    st.session_state["cat_last_processed"] = file_id  # marcar como procesado para no reintentar
            else:
                # Ya procesado: solo mostrar el estado actual
                if st.session_state.get("catalogo_uploaded"):
                    st.success(
                        f"✅ Catálogo actualizado en esta sesión: "
                        f"**{len(st.session_state['catalogo_uploaded']):,} nodos**"
                    )

        if "catalogo_uploaded" in st.session_state and st.session_state["catalogo_uploaded"]:
            if st.button("🔄 Volver al catálogo del repositorio", key="btn_reset_cat"):
                st.session_state["catalogo_uploaded"] = None
                st.session_state["catalogo_uploaded_version"] = ""
                st.session_state["cat_last_processed"] = ""
                st.rerun()

    st.caption("Recurrent Energy · Canadian Solar\n\n"
               "Basado en el CENACE PML Analyzer de Sebastian Roldan (SRF).")


col_left, col_right = st.columns([2, 1])
with col_left:
    st.markdown("### 🚀 Ejecutar consulta")

    modo_descripcion = {
        "🔬 Análisis": "Descargará los PML y mostrará el dashboard: KPIs, mapa de PML, tablas y gráficas.",
        "📊 Datos": "Descargará los PML y mostrará el Centro de Descargas (Excel de datos y de análisis).",
    }
    st.markdown(
        f"<div class='mode-badge'><b>Modo: {modo}</b><br>{modo_descripcion[modo]}</div>",
        unsafe_allow_html=True,
    )

    nodos = []
    if nodos_input.strip():
        nodos_raw = re.split(r"[,\n]+", nodos_input.strip())
        nodos = [n.strip().upper() for n in nodos_raw if n.strip()]
        seen = set()
        nodos = [n for n in nodos if not (n in seen or seen.add(n))]

    if nodos:
        cv = [n for n in nodos if n in catalogo] if catalogo else nodos
        ci = [n for n in nodos if catalogo and n not in catalogo]
        c_a, c_b, c_c = st.columns(3)
        c_a.metric("Solicitados", len(nodos))
        c_b.metric("✅ En catálogo", len(cv))
        c_c.metric("⚠️ Desconocidos", len(ci))
        if ci:
            with st.expander(f"⚠️ {len(ci)} nodos no encontrados"):
                st.code("\n".join(ci))

    if nodos and f_ini <= f_fin:
        n_dias = (f_fin - f_ini).days + 1
        n_bloques = (n_dias + BLOQUE_MAX - 1) // BLOQUE_MAX
        n_lotes = (len(nodos) + 9) // 10
        n_consultas = n_bloques * n_lotes
        tiempo = n_consultas * 1.5 / MAX_WORKERS
        st.caption(f"📊 **{n_dias} días · {n_consultas} consultas · ~{tiempo/60:.1f} min**")

    # ─── WARNINGS PROGRESIVOS POR # NODOS ───
    n_nodos = len(nodos)
    confirmacion_grande = True
    bloqueo_duro = False

    if n_nodos > 300:
        st.error(
            f"🔴 **{n_nodos} nodos** excede el límite recomendado. "
            f"La app puede quedarse sin memoria en Streamlit Cloud. "
            f"Divide la consulta en partes de 100-200 nodos."
        )
        bloqueo_duro = True
    elif n_nodos >= 100:
        st.warning(
            f"🟠 **{n_nodos} nodos** es una consulta grande. "
            f"Puede tomar 5-10 minutos y consumir mucha memoria."
        )
        confirmacion_grande = st.checkbox(
            f"He revisado y quiero proceder con {n_nodos} nodos",
            key="conf_grande"
        )
    elif n_nodos >= 30:
        st.info(f"⚠️ **{n_nodos} nodos** — esto puede tomar 2-5 minutos.")

    boton = st.button(
        "⚡ Ejecutar",
        type="primary",
        disabled=(len(nodos) == 0 or bloqueo_duro or not confirmacion_grande)
    )

with col_right:
    st.markdown("### 📋 Instrucciones")
    if es_solo_datos:
        st.markdown("""
        1. Pega claves de nodos o búscalos con los filtros
        2. Elige proceso, moneda y fechas
        3. Click en **Ejecutar**
        4. Genera y descarga el Excel de datos o el de análisis
        """)
    else:
        st.markdown("""
        1. Pega claves de nodos o búscalos con los filtros
        2. Elige proceso, moneda y fechas
        3. Click en **Ejecutar**
        4. Explora el dashboard interactivo

        💡 Para descargar Excel, usa el modo **📊 Datos** (no vuelve a descargar).
        """)


# ═══════════════════════════════════════════════════════════════════════
# EJECUCIÓN — guarda resultado en session_state
# ═══════════════════════════════════════════════════════════════════════
if boton and nodos:
    fecha_ini = f_ini.strftime("%Y/%m/%d")
    fecha_fin = f_fin.strftime("%Y/%m/%d")
    t0 = time.time()

    # Limpiar estado previo de descargas
    for k in ["excel_datos_bytes", "excel_analisis_bytes"]:
        if k in st.session_state:
            st.session_state[k] = None

    try:
        acumulado = {}
        errores = []
        info_sistemas = {}
        nodos_sin_sis = []
        fx_info = {"moneda": "MXN", "tc_promedio": None, "tc_lookup": {},
                   "fuente": "N/A", "advertencia": ""}

        if f_ini > f_fin:
            st.error("❌ La fecha inicial no puede ser posterior a la final.")
            st.stop()

        progress = st.progress(0, text="Iniciando descarga...")
        def cb(done, total):
            progress.progress(min(done/total, 1.0) if total > 0 else 0,
                              text=f"Descargando: {done}/{total} ({(done/total*100 if total>0 else 0):.0f}%)")

        with st.spinner("Consultando CENACE..."):
            acumulado, errores, info_sistemas, nodos_sin_sis = descargar_pml_auto(
                nodos, fecha_ini, fecha_fin, proceso, catalogo, progress_cb=cb,
            )
        progress.progress(1.0, text="✅ Datos descargados")

        # Mostrar info de sistemas auto-detectados
        if len(info_sistemas) > 1:
            sistemas_str = " · ".join(f"{s}: {n}" for s, n in info_sistemas.items())
            st.info(f"🔄 Multi-sistema detectado: {sistemas_str}")
        if nodos_sin_sis:
            st.caption(f"⚠️ {len(nodos_sin_sis)} nodos no estaban en el catálogo. "
                       f"Se asumió SIN como fallback.")

        if not acumulado:
            st.error("❌ No se recibieron datos PML.")
            st.stop()

        # ─── CONVERSIÓN A USD si aplica ───
        moneda_seleccionada = st.session_state.get("moneda", "MXN")
        if moneda_seleccionada == "USD":
            token = obtener_token_banxico()
            if not token:
                st.warning(
                    "⚠️ No se encontró el token de Banxico en los secrets de Streamlit. "
                    "El análisis se mostrará en MXN. Para usar USD, agrega "
                    "`BANXICO_TOKEN` en Settings → Secrets de tu app."
                )
                st.session_state["moneda"] = "MXN"
                fx_info["advertencia"] = "Sin token Banxico — análisis en MXN."
            else:
                # Descargar serie FIX para el período (con buffer de 5 días para fin de semana)
                f_ini_buffer = (f_ini - timedelta(days=5)).strftime("%Y-%m-%d")
                f_fin_buffer = f_fin.strftime("%Y-%m-%d")

                with st.spinner("📥 Descargando tipo de cambio FIX de Banxico..."):
                    fx_dict = cargar_fx_banxico(f_ini_buffer, f_fin_buffer, token)

                if not fx_dict:
                    st.warning(
                        "⚠️ No se pudo obtener el tipo de cambio de Banxico. "
                        "El análisis se mostrará en MXN."
                    )
                    st.session_state["moneda"] = "MXN"
                    fx_info["advertencia"] = "Banxico no respondió — análisis en MXN."
                else:
                    # Construir lookup completo (incluye fines de semana)
                    fechas_unicas = set()
                    for nodo, filas in acumulado.items():
                        for f in filas:
                            fechas_unicas.add(f.get("fecha", ""))
                    fx_lookup = construir_fx_lookup(fx_dict, fechas_unicas)

                    # NOTA: NO aplicamos conversión al acumulado. Se queda en MXN crudo.
                    # La conversión a USD se hace al construir el DataFrame en
                    # acumulado_a_dataframe(), usando el fx_lookup guardado en session_state.
                    # Esto evita el bug de doble-conversión y permite cambiar moneda
                    # sin re-descargar datos.

                    # Guardar info
                    if fx_lookup:
                        tcs = list(fx_lookup.values())
                        fx_info = {
                            "moneda": "USD",
                            "tc_promedio": sum(tcs)/len(tcs) if tcs else None,
                            "tc_min": min(tcs) if tcs else None,
                            "tc_max": max(tcs) if tcs else None,
                            "tc_lookup": fx_lookup,
                            "fuente": "Banxico FIX (SF43718)",
                            "advertencia": "",
                        }
                        st.success(
                            f"💱 Análisis convertido a USD. TC FIX promedio del período: "
                            f"**{fx_info['tc_promedio']:.4f} MXN/USD** "
                            f"(rango: {fx_info['tc_min']:.4f} – {fx_info['tc_max']:.4f})"
                        )

        # Geocodificación de los nodos con datos para el mapa de PML (modo Análisis)
        matches_df = None
        if necesita_mapa:
            with st.spinner("Cargando subestaciones de OpenStreetMap (primera vez ~30s)..."):
                osm_subs = cargar_osm_subestaciones()
            if osm_subs:
                with st.spinner("Ubicando nodos en OpenStreetMap..."):
                    resultados = matchear_nodos(list(acumulado.keys()), catalogo, osm_subs)
                    matches_df = pd.DataFrame(resultados)
            else:
                st.caption("⚠️ OpenStreetMap no respondió: el dashboard se mostrará sin mapa.")

        # Guardar todo en session_state
        # Sistema: si solo 1 → ese; si multi → "MULTI" + dict
        sistema_str = "MULTI" if len(info_sistemas or {}) > 1 else (
            list(info_sistemas.keys())[0] if info_sistemas else "SIN"
        )
        st.session_state["consulta_ejecutada"] = True
        st.session_state["acumulado"] = acumulado
        st.session_state["matches_df"] = matches_df
        st.session_state["fx_info"] = fx_info
        st.session_state["consulta_params"] = {
            "sistema": sistema_str, "proceso": proceso,
            "info_sistemas": info_sistemas,
            "fecha_ini": fecha_ini, "fecha_fin": fecha_fin,
            "modo": modo, "errores": errores,
            "tiempo": time.time() - t0,
            "moneda_aplicada": st.session_state.get("moneda", "MXN"),
        }

    except Exception as e:
        st.error(f"❌ Error: {type(e).__name__}: {e}")
        st.exception(e)


# ═══════════════════════════════════════════════════════════════════════
# RENDERIZAR RESULTADOS DESDE SESSION_STATE
# Esto se ejecuta cada vez que cambias algo (selector de nodo, moneda, etc.)
# pero SIN volver a descargar datos
# ═══════════════════════════════════════════════════════════════════════
if st.session_state.get("consulta_ejecutada"):
    acumulado = st.session_state["acumulado"]
    matches_df = st.session_state["matches_df"]
    params = st.session_state["consulta_params"]
    fx_info = st.session_state.get("fx_info", {"moneda": "MXN"})
    # Modo actual del sidebar (no el guardado) — esto permite re-render instantáneo
    # al cambiar de modo sin re-ejecutar la descarga
    modo_cur = modo
    modo_consulta_original = params.get("modo", modo)

    # Detectar si parámetros que requieren re-descarga cambiaron
    moneda_actual = st.session_state.get("moneda", "MXN")
    moneda_consulta = params.get("moneda_aplicada", "MXN")

    # ─── Cambio de moneda inteligente ───
    # MXN→USD instantáneo si ya tenemos fx_lookup
    # USD→MXN siempre instantáneo (acumulado siempre está en MXN crudo)
    fx_lookup_disponible = bool(fx_info.get("tc_lookup"))
    cambio_moneda_instantaneo = False
    if moneda_actual != moneda_consulta:
        if moneda_actual == "MXN":
            # USD → MXN: siempre instantáneo
            cambio_moneda_instantaneo = True
        elif moneda_actual == "USD" and fx_lookup_disponible:
            # MXN → USD con TC ya descargado: instantáneo
            cambio_moneda_instantaneo = True

    # Si cambio el modo entre consulta y ahora pero los datos sirven
    cambio_solo_modo = (modo_cur != modo_consulta_original)

    # Cambios que SI requieren re-ejecutar
    parametros_cambiaron = []
    if moneda_actual != moneda_consulta and not cambio_moneda_instantaneo:
        parametros_cambiaron.append(
            f"moneda ({moneda_consulta} → {moneda_actual}) — necesita descargar TC Banxico"
        )
    # Comparar fechas y proceso
    f_ini_str = f_ini.strftime("%Y/%m/%d")
    f_fin_str = f_fin.strftime("%Y/%m/%d")
    if f_ini_str != params.get("fecha_ini") or f_fin_str != params.get("fecha_fin"):
        parametros_cambiaron.append("fechas")
    if proceso != params.get("proceso"):
        parametros_cambiaron.append(f"proceso ({params.get('proceso')} → {proceso})")

    # Validar si los datos actuales sirven para el modo actual
    modo_requiere_datos = True
    modo_requiere_mapa = ("🔬" in modo_cur)
    datos_disponibles = bool(acumulado)
    mapa_disponible = matches_df is not None and not matches_df.empty

    # Warning si cambiaron parámetros
    if parametros_cambiaron:
        st.warning(
            f"⚠️ **Algunos parámetros cambiaron:** {', '.join(parametros_cambiaron)}. "
            f"Los datos mostrados corresponden a la consulta anterior. "
            f"Da click en **⚡ Ejecutar** arriba para refrescar."
        )

    # Warning si modo cambió y faltan datos para ese modo
    if cambio_solo_modo:
        if modo_requiere_datos and not datos_disponibles:
            st.warning(
                f"⚠️ El modo **{modo_cur}** necesita datos PML descargados. "
                f"Da click en **⚡ Ejecutar** para descargarlos."
            )
            st.stop()
        if modo_requiere_mapa and not mapa_disponible:
            st.info(
                "ℹ️ El mapa de PML se arma al ejecutar en modo **🔬 Análisis**. "
                "Da click en **⚡ Ejecutar** para incluirlo."
            )

    # ─── Definir moneda activa para el render ───
    # Si cambio_moneda_instantaneo: usar la moneda del toggle actual
    # Si no: usar la moneda con que se hizo la consulta
    if cambio_moneda_instantaneo:
        moneda_render = moneda_actual
        # Limpiar los Excel generados porque la moneda cambió
        for k in ("excel_datos_bytes", "excel_analisis_bytes"):
            if k in st.session_state:
                st.session_state[k] = None
    else:
        moneda_render = moneda_consulta

    st.session_state["moneda"] = moneda_render  # se usa por get_moneda() en gráficas/tablas

    # Banner FX si está en USD
    if moneda_render == "USD" and fx_info.get("tc_promedio"):
        st.info(
            f"💱 **Análisis en USD** · "
            f"TC promedio del período: **{fx_info['tc_promedio']:.4f} MXN/USD** · "
            f"Fuente: {fx_info['fuente']}"
        )

    # Métricas de la descarga
    if acumulado:
        n_total = sum(len(f) for f in acumulado.values())
        cx, cy, cz = st.columns(3)
        cx.metric("Nodos con datos", f"{len(acumulado)}/{len(nodos) if nodos else len(acumulado)}")
        cy.metric("Total registros", f"{n_total:,}")
        cz.metric("Tiempo descarga", f"{params.get('tiempo', 0):.0f}s")

        if params.get("errores"):
            with st.expander(f"⚠️ {len(params['errores'])} errores en consultas"):
                st.dataframe(pd.DataFrame(params["errores"][:20]), width="stretch")

        # Aviso para consultas pesadas (> 1.3M filas ≈ 150 nodos × 1 año)
        if n_total > 1_300_000:
            st.warning(
                f"⚠️ **Consulta pesada** ({n_total:,} filas). La memoria de Streamlit Cloud "
                f"es limitada: la app puede volverse lenta o reiniciarse al generar gráficas "
                f"o Excel. Para análisis multi-año con ≥150 nodos, divide la consulta."
            )

    # Render según modo
    if "📊" in modo_cur:
        # Modo Datos: Centro de descargas
        if acumulado:
            render_centro_descargas(
                acumulado, catalogo,
                params["sistema"], params["proceso"],
                params["fecha_ini"], params["fecha_fin"],
                moneda=moneda_render,
                tc_info=fx_info,
            )
    else:
        # Modo Análisis: dashboard (para Excel, cambiar a modo Datos sin volver a descargar)
        if acumulado:
            render_dashboard(acumulado, catalogo, matches_df)


st.divider()
st.caption("⚡ PML CENACE · Recurrent Energy / Canadian Solar · Datos: CENACE (SW-PML) y Banxico (FIX) · "
           "Basado en el CENACE PML Analyzer de Sebastian Roldan (SRF)")
