"""
PML CENACE — Streamlit
======================
Recurrent Energy México · Inteligencia de Mercado

Fork de «CENACE PML Analyzer» v9 de Sebastian Roldan (SRF)
https://github.com/sebasrf44/cenace-pml-analyzer
reducido a lo que trata de Precios Marginales Locales (PML).

Qué hace: eliges uno o varios NodosP y la app arma un informe estándar del MTR,
individual (un nodo) o comparativo (dos o más, con la ficha de cada uno), con todo el
historial disponible y un período que se ajusta sin volver a descargar.

Módulos:
- datos_pml.py: descarga del SW-PML, cobertura, almacén en memoria y rama `datos`.
- informe.py: estadísticas, componentes, distribuciones, solar/BESS y lecturas.
- graficas.py, excel_informe.py, ubicacion.py: gráficas, archivos y mapa.
- scripts/actualizar_datos.py + .github/workflows/actualizar-datos.yml: historial
  guardado de los nodos de las listas (rama `datos`, actualización diaria).

Cambios respecto al original:
- Sin BESS de scoring ni KMZ. Un solo flujo, nodos → informe (sin «modo de uso»).
- Solo MTR. El período por defecto es todo el historial del nodo.
- Historial guardado de los nodos de las listas; los demás se descargan bajo demanda
  (12 meses y, si se pide, todo el historial) y quedan en memoria mientras la app
  está despierta.
- USD con el FIX del día de operación (Banxico con token; si no, DOF).
- Buscador por nombre, filtros, importación, listas guardadas, enlaces y novedades
  del catálogo (lo actualiza una GitHub Action; ya no se sube a mano).
"""

import hashlib
import json
import os
import time
from datetime import date, datetime, timedelta
from urllib.parse import urlencode

import pandas as pd
import requests
import streamlit as st

import datos_pml as dp
import graficas as gr
import informe
from catalogo import etiqueta_nodo, normalizar, parse_catalogo, parse_lista_claves
from excel_informe import csv_horario, excel_horario, excel_informe
from ubicacion import descargar_subestaciones, matchear_nodos

st.set_page_config(page_title="PML CENACE · Recurrent Energy", page_icon="⚡", layout="wide",
                   initial_sidebar_state="expanded")

RE_NAVY, RE_RED, RE_BLUE = "#0e346b", "#a0090c", "#2777bd"
st.markdown(f"""
<style>
    .main-header {{
        background: linear-gradient(135deg, {RE_NAVY} 0%, {RE_BLUE} 100%);
        padding: 1.4rem 2rem; border-radius: 8px; margin-bottom: 1.2rem; color: white;
    }}
    .main-header h1 {{ color: white !important; margin: 0; font-size: 1.9rem; }}
    .main-header p {{ color: rgba(255,255,255,0.88); margin: 0.4rem 0 0 0; }}
    .pml-badge {{
        background: {RE_RED}; color: white; padding: 0.25rem 0.7rem; border-radius: 4px;
        font-weight: bold; font-size: 0.85rem; display: inline-block; margin-left: 1rem;
    }}
    .stButton > button[kind="primary"] {{ background: {RE_NAVY}; color: white; border: none; }}
    .stButton > button[kind="primary"]:hover {{ background: {RE_RED}; }}
</style>
""", unsafe_allow_html=True)

WORKERS = 8
MAX_NODOS_INFORME = 30
MAX_AGREGAR_FILTRO = MAX_NODOS_INFORME
VENTANA_INICIAL = 365            # días que se bajan primero si el historial completo tardaría mucho
LIMITE_AUTO = float(os.environ.get("PML_LIMITE_AUTO", 240))   # s: hasta aquí se baja todo de una vez
TOPE_INICIAL = 150               # s: tope de la primera descarga (si el CENACE se pone lento)
TOPE_DESCARGA = 600              # s: tope de «Un año más» / «Completar desde 2017»
REINTENTO_COLA = 3 * 3600        # los días aún no publicados se vuelven a pedir cada 3 h
SEG_POR_CONSULTA = 2.0           # estimación inicial (segundos por consulta y por hilo)
MONEDAS = {"MXN": "🇲🇽 MXN (pesos)", "USD": "🇺🇸 USD (dólares)"}
PERIODOS = {"todo": "Todo el historial", "comun": "Período común", "12m": "Últimos 12 meses",
            "3a": "Últimos 3 años", "anio": "Un año", "pers": "Personalizado"}
COMPARACION = "__comparacion__"
RUTA_DATA = os.path.join(os.path.dirname(__file__), "data")
PLOTLY_CONFIG = {"displaylogo": False, "modeBarButtonsToRemove": ["select2d", "lasso2d"]}


# ═══════════════════════════════════════════════════════════════════════
# CATÁLOGO, LISTAS Y NOVEDADES
# ═══════════════════════════════════════════════════════════════════════
def _firma_archivo(ruta):
    """(mtime, tamaño) del archivo, para renovar la caché cuando la Action lo cambia."""
    try:
        st_ = os.stat(ruta)
        return (st_.st_mtime_ns, st_.st_size)
    except OSError:
        return None


@st.cache_data(show_spinner=False)
def _cargar_catalogo_archivo(ruta, firma):
    return parse_catalogo(ruta)


def cargar_catalogo():
    """(catálogo, versión) de data/catalogo_nodos.xlsx."""
    ruta = os.path.join(RUTA_DATA, "catalogo_nodos.xlsx")
    firma = _firma_archivo(ruta)
    return _cargar_catalogo_archivo(ruta, firma) if firma else ({}, "")


@st.cache_data(show_spinner=False)
def indice_buscador(llave, _catalogo):
    """(claves ordenadas por nombre, {clave: etiqueta}) del catálogo."""
    claves = sorted(_catalogo, key=lambda c: (normalizar(_catalogo[c].get("nombre", "")), c))
    return claves, {c: etiqueta_nodo(c, _catalogo[c]) for c in claves}


@st.cache_data(show_spinner=False)
def catalogo_df(llave, _catalogo):
    df = pd.DataFrame.from_dict(_catalogo, orient="index").reset_index()
    return df.rename(columns={"index": "clave"})


@st.cache_data(show_spinner=False)
def _leer_json_archivo(ruta, firma):
    try:
        with open(ruta, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def _leer_json_data(nombre):
    ruta = os.path.join(RUTA_DATA, nombre)
    firma = _firma_archivo(ruta)
    return _leer_json_archivo(ruta, firma) if firma else {}


def cargar_listas_guardadas():
    """{nombre: [claves]} desde data/listas_nodos.json."""
    datos = _leer_json_data("listas_nodos.json")
    return {str(k): parse_lista_claves(" ".join(map(str, v))) for k, v in datos.items() if isinstance(v, list)}


def cargar_novedades_catalogo():
    return _leer_json_data("catalogo_cambios.json")


def nombre_nodo(clave, catalogo):
    return (catalogo.get(clave) or {}).get("nombre") or clave


# ═══════════════════════════════════════════════════════════════════════
# OPENSTREETMAP
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False, ttl=86400 * 7)
def _subestaciones_osm():
    return descargar_subestaciones()          # RuntimeError si Overpass falla: no se guarda en caché


def subestaciones_osm():
    try:
        return _subestaciones_osm()
    except RuntimeError:
        return []


# ═══════════════════════════════════════════════════════════════════════
# TIPO DE CAMBIO FIX (Banxico SF43718 con token; si no, DOF vía SIDOF)
# ═══════════════════════════════════════════════════════════════════════
BANXICO_API_BASE = "https://www.banxico.org.mx/SieAPIRest/service/v1/series"
SERIE_FIX_USD = "SF43718"
SIDOF_INDICADORES = "https://sidof.segob.gob.mx/dof/sidof/indicadores/158"


@st.cache_data(show_spinner=False, ttl=86400)
def cargar_fx_banxico(fecha_ini, fecha_fin, _token):
    """{AAAA-MM-DD: tc} del FIX por fecha de determinación, o {} si falla."""
    if not _token:
        return {}
    url = f"{BANXICO_API_BASE}/{SERIE_FIX_USD}/datos/{fecha_ini}/{fecha_fin}?mediaType=json"
    try:
        r = requests.get(url, headers={"Bmx-Token": _token, "Accept": "application/json"}, timeout=30)
        if r.status_code != 200:
            return {}
        series = r.json().get("bmx", {}).get("series", [])
        datos = series[0].get("datos", []) if series else []
    except Exception:
        return {}
    fx = {}
    for d in datos:
        try:
            dia, mes, anio = d["fecha"].split("/")
            fx[f"{anio}-{mes.zfill(2)}-{dia.zfill(2)}"] = float(d["dato"])
        except (KeyError, ValueError, AttributeError):
            continue
    return fx


@st.cache_data(show_spinner=False, ttl=86400)
def cargar_fx_dof(fecha_ini, fecha_fin):
    """FIX publicado en el DOF (indicador 158 del SIDOF), sin token.

    El DOF publica el FIX un día hábil después de determinarlo: cada valor se asigna a la
    publicación anterior, que es su fecha de determinación (como la serie de Banxico).
    """
    d0 = date.fromisoformat(fecha_ini) - timedelta(days=10)
    d1 = date.fromisoformat(fecha_fin) + timedelta(days=10)
    try:
        r = requests.get(f"{SIDOF_INDICADORES}/{d0:%d-%m-%Y}/{d1:%d-%m-%Y}",
                         headers={"Accept": "application/json"}, timeout=30)
        if r.status_code != 200:
            return {}
        lista = r.json().get("ListaIndicadores", []) or []
    except Exception:
        return {}
    pubs = []
    for x in lista:
        try:
            pubs.append((datetime.strptime(x["fecha"], "%d-%m-%Y").date(), float(x["valor"])))
        except (KeyError, TypeError, ValueError):
            continue
    pubs.sort()
    return {anterior.isoformat(): valor for (anterior, _), (_, valor) in zip(pubs, pubs[1:])}


def obtener_fx(fecha_ini, fecha_fin, token):
    """(fx, fuente) por fecha de determinación: Banxico con token y, si no, el DOF."""
    if token:
        fx = cargar_fx_banxico(fecha_ini, fecha_fin, token)
        if fx:
            return fx, "Banxico SIE (SF43718)"
    fx = cargar_fx_dof(fecha_ini, fecha_fin)
    if fx:
        return fx, "DOF (SIDOF), FIX en su fecha de determinación"
    return {}, ""


def obtener_token_banxico():
    try:
        return st.secrets.get("BANXICO_TOKEN", "")
    except Exception:
        return ""


def serie_fx(ini, fin, token):
    """(Serie diaria de TC, fuente) para [ini, fin]; (None, '') si falta algún año.

    Se pide año por año (cada año queda en caché). Fines de semana y feriados usan el
    último FIX publicado.
    """
    datos, fuentes, hoy = {}, set(), date.today()
    for anio in range(ini.year, fin.year + 1):
        a = date(anio, 1, 1) - timedelta(days=10)          # años completos: la caché sirve para cualquier período
        b = min(date(anio, 12, 31), hoy)
        fx, fuente = obtener_fx(a.isoformat(), b.isoformat(), token)
        if not fx:
            return None, ""
        datos.update(fx)
        fuentes.add(fuente)
    s = pd.Series(datos, dtype="float64")
    s.index = pd.to_datetime(s.index)
    s = s.sort_index()
    dias_ = pd.date_range(ini - timedelta(days=10), fin, freq="D")
    s = s.reindex(s.index.union(dias_)).ffill().bfill().reindex(dias_)
    return s, " + ".join(sorted(fuentes))


# ═══════════════════════════════════════════════════════════════════════
# DATOS: rama `datos` (historial guardado) + almacén en memoria + descargas
# ═══════════════════════════════════════════════════════════════════════
@st.cache_resource(show_spinner=False)
def almacen():
    return dp.Almacen(max_series=60)


@st.cache_resource(show_spinner=False)
def ritmo_cenace():
    """Segundos por consulta y por hilo medidos en la última descarga (todas las sesiones)."""
    return {"seg": SEG_POR_CONSULTA}


@st.cache_data(show_spinner=False, ttl=1800)
def _indice_remoto():
    return dp.leer_indice_remoto()


def nodos_guardados():
    """{clave: metadatos} de los nodos con historial guardado en la rama `datos`."""
    try:
        return _indice_remoto().get("nodos") or {}
    except RuntimeError:
        return {}


def cargar_guardados(claves):
    """Trae del GitHub el historial guardado de los nodos que lo tengan. Devuelve avisos."""
    meta, alm, avisos = nodos_guardados(), almacen(), []
    for c in claves:
        m = meta.get(c)
        if not m or not m.get("cobertura"):
            continue
        s = alm.get(c)
        if s is not None and s.remoto == m.get("actualizado", ""):
            continue
        df = None
        if m.get("filas"):
            try:
                df = dp.leer_serie_remota(m["archivo"])
            except RuntimeError as e:
                avisos.append(f"No se pudo leer el historial guardado de {c} ({e}); se descarga del CENACE.")
                continue
        alm.agregar(c, df, dp.intervalos_de_json(m["cobertura"]), remoto=m.get("actualizado", ""))
    return avisos


def faltantes_para(claves, modo, ultimo):
    """Rangos por descargar de cada nodo. modo: inicial | anio | todo."""
    alm, ahora, out = almacen(), time.time(), {}
    for c in claves:
        s = alm.get(c)
        cob = s.cobertura if s else []
        if modo == "inicial":       # con historial guardado ya cargado solo faltan los días nuevos
            ini = dp.INICIO_MTR if (s is not None and s.remoto) else ultimo - timedelta(days=VENTANA_INICIAL - 1)
        elif modo == "anio":                     # el año anterior a lo que ya está cargado
            primero = min((a for a, _ in cob), default=ultimo + timedelta(days=1))
            ini = max(dp.INICIO_MTR, primero - timedelta(days=365))
            rangos = dp.restar(ini, primero - timedelta(days=1), cob)
            if rangos:
                out[c] = rangos
            continue
        else:
            ini = dp.INICIO_MTR
        rangos = dp.restar(ini, ultimo, cob)
        # Los días recientes que aún no publica el CENACE no se vuelven a pedir en 3 h
        if (rangos and s is not None and ahora - s.intento_cola < REINTENTO_COLA
                and rangos[-1][0] >= ultimo - timedelta(days=10)):
            rangos = rangos[:-1]
        if rangos:
            out[c] = rangos
    return out


def _duracion(seg):
    return f"{seg:.0f} s" if seg < 90 else f"{seg / 60:.0f} min"


def estimar(faltantes, catalogo):
    """(consultas, segundos estimados) para descargar los faltantes."""
    if not faltantes:
        return 0, 0.0
    sistemas = {c: (catalogo.get(c, {}).get("sistema") or "SIN").upper() for c in faltantes}
    n = len(dp.planificar(faltantes, sistemas, dp.ultimo_dia_mtr()))
    return n, n * ritmo_cenace()["seg"] / WORKERS


def descargar(faltantes, catalogo, texto, tope=TOPE_DESCARGA):
    """Descarga los faltantes con barra de avance y los guarda en el almacén en memoria.

    Las semanas más recientes van primero; pasado `tope` segundos no se lanzan consultas
    nuevas (si el CENACE se pone lento, el informe sale con lo reciente y se completa luego).
    """
    if not faltantes:
        return {"consultas": 0, "errores": [], "segundos": 0.0, "pendientes": 0}
    ultimo = dp.ultimo_dia_mtr()
    sistemas = {c: (catalogo.get(c, {}).get("sistema") or "SIN").upper() for c in faltantes}
    consultas = dp.planificar(faltantes, sistemas, ultimo)
    alm, rec = almacen(), dp.Recolector(ultimo)
    barra = st.progress(0.0, text=f"{texto}…")
    t0 = time.time()

    def volcar():
        for c, (df, cub) in rec.vaciar().items():
            alm.agregar(c, df, cub)

    def al_llegar(consulta, datos):
        rec(consulta, datos)
        if rec.consultas % 25 == 0:          # lo descargado queda guardado aunque se interrumpa
            volcar()

    def progreso(hechas, total):
        trans = time.time() - t0
        resta = trans / hechas * (total - hechas) if hechas else 0
        barra.progress(hechas / total, text=f"{texto}: {hechas}/{total} consultas al CENACE · "
                                            f"faltan ~{_duracion(resta)}")

    try:
        errores, pendientes = dp.ejecutar(consultas, al_llegar, workers=WORKERS, progreso=progreso,
                                          limite=t0 + tope)
    finally:
        volcar()
        for c, rangos in faltantes.items():
            if rangos[-1][1] >= ultimo - timedelta(days=3):
                alm.marcar_intento_cola(c)
    barra.empty()
    seg = time.time() - t0
    hechas = len(consultas) - len(pendientes)
    if hechas >= 2 * WORKERS:                             # con pocas consultas la estimación no sirve
        ritmo_cenace()["seg"] = max(0.2, seg * WORKERS / hechas)
    return {"consultas": hechas, "errores": errores, "segundos": seg, "pendientes": len(pendientes)}


def tabla_errores(errores):
    return pd.DataFrame([{"Semana": f"{e['ini']:%d-%m-%Y} → {e['fin']:%d-%m-%Y}", "Nodos": ", ".join(e["nodos"]),
                          "Error": e["error"]} for e in errores])


def generar_informe(claves, catalogo, modo="inicial"):
    """Junta los datos del informe (historial guardado + descargas) y lo deja en sesión."""
    with st.spinner("Leyendo el historial guardado…"):
        avisos = cargar_guardados(claves)
    ultimo = dp.ultimo_dia_mtr()
    falt = faltantes_para(claves, modo, ultimo)
    texto = "Completando el historial"
    if modo == "inicial":
        # Todo el historial de una vez si el CENACE responde a buen ritmo; si no, 12 meses primero
        completo = faltantes_para(claves, "todo", ultimo)
        if estimar(completo, catalogo)[1] <= LIMITE_AUTO:
            falt = completo
        n = len(falt)
        que = "el historial completo" if falt == completo else "los últimos 12 meses"
        texto = f"Descargando {que} de {n} nodo{'s' if n > 1 else ''} del CENACE"
    res = descargar(falt, catalogo, texto, tope=TOPE_INICIAL if modo == "inicial" else TOPE_DESCARGA)
    if res["pendientes"]:
        avisos.append(f"El CENACE respondió lento y quedaron {res['pendientes']} consultas sin hacer (las semanas "
                      f"más antiguas). El informe usa lo descargado; los botones de historial completan el resto.")
    previo = st.session_state.get("informe") or {}
    alm = almacen()
    errores = [{"ini": e["consulta"].ini, "fin": e["consulta"].fin, "nodos": list(e["consulta"].nodos),
                "error": e["error"]} for e in res["errores"]]
    if modo != "inicial" and previo.get("nodos") == list(claves):
        # Las semanas que fallaron antes y siguen sin datos se conservan; las ya cubiertas, no
        vistos = {(e["ini"], e["fin"], tuple(e["nodos"])) for e in errores}
        errores += [e for e in previo.get("errores", [])
                    if (e["ini"], e["fin"], tuple(e["nodos"])) not in vistos
                    and any(dp.restar(e["ini"], e["fin"], alm.get(c).cobertura if alm.get(c) else [])
                            for c in e["nodos"])]
    st.session_state["informe"] = {
        "nodos": list(claves), "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "avisos": avisos, "errores": errores,
        "consultas": res["consultas"], "segundos": res["segundos"],
    }


# ═══════════════════════════════════════════════════════════════════════
# INFORME: datos del período, fichas y comparación (en caché)
# ═══════════════════════════════════════════════════════════════════════
@st.cache_resource(show_spinner=False, max_entries=48)
def datos_periodo(clave, version, ini, fin, moneda, fx_id, _datos, _fx):
    """Horas del nodo en [ini, fin], en MXN o convertidas a USD con el FIX de su día."""
    m = (_datos["fecha"] >= pd.Timestamp(ini)) & (_datos["fecha"] <= pd.Timestamp(fin))
    df = _datos.loc[m].reset_index(drop=True)
    if moneda == "USD" and _fx is not None and not df.empty:
        tc = df["fecha"].map(_fx)
        df = df.assign(**{c: df[c] / tc for c in dp.PRECIOS}, tc=tc)
    return df


@st.cache_data(show_spinner=False, max_entries=64)
def ficha_nodo(clave, version, ini, fin, moneda, fx_id, _df):
    umbral = informe.UMBRAL_CNG_MXN
    if moneda == "USD" and "tc" in _df.columns and len(_df):
        umbral = informe.UMBRAL_CNG_MXN / float(_df["tc"].mean())
    return informe.ficha(_df, umbral=umbral, unidad=moneda)


@st.cache_data(show_spinner=False, max_entries=16)
def comparacion_nodos(llave, ref, moneda, _dfs, _fichas, _nombres):
    return informe.comparacion(_dfs, _fichas, _nombres, ref, moneda)


def rango_disponible(claves, comun=False):
    """(primer día, último día) con datos de los nodos; con comun=True, el tramo que comparten."""
    desdes, hastas = [], []
    for c in claves:
        s = almacen().get(c)
        if s is not None and not s.datos.empty:
            desdes.append(s.desde())
            hastas.append(s.hasta())
    if not desdes:
        return None, None
    if comun:
        a, b = max(desdes), min(hastas)
        return (a, b) if a < b else (None, None)
    return min(desdes), max(hastas)


def historial_incompleto(claves):
    """{clave: texto} de los nodos con semanas sin consultar entre el inicio del MTR y lo último cargado.

    Pasa si se bajaron solo los últimos 12 meses o si el CENACE no respondió algunas semanas.
    """
    out = {}
    for c in claves:
        s = almacen().get(c)
        cob = s.cobertura if s else []
        if not cob:
            out[c] = "sin datos todavía"
        elif cob[0][0] > dp.INICIO_MTR:
            out[c] = f"cargado desde el {cob[0][0]:%d-%m-%Y}"
        elif dp.restar(dp.INICIO_MTR, cob[-1][1], cob):
            out[c] = f"{dp.dias(dp.restar(dp.INICIO_MTR, cob[-1][1], cob))} días sin respuesta del CENACE"
    return out


def _sanear(clave, opciones, defecto):
    if st.session_state.get(clave) not in opciones:
        st.session_state[clave] = defecto


def controles_periodo(a, b, comun=(None, None)):
    """Selector del período del informe dentro de [a, b]. Devuelve (ini, fin).

    «Período común» (solo en comparaciones de nodos con historias distintas): el tramo
    en que todos tienen datos.
    """
    opciones = [p for p in PERIODOS if p != "comun" or (comun[0] and comun != (a, b))]
    _sanear("periodo_modo", opciones, "todo")
    c1, c2 = st.columns([3, 2])
    with c1:
        modo = st.radio("Período del informe", opciones, horizontal=True, key="periodo_modo",
                        format_func=lambda p: (f"Período común (desde {comun[0]:%m-%Y})" if p == "comun"
                                               else PERIODOS[p]))
    ini, fin = a, b
    if modo == "comun":
        ini, fin = comun
    elif modo == "12m":
        ini = max(a, b - timedelta(days=364))
    elif modo == "3a":
        ini = max(a, b - timedelta(days=3 * 365 - 1))
    elif modo == "anio":
        anios = list(range(b.year, a.year - 1, -1))
        _sanear("periodo_anio", anios, anios[0])
        with c2:
            anio = st.selectbox("Año", anios, key="periodo_anio")
        ini, fin = max(a, date(anio, 1, 1)), min(b, date(anio, 12, 31))
    elif modo == "pers":
        actual = st.session_state.get("periodo_rango")
        if not (isinstance(actual, (tuple, list)) and actual and all(a <= x <= b for x in actual)):
            st.session_state["periodo_rango"] = (a, b)
        with c2:
            rango = st.date_input("Desde / hasta", min_value=a, max_value=b, key="periodo_rango",
                                  format="DD/MM/YYYY")
        if isinstance(rango, (tuple, list)) and len(rango) == 2:
            ini, fin = rango
        elif isinstance(rango, (tuple, list)) and len(rango) == 1:
            ini = rango[0]
    return ini, fin


def token_periodo(ini, fin, a, b):
    modo = st.session_state.get("periodo_modo", "todo")
    if modo in ("todo", "comun", "12m", "3a"):
        return modo
    if modo == "anio":
        return str(st.session_state.get("periodo_anio", fin.year))
    return f"{ini.isoformat()}_{fin.isoformat()}"


def aplicar_token_periodo(token):
    token = (token or "").strip()
    if token in ("todo", "comun", "12m", "3a"):
        st.session_state["periodo_modo"] = token
    elif token.isdigit() and len(token) == 4:
        st.session_state["periodo_modo"] = "anio"
        st.session_state["periodo_anio"] = int(token)
    elif "_" in token:
        try:
            a, b = (date.fromisoformat(x) for x in token.split("_", 1))
            st.session_state["periodo_modo"] = "pers"
            st.session_state["periodo_rango"] = (a, b)
        except ValueError:
            pass


# ═══════════════════════════════════════════════════════════════════════
# RENDER
# ═══════════════════════════════════════════════════════════════════════
def _n(x, dec=None):
    """Número con separador de miles; sin decimales salvo valores chicos (p. ej. en USD)."""
    try:
        if x is None or pd.isna(x):
            return "—"
    except TypeError:
        return str(x)
    if dec is None:
        dec = 1 if abs(x) < 100 else 0
    return f"{x:,.{dec}f}"


def _grafica(fig, key):
    if fig is not None:
        st.plotly_chart(fig, width="stretch", theme=None, key=key, config=PLOTLY_CONFIG)


def render_kpis(k, u):
    uu = f"{u}/MWh"
    c = st.columns(5)
    c[0].metric(f"PML promedio ({uu})", _n(k["promedio"]))
    c[1].metric(f"Mediana ({uu})", _n(k["mediana"]))
    c[2].metric(f"P10 – P90 ({uu})", f"{_n(k['p10'])} – {_n(k['p90'])}",
                help="80 % de las horas del período quedan en este rango.")
    c[3].metric(f"Máximo ({uu})", _n(k["maximo"]), help=f"{k['max_fecha']:%d-%m-%Y}, hora {k['max_hora']}")
    c[4].metric(f"Mínimo ({uu})", _n(k["minimo"]), help=f"{k['min_fecha']:%d-%m-%Y}, hora {k['min_hora']}")
    c = st.columns(5)
    delta = None
    if not pd.isna(k.get("prom_12m_prev", float("nan"))) and k["prom_12m_prev"]:
        delta = f"{(k['prom_12m'] / k['prom_12m_prev'] - 1) * 100:+.1f} % vs. 12 meses previos"
    c[0].metric(f"Últimos 12 meses ({uu})", _n(k["prom_12m"]), delta=delta, delta_color="off")
    c[1].metric(f"Diferencial nodal ({uu})", _n(k["dif"], 1),
                help="PML − componente de energía = pérdidas + congestión. Negativo: zona exportadora.")
    c[2].metric("Horas con congestión", f"{_n(k['pct_cng'], 1)} %",
                help=f"Negativa {_n(k['pct_cng_neg'], 1)} % · positiva {_n(k['pct_cng_pos'], 1)} %")
    c[3].metric("Captura solar", f"{_n(k['solar_factor'])} %",
                help=f"Precio capturado por un perfil solar genérico: {_n(k['solar_captura'])} {uu}.")
    c[4].metric(f"Spread diario 4 h ({uu})", _n(k["spread4"]),
                help="Promedio de las 4 horas más caras menos las 4 más baratas de cada día.")


TABLA_ANUAL = {"horas": ("Horas", "%,d"), "cobertura": ("Cobertura %", "%.0f"), "promedio": ("Promedio", "%,.0f"),
               "mediana": ("Mediana", "%,.0f"), "p10": ("P10", "%,.0f"), "p90": ("P90", "%,.0f"),
               "minimo": ("Mínimo", "%,.0f"), "maximo": ("Máximo", "%,.0f"), "desv": ("Desv. est.", "%,.0f"),
               "pct_neg": ("% negativas", "%.1f"), "energia": ("Energía", "%,.0f"),
               "perdidas": ("Pérdidas", "%,.1f"), "congestion": ("Congestión", "%,.1f"),
               "dif": ("Diferencial", "%,.1f"), "pct_cng": ("% congestión", "%.1f"),
               "solar_factor": ("Captura solar %", "%.0f"), "spread4": ("Spread 4 h", "%,.0f"),
               "margen_bess": ("Margen BESS /MW-año", "%,.0f")}


def tabla_anual(anual):
    df = anual[[c for c in TABLA_ANUAL if c in anual.columns]].copy()
    df.insert(0, "anio", df.index.astype(str))
    st.dataframe(df, width="stretch", hide_index=True,
                 column_config={"anio": "Año", **{c: st.column_config.NumberColumn(t, format=f)
                                                 for c, (t, f) in TABLA_ANUAL.items()}})


def render_ficha(clave, f, u, catalogo, pref):
    info = catalogo.get(clave)
    if f.get("vacio"):
        st.warning(f"Sin datos de {clave} en el período elegido.")
        return
    k = f["kpis"]
    uu = f"{u}/MWh"
    st.markdown(f"#### {nombre_nodo(clave, catalogo)} · `{clave}`")
    st.caption(f"{etiqueta_nodo(clave, info)} · Sistema {(info or {}).get('sistema') or 'SIN'} · "
               f"datos del {k['desde']:%d-%m-%Y} al {k['hasta']:%d-%m-%Y} · {k['horas']:,} horas")
    render_kpis(k, u)
    tabs = st.tabs(["📌 Lecturas", "📈 Historia", "🔌 Congestión y pérdidas", "🕐 Distribuciones",
                    "☀️🔋 Solar y baterías"])
    with tabs[0]:
        st.markdown("\n".join(f"- {t}" for t in f["lecturas"]))
        st.caption("Lecturas automáticas con reglas fijas sobre las cifras del período: revisa las gráficas "
                   "antes de citarlas.")
    with tabs[1]:
        _grafica(gr.historia(f["mensual"], uu), f"{pref}hist")
        _grafica(gr.barras_anuales(f["anual"], "promedio", "PML promedio por año", uu), f"{pref}anual")
        st.markdown("**Indicadores por año**")
        tabla_anual(f["anual"])
    with tabs[2]:
        st.caption("El componente de energía es el mismo para todo el sistema en cada hora. Lo propio del "
                   "nodo es el diferencial: pérdidas + congestión. Negativo = la red castiga inyectar aquí "
                   "(zona exportadora); positivo = la energía llega con restricciones (zona importadora).")
        _grafica(gr.componentes_mensuales(f["mensual"], uu), f"{pref}compm")
        _grafica(gr.mapa_calor(f["mes_hora_cng"], "Congestión promedio por mes y hora", uu, divergente=True),
                 f"{pref}cngmh")
        _grafica(gr.componentes_horarios(f["horario"], uu), f"{pref}comph")
        _grafica(gr.barras_anuales(f["anual"], "pct_cng", "% de horas con congestión, por año", "% de horas",
                                   ",.1f", color=gr.C_CONGESTION), f"{pref}pctcng")
    with tabs[3]:
        _grafica(gr.mapa_calor(f["mes_hora"], "PML promedio por mes y hora", uu), f"{pref}pmlmh")
        _grafica(gr.perfil_horario(f["horario"], uu), f"{pref}perfil")
        if len(f["horario_anual"]) > 1:
            _grafica(gr.perfil_por_anio(f["horario_anual"], uu), f"{pref}perfilanio")
        _grafica(gr.estacionalidad(f["estacional"], uu), f"{pref}estac")
        _grafica(gr.curva_duracion(f["duracion"], uu), f"{pref}durac")
    with tabs[4]:
        c = st.columns(4)
        c[0].metric(f"Captura solar ({uu})", _n(k["solar_captura"]))
        c[1].metric("Factor de captura solar", f"{_n(k['solar_factor'])} %")
        c[2].metric(f"Spread 4 h, P90 ({uu})", _n(k["spread4_p90"]))
        c[3].metric(f"Margen BESS 4 h ({u} mil/MW-año)", _n(k["margen_bess"] / 1000 if k["margen_bess"] else None))
        if len(f["anual"]) > 1:
            _grafica(gr.barras_anuales(f["anual"], "solar_factor", "Factor de captura solar por año",
                                       "% del promedio", ",.0f"), f"{pref}solaranio")
        _grafica(gr.horas_bateria(f["bess_horas"]), f"{pref}bessh")
        _grafica(gr.spread_mensual(f["mensual"], uu), f"{pref}spread")
        if len(f["anual"]) > 1:
            _grafica(gr.barras_anuales(f["anual"].assign(m=f["anual"]["margen_bess"] / 1000), "m",
                                       f"Margen teórico BESS 4 h / 1 ciclo por año (miles de {u}/MW-año)",
                                       f"miles de {u}/MW-año", color=gr.SERIE[1]), f"{pref}margen")
        st.caption("Perfil solar genérico: seno entre 6:30 y 19:00 (hora CENACE), igual todo el año; sirve para "
                   "comparar nodos y años, no para un proyecto concreto. Batería: descarga en las 4 horas más caras "
                   "y carga en las 4 más baratas de cada día, sin restricción de orden, η 85 %: cota superior "
                   "del arbitraje, no ingreso esperado.")


def render_comparacion(comp, fichas, catalogo, colores, u):
    uu = f"{u}/MWh"
    claves = comp["claves"]
    etiquetas = {c: f"{nombre_nodo(c, catalogo)} ({c})" for c in claves}
    st.markdown("**Lecturas de la comparación**")
    st.markdown("\n".join(f"- {t}" for t in comp["lecturas"]))
    t = comp["tabla"].copy()
    t["desde"], t["hasta"] = pd.to_datetime(t["desde"]), pd.to_datetime(t["hasta"])
    cfg = {"nombre": st.column_config.TextColumn("Nombre"),
           "desde": st.column_config.DateColumn("Desde", format="DD/MM/YYYY"),
           "hasta": st.column_config.DateColumn("Hasta", format="DD/MM/YYYY")}
    for col, titulo in informe.COLUMNAS_TABLA.items():
        if col not in cfg:
            fmt = "%.1f" if col.startswith("pct") or col == "solar_factor" else "%,.0f"
            cfg[col] = st.column_config.NumberColumn(titulo, format=fmt)
    st.dataframe(t, width="stretch", column_config=cfg)
    if len(claves) > len(gr.SERIE):
        st.caption(f"En las gráficas, los primeros {len(gr.SERIE)} nodos van en color y el resto en gris.")

    tabs = st.tabs(["📈 Precios", "🔌 Red", f"↔️ Diferencial frente a {comp['ref']}", "☀️🔋 Solar y baterías",
                    "🔗 Correlación"])
    with tabs[0]:
        _grafica(gr.lineas_nodos(comp["mensual"], colores, "PML promedio mensual", uu, etiquetas=etiquetas),
                 "cmp_mensual")
        _grafica(gr.lineas_nodos(comp["horario"], colores, "Perfil intradiario promedio", uu, gr.HORA_TXT,
                                 etiquetas, por_hora=True), "cmp_horario")
        _grafica(gr.barras_nodos(comp["tabla"]["promedio"], colores, "PML promedio del período", uu), "cmp_prom")
    with tabs[1]:
        _grafica(gr.lineas_nodos(comp["horario_dif"], colores, "Diferencial nodal por hora (pérdidas + congestión)",
                                 uu, gr.HORA_TXT, etiquetas, por_hora=True), "cmp_dif_h")
        _grafica(gr.lineas_nodos(comp["horario_cng"], colores, "Congestión promedio por hora", uu, gr.HORA_TXT,
                                 etiquetas, por_hora=True), "cmp_cng_h")
        _grafica(gr.barras_nodos(comp["tabla"]["pct_cng"], colores, "% de horas con congestión", "% de horas",
                                 ",.1f"), "cmp_pct_cng")
    with tabs[2]:
        st.caption("Nodo − referencia, solo en las horas en que ambos tienen datos. La referencia se cambia "
                   "arriba de la tabla.")
        if comp["dif_mensual"].empty:
            st.info("No hay horas en común con la referencia en el período.")
        else:
            _grafica(gr.lineas_nodos(comp["dif_mensual"], colores, f"Diferencial mensual frente a {comp['ref']}",
                                     uu, etiquetas=etiquetas), "cmp_dif_m")
            _grafica(gr.lineas_nodos(comp["dif_horario"], colores, f"Diferencial por hora frente a {comp['ref']}",
                                     uu, gr.HORA_TXT, etiquetas, por_hora=True), "cmp_dif_hr")
    with tabs[3]:
        _grafica(gr.barras_nodos(comp["tabla"]["solar_factor"], colores, "Factor de captura solar",
                                 "% del promedio"), "cmp_solar")
        _grafica(gr.barras_nodos(comp["tabla"]["spread4"], colores, "Spread diario 4 h promedio", uu), "cmp_spread")
        _grafica(gr.barras_nodos(comp["tabla"]["margen_bess"] / 1000, colores,
                                 "Margen teórico BESS 4 h / 1 ciclo", f"miles de {u}/MW-año"), "cmp_margen")
    with tabs[4]:
        if comp["correlacion"].dropna(how="all").shape[0] >= 2:
            _grafica(gr.matriz_correlacion(comp["correlacion"]), "cmp_corr")
        else:
            st.info("Hace falta al menos un mes de días en común para calcular correlaciones.")


def render_mapa(claves, fichas, catalogo, colores, u):
    with st.expander("📍 Ubicación aproximada en el mapa (OpenStreetMap)"):
        st.caption("El CENACE no publica coordenadas: se busca la subestación de OpenStreetMap con nombre "
                   "parecido en el mismo estado. La primera consulta tarda ~30 s.")
        ub = st.session_state.get("ubicaciones")
        if st.button("📍 Ubicar nodos", key="btn_ubicar"):
            with st.spinner("Consultando OpenStreetMap…"):
                subs = subestaciones_osm()
            if not subs:
                st.session_state["ubicaciones"] = ub = None
                st.warning("OpenStreetMap no respondió. Vuelve a intentarlo en unos minutos.")
            else:
                st.session_state["ubicaciones"] = ub = {"claves": tuple(claves),
                                                       "df": pd.DataFrame(matchear_nodos(claves, catalogo, subs))}
        if ub and ub["claves"] == tuple(claves):
            df = ub["df"]
            puntos = df.assign(nombre=df["nombre_cenace"], color=df["clave"].map(colores).fillna(gr.OTROS),
                               promedio=df["clave"].map(lambda c: (fichas.get(c) or {}).get("kpis", {}).get("promedio")))
            _grafica(gr.mapa(puntos, f"{u}/MWh"), "mapa")
            st.dataframe(df[["clave", "nombre_cenace", "nombre_osm", "calidad", "estado"]], hide_index=True,
                         width="stretch", column_config={"clave": "Clave", "nombre_cenace": "Nombre CENACE",
                                                         "nombre_osm": "Subestación OSM", "calidad": "Ubicación",
                                                         "estado": "Estado"})


def render_descargas(claves, dfs, fichas, comp, catalogo, u, ini, fin, fx_txt, umbral_txt):
    nombres = {c: nombre_nodo(c, catalogo) for c in claves}
    meta = {"unidad": u, "periodo": f"{ini:%d-%m-%Y} a {fin:%d-%m-%Y}", "fx": fx_txt, "umbral_txt": umbral_txt,
            "generado": datetime.now().strftime("%Y-%m-%d %H:%M")}
    base = f"PML_MTR_{'_'.join(claves[:4])}{'_etc' if len(claves) > 4 else ''}_{ini:%Y%m%d}-{fin:%Y%m%d}_{u}"
    filas = sum(len(d) for d in dfs.values())
    st.markdown("##### ⬇️ Descargas")
    c = st.columns(3)
    c[0].download_button("📑 Informe (Excel)", data=lambda: excel_informe(meta, fichas, comp, nombres),
                         file_name=f"{base}_informe.xlsx", on_click="ignore", width="stretch",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_informe")
    c[1].download_button(f"🕐 Datos horarios (Excel, {filas:,} filas)", data=lambda: excel_horario(dfs, u, nombres),
                         file_name=f"{base}_horario.xlsx", on_click="ignore", width="stretch",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_horario")
    c[2].download_button("🕐 Datos horarios (CSV)", data=lambda: csv_horario(dfs), file_name=f"{base}_horario.csv",
                         on_click="ignore", width="stretch", mime="text/csv", key="dl_csv")
    st.caption("Los archivos se generan al pulsar; con años de datos horarios el Excel tarda unos segundos.")


# ═══════════════════════════════════════════════════════════════════════
# SELECCIÓN DE NODOS Y ENLACES
# ═══════════════════════════════════════════════════════════════════════
def agregar_nodos(claves, reemplazar=False, limpiar_importacion=False):
    """Programa claves para el buscador y vuelve a ejecutar (se aplican antes de dibujarlo)."""
    st.session_state["_nodos_pendientes"] = {"claves": list(claves), "reemplazar": reemplazar,
                                             "limpiar_importacion": limpiar_importacion}
    st.rerun()


def aplicar_nodos_pendientes():
    pend = st.session_state.pop("_nodos_pendientes", None)
    if not pend:
        return
    actuales = [] if pend["reemplazar"] else list(st.session_state.get("nodos_sel", []))
    for c in pend["claves"]:
        if c not in actuales:
            actuales.append(c)
    st.session_state["nodos_sel"] = actuales
    if pend.get("limpiar_importacion"):
        st.session_state["import_texto"] = ""


def aplicar_query_params():
    """En la primera carga, precarga nodos, moneda y período del enlace y genera el informe."""
    if st.session_state.get("_qp_aplicados"):
        return
    st.session_state["_qp_aplicados"] = True
    qp = st.query_params
    if qp.get("nodos"):
        st.session_state["nodos_sel"] = parse_lista_claves(qp["nodos"])[:MAX_NODOS_INFORME]
        st.session_state["_autogenerar"] = True
    if qp.get("moneda", "").upper() in MONEDAS:
        st.session_state["moneda_sel"] = MONEDAS[qp["moneda"].upper()]
    if qp.get("periodo"):
        aplicar_token_periodo(qp["periodo"])
    elif qp.get("desde") and qp.get("hasta"):                  # enlaces de la versión anterior
        aplicar_token_periodo(f"{qp['desde']}_{qp['hasta']}")


def url_app():
    try:
        url = st.context.url or ""
    except Exception:
        url = ""
    return url.split("?")[0].replace("/~/+", "").rstrip("/")


# ═══════════════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <h1>⚡ PML CENACE <span class="pml-badge">MTR</span></h1>
    <p>Informe de nodos del Mercado de Tiempo Real · precios, congestión, pérdidas, perfiles para solar y
    baterías · MXN o USD — Recurrent Energy / Canadian Solar</p>
</div>
""", unsafe_allow_html=True)

catalogo, catalogo_version = cargar_catalogo()
llave_catalogo = (catalogo_version, len(catalogo))
if catalogo:
    st.caption(f"📚 Catálogo de NodosP del CENACE: **{len(catalogo):,} nodos** · {catalogo_version} "
               f"(se actualiza solo cada semana)")
else:
    st.warning("⚠️ Catálogo no cargado.")

st.session_state.setdefault("moneda_sel", MONEDAS["MXN"])
st.session_state.setdefault("nodos_sel", [])
aplicar_query_params()
aplicar_nodos_pendientes()

novedades = cargar_novedades_catalogo()
if catalogo and novedades.get("version") == catalogo_version:
    altas, bajas, cambios = (novedades.get(k, []) for k in ("altas", "bajas", "cambios"))
    with st.expander(f"🆕 Novedades del catálogo {catalogo_version} frente a "
                     f"{novedades.get('version_anterior', '—')}: {len(altas)} NodosP nuevos · "
                     f"{len(bajas)} bajas · {len(cambios)} cambios de datos"):
        if altas:
            st.markdown("**NodosP nuevos**")
            st.dataframe(pd.DataFrame(altas)[["clave", "nombre", "zona", "kv", "estado", "ccr"]],
                         hide_index=True, width="stretch")
            if st.button(f"➕ Agregar los {len(altas)} nuevos al buscador", key="btn_add_novedades"):
                agregar_nodos([a["clave"] for a in altas])
        if bajas:
            st.markdown("**Dados de baja:** " + ", ".join(f"{b['clave']} ({b.get('nombre', '')})" for b in bajas))
        if cambios:
            st.markdown("**Cambios de datos**")
            st.dataframe(pd.DataFrame(cambios).astype(str), hide_index=True, width="stretch")
        if novedades.get("control_de_cambios"):
            st.markdown("**Control de cambios del CENACE**")
            st.markdown("\n".join(f"- {c['descripcion']} ({c['version']})" for c in novedades["control_de_cambios"]))
        st.caption(f"Fuente: CENACE, Catálogo de NodosP · revisado el {str(novedades.get('revisado', ''))[:10]}")

guardados = nodos_guardados()

with st.sidebar:
    st.markdown(f"<h2 style='color:{RE_BLUE};margin-bottom:0;font-weight:700;'>⚡ Informe de nodos</h2>",
                unsafe_allow_html=True)
    st.caption("PML del Mercado de Tiempo Real (MTR)")

    moneda_sel = st.radio("Moneda", options=list(MONEDAS.values()), key="moneda_sel", horizontal=True,
                          help="**USD**: cada hora se divide entre el FIX de su día de operación (Banxico, o "
                               "DOF si no hay token); fines de semana y feriados usan el último FIX publicado.")
    moneda = "USD" if "USD" in moneda_sel else "MXN"

    st.markdown("**📍 Nodos**")
    caja_buscador = st.container()
    claves_orden, etiquetas = indice_buscador(llave_catalogo, catalogo) if catalogo else ([], {})

    claves_filtradas = None
    with st.expander("🎯 Acotar búsqueda (sistema, estado, zona, CCR, kV)"):
        if catalogo:
            df_f = catalogo_df(llave_catalogo, catalogo)
            sistemas_disp = sorted(s for s in df_f["sistema"].dropna().unique() if s)
            sis_sel = st.selectbox("🌎 Sistema", ["Todos"] + sistemas_disp, key="filt_sis")
            activos = sis_sel != "Todos"
            if activos:
                df_f = df_f[df_f["sistema"] == sis_sel]
            for campo, etiqueta_f, clave_w in (("estado", "🏛️ Estado(s)", "filt_est"),
                                               ("municipio", "🏘️ Municipio(s)", "filt_mun"),
                                               ("zona", "🗺️ Zona de carga", "filt_zona"),
                                               ("ccr", "📍 CCR (Centro de Control Regional)", "filt_ccr")):
                disponibles = sorted(v for v in df_f[campo].dropna().unique() if v)
                elegidos = st.multiselect(etiqueta_f, disponibles, key=clave_w)
                if elegidos:
                    df_f = df_f[df_f[campo].isin(elegidos)]
                    activos = True
            kvs_disp = sorted(int(k) for k in df_f["kv"].dropna().unique() if k > 0)
            kv_sel = st.multiselect("⚡ Tensión (kV)", kvs_disp, key="filt_kv")
            if kv_sel:
                df_f = df_f[df_f["kv"].isin(kv_sel)]
                activos = True
            if activos:
                claves_filtradas = set(df_f["clave"])
                n_f = len(claves_filtradas)
                st.markdown(f"**{n_f:,} nodos** quedan en el buscador")
                if 0 < n_f <= MAX_AGREGAR_FILTRO:
                    if st.button(f"➕ Agregar los {n_f} a la selección", key="btn_add_filtrados", width="stretch"):
                        agregar_nodos([c for c in claves_orden if c in claves_filtradas])
                elif n_f > MAX_AGREGAR_FILTRO:
                    st.caption(f"Acota más para agregarlos de una vez (máximo {MAX_AGREGAR_FILTRO}).")
            else:
                st.caption("Sin filtros: el buscador ofrece todo el catálogo.")
        else:
            st.warning("No hay catálogo cargado.")

    with st.expander("📋 Importar lista de claves"):
        st.text_area("Claves", key="import_texto", height=100, label_visibility="collapsed",
                     placeholder="Una clave por línea o separadas por comas\n04TST-115\n04EMY-115")
        claves_import = parse_lista_claves(st.session_state.get("import_texto", ""))
        c_imp1, c_imp2 = st.columns(2)
        if c_imp1.button("➕ Agregar", key="btn_import_add", width="stretch"):
            if claves_import:
                agregar_nodos(claves_import, limpiar_importacion=True)
            st.caption("Pega al menos una clave.")
        if c_imp2.button("🔁 Reemplazar", key="btn_import_rep", width="stretch"):
            if claves_import:
                agregar_nodos(claves_import, reemplazar=True, limpiar_importacion=True)
            st.caption("Pega al menos una clave.")
        st.caption("Las claves que no estén en el catálogo se marcan y se consultan como SIN.")

    listas = cargar_listas_guardadas()
    if listas:
        with st.expander("📂 Listas guardadas"):
            nombre_lista = st.selectbox("Lista", list(listas), key="lista_sel", label_visibility="collapsed")
            st.caption(", ".join(f"{c} 💾" if c in guardados else c for c in listas[nombre_lista]))
            c_l1, c_l2 = st.columns(2)
            if c_l1.button("➕ Agregar", key="btn_lista_add", width="stretch"):
                agregar_nodos(listas[nombre_lista])
            if c_l2.button("🔁 Reemplazar", key="btn_lista_rep", width="stretch"):
                agregar_nodos(listas[nombre_lista], reemplazar=True)
            st.caption("💾 = historial completo guardado (se actualiza a diario). Las listas se editan en "
                       "`data/listas_nodos.json` del repositorio; un nodo nuevo en una lista queda guardado "
                       "tras la siguiente corrida de la GitHub Action.")

    with caja_buscador:
        seleccion_previa = st.session_state.get("nodos_sel", [])
        base = claves_orden if claves_filtradas is None else [c for c in claves_orden if c in claves_filtradas]
        en_base = set(base)
        st.multiselect(
            "🔎 Buscar y elegir nodos",
            options=base + [c for c in seleccion_previa if c not in en_base],
            format_func=lambda c: ("💾 " if c in guardados else "") + (etiquetas.get(c) or etiqueta_nodo(c, None)),
            key="nodos_sel",
            # «contains»: «Tast» da solo Tastiota y Maniobras Tastiota (el modo fuzzy daba ~980)
            filter_mode="contains",
            placeholder="Nombre, clave, zona o estado (p. ej. Tast)",
            help="Escribe parte del nombre, la clave, la zona o el estado. 💾 = historial guardado (informe "
                 "inmediato). Un nodo: ficha individual. Dos o más: comparación.",
        )
        sel = list(st.session_state.get("nodos_sel", []))
        demasiados = len(sel) > MAX_NODOS_INFORME
        generar = st.button("📑 Generar informe", type="primary", width="stretch", key="btn_generar",
                            disabled=not sel or demasiados)
        if demasiados:
            st.caption(f"⚠️ Máximo {MAX_NODOS_INFORME} nodos por informe.")
        elif sel:
            st.caption("1 nodo: ficha individual · 2 o más: comparación con la ficha de cada uno.")

    st.divider()
    if st.button("🧹 Limpiar todo", key="btn_clear_all", width="stretch",
                 help="Borra la selección, los filtros y el informe en pantalla."):
        for k in ["nodos_sel", "import_texto", "lista_sel", "informe", "ubicaciones", "periodo_modo",
                  "periodo_anio", "periodo_rango", "vista_sel", "ref_sel", "filt_sis", "filt_est", "filt_mun",
                  "filt_zona", "filt_ccr", "filt_kv"]:
            st.session_state.pop(k, None)
        st.query_params.clear()
        st.rerun()
    st.caption("Recurrent Energy · Canadian Solar")


# ─── Generar (botón o enlace) ───
def periodo_por_defecto(claves):
    """Todo el historial; en comparaciones de nodos con historias distintas, el período común."""
    a, _ = rango_disponible(claves)
    c, _ = rango_disponible(claves, comun=True) if len(claves) > 1 else (None, None)
    return "comun" if a and c and (c - a).days > 31 else "todo"


desde_enlace = st.session_state.pop("_autogenerar", False)
if (generar or desde_enlace) and sel and not demasiados:
    generar_informe(sel, catalogo)
    if generar or "periodo_modo" not in st.session_state:
        st.session_state["periodo_modo"] = periodo_por_defecto(sel)
    st.session_state.pop("vista_sel", None)

inf = st.session_state.get("informe")
nodos_informe = (inf or {}).get("nodos") or []

# ─── Sin informe (o selección distinta): resumen de la selección ───
if not inf or set(sel) != set(nodos_informe):
    if sel:
        en_cat = [n for n in sel if n in catalogo]
        c = st.columns(4)
        c[0].metric("Seleccionados", len(sel))
        c[1].metric("✅ En catálogo", len(en_cat))
        c[2].metric("⚠️ Desconocidos", len(sel) - len(en_cat))
        c[3].metric("💾 Con historial guardado", sum(n in guardados for n in sel))
        desconocidos = [n for n in sel if n not in catalogo]
        if desconocidos:
            st.caption("Fuera del catálogo (se consultan como SIN): " + ", ".join(desconocidos))
    if inf and sel:
        st.info("La selección cambió: pulsa **📑 Generar informe** para actualizar el informe.")
    if not inf:
        st.markdown("### 📑 Informe de nodos")
        st.markdown(
            "Elige uno o varios NodosP en la barra lateral y pulsa **Generar informe**.\n\n"
            "- **Un nodo** → ficha completa: estadísticas de precio, historia, congestión y pérdidas (lo que dice "
            "la red en ese punto), distribuciones anual e intradiaria, métricas para solar y baterías y "
            "lecturas automáticas.\n"
            "- **Dos o más** → comparación lado a lado (precios, red, diferenciales, solar, baterías y "
            "correlación) y la ficha de cada uno.\n\n"
            "El informe parte de **todo el historial disponible** del nodo; después ajustas el período sin "
            "volver a descargar.")
        if guardados:
            st.markdown("💾 **Historial guardado** (informe inmediato, se actualiza a diario): "
                        + ", ".join(f"{nombre_nodo(c, catalogo)} ({c})" for c in guardados))
        st.caption(f"Los demás nodos se descargan del CENACE al pedirlos: todo el historial (MTR desde el "
                   f"{dp.INICIO_MTR:%d-%m-%Y}) suele tardar 1-3 min; si el CENACE está lento, primero los "
                   f"últimos 12 meses y un botón completa el resto. Quedan en memoria mientras la app está "
                   f"despierta.")

# ─── Informe ───
if inf:
    claves = [c for c in nodos_informe]
    alm = almacen()
    if any(alm.get(c) is None for c in claves):            # la app se reinició: recuperar lo guardado
        cargar_guardados(claves)
    if any(alm.get(c) is None for c in claves):
        st.warning("Los datos de este informe ya no están en memoria (la app se reinició). Pulsa "
                   "**📑 Generar informe** para volver a armarlo.")
        st.stop()

    st.divider()
    if len(claves) == 1:
        st.markdown(f"### 📑 {nombre_nodo(claves[0], catalogo)} ({claves[0]})")
    else:
        st.markdown(f"### 📑 Comparación de {len(claves)} nodos")

    for aviso in inf.get("avisos", []):
        st.warning(aviso)

    # Ampliar el historial de los nodos sin historial guardado
    incompletos = historial_incompleto(claves)
    if incompletos:
        n_todo, s_todo = estimar(faltantes_para(list(incompletos), "todo", dp.ultimo_dia_mtr()), catalogo)
        n_anio, s_anio = estimar(faltantes_para(list(incompletos), "anio", dp.ultimo_dia_mtr()), catalogo)
        with st.container(border=True):
            st.markdown("⏳ **Historial parcial:** " + "; ".join(f"{c} {t}" for c, t in incompletos.items()) + ".")
            c1, c2, c3 = st.columns([1, 1, 2])
            mas_anio = c1.button(f"⏪ Un año más (~{_duracion(s_anio)})", key="btn_anio", width="stretch",
                                 disabled=n_anio == 0)
            todo = c2.button(f"⏮ Completar desde 2017 (~{_duracion(s_todo)})", key="btn_todo", width="stretch",
                             disabled=n_todo == 0)
            c3.caption(f"{n_todo} consultas al CENACE para el historial completo. Lo descargado queda en "
                       f"memoria mientras la app está despierta; para guardarlo siempre, agrega el nodo a "
                       f"una lista guardada.")
        if mas_anio or todo:
            generar_informe(claves, catalogo, modo="anio" if mas_anio else "todo")
            st.session_state["periodo_modo"] = periodo_por_defecto(claves)
            st.rerun()

    a, b = rango_disponible(claves)
    if a is None:
        st.error("El CENACE no devolvió datos de estos nodos (¿nodo nuevo, de otro sistema o sin MTR "
                 "publicado aún?).")
        if inf.get("errores"):
            st.dataframe(tabla_errores(inf["errores"]), hide_index=True, width="stretch")
        st.stop()

    ini, fin = controles_periodo(a, b, rango_disponible(claves, comun=True) if len(claves) > 1 else (None, None))

    # Moneda y tipo de cambio
    fx, fx_txt, fx_id, u = None, "", "", "MXN"
    if moneda == "USD":
        with st.spinner("Tipo de cambio FIX…"):
            fx, fuente = serie_fx(ini, fin, obtener_token_banxico())
        if fx is None:
            st.error("❌ No se pudo obtener el tipo de cambio FIX (ni de Banxico ni del DOF). **El informe se "
                     "muestra en MXN.** Vuelve a intentarlo en unos minutos.")
        else:
            u = "USD"
            per = fx.loc[pd.Timestamp(ini):pd.Timestamp(fin)]
            fx_id = hashlib.md5(fx.to_numpy().tobytes()).hexdigest()[:12]
            fx_txt = (f"FIX del día de operación (fines de semana y feriados: último FIX publicado) · {fuente} · "
                      f"TC promedio {per.mean():.4f}, rango {per.min():.4f}–{per.max():.4f}")

    # Datos y fichas del período
    dfs, fichas = {}, {}
    with st.spinner("Calculando el informe…"):
        for c in claves:
            s = alm.get(c)
            dfs[c] = datos_periodo(c, s.version, ini, fin, u, fx_id, s.datos, fx)
            fichas[c] = ficha_nodo(c, s.version, ini, fin, u, fx_id, dfs[c])
    con_datos = [c for c in claves if not fichas[c].get("vacio")]
    origen = []
    if any(c in guardados for c in claves):
        origen.append("💾 historial guardado")
    if any(c not in guardados for c in claves):
        origen.append("descarga del CENACE")
    st.caption(f"MTR · {u}/MWh · {ini:%d-%m-%Y} a {fin:%d-%m-%Y} · datos disponibles hasta el {b:%d-%m-%Y} "
               f"(el CENACE publica el MTR 7 días después) · fuente: {' + '.join(origen)}"
               + (f" · {fx_txt}" if fx_txt else ""))
    if inf.get("errores"):
        with st.expander(f"⚠️ {len(inf['errores'])} consultas al CENACE con error (esas semanas faltan)"):
            st.dataframe(tabla_errores(inf["errores"][:200]), hide_index=True, width="stretch")

    params = {"nodos": ",".join(claves), "moneda": u, "periodo": token_periodo(ini, fin, a, b)}
    if dict(st.query_params) != params:
        st.query_params.from_dict(params)
    with st.expander("🔗 Enlace de este informe"):
        st.code(f"{url_app()}/?{urlencode(params)}", language=None)
        st.caption("Abre el informe con los mismos nodos, moneda y período.")

    colores = gr.colores_nodos(claves)
    comp = None
    if len(claves) == 1:
        render_ficha(claves[0], fichas[claves[0]], u, catalogo, pref="f0_")
    elif len(con_datos) < 2:
        st.info("Solo uno de los nodos tiene datos en el período: se muestra su ficha.")
        if con_datos:
            render_ficha(con_datos[0], fichas[con_datos[0]], u, catalogo, pref="f0_")
    else:
        opciones = [COMPARACION] + claves
        _sanear("vista_sel", opciones, COMPARACION)
        vista = st.radio("Ver", opciones, horizontal=True, key="vista_sel",
                         format_func=lambda c: "📊 Comparación" if c == COMPARACION else f"{nombre_nodo(c, catalogo)} ({c})")
        _sanear("ref_sel", con_datos, con_datos[0])
        if vista == COMPARACION:
            st.selectbox("Nodo de referencia para los diferenciales", con_datos, key="ref_sel",
                         format_func=lambda c: f"{nombre_nodo(c, catalogo)} ({c})")
        llave = (tuple((c, alm.get(c).version) for c in claves), ini, fin, fx_id)
        comp = comparacion_nodos(llave, st.session_state["ref_sel"], u, {c: dfs[c] for c in con_datos},
                                 {c: fichas[c] for c in con_datos},
                                 {c: nombre_nodo(c, catalogo) for c in claves})
        if vista == COMPARACION:
            render_comparacion(comp, fichas, catalogo, colores, u)
        else:
            render_ficha(vista, fichas[vista], u, catalogo, pref=f"f_{vista}_")

    st.divider()
    umbral_txt = "1 MXN/MWh" + (f" (≈{1 / per.mean():.3f} USD/MWh)" if u == "USD" else "")
    render_descargas(claves, {c: dfs[c] for c in con_datos}, fichas, comp, catalogo, u, ini, fin, fx_txt, umbral_txt)
    render_mapa(claves, fichas, catalogo, colores, u)

st.divider()
st.caption("⚡ PML CENACE · Recurrent Energy / Canadian Solar · Datos: CENACE (SW-PML, MTR) · FIX: Banxico / DOF")
