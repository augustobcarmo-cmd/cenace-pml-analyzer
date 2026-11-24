"""Datos de PML del CENACE (MTR) para la app y para la GitHub Action.

- Descarga del servicio web SW-PML: cada consulta admite hasta 7 días y 20 NodosP.
  Las consultas siguen una rejilla fija de semanas (desde el inicio del MTR), así los
  nodos con faltantes distintos comparten consultas.
- Cobertura: intervalos de fechas ya consultados de cada nodo. Un día sin datos dentro
  de la cobertura significa que el CENACE no publicó nada para ese nodo ese día (por
  ejemplo, porque el nodo aún no existía).
- Almacén persistente: la rama `datos` del repositorio, con un Parquet por nodo y un
  indice.json. Lo mantiene la GitHub Action con los nodos de las listas guardadas.
- Almacén en memoria: lo que la app descarga mientras está despierta.

Comportamiento del SW-PML verificado el 24-nov-2026:
- Lapso mayor a 7 días o más de 20 nodos: HTTP 400 con mensaje.
- Fechas antes del inicio del mercado (MTR: 27-ene-2017): HTTP 400.
- Un nodo sin datos en el lapso (aún no existía o no está publicado) llega con
  «Valores» vacío; los demás nodos de la consulta llegan completos.
- Un lapso que pasa del último día publicado devuelve solo los días publicados.
- Días de cambio de horario (hasta 2022): 23 o 25 horas.
"""
from __future__ import annotations

import io
import json
import os
import re
import threading
import time
from collections import OrderedDict, defaultdict
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://ws01.cenace.gob.mx:8082/SWPML/SIM"
PROCESO = "MTR"
INICIO_MTR = date(2017, 1, 27)      # inicio del MTR en SIN, BCA y BCS (manual SW-PML)
REZAGO_MTR = 7                      # el MTR del día D se publica el D+7
DIAS_CONSULTA = 7
NODOS_CONSULTA = 20
TIMEOUT = (10, 45)                  # conexión, lectura (s)
SISTEMAS = ("SIN", "BCA", "BCS")

COLUMNAS = ["fecha", "hora", "pml", "pml_ene", "pml_per", "pml_cng"]
PRECIOS = ["pml", "pml_ene", "pml_per", "pml_cng"]

REPO_DATOS = os.environ.get("PML_REPO_DATOS", "augustobcarmo-cmd/cenace-pml-analyzer")
RAMA_DATOS = os.environ.get("PML_RAMA_DATOS", "datos")
INDICE = "indice.json"


def ultimo_dia_mtr(hoy: date | None = None) -> date:
    """Último día de operación con MTR publicado (o por publicarse hoy)."""
    return (hoy or date.today()) - timedelta(days=REZAGO_MTR)


# ─── Intervalos de fechas ────────────────────────────────────────────────
def unir(intervalos) -> list[tuple[date, date]]:
    """Une intervalos [ini, fin] (inclusive) que se traslapan o se tocan."""
    out: list[list[date]] = []
    for a, b in sorted((a, b) for a, b in intervalos if a <= b):
        if out and a <= out[-1][1] + timedelta(days=1):
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return [(a, b) for a, b in out]


def restar(ini: date, fin: date, intervalos) -> list[tuple[date, date]]:
    """Partes de [ini, fin] que no cubren los intervalos."""
    faltan, cursor = [], ini
    for a, b in unir(intervalos):
        if b < cursor:
            continue
        if a > fin:
            break
        if a > cursor:
            faltan.append((cursor, min(fin, a - timedelta(days=1))))
        cursor = max(cursor, b + timedelta(days=1))
        if cursor > fin:
            break
    if cursor <= fin:
        faltan.append((cursor, fin))
    return faltan


def dias(intervalos) -> int:
    return sum((b - a).days + 1 for a, b in unir(intervalos))


def intervalos_a_json(intervalos):
    return [[a.isoformat(), b.isoformat()] for a, b in unir(intervalos)]


def intervalos_de_json(lista):
    out = []
    for par in lista or []:
        try:
            out.append((date.fromisoformat(par[0]), date.fromisoformat(par[1])))
        except (TypeError, ValueError, IndexError):
            continue
    return unir(out)


def bloques(ini: date, fin: date, tope_fin: date | None = None) -> list[tuple[date, date]]:
    """Semanas de la rejilla fija que cubren [ini, fin], recortadas al mercado."""
    ini = max(ini, INICIO_MTR)
    if tope_fin is not None:
        fin = min(fin, tope_fin)
    out = []
    if ini > fin:
        return out
    k = (ini - INICIO_MTR).days // DIAS_CONSULTA
    while True:
        a = INICIO_MTR + timedelta(days=k * DIAS_CONSULTA)
        if a > fin:
            break
        b = a + timedelta(days=DIAS_CONSULTA - 1)
        out.append((a, min(b, tope_fin) if tope_fin else b))
        k += 1
    return out


# ─── Consultas al SW-PML ─────────────────────────────────────────────────
@dataclass(frozen=True)
class Consulta:
    sistema: str
    nodos: tuple
    ini: date
    fin: date

    def url(self, proceso: str = PROCESO) -> str:
        return (f"{BASE_URL}/{self.sistema}/{proceso}/{','.join(self.nodos)}/"
                f"{self.ini:%Y/%m/%d}/{self.fin:%Y/%m/%d}/JSON")


def planificar(faltantes: dict, sistemas: dict, tope_fin: date) -> list[Consulta]:
    """Consultas para cubrir los faltantes {clave: [(ini, fin)]}.

    Agrupa por semana de la rejilla y por sistema, en lotes de 20 nodos. Las
    semanas más recientes van primero: si la descarga se interrumpe, lo reciente
    ya está.
    """
    por_bloque: dict = defaultdict(lambda: defaultdict(list))
    for clave, rangos in faltantes.items():
        sistema = sistemas.get(clave) or "SIN"
        vistos = set()
        for a, b in rangos:
            for blq in bloques(a, b, tope_fin):
                if blq not in vistos:
                    vistos.add(blq)
                    por_bloque[blq][sistema].append(clave)
    consultas = []
    for blq in sorted(por_bloque, reverse=True):
        for sistema, claves in sorted(por_bloque[blq].items()):
            for i in range(0, len(claves), NODOS_CONSULTA):
                consultas.append(Consulta(sistema, tuple(claves[i:i + NODOS_CONSULTA]), *blq))
    return consultas


def normalizar(df: pd.DataFrame) -> pd.DataFrame:
    """Tipos y orden estándar: fecha (día), hora 1-25, precios float64; sin duplicados."""
    if df is None or df.empty:
        return pd.DataFrame({c: pd.Series(dtype="datetime64[ns]" if c == "fecha" else
                                          "int8" if c == "hora" else "float64")
                             for c in COLUMNAS})
    out = df[COLUMNAS].copy()
    out["fecha"] = pd.to_datetime(out["fecha"]).astype("datetime64[ns]")
    out["hora"] = pd.to_numeric(out["hora"], errors="coerce").fillna(0).astype("int8")
    for c in PRECIOS:
        out[c] = pd.to_numeric(out[c], errors="coerce").astype("float64")
    out = out[(out["hora"] >= 1) & out["pml"].notna()]
    return (out.drop_duplicates(["fecha", "hora"], keep="last")
               .sort_values(["fecha", "hora"], ignore_index=True))


def parsear(texto: str) -> dict:
    """{clave: DataFrame} con los nodos que traen datos."""
    obj = json.loads(texto)
    out = {}
    for nd in obj.get("Resultados") or []:
        valores = nd.get("Valores") or []
        if valores:
            out[str(nd.get("clv_nodo", "?"))] = normalizar(pd.DataFrame(valores, columns=COLUMNAS))
    return out


_local = threading.local()


def _sesion() -> requests.Session:
    s = getattr(_local, "sesion", None)
    if s is None:
        s = requests.Session()
        s.verify = False            # la cadena de certificados del CENACE no siempre valida
        _local.sesion = s
    return s


def consultar(consulta: Consulta, reintentos: int = 4, espera: float = 1.0):
    """(datos {clave: df}, None) o (None, error). Reintenta fallas de red y 5xx.

    Tras una falla de red se descarta la sesión: una conexión que el servidor (o un
    intermediario) cerró sin avisar deja colgada la siguiente consulta hasta el timeout.
    """
    error = "sin respuesta"
    for intento in range(reintentos):
        try:
            r = _sesion().get(consulta.url(), timeout=TIMEOUT)
        except requests.RequestException as e:
            error = type(e).__name__
            _local.sesion = None
        except Exception as e:      # noqa: BLE001 — cualquier otra falla cuenta como error de la consulta
            error = f"{type(e).__name__}: {e}"
            _local.sesion = None
        else:
            if r.status_code == 200:
                try:
                    return parsear(r.text), None
                except ValueError:
                    error = "respuesta no es JSON"
            elif r.status_code == 204:
                return {}, None
            elif r.status_code == 400:      # fechas, lapso o nodos inválidos: no se reintenta
                try:
                    msg = r.json().get("Message", "")
                except ValueError:
                    msg = r.text[:120]
                return None, f"HTTP 400 {msg}".strip()
            else:
                error = f"HTTP {r.status_code}"
                if r.status_code == 429:
                    time.sleep(espera * 4 * 2 ** intento)
        if intento < reintentos - 1:
            time.sleep(espera * 2 ** intento)
    return None, error


def ejecutar(consultas, al_llegar, workers: int = 8, progreso=None, limite: float | None = None,
             espera: float = 1.0, gracia: float = 20.0):
    """Corre las consultas en paralelo; al_llegar(consulta, datos) en el hilo que llama.

    limite: time.time() a partir del cual no se lanzan consultas nuevas; las que siguen en
    curso se esperan `gracia` segundos más y luego se abandonan (cuentan como pendientes).
    Devuelve (errores [{consulta, error}], pendientes sin resultado).
    """
    errores, pendientes = [], list(consultas)
    total, hechas = len(pendientes), 0
    pendientes.reverse()                    # pop() saca la primera
    ex = ThreadPoolExecutor(max_workers=max(1, workers))
    en_curso = {}
    try:
        while pendientes or en_curso:
            ahora = time.time()
            lanzar = limite is None or ahora < limite
            while pendientes and len(en_curso) < workers * 2 and lanzar:
                c = pendientes.pop()
                en_curso[ex.submit(consultar, c, 4, espera)] = c
            if not en_curso:
                break
            if limite is None:
                tope = None
            elif lanzar:
                tope = max(0.5, limite - ahora)          # despertar al llegar al límite
            else:
                tope = limite + gracia - ahora
                if tope <= 0:
                    break                                  # se abandonan las que siguen colgadas
            listos, _ = wait(list(en_curso), timeout=tope, return_when=FIRST_COMPLETED)
            for fut in listos:
                c = en_curso.pop(fut)
                datos, error = fut.result()
                hechas += 1
                if error is None:
                    al_llegar(c, datos)
                else:
                    errores.append({"consulta": c, "error": error})
                if progreso:
                    progreso(hechas, total)
    finally:
        ex.shutdown(wait=False, cancel_futures=True)
    pendientes.reverse()
    return errores, list(en_curso.values()) + pendientes


def cubierto(consulta: Consulta, datos: dict, clave: str, ultimo: date):
    """Intervalo que queda consultado para el nodo tras una consulta exitosa.

    Las semanas viejas quedan cubiertas aunque vengan vacías (el nodo no existía).
    En los últimos días, solo hasta el último día con datos: lo no publicado se
    vuelve a pedir.
    """
    if consulta.fin < ultimo - timedelta(days=3):
        return (consulta.ini, consulta.fin)
    df = datos.get(clave)
    if df is None or df.empty:
        return None
    hasta = min(consulta.fin, df["fecha"].max().date())
    return (consulta.ini, hasta) if hasta >= consulta.ini else None


def fusionar(*partes) -> pd.DataFrame:
    partes = [p for p in partes if p is not None and not p.empty]
    if not partes:
        return normalizar(None)
    if len(partes) == 1:
        return partes[0]
    return (pd.concat(partes, ignore_index=True)
              .drop_duplicates(["fecha", "hora"], keep="last")
              .sort_values(["fecha", "hora"], ignore_index=True))


class Recolector:
    """Junta las respuestas de una descarga por nodo, con su cobertura."""

    def __init__(self, ultimo: date):
        self.ultimo = ultimo
        self.partes = defaultdict(list)
        self.cubiertos = defaultdict(list)
        self.consultas = 0

    def __call__(self, consulta: Consulta, datos: dict):
        self.consultas += 1
        for clave in consulta.nodos:
            df = datos.get(clave)
            if df is not None and not df.empty:
                self.partes[clave].append(df)
            iv = cubierto(consulta, datos, clave, self.ultimo)
            if iv:
                self.cubiertos[clave].append(iv)

    def vaciar(self):
        """{clave: (df_nuevo, cubiertos)} y deja el recolector vacío."""
        claves = set(self.partes) | set(self.cubiertos)
        out = {c: (fusionar(*self.partes.get(c, [])), list(self.cubiertos.get(c, [])))
               for c in claves}
        self.partes.clear()
        self.cubiertos.clear()
        return out


# ─── Almacén en memoria (app) ────────────────────────────────────────────
@dataclass
class Serie:
    datos: pd.DataFrame = field(default_factory=lambda: normalizar(None))
    cobertura: list = field(default_factory=list)
    version: int = 0
    remoto: str = ""                # «actualizado» del índice remoto ya cargado
    intento_cola: float = 0.0       # time.time() del último intento de traer los días nuevos

    def desde(self):
        return self.datos["fecha"].iloc[0].date() if not self.datos.empty else None

    def hasta(self):
        return self.datos["fecha"].iloc[-1].date() if not self.datos.empty else None


class Almacen:
    """Series por nodo compartidas por todas las sesiones de la app (con candado)."""

    def __init__(self, max_series: int = 60):
        self._series: OrderedDict = OrderedDict()
        self._lock = threading.RLock()
        self.max_series = max_series

    def get(self, clave: str) -> Serie | None:
        with self._lock:
            s = self._series.get(clave)
            if s is not None:
                self._series.move_to_end(clave)
            return s

    def agregar(self, clave: str, df: pd.DataFrame | None = None, cubiertos=(),
                remoto: str | None = None, reemplazar: bool = False) -> Serie:
        with self._lock:
            s = self._series.get(clave) or Serie()
            if reemplazar:
                s = Serie(intento_cola=s.intento_cola)
            if df is not None and not df.empty:
                s.datos = fusionar(s.datos, df)
            s.cobertura = unir(list(s.cobertura) + list(cubiertos))
            if remoto is not None:
                s.remoto = remoto
            s.version += 1
            self._series[clave] = s
            self._series.move_to_end(clave)
            while len(self._series) > self.max_series:
                self._series.popitem(last=False)
            return s

    def marcar_intento_cola(self, clave: str):
        with self._lock:
            s = self._series.get(clave)
            if s is None:
                s = self._series[clave] = Serie()
            s.intento_cola = time.time()

    def claves(self):
        with self._lock:
            return list(self._series)


# ─── Almacén persistente (rama `datos`) ──────────────────────────────────
def archivo_nodo(clave: str, proceso: str = PROCESO) -> str:
    return f"{proceso}/{re.sub(r'[^A-Za-z0-9_.-]', '_', clave)}.parquet"


def url_remota(ruta: str) -> str:
    return f"https://raw.githubusercontent.com/{REPO_DATOS}/{RAMA_DATOS}/{ruta}"


def leer_indice_remoto(timeout: int = 20) -> dict:
    """indice.json de la rama `datos`; lanza RuntimeError si no se puede leer."""
    try:
        r = requests.get(url_remota(INDICE), timeout=timeout)
    except requests.RequestException as e:
        raise RuntimeError(f"GitHub no respondió: {type(e).__name__}") from e
    if r.status_code == 404:
        return {}
    if r.status_code != 200:
        raise RuntimeError(f"GitHub respondió {r.status_code}")
    try:
        return json.loads(r.text)
    except ValueError as e:
        raise RuntimeError("indice.json inválido") from e


def leer_serie_remota(ruta: str, timeout: int = 60) -> pd.DataFrame:
    """Parquet de un nodo en la rama `datos`; lanza RuntimeError si falla."""
    try:
        r = requests.get(url_remota(ruta), timeout=timeout)
    except requests.RequestException as e:
        raise RuntimeError(f"GitHub no respondió: {type(e).__name__}") from e
    if r.status_code != 200:
        raise RuntimeError(f"GitHub respondió {r.status_code} para {ruta}")
    return normalizar(pd.read_parquet(io.BytesIO(r.content)))


def leer_almacen_local(carpeta: str) -> tuple[dict, dict]:
    """(indice, {clave: df}) de una copia local de la rama `datos`."""
    ruta_ind = os.path.join(carpeta, INDICE)
    indice = {}
    if os.path.exists(ruta_ind):
        with open(ruta_ind, encoding="utf-8") as fh:
            indice = json.load(fh)
    series = {}
    for clave, meta in (indice.get("nodos") or {}).items():
        ruta = os.path.join(carpeta, meta.get("archivo") or archivo_nodo(clave))
        if os.path.exists(ruta):
            series[clave] = normalizar(pd.read_parquet(ruta))
    return indice, series


def escribir_nodo(carpeta: str, clave: str, df: pd.DataFrame) -> tuple[str, int]:
    rel = archivo_nodo(clave)
    ruta = os.path.join(carpeta, rel)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    normalizar(df).to_parquet(ruta, index=False, compression="zstd")
    return rel, os.path.getsize(ruta)


def ahora_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
