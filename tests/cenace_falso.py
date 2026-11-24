"""CENACE (SW-PML, MTR), Banxico, DOF, GitHub y OpenStreetMap simulados para las pruebas.

El CENACE falso se comporta como el real (verificado el 24-nov-2026): máximo 7 días
y 20 nodos por consulta, HTTP 400 antes del inicio del mercado, «Valores» vacío para
nodos sin datos, y solo publica hasta el día de hoy − 7.
"""
import io
import json
import math
import re
from datetime import date, timedelta
from functools import lru_cache

import numpy as np
import pandas as pd

INICIO_MTR = date(2017, 1, 27)
# Fecha desde la que cada nodo tiene datos (los demás nodos del catálogo: desde el inicio)
INICIO_NODO = {"04EMY-115": date(2025, 10, 1), "04MTS-115": date(2023, 3, 15), "01VMA-230": None}
SISTEMA_NODO = {"07RIN-161": "BCA"}
GUARDADOS = ("04TST-115",)           # nodos que la «rama datos» falsa ya tiene
LLAMADAS = {"cenace": 0, "github": 0}


def hoy():
    return date.today()


def ultimo_publicado():
    return hoy() - timedelta(days=7)


def componentes(nodo, dia, hora):
    """(pml, energía, pérdidas, congestión) deterministas, con perfil solar y congestión."""
    sonora = nodo.startswith("04")
    base = 900 + 25 * (dia.year - 2017) + 60 * (dia.month in (6, 7, 8, 9))
    ene = base + 300 * math.cos(2 * math.pi * (hora - 20) / 24) - 220 * math.exp(-((hora - 13) / 2.5) ** 2)
    ene += 15 * ((dia.toordinal() % 11) - 5)
    per = (-0.06 if sonora else 0.02) * ene + (sum(map(ord, nodo)) % 7)
    cng = -180.0 if (sonora and 10 <= hora <= 16 and dia.toordinal() % 3 == 0) else 0.0
    if nodo.startswith("06") and hora in (19, 20, 21) and dia.day % 4 == 0:
        cng = 120.0
    ene, per = round(ene, 2), round(per, 2)
    return round(ene + per + cng, 2), ene, per, cng


class Resp:
    def __init__(self, status=200, text="", payload=None, content=None):
        self.status_code = status
        self.text = text if payload is None else json.dumps(payload)
        self._payload = payload
        self.content = content if content is not None else self.text.encode()

    def json(self):
        return self._payload if self._payload is not None else json.loads(self.text)


def _inicio(nodo):
    return INICIO_NODO.get(nodo, INICIO_MTR)


def session_get(self, url, timeout=None, **kw):
    LLAMADAS["cenace"] += 1
    m = re.search(r"/SIM/(\w+)/(\w+)/([^/]+)/(\d{4})/(\d{2})/(\d{2})/(\d{4})/(\d{2})/(\d{2})/JSON$", url)
    assert m, url
    sistema, proceso, nodos = m.group(1), m.group(2), m.group(3).split(",")
    assert proceso == "MTR"
    d0 = date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    d1 = date(int(m.group(7)), int(m.group(8)), int(m.group(9)))
    if len(nodos) > 20:
        return Resp(400, payload={"Message": "No se permite ingresar más de 20 nodos"})
    if (d1 - d0).days + 1 > 7 or d1 < d0:
        return Resp(400, payload={"Message": "No se pueden mostrar datos con un lapso mayor a 7 dias. Favor de verificar."})
    if d0 < INICIO_MTR:
        return Resp(400, payload={"Message": "Las fechas son menores al inicio del mercado."})
    resultados = []
    for n in nodos:
        valores, ini = [], _inicio(n)
        if ini is not None and SISTEMA_NODO.get(n, "SIN") == sistema:
            d = max(d0, ini)
            while d <= min(d1, ultimo_publicado()):
                for h in range(1, 25):
                    p, e, pe, c = componentes(n, d, h)
                    valores.append({"fecha": d.isoformat(), "hora": str(h), "pml": f"{p:.2f}",
                                    "pml_ene": f"{e:.2f}", "pml_per": f"{pe:.2f}", "pml_cng": f"{c:.2f}"})
                d += timedelta(days=1)
        resultados.append({"clv_nodo": n, "Valores": valores})
    status = "OK" if any(r["Valores"] for r in resultados) else "ZERO_RESULTS"
    return Resp(200, payload={"nombre": "PML", "proceso": "MTR", "sistema": sistema, "area": "PÚBLICA",
                              "status": status, "Resultados": resultados})


@lru_cache(maxsize=None)
def serie_guardada(nodo, hasta_iso):
    """Lo que la rama datos falsa tiene del nodo: desde su inicio hasta `hasta`."""
    hasta = date.fromisoformat(hasta_iso)
    fechas = pd.date_range(max(_inicio(nodo), INICIO_MTR), hasta, freq="D")
    filas = [(f, h, *componentes(nodo, f.date(), h)) for f in fechas for h in range(1, 25)]
    df = pd.DataFrame(filas, columns=["fecha", "hora", "pml", "pml_ene", "pml_per", "pml_cng"])
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    return buf.getvalue(), len(df)


def indice_guardado():
    hasta = ultimo_publicado().isoformat()
    nodos = {}
    for n in GUARDADOS:
        _, filas = serie_guardada(n, hasta)
        nodos[n] = {"sistema": SISTEMA_NODO.get(n, "SIN"), "cobertura": [[INICIO_MTR.isoformat(), hasta]],
                    "filas": filas, "desde": max(_inicio(n), INICIO_MTR).isoformat(), "hasta": hasta,
                    "archivo": f"MTR/{n}.parquet", "actualizado": f"{hasta}T15:40:00Z"}
    return {"version": 1, "proceso": "MTR", "actualizado": f"{hasta}T15:40:00Z", "nodos": nodos}


def github_get(url):
    LLAMADAS["github"] += 1
    if url.endswith("/indice.json"):
        return Resp(200, payload=indice_guardado())
    m = re.search(r"/MTR/([^/]+)\.parquet$", url)
    if m and m.group(1) in GUARDADOS:
        contenido, _ = serie_guardada(m.group(1), ultimo_publicado().isoformat())
        return Resp(200, text="", content=contenido)
    return Resp(404, text="404: Not Found")


def fix(d):
    return round(17.0 + (d.toordinal() % 97) / 50, 4)


FERIADOS = {date(2026, 1, 1), date(2025, 1, 1)}


def habil(d):
    return d.weekday() < 5 and d not in FERIADOS


def fix_aplicable(d):
    while not habil(d):
        d -= timedelta(days=1)
    return fix(d)


def banxico_get(url):
    m = re.search(r"/datos/(\d{4}-\d{2}-\d{2})/(\d{4}-\d{2}-\d{2})", url)
    d0, d1 = date.fromisoformat(m.group(1)), date.fromisoformat(m.group(2))
    datos, d = [], d0
    while d <= d1:
        if habil(d):
            datos.append({"fecha": d.strftime("%d/%m/%Y"), "dato": f"{fix(d):.4f}"})
        d += timedelta(days=1)
    return Resp(200, payload={"bmx": {"series": [{"idSerie": "SF43718", "datos": datos}]}})


def sidof_get(url):
    """DOF: publica el FIX de cada día hábil el siguiente día hábil."""
    m = re.search(r"/indicadores/158/(\d{2})-(\d{2})-(\d{4})/(\d{2})-(\d{2})-(\d{4})$", url)
    d0 = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    d1 = date(int(m.group(6)), int(m.group(5)), int(m.group(4)))
    lista, d = [], d0 - timedelta(days=10)
    while d <= d1:
        if habil(d):
            pub = d + timedelta(days=1)
            while not habil(pub):
                pub += timedelta(days=1)
            if d0 <= pub <= d1:
                lista.append({"codTipoIndicador": 158, "fecha": pub.strftime("%d-%m-%Y"), "valor": f"{fix(d):.4f}"})
        d += timedelta(days=1)
    return Resp(200, payload={"messageCode": 200, "ListaIndicadores": lista})


OSM = {"elements": [
    {"type": "node", "id": 1, "lat": 20.27, "lon": -98.94,
     "tags": {"power": "substation", "name": "Actopan", "voltage": "85000"}},
    {"type": "node", "id": 2, "lat": 28.37, "lon": -111.42,
     "tags": {"power": "substation", "name": "Tastiota", "voltage": "115000"}},
    {"type": "node", "id": 3, "lat": 25.73, "lon": -100.28,
     "tags": {"power": "substation", "name": "Monterrey Potencia", "voltage": "115000"}},
]}


def requests_get(url, headers=None, timeout=None, **kw):
    if "raw.githubusercontent.com" in url:
        return github_get(url)
    if "sidof" in url:
        return sidof_get(url)
    if "banxico" in url:
        return banxico_get(url)
    raise AssertionError(f"URL inesperada: {url}")


def requests_post(url, data=None, headers=None, timeout=None, **kw):
    assert "overpass" in url
    return Resp(200, payload=OSM)


def pml_esperado(nodo, d0, d1, usd=False):
    """Promedio del PML de un nodo en [d0, d1] según el CENACE falso."""
    vals = []
    d = max(d0, _inicio(nodo) or d1 + timedelta(days=1))
    while d <= min(d1, ultimo_publicado()):
        for h in range(1, 25):
            vals.append(componentes(nodo, d, h)[0] / (fix_aplicable(d) if usd else 1))
        d += timedelta(days=1)
    return float(np.mean(vals)) if vals else float("nan")
