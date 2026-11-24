"""Prueba de humo de la app con el CENACE, Banxico y OpenStreetMap simulados.

    pytest -q tests

Recorre el modo Análisis (MXN y USD) y el modo Datos (Excel de datos y de
análisis) y verifica que no quede nada de BESS ni de KMZ.
"""
import io
import json
import math
import re
from datetime import date, timedelta
from pathlib import Path
from unittest import mock

import openpyxl
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parents[1] / "app.py")

# SIN (Central, Noroeste, Noreste) + BCA, para probar multi-sistema y spread por CCR
NODOS = ["01AAN-85", "04TST-115", "06MTP-115", "07RIN-161"]
F_INI, F_FIN = date(2025, 12, 20), date(2026, 1, 12)   # cruza de año: multi-año

OSM = {"elements": [
    {"type": "node", "id": 1, "lat": 20.27, "lon": -98.94,
     "tags": {"power": "substation", "name": "Actopan", "voltage": "85000"}},
    {"type": "node", "id": 2, "lat": 28.37, "lon": -111.42,
     "tags": {"power": "substation", "name": "Tastiota", "voltage": "115000"}},
    {"type": "node", "id": 3, "lat": 25.73, "lon": -100.28,
     "tags": {"power": "substation", "name": "Monterrey Potencia", "voltage": "115000"}},
]}


class _Resp:
    def __init__(self, status=200, text="", payload=None):
        self.status_code = status
        self.text = text if payload is None else json.dumps(payload)
        self._payload = payload

    def json(self):
        return self._payload if self._payload is not None else json.loads(self.text)


def _pml(nodo, dia, hora):
    """PML sintético: perfil diario, nivel por nodo y algunas horas negativas."""
    base = 900 + 150 * (sum(map(ord, nodo)) % 7)
    v = base + 400 * math.sin((hora - 6) / 24 * 2 * math.pi) + 10 * dia.day
    return round(v - 2000 if (hora in (12, 13) and dia.day % 5 == 0) else v, 2)


def fake_session_get(self, url, timeout=None, **kw):
    m = re.search(r"/SIM/(\w+)/(\w+)/([^/]+)/(\d{4})/(\d{2})/(\d{2})/(\d{4})/(\d{2})/(\d{2})/JSON$", url)
    assert m, url
    nodos = m.group(3).split(",")
    assert len(nodos) <= 20
    d0 = date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    d1 = date(int(m.group(7)), int(m.group(8)), int(m.group(9)))
    assert 1 <= (d1 - d0).days + 1 <= 7
    resultados = []
    for n in nodos:
        valores, d = [], d0
        while d <= d1:
            for h in range(1, 25):
                p = _pml(n, d, h)
                valores.append({"fecha": d.isoformat(), "hora": str(h), "pml": f"{p:.2f}",
                                "pml_ene": f"{p * 0.95:.2f}", "pml_per": f"{p * 0.03:.2f}",
                                "pml_cng": f"{p * 0.02:.2f}"})
            d += timedelta(days=1)
        resultados.append({"clv_nodo": n, "Valores": valores})
    return _Resp(200, payload={"nombre": "PML", "proceso": m.group(2),
                               "sistema": m.group(1), "Resultados": resultados})


def fake_requests_get(url, headers=None, timeout=None, **kw):
    assert "banxico" in url
    datos, d = [], F_INI - timedelta(days=5)
    while d <= F_FIN:
        if d.weekday() < 5:
            datos.append({"fecha": d.strftime("%d/%m/%Y"), "dato": "18.5000"})
        d += timedelta(days=1)
    return _Resp(200, payload={"bmx": {"series": [{"idSerie": "SF43718", "datos": datos}]}})


def fake_requests_post(url, data=None, headers=None, timeout=None, **kw):
    assert "overpass" in url
    return _Resp(200, payload=OSM)


@pytest.fixture()
def app():
    with mock.patch("requests.Session.get", fake_session_get), \
         mock.patch("requests.get", fake_requests_get), \
         mock.patch("requests.post", fake_requests_post):
        st.cache_data.clear()
        at = AppTest.from_file(APP, default_timeout=180)
        at.secrets["BANXICO_TOKEN"] = "token-de-prueba"
        at.run()
        yield at


def _boton(at, etiqueta):
    return next(b for b in at.button if b.label == etiqueta)


def _radio(at, etiqueta):
    return next(r for r in at.radio if r.label == etiqueta)


def _consultar(at):
    at.multiselect(key="nodos_sel").set_value(NODOS)
    at.date_input(key="f_ini").set_value(F_INI)
    at.date_input(key="f_fin").set_value(F_FIN)
    at.run()
    _boton(at, "⚡ Ejecutar").click()
    at.run()
    assert not at.exception, [e.message for e in at.exception]


def _textos(at):
    partes = [m.value for m in at.markdown] + [c.value for c in at.caption]
    partes += [i.value for i in at.info] + [w.value for w in at.warning]
    return "\n".join(str(p) for p in partes)


def test_portada_sin_bess(app):
    assert not app.exception
    assert "PML CENACE" in _textos(app)
    assert "v2026-07-16" in _textos(app)
    assert _radio(app, "Selecciona qué hacer:").options == ["🔬 Análisis", "📊 Datos"]
    assert not re.search(r"bess|kmz|scoring", _textos(app), re.I)


def test_analisis_mxn(app):
    _consultar(app)
    acumulado = app.session_state["acumulado"]
    assert set(acumulado) == set(NODOS)
    for filas in acumulado.values():          # 24 días × 24 horas, ordenados
        assert len(filas) == 24 * 24
        claves = [(f["fecha"], f["hora"]) for f in filas]
        assert claves == sorted(claves)
    params = app.session_state["consulta_params"]
    assert params["sistema"] == "MULTI" and params["info_sistemas"] == {"SIN": 3, "BCA": 1}
    matches = app.session_state["matches_df"]
    assert matches["lat"].notna().sum() == 3     # 07RIN-161 queda sin ubicar
    labels = [m.label for m in app.metric]
    assert "PML promedio global" in labels and "Ubicados" in labels
    assert not re.search(r"bess|scoring", _textos(app), re.I)


def test_analisis_usd(app):
    _radio(app, "Moneda:").set_value("🇺🇸 USD (dólares)")
    _consultar(app)
    fx = app.session_state["fx_info"]
    assert fx["moneda"] == "USD" and abs(fx["tc_promedio"] - 18.5) < 1e-9
    assert "Análisis en USD" in _textos(app)


def test_datos_excel(app):
    _consultar(app)
    _radio(app, "Selecciona qué hacer:").set_value("📊 Datos")
    app.run()
    assert not app.exception, [e.message for e in app.exception]
    assert "Centro de descargas" in _textos(app)

    _boton(app, "⚡ Generar Excel de Datos").click()
    app.run()
    _boton(app, "⚡ Generar Excel de Análisis").click()
    app.run()
    assert not app.exception, [e.message for e in app.exception]

    wb = openpyxl.load_workbook(io.BytesIO(app.session_state["excel_datos_bytes"]))
    assert wb.sheetnames[:2] == ["■ Portada", "Resumen"] and set(NODOS) <= set(wb.sheetnames)
    ws = wb["04TST-115"]
    fechas = [ws.cell(row=r, column=1).value for r in range(4, ws.max_row + 1)]
    assert fechas[0] == F_INI.isoformat() and fechas[-1] == F_FIN.isoformat()
    assert isinstance(wb["Resumen"].cell(row=3, column=3).value, float)

    wb = openpyxl.load_workbook(io.BytesIO(app.session_state["excel_analisis_bytes"]))
    nombres = wb.sheetnames
    assert {"■ Portada", "📊 Resumen", "🏆 Rankings", "📊 Spread CCR"} <= set(nombres)
    assert sum(n.startswith("📈 ") for n in nombres) == len(NODOS)   # multi-año
    assert not any(re.search(r"bess", n, re.I) for n in nombres)
    rk = [wb["🏆 Rankings"].cell(row=r, column=1).value for r in range(1, wb["🏆 Rankings"].max_row + 1)]
    assert "🔻 Top horas negativas" in rk
