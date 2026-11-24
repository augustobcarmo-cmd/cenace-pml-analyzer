"""Conversión a USD con el FIX de cada día y caché del DataFrame entre consultas.

    pytest -q tests

El FIX simulado cambia cada día hábil y no existe en fines de semana ni el
1-ene-2026 (feriado): esos días deben usar el último FIX publicado.
"""
import io
import re
from datetime import date, timedelta
from statistics import mean
from unittest import mock

import openpyxl
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

from test_app import APP, NODOS, _Resp, _boton, _radio, fake_requests_post

F_INI, F_FIN = date(2025, 12, 26), date(2026, 1, 8)    # fin de semana y feriado al inicio
FERIADOS = {date(2026, 1, 1)}


def _pml(nodo, dia, hora):
    """PML sintético en MXN que cambia por nodo, mes, día y hora."""
    return round(800 + 37 * (sum(map(ord, nodo)) % 9) + 50 * dia.month + 7 * dia.day + 11 * hora, 2)


def _fix(d):
    return round(17.0 + (d.toordinal() % 97) / 50, 4)


def _fix_aplicable(d):
    """Último FIX publicado en o antes del día d."""
    while d.weekday() >= 5 or d in FERIADOS:
        d -= timedelta(days=1)
    return _fix(d)


def _dias(d0, d1):
    return [d0 + timedelta(days=i) for i in range((d1 - d0).days + 1)]


def fake_session_get(self, url, timeout=None, **kw):
    m = re.search(r"/SIM/(\w+)/(\w+)/([^/]+)/(\d{4})/(\d{2})/(\d{2})/(\d{4})/(\d{2})/(\d{2})/JSON$", url)
    d0 = date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    d1 = date(int(m.group(7)), int(m.group(8)), int(m.group(9)))
    resultados = [{"clv_nodo": n, "Valores": [
        {"fecha": d.isoformat(), "hora": str(h), "pml": f"{_pml(n, d, h):.2f}",
         "pml_ene": f"{_pml(n, d, h) * 0.9:.2f}", "pml_per": "0", "pml_cng": "0"}
        for d in _dias(d0, d1) for h in range(1, 25)]} for n in m.group(3).split(",")]
    return _Resp(200, payload={"Resultados": resultados})


def _habil(d):
    return d.weekday() < 5 and d not in FERIADOS


def fake_sidof(url):
    """DOF: publica el FIX de cada día hábil el siguiente día hábil."""
    m = re.search(r"/indicadores/158/(\d{2})-(\d{2})-(\d{4})/(\d{2})-(\d{2})-(\d{4})$", url)
    d0 = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    d1 = date(int(m.group(6)), int(m.group(5)), int(m.group(4)))
    lista = []
    for d in _dias(d0 - timedelta(days=10), d1):
        if not _habil(d):
            continue
        pub = d + timedelta(days=1)
        while not _habil(pub):
            pub += timedelta(days=1)
        if d0 <= pub <= d1:
            lista.append({"codTipoIndicador": 158, "fecha": pub.strftime("%d-%m-%Y"), "valor": f"{_fix(d):.4f}"})
    return _Resp(200, payload={"messageCode": 200, "ListaIndicadores": lista})


def fake_requests_get(url, headers=None, timeout=None, **kw):
    if "sidof" in url:
        return fake_sidof(url)
    m = re.search(r"/datos/(\d{4}-\d{2}-\d{2})/(\d{4}-\d{2}-\d{2})", url)
    d0, d1 = date.fromisoformat(m.group(1)), date.fromisoformat(m.group(2))
    datos = [{"fecha": d.strftime("%d/%m/%Y"), "dato": f"{_fix(d):.4f}"}
             for d in _dias(d0, d1) if d.weekday() < 5 and d not in FERIADOS]
    return _Resp(200, payload={"bmx": {"series": [{"idSerie": "SF43718", "datos": datos}]}})


def _nueva_app(token=True):
    st.cache_data.clear()
    at = AppTest.from_file(APP, default_timeout=180)
    if token:
        at.secrets["BANXICO_TOKEN"] = "token-de-prueba"
    at.run()
    return at


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


def _consultar(at, f_ini=F_INI, f_fin=F_FIN):
    at.multiselect(key="nodos_sel").set_value(NODOS)
    at.date_input(key="f_ini").set_value(f_ini)
    at.date_input(key="f_fin").set_value(f_fin)
    at.run()
    _boton(at, "⚡ Ejecutar").click()
    at.run()
    assert not at.exception, [e.message for e in at.exception]


def _kpi(at, etiqueta):
    valor = next(m.value for m in at.metric if m.label == etiqueta)
    return float(re.sub(r"[^0-9.\-]", "", str(valor)))


def _promedio(f_ini, f_fin, usd=False):
    return mean(_pml(n, d, h) / (_fix_aplicable(d) if usd else 1)
                for n in NODOS for d in _dias(f_ini, f_fin) for h in range(1, 25))


def test_fix_del_dia_con_fines_de_semana_y_feriados(app):
    _radio(app, "Moneda:").set_value("🇺🇸 USD (dólares)")
    _consultar(app)
    lookup = app.session_state["fx_info"]["tc_lookup"]
    assert lookup["2025-12-26"] == _fix(date(2025, 12, 26))                      # viernes
    assert lookup["2025-12-27"] == lookup["2025-12-28"] == _fix(date(2025, 12, 26))
    assert lookup["2026-01-01"] == _fix(date(2025, 12, 31))                      # feriado
    assert lookup["2026-01-05"] == _fix(date(2026, 1, 5))
    assert abs(_kpi(app, "PML promedio global") - _promedio(F_INI, F_FIN, usd=True)) < 0.01


def test_excel_de_datos_en_usd(app):
    _radio(app, "Moneda:").set_value("🇺🇸 USD (dólares)")
    _consultar(app)
    _radio(app, "Selecciona qué hacer:").set_value("📊 Datos")
    app.run()
    _boton(app, "⚡ Generar Excel de Datos").click()
    app.run()
    assert not app.exception, [e.message for e in app.exception]

    wb = openpyxl.load_workbook(io.BytesIO(app.session_state["excel_datos_bytes"]))
    portada = wb["■ Portada"]
    moneda = next(portada[f"E{r}"].value for r in range(8, 17) if portada[f"B{r}"].value == "Moneda")
    assert moneda.startswith("USD") and "último FIX" in moneda
    ws = wb["04TST-115"]
    assert ws.cell(row=3, column=3).value == "PML (USD$/MWh)"
    assert ws.cell(row=3, column=7).value == "TC FIX (MXN/USD)"
    assert ws.max_row - 3 == 24 * len(_dias(F_INI, F_FIN))
    for r in range(4, ws.max_row + 1):
        d, h = date.fromisoformat(ws.cell(row=r, column=1).value), ws.cell(row=r, column=2).value
        assert ws.cell(row=r, column=7).value == _fix_aplicable(d)
        assert abs(ws.cell(row=r, column=3).value - _pml("04TST-115", d, h) / _fix_aplicable(d)) < 1e-3


def test_consulta_nueva_del_mismo_largo_no_reutiliza_la_anterior(app):
    _consultar(app)                                          # 14 días: dic-ene
    assert abs(_kpi(app, "PML promedio global") - _promedio(F_INI, F_FIN)) < 0.01
    otra_ini, otra_fin = date(2026, 3, 2), date(2026, 3, 15)  # 14 días: marzo
    _consultar(app, otra_ini, otra_fin)
    assert abs(_kpi(app, "PML promedio global") - _promedio(otra_ini, otra_fin)) < 0.01


def test_usd_sin_token_usa_el_dof_con_los_mismos_valores():
    with mock.patch("requests.Session.get", fake_session_get), \
         mock.patch("requests.get", fake_requests_get), \
         mock.patch("requests.post", fake_requests_post):
        at = _nueva_app(token=False)
        _radio(at, "Moneda:").set_value("🇺🇸 USD (dólares)")
        _consultar(at)
        fx = at.session_state["fx_info"]
        assert fx["moneda"] == "USD" and fx["fuente"].startswith("DOF")
        assert fx["tc_lookup"]["2026-01-01"] == _fix(date(2025, 12, 31))
        assert fx["tc_lookup"]["2026-01-05"] == _fix(date(2026, 1, 5))
        assert abs(_kpi(at, "PML promedio global") - _promedio(F_INI, F_FIN, usd=True)) < 0.01


def test_sin_tipo_de_cambio_avisa_con_error_y_queda_en_mxn():
    def caido(url, headers=None, timeout=None, **kw):
        raise ConnectionError("sin red")
    with mock.patch("requests.Session.get", fake_session_get), \
         mock.patch("requests.get", caido), \
         mock.patch("requests.post", fake_requests_post):
        at = _nueva_app()
        _radio(at, "Moneda:").set_value("🇺🇸 USD (dólares)")
        _consultar(at)
        assert any("No se pudo obtener el tipo de cambio" in e.value for e in at.error)
        assert at.session_state["fx_info"]["moneda"] == "MXN"
        assert abs(_kpi(at, "PML promedio global") - _promedio(F_INI, F_FIN)) < 0.01
