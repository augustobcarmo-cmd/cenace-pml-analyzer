"""Buscador de nodos, importación, listas guardadas, enlaces, novedades del
catálogo y funciones de catalogo.py (sin red)."""
import json
import os
from pathlib import Path
from unittest import mock

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

import catalogo as cat
import cenace_falso as cf
from test_app import APP

RAIZ = Path(APP).parent
CAMBIOS = RAIZ / "data" / "catalogo_cambios.json"


@pytest.fixture()
def app():
    with mock.patch("requests.Session.get", cf.session_get), \
         mock.patch("requests.get", cf.requests_get), \
         mock.patch("requests.post", cf.requests_post):
        st.cache_data.clear()
        st.cache_resource.clear()
        at = AppTest.from_file(APP, default_timeout=180)
        yield at


def _boton(at, etiqueta):
    return next(b for b in at.button if b.label == etiqueta)


# ─── catalogo.py ───
def test_etiqueta_y_lista_de_claves():
    catalogo, version = cat.parse_catalogo(RAIZ / "data" / "catalogo_nodos.xlsx")
    assert version.startswith("v20") and len(catalogo) > 2000
    assert cat.etiqueta_nodo("04TST-115", catalogo["04TST-115"]) == \
        "Tastiota · 04TST-115 · 115 kV · Hermosillo, Sonora · NOROESTE"
    assert "fuera del catálogo" in cat.etiqueta_nodo("XXX-999", None)
    assert cat.parse_lista_claves(" 04tst-115,04EMY-115\n04TST-115\t01VAJ-230; ") == \
        ["04TST-115", "04EMY-115", "01VAJ-230"]
    assert cat.normalizar("Benito Juárez\xa0 Peñitas") == "benito juarez penitas"


def test_versiones_publicadas_toma_la_mas_nueva_y_codifica_la_url():
    html = ('<a href="/Docs/01_MECP/CatalogoNodosP/2026/Catálogo NodosP Sistema Eléctrico Nacional (v2026-10-23).xlsx">x</a>'
            '<a href="/Docs/01_MECP/CatalogoNodosP/2026/Catálogo NodosP Sistema Eléctrico Nacional (v2026-11-18).xlsx">y</a>')
    (v, url), (v2, _) = cat.versiones_publicadas(html)
    assert (v, v2) == ("v2026-11-18", "v2026-10-23")
    assert url == ("https://www.cenace.gob.mx/Docs/01_MECP/CatalogoNodosP/2026/"
                   "Cat%C3%A1logo%20NodosP%20Sistema%20El%C3%A9ctrico%20Nacional%20(v2026-11-18).xlsx")


def test_diff_y_control_de_cambios():
    viejo = {"A": {"nombre": "Uno", "kv": 115}, "B": {"nombre": "Dos", "kv": 230}}
    nuevo = {"A": {"nombre": "Uno", "kv": 230}, "C": {"nombre": "Tres", "kv": 115}}
    d = cat.diff_catalogos(viejo, nuevo)
    assert [a["clave"] for a in d["altas"]] == ["C"] and [b["clave"] for b in d["bajas"]] == ["B"]
    assert d["cambios"] == [{"clave": "A", "campo": "kv", "antes": 115, "despues": 230}]
    filas = cat.leer_control_de_cambios(RAIZ / "data" / "catalogo_nodos.xlsx")
    assert len(filas) > 100 and filas[0]["numero"] == "1" and filas[-1]["version"].startswith("20")


# ─── App ───
def test_buscador_ofrece_todo_el_catalogo_con_etiquetas(app):
    app.run()
    ms = app.multiselect(key="nodos_sel")
    assert len(ms.options) > 2000
    # 💾 = historial guardado en la rama datos (en la prueba, solo 04TST-115)
    assert "💾 Tastiota · 04TST-115 · 115 kV · Hermosillo, Sonora · NOROESTE" in ms.options
    assert "Maniobras Tastiota · 04MTS-115 · 115 kV · Hermosillo, Sonora · NOROESTE" in ms.options


def test_filtros_acotan_y_conservan_lo_elegido(app):
    app.run()
    app.multiselect(key="nodos_sel").set_value(["01VAJ-230"])
    app.multiselect(key="filt_est").set_value(["SONORA"])
    app.run()
    opciones = app.multiselect(key="nodos_sel").options
    assert all("Sonora" in o for o in opciones[:-1]) and "01VAJ-230" in opciones[-1]
    assert app.multiselect(key="nodos_sel").value == ["01VAJ-230"]


def test_importar_lista_agrega_y_marca_desconocidas(app):
    app.run()
    app.text_area(key="import_texto").set_value("04tst-115, 04EMY-115\nXXX-999")
    _boton(app, "➕ Agregar").click()
    app.run()
    assert app.session_state["nodos_sel"] == ["04TST-115", "04EMY-115", "XXX-999"]
    assert app.text_area(key="import_texto").value == ""
    assert {m.label: m.value for m in app.metric}["⚠️ Desconocidos"] == "1"


def test_lista_guardada(app):
    app.run()
    app.selectbox(key="lista_sel").set_value("Tastiota / El Mayo (SAEE)")
    next(b for b in app.button if b.key == "btn_lista_add").click()
    app.run()
    assert app.session_state["nodos_sel"] == ["04TST-115", "04MTS-115", "04EMY-115"]


def test_novedades_del_catalogo(app):
    _, version = cat.parse_catalogo(RAIZ / "data" / "catalogo_nodos.xlsx")
    creado = not CAMBIOS.exists()
    if creado:
        CAMBIOS.write_text(json.dumps({
            "version": version, "version_anterior": "v2000-01-01", "revisado": "2026-11-24T00:00:00",
            "altas": [{"clave": "04TST-115", "nombre": "Tastiota", "zona": "HERMOSILLO", "kv": 115,
                       "estado": "SONORA", "ccr": "NOROESTE", "sistema": "SIN", "municipio": "HERMOSILLO"}],
            "bajas": [], "cambios": [], "control_de_cambios": [],
        }), encoding="utf-8")
    try:
        datos = json.loads(CAMBIOS.read_text(encoding="utf-8"))
        app.run()
        assert any(e.label.startswith(f"🆕 Novedades del catálogo {version}") for e in app.expander)
        if datos["version"] == version and datos["altas"]:
            next(b for b in app.button if b.key == "btn_add_novedades").click()
            app.run()
            assert set(a["clave"] for a in datos["altas"]) <= set(app.session_state["nodos_sel"])
    finally:
        if creado:
            os.remove(CAMBIOS)
