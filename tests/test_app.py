"""App completa con el CENACE, Banxico, DOF, GitHub y OpenStreetMap simulados.

    python -m pytest -q tests

Un nodo con historial guardado (04TST-115) sale sin consultar al CENACE; los demás
se descargan (12 meses y luego todo el historial). Se revisan la ficha, la
comparación, el período, la conversión a USD con el FIX de cada día, los enlaces y el
mapa.
"""
import re
from datetime import date, timedelta
from pathlib import Path
from statistics import mean
from unittest import mock

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

import cenace_falso as cf
import datos_pml as dp

APP = str(Path(__file__).resolve().parents[1] / "app.py")
NODOS = ["04TST-115", "04EMY-115", "06MTP-115"]


def _parches(get=cf.requests_get, post=cf.requests_post):
    return (mock.patch("requests.Session.get", cf.session_get), mock.patch("requests.get", get),
            mock.patch("requests.post", post))


def nueva_app(token=True, query=None):
    st.cache_data.clear()
    st.cache_resource.clear()
    at = AppTest.from_file(APP, default_timeout=600)
    if token:
        at.secrets["BANXICO_TOKEN"] = "token-de-prueba"
    for k, v in (query or {}).items():
        at.query_params[k] = v
    at.run()
    return at


@pytest.fixture()
def app():
    p1, p2, p3 = _parches()
    with p1, p2, p3:
        yield nueva_app()


def boton(at, key):
    return next(b for b in at.button if b.key == key)


def generar(at, nodos):
    at.multiselect(key="nodos_sel").set_value(nodos)
    at.run()
    cf.LLAMADAS["cenace"] = 0
    boton(at, "btn_generar").click()
    at.run()
    sin_errores(at)


def sin_errores(at):
    assert not at.exception, [e.message for e in at.exception]


def textos(at):
    partes = [m.value for m in at.markdown] + [c.value for c in at.caption]
    partes += [i.value for i in at.info] + [w.value for w in at.warning] + [e.value for e in at.error]
    return "\n".join(str(p) for p in partes)


def metrica(at, etiqueta):
    valor = next(m.value for m in at.metric if m.label == etiqueta)
    return float(re.sub(r"[^0-9.\-]", "", str(valor)))


def periodo_mostrado(at):
    cap = next(c.value for c in at.caption if c.value.startswith("MTR ·"))
    d0, d1 = re.search(r"(\d{2}-\d{2}-\d{4}) a (\d{2}-\d{2}-\d{4})", cap).groups()
    conv = lambda s: date(int(s[6:]), int(s[3:5]), int(s[:2]))          # noqa: E731
    return conv(d0), conv(d1)


def tabla_anual(at):
    return next(d.value for d in at.dataframe if "anio" in d.value.columns)


def test_portada_sin_modos_proceso_ni_subida_de_catalogo(app):
    sin_errores(app)
    claves_widgets = {w.key for w in app.radio} | {w.key for w in app.selectbox}
    assert not {"modo_sel", "proceso_sel"} & claves_widgets
    assert not app.get("file_uploader")
    from catalogo import parse_catalogo
    version = parse_catalogo(Path(APP).parent / "data" / "catalogo_nodos.xlsx")[1]
    t = textos(app)
    assert version in t and "Generar informe" in t and "Tastiota (04TST-115)" in t
    assert not re.search(r"kmz|scoring|modo de uso|actualizar catálogo", t, re.I)


def test_ficha_de_nodo_guardado_sale_sin_consultar_al_cenace(app):
    generar(app, ["04TST-115"])
    assert cf.LLAMADAS["cenace"] == 0
    assert periodo_mostrado(app) == (dp.INICIO_MTR, dp.ultimo_dia_mtr())
    esperado = cf.pml_esperado("04TST-115", dp.INICIO_MTR, dp.ultimo_dia_mtr())
    assert metrica(app, "PML promedio (MXN/MWh)") == pytest.approx(esperado, abs=0.6)
    assert [t.label for t in app.tabs][:5] == ["📌 Lecturas", "📈 Historia", "🔌 Congestión y pérdidas",
                                               "🕐 Distribuciones", "☀️🔋 Solar y baterías"]
    t = textos(app)
    for palabra in ("**Nivel.**", "**Red.**", "zona exportadora", "**Solar.**", "**Baterías.**"):
        assert palabra in t, palabra
    anual = tabla_anual(app)
    assert list(anual["anio"]) == [str(a) for a in range(2017, dp.ultimo_dia_mtr().year + 1)]
    assert [d.label for d in app.get("download_button")][0] == "📑 Informe (Excel)"
    assert not any(b.key in ("btn_anio", "btn_todo") for b in app.button)


def test_nodo_sin_historial_baja_todo_el_historial(app):
    generar(app, ["06MTP-115"])
    ultimo = dp.ultimo_dia_mtr()
    assert cf.LLAMADAS["cenace"] == len(dp.bloques(dp.INICIO_MTR, ultimo, ultimo))
    assert periodo_mostrado(app) == (dp.INICIO_MTR, ultimo)
    assert "Historial parcial" not in textos(app)
    esperado = cf.pml_esperado("06MTP-115", dp.INICIO_MTR, ultimo)
    assert metrica(app, "PML promedio (MXN/MWh)") == pytest.approx(esperado, abs=0.6)


def test_con_el_cenace_lento_baja_12_meses_y_luego_completa():
    p1, p2, p3 = _parches()
    with p1, p2, p3, mock.patch.dict("os.environ", {"PML_LIMITE_AUTO": "0"}):
        app = nueva_app()
        generar(app, ["06MTP-115"])
        ultimo = dp.ultimo_dia_mtr()
        assert cf.LLAMADAS["cenace"] == len(dp.bloques(ultimo - timedelta(days=364), ultimo, ultimo))
        d0, d1 = periodo_mostrado(app)
        assert d1 == ultimo and timedelta(days=364) <= d1 - d0 <= timedelta(days=371)
        assert "Historial parcial" in textos(app)
        cf.LLAMADAS["cenace"] = 0
        boton(app, "btn_todo").click()
        app.run()
        sin_errores(app)
        assert periodo_mostrado(app) == (dp.INICIO_MTR, ultimo)
        assert cf.LLAMADAS["cenace"] == len(dp.bloques(dp.INICIO_MTR, d0 - timedelta(days=1), ultimo))
        assert not any(b.key in ("btn_anio", "btn_todo") for b in app.button)


def test_comparacion_periodo_comun_y_ficha_de_cada_nodo(app):
    generar(app, NODOS)
    assert app.radio(key="vista_sel").options == ["📊 Comparación", "Tastiota (04TST-115)", "El Mayo (04EMY-115)",
                                                  "Monterrey Potencia (06MTP-115)"]
    tabla = next(d.value for d in app.dataframe if "solar_factor" in d.value.columns)
    assert list(tabla.index) == NODOS
    assert app.radio(key="periodo_modo").value == "comun"       # historias distintas: se compara a la par
    app.radio(key="periodo_modo").set_value("todo")
    app.run()
    assert "Ojo con los períodos" in textos(app) and "(tramo común)" in textos(app)
    app.radio(key="periodo_modo").set_value("comun")
    app.run()
    sin_errores(app)
    ultimo = dp.ultimo_dia_mtr()
    assert periodo_mostrado(app) == (cf.INICIO_NODO["04EMY-115"], ultimo)
    assert "Ojo con los períodos" not in textos(app)
    app.selectbox(key="ref_sel").set_value("06MTP-115")
    app.run()
    assert "04TST-115 − 06MTP-115" in textos(app)
    app.radio(key="vista_sel").set_value("04EMY-115")
    app.run()
    sin_errores(app)
    esperado = cf.pml_esperado("04EMY-115", cf.INICIO_NODO["04EMY-115"], ultimo)
    assert metrica(app, "PML promedio (MXN/MWh)") == pytest.approx(esperado, abs=0.6)


def _usd_en_periodo(at, d0, d1):
    at.radio(key="moneda_sel").set_value("🇺🇸 USD (dólares)")
    at.radio(key="periodo_modo").set_value("pers")
    at.run()
    at.date_input(key="periodo_rango").set_value((d0, d1))
    at.run()
    sin_errores(at)
    return tabla_anual(at).set_index("anio")["promedio"]


def _esperado_usd(nodo, d0, d1):
    return mean(cf.componentes(nodo, d, h)[0] / cf.fix_aplicable(d)
                for d in (d0 + timedelta(days=i) for i in range((d1 - d0).days + 1)) for h in range(1, 25))


def test_usd_con_el_fix_de_cada_dia_fines_de_semana_y_feriados(app):
    generar(app, ["04TST-115"])
    d0, d1 = date(2025, 12, 26), date(2026, 1, 8)            # fin de semana y feriado del 1-ene
    prom = _usd_en_periodo(app, d0, d1)
    assert prom["2025"] == pytest.approx(_esperado_usd("04TST-115", d0, date(2025, 12, 31)), rel=1e-9)
    assert prom["2026"] == pytest.approx(_esperado_usd("04TST-115", date(2026, 1, 1), d1), rel=1e-9)
    assert "Banxico SIE" in textos(app)


def test_usd_sin_token_usa_el_dof_con_los_mismos_valores():
    p1, p2, p3 = _parches()
    with p1, p2, p3:
        at = nueva_app(token=False)
        generar(at, ["04TST-115"])
        d0, d1 = date(2025, 12, 26), date(2026, 1, 8)
        prom = _usd_en_periodo(at, d0, d1)
        assert prom["2026"] == pytest.approx(_esperado_usd("04TST-115", date(2026, 1, 1), d1), rel=1e-9)
        assert "DOF (SIDOF)" in textos(at)


def test_sin_tipo_de_cambio_avisa_y_queda_en_mxn():
    def solo_github(url, headers=None, timeout=None, **kw):
        if "raw.githubusercontent.com" in url:
            return cf.github_get(url)
        raise ConnectionError("sin red")
    p1, p2, p3 = _parches(get=solo_github)
    with p1, p2, p3:
        at = nueva_app()
        generar(at, ["04TST-115"])
        at.radio(key="moneda_sel").set_value("🇺🇸 USD (dólares)")
        at.run()
        sin_errores(at)
        assert any("No se pudo obtener el tipo de cambio" in e.value for e in at.error)
        assert any(m.label == "PML promedio (MXN/MWh)" for m in at.metric)


def test_enlace_genera_el_informe_con_moneda_y_periodo():
    p1, p2, p3 = _parches()
    with p1, p2, p3:
        anio = dp.ultimo_dia_mtr().year - 1
        at = nueva_app(query={"nodos": "04TST-115,06MTP-115", "moneda": "USD", "periodo": str(anio)})
        sin_errores(at)
        assert at.session_state["nodos_sel"] == ["04TST-115", "06MTP-115"]
        assert periodo_mostrado(at) == (date(anio, 1, 1), date(anio, 12, 31))
        assert any(m.label.endswith("(USD/MWh)") for m in at.metric) or "USD/MWh" in textos(at)
        assert any(f"periodo={anio}" in c.value and "moneda=USD" in c.value for c in at.code)


def test_enlace_de_la_version_anterior_respeta_las_fechas():
    p1, p2, p3 = _parches()
    with p1, p2, p3:
        at = nueva_app(query={"nodos": "04TST-115", "proceso": "MDA", "modo": "datos",
                              "desde": "2026-01-05", "hasta": "2026-01-11"})
        sin_errores(at)
        assert periodo_mostrado(at) == (date(2026, 1, 5), date(2026, 1, 11))
        assert at.query_params["periodo"] in (["2026-01-05_2026-01-11"], "2026-01-05_2026-01-11")


def test_reinicio_de_la_app_recupera_el_historial_guardado(app):
    generar(app, ["04TST-115", "06MTP-115"])
    st.cache_resource.clear()                                  # la app se durmió: memoria vacía
    app.run()
    assert "ya no están en memoria" in textos(app)              # 06MTP-115 no estaba guardado
    app.multiselect(key="nodos_sel").set_value(["04TST-115"])
    boton(app, "btn_generar").click()
    app.run()
    st.cache_resource.clear()
    app.run()
    sin_errores(app)
    assert "ya no están en memoria" not in textos(app) and any(m.label.startswith("PML promedio") for m in app.metric)


def test_seleccion_cambiada_y_limpiar_todo(app):
    generar(app, ["04TST-115"])
    app.multiselect(key="nodos_sel").set_value(["04TST-115", "04MTS-115"])
    app.run()
    assert "La selección cambió" in textos(app)
    boton(app, "btn_clear_all").click()
    app.run()
    sin_errores(app)
    assert app.session_state["nodos_sel"] == [] and "informe" not in app.session_state


def test_fallo_de_openstreetmap_no_queda_en_cache(app):
    def overpass_caido(url, data=None, headers=None, timeout=None, **kw):
        raise ConnectionError("Overpass caído")
    generar(app, ["04TST-115", "01AAN-85"])
    with mock.patch("requests.post", overpass_caido):
        boton(app, "btn_ubicar").click()
        app.run()
    assert app.session_state["ubicaciones"] is None and "OpenStreetMap no respondió" in textos(app)
    boton(app, "btn_ubicar").click()
    app.run()
    sin_errores(app)
    ub = app.session_state["ubicaciones"]["df"].set_index("clave")
    assert ub.loc["04TST-115", "lat"] == pytest.approx(28.37) and ub.loc["01AAN-85", "lat"] == pytest.approx(20.27)


def test_semanas_sin_respuesta_quedan_como_huecos_y_se_completan():
    semana = "/2019/03/"

    def con_fallas(self, url, timeout=None, **kw):
        if semana in url.split("/JSON")[0][-22:]:            # una semana de marzo de 2019 no responde
            return cf.Resp(503, text="Service Unavailable")
        return cf.session_get(self, url, timeout=timeout)

    p1, p2, p3 = _parches()
    with p1, p2, p3, mock.patch("datos_pml.time.sleep", lambda s: None):
        app = nueva_app()
        with mock.patch("requests.Session.get", con_fallas):
            generar(app, ["06MTP-115"])
        assert "días sin respuesta del CENACE" in textos(app)
        assert any("consultas al CENACE con error" in e.label for e in app.expander)
        assert periodo_mostrado(app)[0] == dp.INICIO_MTR
        cf.LLAMADAS["cenace"] = 0
        boton(app, "btn_todo").click()
        app.run()
        sin_errores(app)
        assert 1 <= cf.LLAMADAS["cenace"] <= 6 and "Historial parcial" not in textos(app)
        assert not any("consultas al CENACE con error" in e.label for e in app.expander)
        assert app.session_state["informe"]["errores"] == []
