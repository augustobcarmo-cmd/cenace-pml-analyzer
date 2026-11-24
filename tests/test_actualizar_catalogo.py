"""scripts/actualizar_catalogo.py con la página y el archivo del CENACE simulados."""
import importlib.util
import io
import json
import shutil
from pathlib import Path
from unittest import mock

import openpyxl
import pytest

import catalogo as cat

RAIZ = Path(__file__).resolve().parents[1]
NUEVA = "v2026-11-18"
HTML = ('<a href="/Docs/01_MECP/CatalogoNodosP/2026/'
        'Catálogo NodosP Sistema Eléctrico Nacional (v2026-11-18).xlsx">Excel</a>')


def _catalogo_nuevo(quitar_casi_todo=False):
    """El catálogo del repo con hoja renombrada, una alta, una baja y un cambio de kV."""
    wb = openpyxl.load_workbook(RAIZ / "data" / "catalogo_nodos.xlsx")
    ws = wb[wb.sheetnames[0]]
    ws.title = "Catalogo de NodosP V20261118"
    filas = {str(ws.cell(row=r, column=4).value): r for r in range(3, ws.max_row + 1)}
    ws.cell(row=filas["04TST-115"], column=6).value = 230
    ws.append([c.value for c in ws[filas["04EMY-115"]]])
    ws.cell(row=ws.max_row, column=4).value = "04NEW-115"
    ws.cell(row=ws.max_row, column=5).value = "Nodo Nuevo"
    ws.delete_rows(filas["07RIN-161"])
    if quitar_casi_todo:
        ws.delete_rows(10, ws.max_row)
    cc = wb[wb.sheetnames[1]]
    cc.append([999, "Alta del nuevo NodoP 04NEW-115", "2026-11-18", "USO"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture()
def repo(tmp_path):
    (tmp_path / "scripts").mkdir()
    (tmp_path / "data").mkdir()
    shutil.copy(RAIZ / "scripts" / "actualizar_catalogo.py", tmp_path / "scripts")
    shutil.copy(RAIZ / "catalogo.py", tmp_path)
    shutil.copy(RAIZ / "data" / "catalogo_nodos.xlsx", tmp_path / "data")
    spec = importlib.util.spec_from_file_location("actualizar_tmp", tmp_path / "scripts" / "actualizar_catalogo.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return tmp_path, mod


def _get_falso(contenido):
    def get(url, headers=None, timeout=None, **kw):
        r = mock.Mock(status_code=200)
        r.raise_for_status = lambda: None
        if url == cat.URL_PAGINA_CATALOGO:
            r.text = HTML
        else:
            assert url.endswith("(v2026-11-18).xlsx") and "%20" in url
            r.content = contenido
        return r
    return get


def test_actualiza_y_registra_novedades(repo, monkeypatch, tmp_path):
    raiz, mod = repo
    salida = tmp_path / "gh_output"
    monkeypatch.setenv("GITHUB_OUTPUT", str(salida))
    _, version_antes = cat.parse_catalogo(raiz / "data" / "catalogo_nodos.xlsx")
    with mock.patch.object(mod.requests, "get", _get_falso(_catalogo_nuevo())):
        assert mod.main([]) == 0
    datos = json.loads((raiz / "data" / "catalogo_cambios.json").read_text(encoding="utf-8"))
    assert datos["version"] == NUEVA and datos["version_anterior"] == version_antes
    assert [a["clave"] for a in datos["altas"]] == ["04NEW-115"]
    assert [b["clave"] for b in datos["bajas"]] == ["07RIN-161"]
    assert {"clave": "04TST-115", "campo": "kv", "antes": 115, "despues": 230} in datos["cambios"]
    assert datos["control_de_cambios"][-1]["descripcion"] == "Alta del nuevo NodoP 04NEW-115"
    assert cat.parse_catalogo(raiz / "data" / "catalogo_nodos.xlsx")[1] == NUEVA
    assert "actualizado=true" in salida.read_text() and f"version={NUEVA}" in salida.read_text()

    # Segunda corrida: ya está al día
    with mock.patch.object(mod.requests, "get", _get_falso(b"")):
        assert mod.main([]) == 0
    assert salida.read_text().strip().endswith("actualizado=false")


def test_rechaza_un_archivo_que_no_pasa_la_validacion(repo):
    raiz, mod = repo
    antes = (raiz / "data" / "catalogo_nodos.xlsx").read_bytes()
    with mock.patch.object(mod.requests, "get", _get_falso(_catalogo_nuevo(quitar_casi_todo=True))):
        assert mod.main([]) == 1
    assert (raiz / "data" / "catalogo_nodos.xlsx").read_bytes() == antes
    assert not (raiz / "data" / "catalogo_cambios.json").exists()
