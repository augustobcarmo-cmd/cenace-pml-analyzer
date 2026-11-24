"""Descarga, cobertura, almacén e informe (sin Streamlit), con el CENACE simulado."""
import json
import os

import sys
from datetime import date, timedelta
from pathlib import Path
from unittest import mock


import pandas as pd
import pytest

import cenace_falso as cf
import datos_pml as dp
import informe

RAIZ = Path(__file__).resolve().parents[1]
D = date


def test_unir_y_restar_intervalos():
    iv = dp.unir([(D(2020, 1, 1), D(2020, 1, 10)), (D(2020, 1, 11), D(2020, 1, 20)), (D(2020, 2, 1), D(2020, 2, 5))])
    assert iv == [(D(2020, 1, 1), D(2020, 1, 20)), (D(2020, 2, 1), D(2020, 2, 5))]
    assert dp.restar(D(2019, 12, 25), D(2020, 2, 10), iv) == [
        (D(2019, 12, 25), D(2019, 12, 31)), (D(2020, 1, 21), D(2020, 1, 31)), (D(2020, 2, 6), D(2020, 2, 10))]
    assert dp.restar(D(2020, 1, 2), D(2020, 1, 5), iv) == []
    assert dp.intervalos_de_json(dp.intervalos_a_json(iv)) == iv


def test_rejilla_de_semanas_y_plan():
    b = dp.bloques(D(2017, 1, 1), D(2017, 2, 20), tope_fin=D(2017, 2, 15))
    assert b[0] == (D(2017, 1, 27), D(2017, 2, 2)) and b[-1] == (D(2017, 2, 10), D(2017, 2, 15))
    assert all((y - x).days < 7 for x, y in b)
    claves = [f"04X{i:02d}-115" for i in range(25)] + ["07RIN-161"]
    falt = {c: [(D(2024, 1, 1), D(2024, 1, 31))] for c in claves}
    plan = dp.planificar(falt, {"07RIN-161": "BCA"}, tope_fin=D(2024, 1, 20))
    assert plan[0].fin == D(2024, 1, 20)                      # lo más reciente primero
    assert all(len(c.nodos) <= 20 for c in plan)
    assert {c.sistema for c in plan if "07RIN-161" in c.nodos} == {"BCA"}
    assert all("07RIN-161" not in c.nodos for c in plan if c.sistema == "SIN")
    semanas = {(c.ini, c.fin) for c in plan}
    assert len(plan) == 3 * len(semanas)                      # 25 SIN → 2 lotes + 1 BCA por semana


def _descargar(claves, ini, fin):
    ultimo = dp.ultimo_dia_mtr()
    faltan = {c: [(ini, fin)] for c in claves}
    rec = dp.Recolector(ultimo)
    with mock.patch("requests.Session.get", cf.session_get):
        errores, pend = dp.ejecutar(dp.planificar(faltan, {}, ultimo), rec, workers=4, espera=0)
    assert not errores and not pend
    return rec.vaciar(), ultimo


def test_descarga_cobertura_y_nodos_nuevos():
    ultimo = dp.ultimo_dia_mtr()
    ini = ultimo - timedelta(days=364)
    nuevos, _ = _descargar(["04TST-115", "04EMY-115", "01VMA-230"], ini, ultimo)
    df, cub = nuevos["04TST-115"]
    assert df["fecha"].max().date() == ultimo and df["fecha"].min().date() <= ini
    assert dp.unir(cub)[-1][1] == ultimo and dp.restar(ini, ultimo, cub) == []
    p = df.iloc[0]
    assert (p.pml, p.pml_ene, p.pml_per, p.pml_cng) == pytest.approx(
        cf.componentes("04TST-115", p.fecha.date(), int(p.hora)))
    df_m, cub_m = nuevos["04EMY-115"]                          # nodo nuevo: datos desde su alta
    assert df_m["fecha"].min().date() == max(cf.INICIO_NODO["04EMY-115"], df["fecha"].min().date())
    assert dp.restar(ini, ultimo, cub_m) == []                 # las semanas vacías también quedan consultadas
    df_v, cub_v = nuevos.get("01VMA-230", (dp.normalizar(None), []))
    assert df_v.empty and dp.unir(cub_v)[-1][1] < ultimo      # sin datos recientes: se vuelve a pedir


def test_almacen_en_memoria_fusiona_y_cuenta_versiones():
    alm = dp.Almacen(max_series=2)
    ultimo = dp.ultimo_dia_mtr()
    nuevos, _ = _descargar(["04TST-115"], ultimo - timedelta(days=20), ultimo)
    df, cub = nuevos["04TST-115"]
    s = alm.agregar("04TST-115", df.iloc[: len(df) // 2], cub[:1])
    s = alm.agregar("04TST-115", df, cub)
    assert len(s.datos) == len(df) and s.version == 2 and s.hasta() == ultimo
    alm.agregar("A")
    alm.agregar("B")
    assert alm.get("04TST-115") is None and alm.claves() == ["A", "B"]


def test_script_de_la_action_descarga_todo_y_luego_solo_lo_nuevo(tmp_path, capsys):
    listas = tmp_path / "listas.json"
    listas.write_text(json.dumps({"Prueba": ["04MTS-115", "01VMA-230"]}), encoding="utf-8")
    destino = tmp_path / "datos"
    sys.path.insert(0, str(RAIZ / "scripts"))
    import actualizar_datos
    salida = tmp_path / "salida.txt"
    with mock.patch("requests.Session.get", cf.session_get), \
         mock.patch.dict(os.environ, {"GITHUB_OUTPUT": str(salida)}), \
         mock.patch.object(dp, "INICIO_MTR", D(2022, 6, 3)):     # historia corta para la prueba
        assert actualizar_datos.main(["--destino", str(destino), "--listas", str(listas), "--workers", "4"]) == 0
        indice = json.loads((destino / "indice.json").read_text(encoding="utf-8"))
        m = indice["nodos"]["04MTS-115"]
        assert m["desde"] == cf.INICIO_NODO["04MTS-115"].isoformat() and m["hasta"] == dp.ultimo_dia_mtr().isoformat()
        df = pd.read_parquet(destino / m["archivo"])
        assert len(df) == m["filas"] == 24 * ((dp.ultimo_dia_mtr() - cf.INICIO_NODO["04MTS-115"]).days + 1)
        assert indice["nodos"]["01VMA-230"]["filas"] == 0
        assert "cambios=true" in salida.read_text()
        assert (destino / "README.md").read_text(encoding="utf-8").count("|") > 10
        cf.LLAMADAS["cenace"] = 0
        salida.write_text("")
        assert actualizar_datos.main(["--destino", str(destino), "--listas", str(listas)]) == 0
        assert cf.LLAMADAS["cenace"] <= 2                       # solo la cola del nodo sin datos
        assert len(pd.read_parquet(destino / m["archivo"])) == m["filas"]


def _df_sintetico(nodo, d0, d1):
    filas = [(pd.Timestamp(d), h, *cf.componentes(nodo, d, h))
             for d in (d0 + timedelta(days=i) for i in range((d1 - d0).days + 1)) for h in range(1, 25)]
    return dp.normalizar(pd.DataFrame(filas, columns=dp.COLUMNAS))


def test_ficha_de_un_nodo_de_sonora():
    df = _df_sintetico("04TST-115", D(2023, 1, 1), D(2025, 12, 31))
    f = informe.ficha(df)
    k = f["kpis"]
    assert k["promedio"] == pytest.approx(df["pml"].mean())
    assert k["dif"] == pytest.approx((df["pml"] - df["pml_ene"]).mean())
    assert k["dif"] < 0 and k["pct_cng_neg"] == pytest.approx(7 / 24 / 3 * 100, abs=0.5) and k["pct_cng_pos"] == 0
    assert k["cobertura"] == pytest.approx(100)
    assert list(f["anual"].index) == [2023, 2024, 2025]
    assert len(f["mensual"]) == 36 and f["horario"].shape[0] == 24 and f["mes_hora"].shape == (12, 24)
    cng_h = f["horario"]["congestion"]
    assert set(cng_h[cng_h < -1].index) == set(range(10, 17))
    bh = f["bess_horas"]
    assert bh["carga"].sum() == pytest.approx(400) and bh["descarga"].sum() == pytest.approx(400)
    assert set(bh["carga"].nlargest(4).index) <= set(range(7, 17))
    dia = f["diario"]
    uno = df[df["fecha"] == pd.Timestamp(2024, 5, 10)]["pml"].sort_values()
    fila = dia[dia["fecha"] == pd.Timestamp(2024, 5, 10)].iloc[0]
    assert fila["spread4"] == pytest.approx(uno.tail(4).mean() - uno.head(4).mean())
    assert 50 < k["solar_factor"] < 100
    texto = "\n".join(f["lecturas"])
    for palabra in ("Nivel", "Red", "Congestión", "Solar", "Baterías", "zona exportadora", "10-16 h"):
        assert palabra in texto, palabra


def test_comparacion_entre_nodos():
    d0, d1 = D(2025, 1, 1), D(2025, 12, 31)
    dfs = {"04TST-115": _df_sintetico("04TST-115", d0, d1), "06MTP-115": _df_sintetico("06MTP-115", d0, d1),
           "04EMY-115": _df_sintetico("04EMY-115", D(2025, 10, 1), d1)}
    fichas = {c: informe.ficha(df) for c, df in dfs.items()}
    comp = informe.comparacion(dfs, fichas, {}, "04TST-115")
    t = comp["tabla"]
    assert list(t.index) == list(dfs) and t.at["06MTP-115", "promedio"] == pytest.approx(dfs["06MTP-115"]["pml"].mean())
    dif = comp["dif_ref"].loc["06MTP-115"]
    esperado = (dfs["06MTP-115"].set_index(["fecha", "hora"])["pml"]
                - dfs["04TST-115"].set_index(["fecha", "hora"])["pml"]).mean()
    assert dif["promedio"] == pytest.approx(esperado) and dif["horas"] == 24 * 365
    assert comp["dif_ref"].loc["04EMY-115", "horas"] == 24 * 92
    texto = "\n".join(comp["lecturas"])
    assert "Ojo con los períodos" in texto and "Correlación diaria" in texto


def test_excel_del_informe_y_datos_horarios():
    import openpyxl
    import io
    from excel_informe import csv_horario, excel_horario, excel_informe
    d0, d1 = D(2025, 6, 1), D(2025, 8, 31)
    dfs = {"04TST-115": _df_sintetico("04TST-115", d0, d1), "06MTP-115": _df_sintetico("06MTP-115", d0, d1)}
    fichas = {c: informe.ficha(df) for c, df in dfs.items()}
    comp = informe.comparacion(dfs, fichas, {"04TST-115": "Tastiota"}, "04TST-115")
    meta = {"unidad": "MXN", "periodo": "01-06-2025 a 31-08-2025"}
    wb = openpyxl.load_workbook(io.BytesIO(excel_informe(meta, fichas, comp, {"04TST-115": "Tastiota"})))
    assert wb.sheetnames == ["Portada", "Lecturas", "Resumen", "Anual", "Mensual", "Perfil horario",
                             "Mes x hora", "Comparación"]
    res = wb["Resumen"]
    fila = next(r for r in range(4, res.max_row + 1) if res.cell(r, 1).value == "PML promedio")
    assert res.cell(fila, 2).value == pytest.approx(dfs["04TST-115"]["pml"].mean())
    wb = openpyxl.load_workbook(io.BytesIO(excel_horario(dfs, "MXN", {})))
    ws = wb["04TST-115"]
    assert ws.max_row == 2 + len(dfs["04TST-115"]) and ws.cell(3, 3).value == pytest.approx(dfs["04TST-115"]["pml"].iloc[0])
    csv = pd.read_csv(io.BytesIO(csv_horario(dfs)))
    assert len(csv) == sum(len(x) for x in dfs.values()) and list(csv.columns)[:3] == ["clave", "fecha", "hora"]


def test_con_el_cenace_colgado_la_descarga_respeta_el_limite():
    import threading
    import time
    lento = threading.Event()

    def colgado(self, url, timeout=None, **kw):
        if "/2026/" not in url and not lento.wait(3):          # semanas viejas: no responden
            pass
        return cf.session_get(self, url, timeout=timeout)

    ultimo = dp.ultimo_dia_mtr()
    plan = dp.planificar({"04TST-115": [(D(2024, 1, 1), ultimo)]}, {}, ultimo)
    rec = dp.Recolector(ultimo)
    t0 = time.time()
    with mock.patch("requests.Session.get", colgado):
        errores, pendientes = dp.ejecutar(plan, rec, workers=4, espera=0, limite=time.time() + 1.0, gracia=0.5)
    assert time.time() - t0 < 4                                 # no espera a las colgadas
    assert pendientes and len(pendientes) + rec.consultas + len(errores) == len(plan)
    assert all(c.fin.year == 2026 or c.ini.year == 2026 for c in plan[:rec.consultas])   # lo reciente primero
    lento.set()
