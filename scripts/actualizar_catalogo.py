#!/usr/bin/env python3
"""Actualiza el catálogo de NodosP desde la página del CENACE.

Lee https://www.cenace.gob.mx/Paginas/SIM/NodosP.aspx, toma la versión más
nueva y, si es posterior a la de data/catalogo_nodos.xlsx, la descarga, la
valida y escribe:
  - data/catalogo_nodos.xlsx     (el archivo oficial, tal cual)
  - data/catalogo_cambios.json   (altas, bajas y cambios frente a la versión
                                  anterior, más el «Control de Cambios» del CENACE)

Sale con código 0 haya o no versión nueva (el workflow hace commit solo si
cambió algo) y con código 1 si la descarga o la validación fallan.

    python scripts/actualizar_catalogo.py            # revisar y actualizar
    python scripts/actualizar_catalogo.py --revisar  # solo informar
"""
import argparse
import io
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from catalogo import (URL_PAGINA_CATALOGO, diff_catalogos, leer_control_de_cambios,  # noqa: E402
                      parse_catalogo, versiones_publicadas)

RUTA_XLSX = RAIZ / "data" / "catalogo_nodos.xlsx"
RUTA_CAMBIOS = RAIZ / "data" / "catalogo_cambios.json"
CABECERAS = {"User-Agent": "PML-CENACE/1.0 (github.com/augustobcarmo-cmd/cenace-pml-analyzer)"}
MIN_NODOS, MAX_VARIACION = 2000, 0.05


def _get(url, timeout):
    """GET con verificación TLS; si el certificado del CENACE falla, reintenta sin ella."""
    try:
        r = requests.get(url, headers=CABECERAS, timeout=timeout)
    except requests.exceptions.SSLError:
        print("::warning::Certificado TLS del CENACE no verificable; se reintenta sin verificación.")
        requests.packages.urllib3.disable_warnings()
        r = requests.get(url, headers=CABECERAS, timeout=timeout, verify=False)
    r.raise_for_status()
    return r


def _salida_github(**valores):
    ruta = os.environ.get("GITHUB_OUTPUT")
    if ruta:
        with open(ruta, "a", encoding="utf-8") as fh:
            for k, v in valores.items():
                fh.write(f"{k}={v}\n")


def validar(nuevo, version_nueva, version_esperada, actual):
    errores = []
    if version_nueva != version_esperada:
        errores.append(f"la hoja dice {version_nueva or '(sin versión)'} y el enlace {version_esperada}")
    if len(nuevo) < MIN_NODOS:
        errores.append(f"solo {len(nuevo)} nodos (mínimo {MIN_NODOS})")
    if actual and abs(len(nuevo) - len(actual)) > MAX_VARIACION * len(actual):
        errores.append(f"{len(nuevo)} nodos frente a {len(actual)} (cambio > {MAX_VARIACION:.0%})")
    incompletos = sum(1 for v in nuevo.values() if not (v["sistema"] and v["nombre"] and v["ccr"]))
    if incompletos > 0.01 * max(len(nuevo), 1):
        errores.append(f"{incompletos} nodos sin sistema, nombre o CCR")
    return errores


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--revisar", action="store_true", help="solo informar, sin escribir archivos")
    args = ap.parse_args(argv)

    actual, version_actual = parse_catalogo(RUTA_XLSX) if RUTA_XLSX.exists() else ({}, "")
    print(f"Catálogo del repositorio: {version_actual or '—'} ({len(actual)} nodos)")

    versiones = versiones_publicadas(_get(URL_PAGINA_CATALOGO, 60).text)
    if not versiones:
        print("::warning::La página del CENACE no lista catálogos (¿cambió su formato o es inicio de año?).")
        _salida_github(actualizado="false")
        return 0
    version_ultima, url = versiones[0]
    print(f"Última versión publicada: {version_ultima} → {url}")
    if version_actual and version_ultima <= version_actual:
        print("Sin cambios.")
        _salida_github(actualizado="false")
        return 0
    if args.revisar:
        print(f"Hay una versión nueva: {version_ultima}.")
        return 0

    contenido = _get(url, 120).content
    nuevo, version_nueva = parse_catalogo(io.BytesIO(contenido))
    errores = validar(nuevo, version_nueva, version_ultima, actual)
    if errores:
        for e in errores:
            print(f"::error::{e}")
        return 1

    diff = diff_catalogos(actual, nuevo)
    desde = version_actual[1:] if version_actual else ""
    control = [c for c in leer_control_de_cambios(io.BytesIO(contenido)) if c["version"] > desde]
    RUTA_XLSX.write_bytes(contenido)
    RUTA_CAMBIOS.write_text(json.dumps({
        "version": version_nueva,
        "version_anterior": version_actual,
        "revisado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "fuente": url,
        "nodos": len(nuevo),
        **diff,
        "control_de_cambios": control,
    }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Actualizado a {version_nueva}: {len(nuevo)} nodos · {len(diff['altas'])} altas · "
          f"{len(diff['bajas'])} bajas · {len(diff['cambios'])} cambios")
    _salida_github(actualizado="true", version=version_nueva)
    return 0


if __name__ == "__main__":
    sys.exit(main())
