"""Mantiene el historial de PML (MTR) de los nodos de las listas guardadas.

Lee data/listas_nodos.json, compara con lo que ya hay en la copia local de la
rama `datos` (indice.json + MTR/*.parquet), descarga del CENACE solo lo que falta
(la primera vez, todo el historial desde el 27-ene-2017; después, los días nuevos)
y escribe los archivos. La GitHub Action publica la carpeta en la rama `datos`.

    python scripts/actualizar_datos.py --destino datos_rama [--nodos 04TST-115,04EMY-115]

Escribe en $GITHUB_OUTPUT: cambios=true|false, hasta=AAAA-MM-DD, resumen=texto.
Sale con código 1 si ninguna consulta al CENACE funcionó.
"""
import argparse
import json
import os
import sys
import time
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import datos_pml as dp                                   # noqa: E402
from catalogo import parse_catalogo, parse_lista_claves  # noqa: E402


def claves_vigiladas(ruta_listas, extra=""):
    claves = []
    if os.path.exists(ruta_listas):
        with open(ruta_listas, encoding="utf-8") as fh:
            listas = json.load(fh)
        for valores in listas.values():
            if isinstance(valores, list):
                claves += parse_lista_claves(" ".join(map(str, valores)))
    claves += parse_lista_claves(extra or "")
    return list(dict.fromkeys(claves))


def escribir_readme(destino, indice):
    filas = [f"| {c} | {m.get('nombre', '')} | {m.get('sistema', '')} | {m.get('desde') or '—'} | "
             f"{m.get('hasta') or '—'} | {m.get('filas', 0):,} |"
             for c, m in indice["nodos"].items()]
    texto = "\n".join([
        "# Historial de PML (MTR) — rama generada",
        "",
        "La escribe la GitHub Action «Actualizar historial de PML (MTR)» con los nodos de",
        "`data/listas_nodos.json` (rama main). No se edita a mano: cada corrida la reemplaza.",
        "",
        f"Fuente: {indice['fuente']} · actualizado {indice['actualizado']} · "
        f"último día esperado {indice['ultimo_dia_esperado']}.",
        "",
        "Un Parquet por nodo en `MTR/` (fecha, hora, pml, pml_ene, pml_per, pml_cng, en MXN/MWh)",
        "y `indice.json` con la cobertura consultada de cada nodo.",
        "",
        "| Clave | Nombre | Sistema | Desde | Hasta | Horas |",
        "|---|---|---|---|---|---|",
        *filas,
        "",
    ])
    with open(os.path.join(destino, "README.md"), "w", encoding="utf-8") as fh:
        fh.write(texto)


def salida(**kw):
    ruta = os.environ.get("GITHUB_OUTPUT")
    lineas = [f"{k}={v}" for k, v in kw.items()]
    print("\n".join(lineas))
    if ruta:
        with open(ruta, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lineas) + "\n")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--destino", required=True, help="copia local de la rama datos")
    ap.add_argument("--listas", default=os.path.join(RAIZ, "data", "listas_nodos.json"))
    ap.add_argument("--catalogo", default=os.path.join(RAIZ, "data", "catalogo_nodos.xlsx"))
    ap.add_argument("--nodos", default="", help="claves extra separadas por coma")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--max-minutos", type=float, default=240)
    ap.add_argument("--hoy", default="", help="AAAA-MM-DD (pruebas)")
    args = ap.parse_args(argv)

    hoy = date.fromisoformat(args.hoy) if args.hoy else date.today()
    ultimo = dp.ultimo_dia_mtr(hoy)
    claves = claves_vigiladas(args.listas, args.nodos)
    if not claves:
        print("No hay nodos en las listas guardadas.")
        salida(cambios="false", hasta=ultimo.isoformat(), resumen="sin nodos")
        return 0

    catalogo, version = parse_catalogo(args.catalogo) if os.path.exists(args.catalogo) else ({}, "")
    sistemas = {c: (catalogo.get(c, {}).get("sistema") or "SIN").upper() for c in claves}
    for c in claves:
        if c not in catalogo:
            print(f"⚠️ {c} no está en el catálogo {version}: se consulta como SIN")

    os.makedirs(args.destino, exist_ok=True)
    indice, series = dp.leer_almacen_local(args.destino)
    meta_nodos = indice.get("nodos") or {}
    coberturas = {c: dp.intervalos_de_json((meta_nodos.get(c) or {}).get("cobertura")) for c in claves}
    faltantes = {c: dp.restar(dp.INICIO_MTR, ultimo, coberturas[c]) for c in claves}
    faltantes = {c: r for c, r in faltantes.items() if r}
    consultas = dp.planificar(faltantes, sistemas, ultimo)
    print(f"{len(claves)} nodos vigilados · último MTR esperado {ultimo} · "
          f"{len(consultas)} consultas para {len(faltantes)} nodos con faltantes")

    recolector = dp.Recolector(ultimo)
    t0 = time.time()

    def progreso(hechas, total):
        if hechas % 25 == 0 or hechas == total:
            print(f"  {hechas}/{total} consultas · {time.time() - t0:.0f} s", flush=True)

    errores, pendientes = dp.ejecutar(consultas, recolector, workers=args.workers, progreso=progreso,
                                      limite=t0 + args.max_minutos * 60)
    nuevos = recolector.vaciar()

    ahora = dp.ahora_iso()
    cambios = False
    for c in claves:
        meta = dict(meta_nodos.get(c) or {})
        df_nuevo, cubiertos = nuevos.get(c, (None, []))
        viejo = series.get(c)
        df = dp.fusionar(viejo, df_nuevo)
        cobertura = dp.unir(coberturas[c] + cubiertos)
        cambio_datos = (len(df) > 0) if viejo is None else not df.equals(viejo)
        if not cambio_datos and cobertura == coberturas[c] and c in meta_nodos:
            continue
        cambios = True
        info = catalogo.get(c, {})
        meta.update({
            "sistema": sistemas[c], "nombre": info.get("nombre", ""),
            "cobertura": dp.intervalos_a_json(cobertura),
            "filas": int(len(df)),
            "desde": df["fecha"].iloc[0].date().isoformat() if len(df) else None,
            "hasta": df["fecha"].iloc[-1].date().isoformat() if len(df) else None,
            "dias_con_datos": int(df["fecha"].nunique()) if len(df) else 0,
        })
        if cambio_datos or "archivo" not in meta:
            meta["archivo"], meta["bytes"] = dp.escribir_nodo(args.destino, c, df)
            meta["actualizado"] = ahora
        meta_nodos[c] = meta

    errores_txt = [{"sistema": e["consulta"].sistema, "nodos": list(e["consulta"].nodos),
                    "ini": e["consulta"].ini.isoformat(), "fin": e["consulta"].fin.isoformat(),
                    "error": e["error"]} for e in errores]
    if cambios or errores_txt != indice.get("errores_ultima_corrida", []):
        cambios = True
        indice = {
            "version": 1, "proceso": dp.PROCESO, "actualizado": ahora,
            "ultimo_dia_esperado": ultimo.isoformat(), "catalogo": version,
            "fuente": "CENACE, servicio web SW-PML (MTR)",
            "nodos": {c: meta_nodos[c] for c in sorted(meta_nodos)},
            "errores_ultima_corrida": errores_txt[:200],
            "consultas_pendientes": len(pendientes),
        }
        with open(os.path.join(args.destino, dp.INDICE), "w", encoding="utf-8") as fh:
            json.dump(indice, fh, ensure_ascii=False, indent=1)
        escribir_readme(args.destino, indice)

    hasta = max((m.get("hasta") or "" for m in meta_nodos.values()), default="")
    resumen = (f"{len(consultas) - len(pendientes)} consultas en {time.time() - t0:.0f} s, "
               f"{len(errores)} con error, {len(pendientes)} pendientes")
    for c in claves:
        m = meta_nodos.get(c) or {}
        print(f"  {c}: {m.get('filas', 0):,} horas · {m.get('desde')} → {m.get('hasta')}")
    print(resumen)
    salida(cambios=str(cambios).lower(), hasta=hasta or ultimo.isoformat(), resumen=resumen)
    if errores and len(errores) == len(consultas) - len(pendientes):
        print("❌ Ninguna consulta al CENACE funcionó.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
