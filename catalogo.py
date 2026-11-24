"""Catálogo de NodosP del CENACE.

Lectura del xlsx oficial, etiquetas para el buscador, diferencias entre
versiones y descubrimiento de la última versión publicada. No depende de
Streamlit: lo usan app.py y scripts/actualizar_catalogo.py.
"""
import re
import unicodedata
from urllib.parse import quote, unquote, urljoin

URL_PAGINA_CATALOGO = "https://www.cenace.gob.mx/Paginas/SIM/NodosP.aspx"
CAMPOS = ("sistema", "ccr", "zona", "nombre", "kv", "estado", "municipio")


def _texto(v):
    """Texto limpio: sin espacios duros ni dobles."""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip() if v is not None else ""


def normalizar(s):
    """Minúsculas y sin acentos, para comparar textos."""
    s = unicodedata.normalize("NFKD", _texto(s))
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def version_de_hoja(nombre_hoja):
    """'Catalogo de NodosP V20260716' → 'v2026-07-16' ('' si no se reconoce)."""
    m = re.search(r"V(\d{4})(\d{2})(\d{2})", nombre_hoja or "", flags=re.IGNORECASE)
    return f"v{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""


def parse_catalogo(path_or_buffer):
    """Catálogo oficial de NodosP (xlsx) → ({clave: {campos}}, version)."""
    import openpyxl
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True, read_only=True)
    hoja = wb.sheetnames[0]
    ws = wb[hoja]
    catalogo = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 6 or not r[3]:
            continue
        clave = _texto(r[3]).upper()
        try:
            kv = int(float(r[5])) if r[5] not in (None, "") else 0
        except (TypeError, ValueError):
            kv = 0
        catalogo[clave] = {
            "sistema":   _texto(r[0]),
            "ccr":       _texto(r[1]),
            "zona":      _texto(r[2]),
            "nombre":    _texto(r[4]),
            "kv":        kv,
            "estado":    _texto(r[15]) if len(r) > 15 else "",
            "municipio": _texto(r[17]) if len(r) > 17 else "",
        }
    wb.close()
    return catalogo, version_de_hoja(hoja)


def leer_control_de_cambios(path_or_buffer):
    """Filas de la hoja «Control de Cambios»: [{numero, descripcion, version, responsable}]."""
    import openpyxl
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True, read_only=True)
    hojas = [h for h in wb.sheetnames if "cambio" in normalizar(h)]
    filas = []
    if hojas:
        en_tabla = False
        for r in wb[hojas[0]].iter_rows(values_only=True):
            celdas = [c for c in r if c is not None]
            if not celdas:
                continue
            if not en_tabla:   # la tabla empieza en la fila de encabezados, no en la de definiciones
                en_tabla = (len(celdas) > 1 and normalizar(celdas[0]).startswith("numero de cambio")
                            and normalizar(celdas[1]).startswith("descripcion del cambio"))
                continue
            if len(celdas) < 3 or not re.fullmatch(r"\d+", _texto(celdas[0])):
                continue
            fecha = celdas[2]
            version = fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else _texto(fecha)[:10]
            filas.append({"numero": _texto(celdas[0]), "descripcion": _texto(celdas[1]),
                          "version": version,
                          "responsable": _texto(celdas[3]) if len(celdas) > 3 else ""})
    wb.close()
    return filas


def diff_catalogos(viejo, nuevo):
    """Altas, bajas y cambios de campo entre dos catálogos {clave: campos}."""
    altas = [{"clave": c, **nuevo[c]} for c in sorted(set(nuevo) - set(viejo))]
    bajas = [{"clave": c, **viejo[c]} for c in sorted(set(viejo) - set(nuevo))]
    cambios = []
    for c in sorted(set(viejo) & set(nuevo)):
        for campo in CAMPOS:
            antes, despues = viejo[c].get(campo), nuevo[c].get(campo)
            if antes != despues:
                cambios.append({"clave": c, "campo": campo, "antes": antes, "despues": despues})
    return {"altas": altas, "bajas": bajas, "cambios": cambios}


def versiones_publicadas(html, base=URL_PAGINA_CATALOGO):
    """Enlaces de la página NodosP.aspx → [(version, url_absoluta)], la más nueva primero."""
    vistos = {}
    for href in re.findall(r"""href\s*=\s*["']([^"']*CatalogoNodosP[^"']*\.xlsx)["']""", html, flags=re.I):
        ruta = unquote(href)
        m = re.search(r"\(v(\d{4}-\d{2}-\d{2})\)\.xlsx$", ruta, flags=re.I)
        if not m:
            continue
        absoluta = urljoin(base, ruta)
        esquema, resto = absoluta.split("://", 1)
        dominio, _, camino = resto.partition("/")
        vistos[f"v{m.group(1)}"] = f"{esquema}://{dominio}/{quote(camino, safe='/()')}"
    return sorted(vistos.items(), reverse=True)


def etiqueta_nodo(clave, info):
    """Texto de cada opción del buscador: nombre · clave · kV · zona, estado · CCR."""
    if not info:
        return f"{clave} · ⚠️ fuera del catálogo"
    partes = [info.get("nombre") or clave, clave]
    if info.get("kv"):
        partes.append(f"{info['kv']} kV")
    lugar = ", ".join(x.title() for x in (info.get("zona"), info.get("estado")) if x)
    if lugar:
        partes.append(lugar)
    if info.get("ccr"):
        partes.append(info["ccr"])
    return " · ".join(partes)


def parse_lista_claves(texto):
    """Claves pegadas (saltos de línea, comas, punto y coma, tabs o espacios) → lista sin duplicados."""
    vistas, claves = set(), []
    for c in re.split(r"[\s,;]+", texto or ""):
        c = c.strip().upper()
        if c and c not in vistas:
            vistas.add(c)
            claves.append(c)
    return claves
