"""Archivos del informe: Excel del informe, datos horarios en Excel y en CSV."""
from __future__ import annotations

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from informe import COLUMNAS_TABLA

NAVY, AZUL, GRIS_CLARO, BLANCO = "0E346B", "2777BD", "EBF3FB", "FFFFFF"
_LADO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=_LADO, right=_LADO, top=_LADO, bottom=_LADO)

ETIQUETAS_KPI = [
    ("desde", "Desde", None), ("hasta", "Hasta", None), ("horas", "Horas con datos", "#,##0"),
    ("cobertura", "Cobertura del período (%)", "0.0"), ("promedio", "PML promedio", "#,##0.00"),
    ("mediana", "Mediana", "#,##0.00"), ("p10", "P10", "#,##0.00"), ("p90", "P90", "#,##0.00"),
    ("minimo", "Mínimo", "#,##0.00"), ("maximo", "Máximo", "#,##0.00"), ("desv", "Desviación estándar", "#,##0.00"),
    ("pct_neg", "% horas con PML negativo", "0.0"), ("energia", "Componente de energía (prom.)", "#,##0.00"),
    ("perdidas", "Componente de pérdidas (prom.)", "#,##0.00"), ("congestion", "Componente de congestión (prom.)", "#,##0.00"),
    ("dif", "Diferencial nodal (pérdidas + congestión)", "#,##0.00"), ("pct_cng", "% horas con congestión", "0.0"),
    ("pct_cng_neg", "% horas con congestión negativa", "0.0"), ("pct_cng_pos", "% horas con congestión positiva", "0.0"),
    ("cng_cuando", "Congestión promedio cuando hay", "#,##0.00"), ("prom_12m", "Promedio últimos 12 meses", "#,##0.00"),
    ("prom_12m_prev", "Promedio 12 meses previos", "#,##0.00"), ("solar_captura", "Precio capturado perfil solar", "#,##0.00"),
    ("solar_factor", "Factor de captura solar (%)", "0.0"), ("spread1", "Spread diario máx − mín", "#,##0.00"),
    ("spread4", "Spread diario 4 h", "#,##0.00"), ("spread4_p90", "Spread diario 4 h, P90", "#,##0.00"),
    ("margen_bess", "Margen teórico BESS 4 h (por MW-año)", "#,##0"),
]

COLS_ANUAL = [("horas", "Horas", "#,##0"), ("cobertura", "Cobertura %", "0.0"), ("promedio", "Promedio", "#,##0.00"),
              ("mediana", "Mediana", "#,##0.00"), ("p10", "P10", "#,##0.00"), ("p90", "P90", "#,##0.00"),
              ("minimo", "Mínimo", "#,##0.00"), ("maximo", "Máximo", "#,##0.00"), ("desv", "Desv. est.", "#,##0.00"),
              ("pct_neg", "% horas negativas", "0.0"), ("energia", "Energía", "#,##0.00"),
              ("perdidas", "Pérdidas", "#,##0.00"), ("congestion", "Congestión", "#,##0.00"),
              ("dif", "Diferencial nodal", "#,##0.00"), ("pct_cng", "% horas congestión", "0.0"),
              ("solar_captura", "Captura solar", "#,##0.00"), ("solar_factor", "Factor captura %", "0.0"),
              ("spread4", "Spread 4 h", "#,##0.00"), ("margen_bess", "Margen BESS /MW-año", "#,##0")]

COLS_MENSUAL = [("promedio", "Promedio", "#,##0.00"), ("p10", "P10", "#,##0.00"), ("p90", "P90", "#,##0.00"),
                ("energia", "Energía", "#,##0.00"), ("perdidas", "Pérdidas", "#,##0.00"),
                ("congestion", "Congestión", "#,##0.00"), ("dif", "Diferencial nodal", "#,##0.00"),
                ("pct_cng", "% horas congestión", "0.0"), ("spread4", "Spread 4 h", "#,##0.00"),
                ("horas", "Horas", "#,##0")]

COLS_HORARIO = [("promedio", "Promedio", "#,##0.00"), ("p10", "P10", "#,##0.00"), ("p25", "P25", "#,##0.00"),
                ("mediana", "Mediana", "#,##0.00"), ("p75", "P75", "#,##0.00"), ("p90", "P90", "#,##0.00"),
                ("energia", "Energía", "#,##0.00"), ("perdidas", "Pérdidas", "#,##0.00"),
                ("congestion", "Congestión", "#,##0.00"), ("dif", "Diferencial nodal", "#,##0.00"),
                ("pct_cng", "% horas congestión", "0.0"), ("pct_neg", "% horas negativas", "0.0")]


def _sin_markdown(texto: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"\1", texto)


def _valor(v):
    if v is None:
        return None
    if isinstance(v, (float, np.floating)):
        return None if np.isnan(v) else float(v)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime().date()
    return v


def _titulo(ws, texto, ancho_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ancho_cols)
    c = ws.cell(row=1, column=1, value=texto)
    c.font = Font(bold=True, color=BLANCO, size=13, name="Arial")
    c.fill = PatternFill("solid", start_color=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26


def _encabezados(ws, fila, textos):
    for j, t in enumerate(textos, 1):
        c = ws.cell(row=fila, column=j, value=t)
        c.font = Font(bold=True, color=BLANCO, size=10, name="Arial")
        c.fill = PatternFill("solid", start_color=AZUL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fila].height = 30


def _fila(ws, fila, valores, formatos):
    for j, (v, fmt) in enumerate(zip(valores, formatos), 1):
        c = ws.cell(row=fila, column=j, value=_valor(v))
        c.font = Font(size=10, name="Arial")
        c.border = BORDE
        if fmt:
            c.number_format = fmt
        if fila % 2 == 0:
            c.fill = PatternFill("solid", start_color=GRIS_CLARO)


def _anchos(ws, anchos):
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def excel_informe(meta: dict, fichas: dict, comp: dict | None, nombres: dict) -> bytes:
    """Excel con portada, lecturas, resumen, anual, mensual, perfil horario y mes × hora."""
    u = meta["unidad"]
    claves = [c for c in fichas if not fichas[c].get("vacio")]
    wb = Workbook()

    ws = wb.active
    ws.title = "Portada"
    ws.sheet_view.showGridLines = False
    _titulo(ws, "Informe de PML · Mercado de Tiempo Real (MTR) · CENACE", 4)
    filas = [
        ("Nodos", ", ".join(f"{nombres.get(c, c)} ({c})" for c in claves)),
        ("Período", meta["periodo"]),
        ("Moneda", f"{u}/MWh" + (f" · {meta['fx']}" if meta.get("fx") else "")),
        ("Fuente", "CENACE, servicio web SW-PML (MTR, ex post)"),
        ("Generado", meta.get("generado") or datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Diferencial nodal", "PML − componente de energía = pérdidas + congestión (lo propio del punto de la red)"),
        ("Hora congestionada", f"|componente de congestión| > {meta.get('umbral_txt', '1 MXN/MWh')}"),
        ("Perfil solar", "Genérico: seno entre 6:30 y 19:00 (hora CENACE), igual todo el año. Ilustrativo."),
        ("Batería (BESS)", "4 h / 1 ciclo diario: descarga en las 4 horas más caras y carga en las 4 más "
                           "baratas del día, sin restricción de orden, η 85 %. Cota superior del arbitraje."),
        ("Horas", "Hora 1 = 00:00-01:00. Días de cambio de horario (hasta 2022) con 23 o 25 horas."),
    ]
    for i, (k, v) in enumerate(filas, 3):
        a = ws.cell(row=i, column=1, value=k)
        a.font = Font(bold=True, color=NAVY, name="Arial", size=10)
        a.alignment = Alignment(vertical="top")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        b = ws.cell(row=i, column=2, value=v)
        b.font = Font(name="Arial", size=10)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 30 if len(str(v)) > 70 else 18
    ws.cell(row=len(filas) + 4, column=1, value="Recurrent Energy México · Inteligencia de Mercado").font = \
        Font(italic=True, color=NAVY, name="Arial", size=10)
    _anchos(ws, [22, 40, 40, 40])

    ws = wb.create_sheet("Lecturas")
    _titulo(ws, "Lecturas automáticas", 2)
    r = 3
    bloques = []
    if comp and not comp.get("vacio"):
        bloques.append(("Comparación", comp["lecturas"]))
    bloques += [(f"{nombres.get(c, c)} ({c})", fichas[c]["lecturas"]) for c in claves]
    for titulo, textos in bloques:
        c = ws.cell(row=r, column=1, value=titulo)
        c.font = Font(bold=True, color=NAVY, name="Arial", size=11)
        r += 1
        for t in textos:
            c = ws.cell(row=r, column=1, value="• " + _sin_markdown(t))
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.row_dimensions[r].height = 15 * (1 + len(t) // 120)
            r += 1
        r += 1
    _anchos(ws, [90, 60])

    ws = wb.create_sheet("Resumen")
    _titulo(ws, f"Indicadores del período ({u}/MWh salvo indicación)", 1 + len(claves))
    _encabezados(ws, 3, ["Indicador"] + [f"{nombres.get(c, c)}\n{c}" for c in claves])
    for i, (k, etiqueta, fmt) in enumerate(ETIQUETAS_KPI, 4):
        _fila(ws, i, [etiqueta] + [fichas[c]["kpis"].get(k) for c in claves],
              [None] + [fmt or "yyyy-mm-dd" for _ in claves])
    _anchos(ws, [40] + [20] * len(claves))
    ws.freeze_panes = "B4"

    def tabla_larga(nombre, titulo, extraer, primera, cols, fmt_primera):
        hoja = wb.create_sheet(nombre)
        _titulo(hoja, titulo, 2 + len(cols))
        _encabezados(hoja, 3, ["Nodo", primera] + [e for _, e, _ in cols])
        fila = 4
        for c in claves:
            df = extraer(fichas[c])
            for idx, row in df.iterrows():
                _fila(hoja, fila, [c, idx] + [row.get(k) for k, _, _ in cols],
                      [None, fmt_primera] + [f for _, _, f in cols])
                fila += 1
        _anchos(hoja, [14, 12] + [13] * len(cols))
        hoja.freeze_panes = "C4"

    tabla_larga("Anual", f"Indicadores por año ({u}/MWh)", lambda f: f["anual"], "Año", COLS_ANUAL, "0")
    tabla_larga("Mensual", f"Indicadores por mes ({u}/MWh)", lambda f: f["mensual"], "Mes", COLS_MENSUAL,
                "yyyy-mm")
    cols_h = COLS_HORARIO + [("carga", "% días entre las 4 más baratas", "0.0"),
                             ("descarga", "% días entre las 4 más caras", "0.0")]
    tabla_larga("Perfil horario", f"Perfil intradiario ({u}/MWh)",
                lambda f: f["horario"].join(f["bess_horas"]), "Hora", cols_h, "0")

    ws = wb.create_sheet("Mes x hora")
    _titulo(ws, f"Promedio por mes del año y hora ({u}/MWh): PML y congestión", 26)
    r = 3
    for c in claves:
        for clave_m, nombre_m in (("mes_hora", "PML promedio"), ("mes_hora_cng", "Congestión promedio")):
            m = fichas[c][clave_m]
            ws.cell(row=r, column=1, value=f"{nombres.get(c, c)} ({c}) · {nombre_m}").font = \
                Font(bold=True, color=NAVY, name="Arial", size=10)
            r += 1
            _encabezados(ws, r, ["Mes"] + [str(h) for h in m.columns])
            r += 1
            for mes, fila_m in m.iterrows():
                _fila(ws, r, [mes] + list(fila_m.values), [None] + ["#,##0.0"] * len(m.columns))
                r += 1
            r += 1
    _anchos(ws, [8] + [8] * 24)

    if comp and not comp.get("vacio"):
        ws = wb.create_sheet("Comparación")
        t = comp["tabla"]
        cols = [k for k in COLUMNAS_TABLA if k != "nombre"]
        _titulo(ws, f"Comparación entre nodos ({u}/MWh)", 2 + len(cols))
        _encabezados(ws, 3, ["Clave", "Nombre"] + [COLUMNAS_TABLA[k] for k in cols])
        for i, (c, row) in enumerate(t.iterrows(), 4):
            _fila(ws, i, [c, row["nombre"]] + [row[k] for k in cols],
                  [None, None] + ["yyyy-mm-dd" if k in ("desde", "hasta") else "0.0" if k.startswith(("pct", "solar_f"))
                                  else "#,##0.00" for k in cols])
        r = 5 + len(t)
        if not comp["dif_ref"].empty:
            ws.cell(row=r, column=1, value=f"Diferencial horario frente a {comp['ref']} (nodo − referencia)").font = \
                Font(bold=True, color=NAVY, name="Arial", size=10)
            _encabezados(ws, r + 1, ["Clave", "Horas comunes", "Promedio", "P10", "P90", "% horas por encima"])
            for i, (c, row) in enumerate(comp["dif_ref"].iterrows(), r + 2):
                _fila(ws, i, [c, row["horas"], row["promedio"], row["p10"], row["p90"], row["pct_mayor"]],
                      [None, "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "0.0"])
            r = r + 3 + len(comp["dif_ref"])
        corr = comp["correlacion"]
        ws.cell(row=r, column=1, value="Correlación de los promedios diarios").font = \
            Font(bold=True, color=NAVY, name="Arial", size=10)
        _encabezados(ws, r + 1, [""] + list(corr.columns))
        for i, (c, row) in enumerate(corr.iterrows(), r + 2):
            _fila(ws, i, [c] + list(row.values), [None] + ["0.00"] * len(corr.columns))
        _anchos(ws, [14, 28] + [14] * len(cols))
        ws.freeze_panes = "C4"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_horario(dfs: dict, unidad: str, nombres: dict) -> bytes:
    """Datos horarios: una hoja por nodo (modo de solo escritura, rápido con años de datos)."""
    wb = Workbook(write_only=True)
    con_tc = any("tc" in df.columns for df in dfs.values())
    for c, df in dfs.items():
        ws = wb.create_sheet(title=re.sub(r"[\[\]:*?/\\]", "_", c)[:31])
        ws.freeze_panes = "A3"
        for j, w in enumerate([12, 7, 14, 14, 14, 14, 12], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        cab = WriteOnlyCell(ws, value=f"{nombres.get(c, c)} ({c}) · PML MTR en {unidad}/MWh · Fuente: CENACE SW-PML")
        cab.font = Font(bold=True, color=NAVY, name="Arial")
        ws.append([cab])
        enc = ["Fecha", "Hora", "PML", "Energía", "Pérdidas", "Congestión"] + (["TC FIX (MXN/USD)"] if con_tc else [])
        celdas = []
        for t in enc:
            x = WriteOnlyCell(ws, value=t)
            x.font = Font(bold=True, color=BLANCO, name="Arial")
            x.fill = PatternFill("solid", start_color=AZUL)
            celdas.append(x)
        ws.append(celdas)
        cols = ["pml", "pml_ene", "pml_per", "pml_cng"] + (["tc"] if con_tc else [])
        fechas = df["fecha"].dt.date.to_numpy()
        horas = df["hora"].to_numpy()
        valores = df[cols].round(4).to_numpy()
        for f, h, v in zip(fechas, horas, valores):
            celda = WriteOnlyCell(ws, value=f)
            celda.number_format = "yyyy-mm-dd"
            ws.append([celda, int(h), *[None if np.isnan(x) else float(x) for x in v]])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def csv_horario(dfs: dict) -> bytes:
    partes = []
    for c, df in dfs.items():
        x = df.copy()
        x.insert(0, "clave", c)
        x["fecha"] = x["fecha"].dt.strftime("%Y-%m-%d")
        partes.append(x)
    if not partes:
        return b""
    return pd.concat(partes, ignore_index=True).to_csv(index=False, float_format="%.4f").encode("utf-8")
