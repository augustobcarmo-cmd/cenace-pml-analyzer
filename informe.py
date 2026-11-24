"""Informe estándar de un nodo y comparación entre nodos (cálculos, sin Streamlit).

Entrada: DataFrame de datos_pml (fecha, hora, pml, pml_ene, pml_per, pml_cng), ya en
la moneda del informe ($/MWh en MXN o USD).

Convenciones:
- Diferencial nodal = PML − componente de energía = pérdidas + congestión. El
  componente de energía es el mismo para todos los nodos de un sistema en una hora,
  así que el diferencial mide lo propio del punto de la red.
- Hora congestionada: |componente de congestión| mayor al umbral (1 MXN/MWh).
- Perfil solar genérico: seno entre las 6:30 y las 19:00 (hora del CENACE), igual
  todo el año. Es ilustrativo: sirve para comparar nodos y años, no para un
  proyecto concreto.
- Batería 4 h / 1 ciclo diario: descarga en las 4 horas más caras del día y carga
  en las 4 más baratas, sin restricción de orden; eficiencia de ida y vuelta 85 %.
  Es una cota superior del arbitraje, no un ingreso esperado.
"""
from __future__ import annotations

import math

import numpy as np
import pandas as pd

HORAS = list(range(1, 25))
UMBRAL_CNG_MXN = 1.0
HORAS_BESS = 4
EFICIENCIA_BESS = 0.85
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MIN_HORAS_DIA = 20          # días con menos horas no entran en spreads ni márgenes


def peso_solar(hora: int) -> float:
    t = hora - 0.5                           # mitad de la hora (hora 13 = 12:00-13:00)
    return max(0.0, math.sin(math.pi * (t - 6.5) / 12.5))


PESOS_SOLARES = {h: peso_solar(h) for h in range(1, 26)}


# ─── Utilidades ──────────────────────────────────────────────────────────
def _q(s: pd.Series, q: float):
    return float(s.quantile(q)) if len(s) else np.nan


def _pct(mask) -> float:
    return float(mask.mean() * 100) if len(mask) else np.nan


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["anio"] = d["fecha"].dt.year.astype("int16")
    d["mes"] = d["fecha"].dt.month.astype("int8")
    d["dif"] = d["pml"] - d["pml_ene"]
    d["w_solar"] = d["hora"].map(PESOS_SOLARES).fillna(0.0)
    return d


def _captura(d: pd.DataFrame) -> float:
    w = d["w_solar"]
    return float((d["pml"] * w).sum() / w.sum()) if w.sum() > 0 else np.nan


def diario(d: pd.DataFrame) -> pd.DataFrame:
    """Una fila por día: promedio, extremos, spreads y margen teórico de una batería 4 h."""
    if d.empty:
        return pd.DataFrame(columns=["fecha", "promedio", "maximo", "minimo", "horas",
                                     "spread1", "spread4", "margen"])
    g = d.groupby("fecha")["pml"]
    out = pd.DataFrame({"promedio": g.mean(), "maximo": g.max(), "minimo": g.min(), "horas": g.size()})
    orden = d[["fecha", "pml"]].sort_values(["fecha", "pml"], kind="mergesort")
    go = orden.groupby("fecha")["pml"]
    baratas = go.head(HORAS_BESS).groupby(orden["fecha"]).agg(["mean", "sum"])
    caras = go.tail(HORAS_BESS).groupby(orden["fecha"]).agg(["mean", "sum"])
    out["spread1"] = out["maximo"] - out["minimo"]
    out["spread4"] = caras["mean"] - baratas["mean"]
    out["margen"] = (caras["sum"] - baratas["sum"] / EFICIENCIA_BESS).clip(lower=0)
    incompletos = out["horas"] < MIN_HORAS_DIA
    out.loc[incompletos, ["spread1", "spread4", "margen"]] = np.nan
    return out.reset_index()


def horas_bess(d: pd.DataFrame) -> pd.DataFrame:
    """% de los días en que cada hora está entre las 4 más baratas (carga) o más caras (descarga)."""
    if d.empty:
        return pd.DataFrame(index=pd.Index(HORAS, name="hora"), columns=["carga", "descarga"], dtype=float)
    x = d[["fecha", "hora", "pml"]].copy()
    x = x[x.groupby("fecha")["pml"].transform("size") >= MIN_HORAS_DIA]
    x["asc"] = x.groupby("fecha")["pml"].rank(method="first")
    x["desc"] = x.groupby("fecha")["pml"].rank(method="first", ascending=False)
    n_dias = x["fecha"].nunique() or 1
    out = pd.DataFrame({
        "carga": x[x["asc"] <= HORAS_BESS].groupby("hora").size() / n_dias * 100,
        "descarga": x[x["desc"] <= HORAS_BESS].groupby("hora").size() / n_dias * 100,
    }).reindex(HORAS).fillna(0.0)
    out.index.name = "hora"
    return out


def _estadisticas(d: pd.DataFrame, dia: pd.DataFrame, umbral: float) -> dict:
    """Estadísticas de un bloque de horas (todo el período, un año o un mes)."""
    p = d["pml"]
    cng = d["pml_cng"]
    congest = cng.abs() > umbral
    prom = float(p.mean()) if len(p) else np.nan
    captura = _captura(d)
    return {
        "horas": int(len(d)),
        "promedio": prom,
        "mediana": _q(p, 0.5), "p10": _q(p, 0.10), "p90": _q(p, 0.90),
        "minimo": float(p.min()) if len(p) else np.nan,
        "maximo": float(p.max()) if len(p) else np.nan,
        "desv": float(p.std()) if len(p) > 1 else np.nan,
        "pct_neg": _pct(p < 0),
        "energia": float(d["pml_ene"].mean()) if len(d) else np.nan,
        "perdidas": float(d["pml_per"].mean()) if len(d) else np.nan,
        "congestion": float(cng.mean()) if len(d) else np.nan,
        "dif": float(d["dif"].mean()) if len(d) else np.nan,
        "pct_cng": _pct(congest),
        "pct_cng_neg": _pct(cng < -umbral),
        "pct_cng_pos": _pct(cng > umbral),
        "cng_cuando": float(cng[congest].mean()) if congest.any() else 0.0,
        "solar_captura": captura,
        "solar_factor": captura / prom * 100 if prom and not np.isnan(captura) and prom > 0 else np.nan,
        "spread1": float(dia["spread1"].mean()) if len(dia) else np.nan,
        "spread4": float(dia["spread4"].mean()) if len(dia) else np.nan,
        "spread4_p90": _q(dia["spread4"].dropna(), 0.90) if len(dia) else np.nan,
        "margen_bess": float(dia["margen"].mean() * 365) if len(dia) and dia["margen"].notna().any() else np.nan,
    }


# ─── Ficha de un nodo ────────────────────────────────────────────────────
def ficha(df: pd.DataFrame, umbral: float = UMBRAL_CNG_MXN, unidad: str = "MXN") -> dict:
    """Todo lo que muestra el informe de un nodo para el período de df."""
    if df is None or df.empty:
        return {"vacio": True, "kpis": {"horas": 0}, "lecturas": []}
    d = preparar(df)
    dia = diario(d)
    desde, hasta = d["fecha"].iloc[0], d["fecha"].iloc[-1]
    dias_periodo = (hasta - desde).days + 1

    k = _estadisticas(d, dia, umbral)
    imax, imin = d["pml"].idxmax(), d["pml"].idxmin()
    k.update({
        "desde": desde.date(), "hasta": hasta.date(),
        "dias": int(d["fecha"].nunique()), "dias_periodo": int(dias_periodo),
        "cobertura": len(d) / (24 * dias_periodo) * 100,
        "max_fecha": d.at[imax, "fecha"].date(), "max_hora": int(d.at[imax, "hora"]),
        "min_fecha": d.at[imin, "fecha"].date(), "min_hora": int(d.at[imin, "hora"]),
    })
    # Últimos 12 meses frente a los 12 previos (si hay datos suficientes)
    corte = hasta - pd.Timedelta(days=365)
    ult = d[d["fecha"] > corte]
    prev = d[(d["fecha"] <= corte) & (d["fecha"] > corte - pd.Timedelta(days=365))]
    k["prom_12m"] = float(ult["pml"].mean()) if len(ult) else np.nan
    k["prom_12m_prev"] = float(prev["pml"].mean()) if len(prev) > 24 * 300 else np.nan
    k["pct_cng_12m"] = _pct(ult["pml_cng"].abs() > umbral) if len(ult) else np.nan
    k["pct_cng_12m_prev"] = _pct(prev["pml_cng"].abs() > umbral) if len(prev) > 24 * 300 else np.nan

    # Por año
    anual = []
    for anio, g in d.groupby("anio"):
        gd = dia[dia["fecha"].dt.year == anio]
        fila = {"anio": int(anio), **_estadisticas(g, gd, umbral)}
        ini_a = max(pd.Timestamp(int(anio), 1, 1), desde)
        fin_a = min(pd.Timestamp(int(anio), 12, 31), hasta)
        fila["cobertura"] = len(g) / (24 * ((fin_a - ini_a).days + 1)) * 100
        anual.append(fila)
    anual = pd.DataFrame(anual).set_index("anio")

    # Por mes calendario de cada año
    d["periodo"] = d["fecha"].dt.to_period("M").dt.to_timestamp()
    gm = d.groupby("periodo")
    dia["periodo"] = dia["fecha"].dt.to_period("M").dt.to_timestamp()
    mensual = pd.DataFrame({
        "promedio": gm["pml"].mean(),
        "p10": gm["pml"].quantile(0.10), "p90": gm["pml"].quantile(0.90),
        "energia": gm["pml_ene"].mean(), "perdidas": gm["pml_per"].mean(),
        "congestion": gm["pml_cng"].mean(), "dif": gm["dif"].mean(),
        "pct_cng": gm["pml_cng"].apply(lambda s: _pct(s.abs() > umbral)),
        "horas": gm.size(),
        "spread4": dia.groupby("periodo")["spread4"].mean(),
    })
    mensual.index.name = "mes"

    # Perfil intradiario (horas 1-24; la hora 25 de los cambios de horario queda fuera)
    d24 = d[d["hora"] <= 24]
    gh = d24.groupby("hora")
    horario = pd.DataFrame({
        "promedio": gh["pml"].mean(),
        "p10": gh["pml"].quantile(0.10), "p25": gh["pml"].quantile(0.25),
        "mediana": gh["pml"].median(),
        "p75": gh["pml"].quantile(0.75), "p90": gh["pml"].quantile(0.90),
        "energia": gh["pml_ene"].mean(), "perdidas": gh["pml_per"].mean(),
        "congestion": gh["pml_cng"].mean(), "dif": gh["dif"].mean(),
        "pct_cng": gh["pml_cng"].apply(lambda s: _pct(s.abs() > umbral)),
        "pct_neg": gh["pml"].apply(lambda s: _pct(s < 0)),
    }).reindex(HORAS)
    horario.index.name = "hora"
    horario_anual = d24.pivot_table(index="anio", columns="hora", values="pml", aggfunc="mean")

    mes_hora = d24.pivot_table(index="mes", columns="hora", values="pml", aggfunc="mean")
    mes_hora_cng = d24.pivot_table(index="mes", columns="hora", values="pml_cng", aggfunc="mean")
    for m in (mes_hora, mes_hora_cng):
        m.index = [MESES[i - 1] for i in m.index]

    ge = d.groupby("mes")
    estacional = pd.DataFrame({
        "promedio": ge["pml"].mean(), "p10": ge["pml"].quantile(0.10), "p90": ge["pml"].quantile(0.90),
        "energia": ge["pml_ene"].mean(), "dif": ge["dif"].mean(),
        "pct_cng": ge["pml_cng"].apply(lambda s: _pct(s.abs() > umbral)),
    })
    estacional.index = [MESES[i - 1] for i in estacional.index]

    cuantiles = np.linspace(0, 1, 201)
    duracion = pd.DataFrame({"pct_horas": (1 - cuantiles) * 100,
                             "pml": np.quantile(d["pml"].to_numpy(), cuantiles)})

    f = {
        "vacio": False, "kpis": k, "anual": anual, "mensual": mensual, "horario": horario,
        "horario_anual": horario_anual, "mes_hora": mes_hora, "mes_hora_cng": mes_hora_cng,
        "estacional": estacional, "duracion": duracion,
        "diario": dia.drop(columns=["periodo"]), "bess_horas": horas_bess(d24),
    }
    f["lecturas"] = lecturas(f, unidad)
    return f


# ─── Lecturas automáticas ────────────────────────────────────────────────
def _n(x, dec=0) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{x:,.{dec}f}"


def _rango_horas(horas) -> str:
    """[11, 12, 13, 15] → «11-13 y 15 h»."""
    horas = sorted(int(h) for h in horas)
    if not horas:
        return ""
    tramos, ini, prev = [], horas[0], horas[0]
    for h in horas[1:] + [None]:
        if h is not None and h == prev + 1:
            prev = h
            continue
        tramos.append(f"{ini}" if ini == prev else f"{ini}-{prev}")
        if h is not None:
            ini = prev = h
    txt = tramos[0] if len(tramos) == 1 else ", ".join(tramos[:-1]) + " y " + tramos[-1]
    return txt + " h"


def lecturas(f: dict, unidad: str = "MXN") -> list[str]:
    """Conclusiones en texto, con reglas fijas y las cifras del período."""
    k = f["kpis"]
    u = f"{unidad}/MWh"
    out = []
    prom = k["promedio"]

    out.append(f"**Nivel.** PML promedio de **{_n(prom)} {u}** entre {k['desde']:%d-%m-%Y} y "
               f"{k['hasta']:%d-%m-%Y} (mediana {_n(k['mediana'])}; 80 % de las horas entre "
               f"{_n(k['p10'])} y {_n(k['p90'])}).")

    if not np.isnan(k.get("prom_12m_prev", np.nan)) and k["prom_12m_prev"]:
        var = (k["prom_12m"] / k["prom_12m_prev"] - 1) * 100
        sentido = "subió" if var > 0 else "bajó"
        out.append(f"**Tendencia.** En los últimos 12 meses el promedio {sentido} {_n(abs(var), 1)} % "
                   f"({_n(k['prom_12m'])} frente a {_n(k['prom_12m_prev'])} {u} en los 12 previos).")

    dif, per, cng = k["dif"], k["perdidas"], k["congestion"]
    if not np.isnan(dif):
        if abs(dif) < 0.02 * max(abs(k["energia"]), 1):
            out.append(f"**Red.** El nodo cotiza prácticamente al componente de energía del sistema "
                       f"(diferencial nodal {_n(dif, 1)} {u}).")
        else:
            lado = "por debajo" if dif < 0 else "por encima"
            lectura = ("zona exportadora: sobra generación local y sacarla cuesta (pérdidas o congestión)"
                       if dif < 0 else
                       "zona importadora: la energía llega de lejos o con restricciones de entrada")
            pct = dif / k["energia"] * 100 if k["energia"] else np.nan
            out.append(f"**Red.** El nodo cotiza **{_n(abs(dif), 1)} {u} {lado}** del componente de energía "
                       f"del sistema ({_n(pct, 1)} %; pérdidas {_n(per, 1)}, congestión {_n(cng, 1)}): "
                       f"señal de {lectura}.")

    pct_cng = k["pct_cng"]
    if not np.isnan(pct_cng):
        if pct_cng < 1:
            out.append(f"**Congestión.** Casi no hay congestión ({_n(pct_cng, 1)} % de las horas).")
        else:
            frec = "frecuente" if pct_cng >= 20 else "ocasional" if pct_cng >= 5 else "esporádica"
            signo = ("sobre todo negativa (la red no puede sacar toda la energía del nodo)"
                     if k["pct_cng_neg"] > 2 * k["pct_cng_pos"] else
                     "sobre todo positiva (la red no puede traer toda la energía que el nodo pide)"
                     if k["pct_cng_pos"] > 2 * k["pct_cng_neg"] else "en ambos sentidos")
            texto = (f"**Congestión {frec}:** {_n(pct_cng, 1)} % de las horas, {signo}; cuando aparece "
                     f"mueve el precio {_n(k['cng_cuando'], 1)} {u} en promedio.")
            h = f["horario"]["congestion"]
            fuertes = h[h.abs() >= max(abs(h).max() * 0.5, umbral_horas(prom))].index
            if len(fuertes) and len(fuertes) < 24:
                texto += f" Se concentra en las {_rango_horas(fuertes)}."
            a, b = k.get("pct_cng_12m", np.nan), k.get("pct_cng_12m_prev", np.nan)
            if not np.isnan(a) and not np.isnan(b) and abs(a - b) >= 5:
                texto += (f" Va {'en aumento' if a > b else 'a la baja'}: {_n(a, 0)} % de las horas en los "
                          f"últimos 12 meses frente a {_n(b, 0)} % antes.")
            out.append(texto)

    if k["pct_neg"] >= 0.5:
        h = f["horario"]["pct_neg"]
        horas_neg = h[h >= max(h.max() * 0.5, 1)].index
        out.append(f"**Precios negativos** en {_n(k['pct_neg'], 1)} % de las horas"
                   + (f", sobre todo a las {_rango_horas(horas_neg)}" if len(horas_neg) else "")
                   + ": riesgo de vertimiento o de vender a pérdida en esas horas.")

    h = f["horario"]["promedio"].dropna()
    if len(h) >= 20:
        baratas = h.nsmallest(4).index
        caras = h.nlargest(4).index
        out.append(f"**Perfil del día.** Horas más baratas en promedio: {_rango_horas(baratas)}; más caras: "
                   f"{_rango_horas(caras)} (diferencia {_n(h.max() - h.min())} {u}).")
        mediodia = h.loc[[x for x in range(11, 17) if x in h.index]].mean()
        if prom > 0 and mediodia < 0.9 * prom:
            out.append(f"**Curva de pato.** A mediodía (11-16 h) el precio queda {_n((1 - mediodia / prom) * 100)} % "
                       f"por debajo del promedio del día: la generación solar ya deprime esas horas.")

    e = f["estacional"]["promedio"].dropna()
    if len(e) >= 6:
        out.append(f"**Estacionalidad.** Meses más caros: {', '.join(e.nlargest(2).index)}; "
                   f"más baratos: {', '.join(e.nsmallest(2).index)}.")

    if not np.isnan(k["solar_factor"]):
        texto = (f"**Solar.** Un perfil solar genérico captura **{_n(k['solar_captura'])} {u}**, "
                 f"{_n(k['solar_factor'])} % del promedio (factor de captura).")
        an = f["anual"]
        an = an[an["horas"] >= 24 * 200]["solar_factor"].dropna()
        if len(an) >= 3 and abs(an.iloc[-1] - an.iloc[0]) >= 5:
            sentido = "cayó" if an.iloc[-1] < an.iloc[0] else "subió"
            texto += (f" El factor {sentido} de {_n(an.iloc[0])} % en {an.index[0]} a {_n(an.iloc[-1])} % "
                      f"en {an.index[-1]}" + (": canibalización solar." if sentido == "cayó" else "."))
        out.append(texto)

    if not np.isnan(k["spread4"]):
        bh = f["bess_horas"]
        carga = bh["carga"].nlargest(4).index
        descarga = bh["descarga"].nlargest(4).index
        out.append(f"**Baterías.** Spread diario 4 h de **{_n(k['spread4'])} {u}** en promedio (P90 "
                   f"{_n(k['spread4_p90'])}). Lo típico es cargar a las {_rango_horas(carga)} y descargar a las "
                   f"{_rango_horas(descarga)}. Margen teórico de una batería 4 h / 1 ciclo (η 85 %): "
                   f"{_n(k['margen_bess'] / 1000, 0)} mil {unidad}/MW-año (cota superior).")

    out.append(f"**Volatilidad.** Desviación estándar {_n(k['desv'])} {u}; máximo {_n(k['maximo'])} el "
               f"{k['max_fecha']:%d-%m-%Y} a las {k['max_hora']} h y mínimo {_n(k['minimo'])} el "
               f"{k['min_fecha']:%d-%m-%Y} a las {k['min_hora']} h.")

    if k["cobertura"] < 97:
        out.append(f"**Datos.** Hay {_n(k['cobertura'], 1)} % de las horas del período "
                   f"({k['dias']:,} de {k['dias_periodo']:,} días con datos); el resto no está publicado "
                   f"o el nodo no existía.")
    return out


def umbral_horas(prom: float) -> float:
    """Congestión horaria promedio mínima para decir que «se concentra» en esa hora."""
    return max(abs(prom) * 0.01, 1e-9)


# ─── Comparación entre nodos ─────────────────────────────────────────────
COLUMNAS_TABLA = {
    "nombre": "Nombre", "desde": "Desde", "hasta": "Hasta", "promedio": "Promedio",
    "mediana": "Mediana", "p10": "P10", "p90": "P90", "desv": "Desv. est.",
    "pct_neg": "% horas negativas", "energia": "Energía", "perdidas": "Pérdidas",
    "congestion": "Congestión", "dif": "Diferencial nodal", "pct_cng": "% horas congestión",
    "solar_captura": "Captura solar", "solar_factor": "Factor captura %",
    "spread4": "Spread 4 h", "margen_bess": "Margen BESS /MW-año", "prom_12m": "Prom. 12 meses",
}


def comparacion(dfs: dict, fichas: dict, nombres: dict, ref: str, unidad: str = "MXN") -> dict:
    claves = [c for c in dfs if not fichas[c].get("vacio")]
    if len(claves) < 2:
        return {"vacio": True, "lecturas": []}
    ref = ref if ref in claves else claves[0]
    tabla = pd.DataFrame([{**{c_: fichas[c]["kpis"].get(c_) for c_ in COLUMNAS_TABLA if c_ != "nombre"},
                           "nombre": nombres.get(c, c), "clave": c} for c in claves]).set_index("clave")
    tabla = tabla[list(COLUMNAS_TABLA)]

    mensual = pd.DataFrame({c: fichas[c]["mensual"]["promedio"] for c in claves})
    horario = pd.DataFrame({c: fichas[c]["horario"]["promedio"] for c in claves})
    horario_cng = pd.DataFrame({c: fichas[c]["horario"]["congestion"] for c in claves})
    horario_dif = pd.DataFrame({c: fichas[c]["horario"]["dif"] for c in claves})

    base = dfs[ref].set_index(["fecha", "hora"])["pml"]
    difs, dif_mensual, dif_horario = {}, {}, {}
    for c in claves:
        if c == ref:
            continue
        otro = dfs[c].set_index(["fecha", "hora"])["pml"]
        dd = (otro - base).dropna()
        if dd.empty:
            continue
        x = dd.reset_index()
        x.columns = ["fecha", "hora", "dif"]
        difs[c] = {"horas": int(len(dd)), "promedio": float(dd.mean()), "p10": _q(dd, .1),
                   "p90": _q(dd, .9), "pct_mayor": _pct(dd > 0)}
        dif_mensual[c] = x.groupby(x["fecha"].dt.to_period("M").dt.to_timestamp())["dif"].mean()
        dif_horario[c] = x[x["hora"] <= 24].groupby("hora")["dif"].mean()

    diarios = pd.DataFrame({c: fichas[c]["diario"].set_index("fecha")["promedio"] for c in claves})
    correlacion = diarios.corr(min_periods=30)

    comp = {
        "vacio": False, "ref": ref, "claves": claves, "tabla": tabla, "mensual": mensual,
        "horario": horario, "horario_cng": horario_cng, "horario_dif": horario_dif,
        "dif_ref": pd.DataFrame(difs).T, "dif_mensual": pd.DataFrame(dif_mensual),
        "dif_horario": pd.DataFrame(dif_horario).reindex(HORAS), "correlacion": correlacion,
    }
    comp["lecturas"] = lecturas_comparacion(comp, nombres, unidad)
    return comp


def lecturas_comparacion(comp: dict, nombres: dict, unidad: str = "MXN") -> list[str]:
    t = comp["tabla"]
    u = f"{unidad}/MWh"
    nom = lambda c: f"{nombres.get(c, c)} ({c})"          # noqa: E731
    out = []
    prom = t["promedio"].dropna()
    if len(prom) >= 2:
        alto, bajo = prom.idxmax(), prom.idxmin()
        out.append(f"**Precio.** El más caro en promedio es {nom(alto)} con {_n(prom[alto])} {u}; el más barato, "
                   f"{nom(bajo)} con {_n(prom[bajo])} {u} ({_n((prom[alto] / prom[bajo] - 1) * 100, 1)} % de "
                   f"diferencia en sus períodos).")
    desde = t["desde"].dropna()
    if len(desde) and desde.nunique() > 1:
        out.append("**Ojo con los períodos:** los nodos no tienen la misma historia ("
                   + "; ".join(f"{c} desde {desde[c]:%m-%Y}" for c in desde.index)
                   + "). Los promedios de los nodos más nuevos no cubren los mismos años: "
                     "elige «Período común» para comparar a la par.")
    cng = t["pct_cng"].dropna()
    if len(cng) >= 2 and cng.max() >= 5:
        c = cng.idxmax()
        out.append(f"**Congestión.** {nom(c)} es el más congestionado ({_n(cng[c], 1)} % de las horas; "
                   f"diferencial nodal {_n(t.at[c, 'dif'], 1)} {u}).")
    sf = t["solar_factor"].dropna()
    if len(sf) >= 2:
        out.append(f"**Solar.** Mejor factor de captura: {nom(sf.idxmax())} ({_n(sf.max())} %); peor: "
                   f"{nom(sf.idxmin())} ({_n(sf.min())} %).")
    sp = t["spread4"].dropna()
    if len(sp) >= 2:
        out.append(f"**Baterías.** Mayor spread diario 4 h: {nom(sp.idxmax())} ({_n(sp.max())} {u}); "
                   f"menor: {nom(sp.idxmin())} ({_n(sp.min())} {u}).")
    ref = comp["ref"]
    for c, fila in comp["dif_ref"].iterrows():
        out.append(f"**{c} − {ref}.** En {_n(fila['horas'])} horas comunes, {c} cotiza {_n(fila['promedio'], 1)} {u} "
                   f"frente a {ref} en promedio (80 % de las horas entre {_n(fila['p10'], 1)} y "
                   f"{_n(fila['p90'], 1)}); está por encima en {_n(fila['pct_mayor'])} % de las horas.")
    corr = comp["correlacion"]
    if corr.shape[0] >= 2:
        m = corr.where(~np.eye(len(corr), dtype=bool)).stack().dropna()
        if len(m):
            a, b = m.idxmin()
            out.append(f"**Correlación diaria.** El par menos correlacionado es {a} / {b} "
                       f"(r = {m.min():.2f}): ahí está el mayor riesgo de base entre nodos.")
    return out
