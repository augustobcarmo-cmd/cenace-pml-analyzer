"""Gráficas del informe (Plotly), sin Streamlit.

Reglas: un solo eje Y por gráfica; los nodos conservan su color en todas las
gráficas (paleta categórica validada para daltonismo, en orden fijo); magnitudes con
una sola tonalidad de claro a oscuro; la congestión (con signo) en escala divergente
con gris en cero; líneas de 2 px y leyenda siempre que haya 2 series o más.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.graph_objects as go

# Paleta categórica de referencia (skill dataviz, validada: ΔE CVD adyacente ≥ 9)
SERIE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
OTROS = "#b9b8b3"                 # nodos más allá del 8.º: gris, sin color propio
TEXTO = "#1a1a1a"
TEXTO_2 = "#52514e"
REJILLA = "#e8e7e3"
EJE = "#bdbcb7"
AZUL = SERIE[0]
C_ENERGIA, C_CONGESTION, C_PERDIDAS, C_DIF = SERIE[0], SERIE[1], SERIE[2], TEXTO_2
SECUENCIAL = [[0.0, "#f1f6fc"], [0.25, "#bcd5f2"], [0.5, "#6fa3e2"], [0.75, "#2a6fc4"], [1.0, "#0d3a73"]]
DIVERGENTE = [[0.0, "#1d5fae"], [0.25, "#7fb0e6"], [0.5, "#e7e6e2"], [0.75, "#f09a92"], [1.0, "#c22e2d"]]
HORA_TXT = "Hora (1 = 00:00-01:00)"


def _rgba(hexcolor: str, alpha: float) -> str:
    h = hexcolor.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


def colores_nodos(claves) -> dict:
    """Color por nodo según el orden de selección (el 9.º en adelante, gris)."""
    return {c: SERIE[i] if i < len(SERIE) else OTROS for i, c in enumerate(claves)}


def _estilo(fig: go.Figure, titulo: str, y: str = "", x: str = "", altura: int = 380,
            hover: str = "x unified") -> go.Figure:
    fig.update_layout(
        title=dict(text=titulo, x=0, xanchor="left", font=dict(size=15, color=TEXTO)),
        height=altura, margin=dict(l=8, r=8, t=56, b=8),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial, sans-serif", size=12, color=TEXTO),
        hovermode=hover, barcornerradius=4,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                    font=dict(color=TEXTO_2), title=None),
        hoverlabel=dict(bgcolor="white", font_color=TEXTO, bordercolor=EJE),
    )
    fig.update_xaxes(showgrid=False, linecolor=EJE, ticks="outside", tickcolor=EJE,
                     title=dict(text=x, font=dict(color=TEXTO_2)), tickfont=dict(color=TEXTO_2))
    fig.update_yaxes(gridcolor=REJILLA, zeroline=True, zerolinecolor="#8f8e89", zerolinewidth=1,
                     linecolor=EJE, title=dict(text=y, font=dict(color=TEXTO_2)),
                     tickfont=dict(color=TEXTO_2), separatethousands=True)
    return fig


def _banda(fig, x, bajo, alto, color, alpha, nombre, mostrar=True):
    fig.add_trace(go.Scatter(x=x, y=alto, mode="lines", line=dict(width=0), hoverinfo="skip",
                             showlegend=False, legendgroup=nombre))
    fig.add_trace(go.Scatter(x=x, y=bajo, mode="lines", line=dict(width=0), fill="tonexty",
                             fillcolor=_rgba(color, alpha), name=nombre, hoverinfo="skip",
                             showlegend=mostrar, legendgroup=nombre))


# ─── Ficha de un nodo ────────────────────────────────────────────────────
def historia(mensual: pd.DataFrame, u: str) -> go.Figure:
    fig = go.Figure()
    x = mensual.index
    _banda(fig, x, mensual["p10"], mensual["p90"], AZUL, 0.16, "80 % de las horas (P10-P90)")
    fig.add_trace(go.Scatter(x=x, y=mensual["energia"], name="Componente de energía del sistema",
                             mode="lines", line=dict(color=TEXTO_2, width=2, dash="dot"),
                             hovertemplate="%{y:,.0f}"))
    fig.add_trace(go.Scatter(x=x, y=mensual["promedio"], name="PML promedio del nodo", mode="lines",
                             line=dict(color=AZUL, width=2), hovertemplate="%{y:,.0f}"))
    return _estilo(fig, "PML promedio mensual", u)


def barras_anuales(anual: pd.DataFrame, columna: str, titulo: str, y: str, formato: str = ",.0f",
                   color: str = AZUL) -> go.Figure:
    s = anual[columna].dropna()
    fig = go.Figure(go.Bar(x=s.index.astype(str), y=s.values, marker_color=color,
                           text=[f"{v:{formato}}" for v in s.values], textposition="outside",
                           textfont=dict(color=TEXTO_2), cliponaxis=False,
                           hovertemplate="%{x}: %{y:" + formato + "}<extra></extra>"))
    return _estilo(fig, titulo, y, hover="closest")


def componentes_mensuales(mensual: pd.DataFrame, u: str) -> go.Figure:
    x = mensual.index
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x, y=mensual["congestion"], name="Congestión", marker_color=C_CONGESTION,
                         hovertemplate="%{y:,.1f}"))
    fig.add_trace(go.Bar(x=x, y=mensual["perdidas"], name="Pérdidas", marker_color=C_PERDIDAS,
                         hovertemplate="%{y:,.1f}"))
    fig.add_trace(go.Scatter(x=x, y=mensual["dif"], name="Diferencial nodal (suma)", mode="lines",
                             line=dict(color=C_DIF, width=2), hovertemplate="%{y:,.1f}"))
    fig.update_layout(barmode="relative", bargap=0.15)
    return _estilo(fig, "Diferencial nodal mensual: PML − energía = pérdidas + congestión", u)


def componentes_horarios(horario: pd.DataFrame, u: str) -> go.Figure:
    fig = go.Figure()
    for col, nombre, color in (("congestion", "Congestión", C_CONGESTION), ("perdidas", "Pérdidas", C_PERDIDAS),
                               ("dif", "Diferencial nodal", C_DIF)):
        fig.add_trace(go.Scatter(x=horario.index, y=horario[col], name=nombre, mode="lines+markers",
                                 line=dict(color=color, width=2), marker=dict(size=6),
                                 hovertemplate="%{y:,.1f}"))
    fig.update_xaxes(dtick=1)
    return _estilo(fig, "Pérdidas y congestión por hora del día (promedio)", u, HORA_TXT)


def mapa_calor(m: pd.DataFrame, titulo: str, u: str, divergente: bool = False) -> go.Figure:
    z = m.to_numpy(dtype=float)
    kw = {}
    if divergente:
        lim = float(np.nanmax(np.abs(z))) if np.isfinite(z).any() else 1.0
        kw = dict(zmid=0, zmin=-lim or -1, zmax=lim or 1)
    fig = go.Figure(go.Heatmap(
        z=z, x=[str(c) for c in m.columns], y=list(m.index),
        colorscale=DIVERGENTE if divergente else SECUENCIAL, xgap=1, ygap=1,
        colorbar=dict(title=dict(text=u, font=dict(color=TEXTO_2)), thickness=12,
                      tickfont=dict(color=TEXTO_2)),
        hovertemplate="%{y}, hora %{x}: %{z:,.1f}<extra></extra>", **kw))
    fig.update_yaxes(autorange="reversed", showgrid=False, zeroline=False)
    fig.update_xaxes(dtick=1)
    return _estilo(fig, titulo, "", HORA_TXT, altura=420, hover="closest")


def perfil_horario(horario: pd.DataFrame, u: str) -> go.Figure:
    x = horario.index
    fig = go.Figure()
    _banda(fig, x, horario["p10"], horario["p90"], AZUL, 0.14, "P10-P90")
    _banda(fig, x, horario["p25"], horario["p75"], AZUL, 0.28, "P25-P75")
    fig.add_trace(go.Scatter(x=x, y=horario["mediana"], name="Mediana", mode="lines",
                             line=dict(color=TEXTO_2, width=2, dash="dot"), hovertemplate="%{y:,.0f}"))
    fig.add_trace(go.Scatter(x=x, y=horario["promedio"], name="Promedio", mode="lines+markers",
                             line=dict(color=AZUL, width=2), marker=dict(size=6), hovertemplate="%{y:,.0f}"))
    fig.update_xaxes(dtick=1)
    return _estilo(fig, "Perfil intradiario: distribución del PML por hora", u, HORA_TXT)


def perfil_por_anio(horario_anual: pd.DataFrame, u: str) -> go.Figure:
    fig = go.Figure()
    anios = list(horario_anual.index)
    n = len(anios)
    for i, anio in enumerate(anios):
        t = 0.35 + 0.65 * (i / (n - 1) if n > 1 else 1)          # más viejo, más claro
        color = _mezcla("#d6e5f7", "#0d3a73", t)
        fig.add_trace(go.Scatter(x=horario_anual.columns, y=horario_anual.loc[anio], name=str(anio),
                                 mode="lines", line=dict(color=color, width=2),
                                 hovertemplate=f"{anio}: " + "%{y:,.0f}<extra></extra>"))
    fig.update_xaxes(dtick=1)
    return _estilo(fig, "Perfil intradiario promedio por año", u, HORA_TXT, hover="closest")


def _mezcla(a: str, b: str, t: float) -> str:
    ca = [int(a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(ca, cb))


def estacionalidad(estacional: pd.DataFrame, u: str) -> go.Figure:
    e = estacional
    fig = go.Figure(go.Bar(
        x=list(e.index), y=e["promedio"], marker_color=AZUL, name="Promedio",
        error_y=dict(type="data", symmetric=False, array=e["p90"] - e["promedio"],
                     arrayminus=e["promedio"] - e["p10"], color=TEXTO_2, thickness=1.2, width=4),
        hovertemplate="%{x}: %{y:,.0f}<extra></extra>"))
    return _estilo(fig, "Estacionalidad: promedio por mes del año (barras: P10-P90)", u, hover="closest")


def curva_duracion(duracion: pd.DataFrame, u: str) -> go.Figure:
    fig = go.Figure(go.Scatter(x=duracion["pct_horas"], y=duracion["pml"], mode="lines",
                               line=dict(color=AZUL, width=2), name="PML",
                               hovertemplate="%{x:.1f} % de las horas ≥ %{y:,.0f}<extra></extra>"))
    fig.update_xaxes(ticksuffix=" %")
    return _estilo(fig, "Curva de duración de precios", u, "% de las horas con precio mayor o igual",
                   hover="closest")


def horas_bateria(bess_horas: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=bess_horas.index, y=bess_horas["carga"], name="Entre las 4 más baratas (carga)",
                         marker_color=SERIE[0], hovertemplate="%{y:.0f} % de los días"))
    fig.add_trace(go.Bar(x=bess_horas.index, y=bess_horas["descarga"], name="Entre las 4 más caras (descarga)",
                         marker_color=SERIE[1], hovertemplate="%{y:.0f} % de los días"))
    fig.update_layout(barmode="group", bargap=0.2, bargroupgap=0.08)
    fig.update_xaxes(dtick=1)
    fig.update_yaxes(ticksuffix=" %")
    return _estilo(fig, "¿Cuándo cargar y descargar? Frecuencia de cada hora", "% de los días", HORA_TXT)


def spread_mensual(mensual: pd.DataFrame, u: str) -> go.Figure:
    s = mensual["spread4"]
    fig = go.Figure(go.Scatter(x=s.index, y=s, mode="lines", line=dict(color=SERIE[1], width=2),
                               name="Spread 4 h", hovertemplate="%{y:,.0f}<extra></extra>"))
    return _estilo(fig, "Spread diario 4 h (más caras − más baratas), promedio mensual", u, hover="x")


# ─── Comparación ─────────────────────────────────────────────────────────
def lineas_nodos(df: pd.DataFrame, colores: dict, titulo: str, y: str, x: str = "",
                 etiquetas: dict | None = None, por_hora: bool = False) -> go.Figure:
    fig = go.Figure()
    for c in df.columns:
        fig.add_trace(go.Scatter(
            x=df.index, y=df[c], name=(etiquetas or {}).get(c, c), mode="lines+markers" if por_hora else "lines",
            line=dict(color=colores.get(c, OTROS), width=2 if colores.get(c) != OTROS else 1.2),
            marker=dict(size=6), hovertemplate="%{y:,.1f}"))
    if por_hora:
        fig.update_xaxes(dtick=1)
    return _estilo(fig, titulo, y, x)


def barras_nodos(s: pd.Series, colores: dict, titulo: str, y: str, formato: str = ",.0f") -> go.Figure:
    s = s.dropna()
    fig = go.Figure(go.Bar(x=list(s.index), y=s.values, marker_color=[colores.get(c, OTROS) for c in s.index],
                           text=[f"{v:{formato}}" for v in s.values], textposition="outside",
                           textfont=dict(color=TEXTO_2), cliponaxis=False,
                           hovertemplate="%{x}: %{y:" + formato + "}<extra></extra>"))
    return _estilo(fig, titulo, y, hover="closest")


def matriz_correlacion(corr: pd.DataFrame) -> go.Figure:
    fig = go.Figure(go.Heatmap(
        z=corr.to_numpy(dtype=float), x=list(corr.columns), y=list(corr.index), zmin=0, zmax=1,
        colorscale=SECUENCIAL, xgap=2, ygap=2, text=corr.round(2).to_numpy(), texttemplate="%{text:.2f}",
        colorbar=dict(thickness=12), hovertemplate="%{y} / %{x}: r = %{z:.2f}<extra></extra>"))
    fig.update_yaxes(autorange="reversed", showgrid=False, zeroline=False)
    return _estilo(fig, "Correlación de los promedios diarios", altura=max(300, 60 * len(corr) + 120),
                   hover="closest")


def mapa(puntos: pd.DataFrame, u: str) -> go.Figure | None:
    """puntos: clave, nombre, lat, lon, promedio, color (por nodo)."""
    p = puntos.dropna(subset=["lat", "lon"])
    if p.empty:
        return None
    lat_min, lat_max = p["lat"].min(), p["lat"].max()
    lon_min, lon_max = p["lon"].min(), p["lon"].max()
    extension = max(lat_max - lat_min, lon_max - lon_min)
    zoom = 4 if extension > 15 else 5 if extension > 8 else 6 if extension > 4 else 7 if extension > 2 else \
        8 if extension > 1 else 9
    fig = go.Figure()
    fig.add_trace(go.Scattermap(lat=p["lat"], lon=p["lon"], mode="markers", hoverinfo="skip", showlegend=False,
                                marker=dict(size=22, color="white", opacity=0.95)))
    fig.add_trace(go.Scattermap(
        lat=p["lat"], lon=p["lon"], mode="markers+text", text=p["clave"], textposition="top right",
        textfont=dict(size=12, color=TEXTO, family="Open Sans Bold"),
        marker=dict(size=15, color=list(p["color"]), opacity=1),
        customdata=[[n, v] for n, v in zip(p["nombre"], p["promedio"])],
        hovertemplate="<b>%{text}</b> · %{customdata[0]}<br>PML promedio %{customdata[1]:,.0f} " + u
                      + "<extra></extra>", showlegend=False))
    fig.update_layout(map=dict(style="open-street-map", zoom=zoom,
                               center=dict(lat=(lat_min + lat_max) / 2, lon=(lon_min + lon_max) / 2)),
                      height=460, margin=dict(l=0, r=0, t=8, b=0), paper_bgcolor="white")
    return fig
