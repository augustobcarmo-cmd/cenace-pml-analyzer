# ⚡ PML CENACE

Aplicación web que arma un **informe estándar de NodosP** con los Precios Marginales Locales (PML) del Mercado de Tiempo Real (MTR) del CENACE: estadísticas de precio, congestión y pérdidas, distribuciones anual e intradiaria, métricas para solar y baterías y lecturas automáticas.

**Recurrent Energy / Canadian Solar — México · Inteligencia de Mercado**

Es un fork de [CENACE PML Analyzer](https://github.com/sebasrf44/cenace-pml-analyzer) de Sebastian Roldan (SRF), reducido a lo que trata de PML.

App publicada: https://pml-cenace-mx.streamlit.app

## Cómo se usa

1. Busca uno o varios nodos en la barra lateral (por nombre, clave, zona o estado; «Tast» muestra Tastiota y Maniobras Tastiota).
2. Pulsa **📑 Generar informe**.
   - **Un nodo:** ficha individual.
   - **Dos o más:** comparación lado a lado, con la ficha de cada nodo a un clic.
3. Ajusta el **período** (todo el historial, período común, últimos 12 meses, últimos 3 años, un año o personalizado) y la **moneda** (MXN o USD): el informe se recalcula sin volver a descargar.
4. Descarga el **informe en Excel** o los **datos horarios** (Excel o CSV) al final del informe.

## Qué trae el informe

| Sección | Contenido |
|---|---|
| Indicadores | Promedio, mediana, P10–P90, máximo y mínimo (con fecha y hora), últimos 12 meses frente a los 12 previos, diferencial nodal, % de horas con congestión, captura solar y spread diario 4 h |
| 📌 Lecturas | Conclusiones automáticas: nivel, tendencia, señal de red (zona exportadora o importadora), congestión y sus horas, precios negativos, perfil del día y curva de pato, estacionalidad, canibalización solar, oportunidad para baterías, volatilidad y calidad de los datos |
| 📈 Historia | PML mensual con banda P10–P90 y componente de energía del sistema, promedio por año y tabla anual completa |
| 🔌 Congestión y pérdidas | Diferencial nodal mensual (pérdidas + congestión), mapa de calor de la congestión por mes y hora, componentes por hora y % de horas congestionadas por año |
| 🕐 Distribuciones | Mapa de calor del PML por mes y hora, perfil intradiario (P10–P90, P25–P75, mediana y promedio), perfil por año, estacionalidad y curva de duración |
| ☀️🔋 Solar y baterías | Precio capturado y factor de captura de un perfil solar genérico (y su evolución anual), frecuencia de cada hora entre las 4 más baratas o más caras del día, spread 4 h mensual y margen teórico de una batería 4 h / 1 ciclo |
| Comparación | Tabla de indicadores por nodo, PML mensual y perfil intradiario por nodo, diferencial nodal y congestión por hora, diferencial horario frente a un nodo de referencia, solar, baterías y correlación de los promedios diarios |

Convenciones (también en la portada del Excel):

- **Diferencial nodal** = PML − componente de energía = pérdidas + congestión. El componente de energía es el mismo para todo el sistema en cada hora; el diferencial es lo propio del punto de la red. Negativo: zona exportadora (la red castiga inyectar ahí). Positivo: zona importadora.
- **Hora congestionada:** |componente de congestión| > 1 MXN/MWh.
- **Perfil solar genérico:** seno entre 6:30 y 19:00 (hora CENACE), igual todo el año. Sirve para comparar nodos y años, no para un proyecto concreto.
- **Batería 4 h / 1 ciclo:** descarga en las 4 horas más caras y carga en las 4 más baratas de cada día, sin restricción de orden, eficiencia de ida y vuelta 85 %. Es una cota superior del arbitraje, no un ingreso esperado.
- **USD:** cada hora se divide entre el FIX de su día de operación; fines de semana y feriados usan el último FIX publicado. Fuente: Banxico SIE (SF43718) si hay token; si no, el FIX publicado en el DOF (SIDOF), asignado a su fecha de determinación. Si ninguna responde, la app lo dice y deja el informe en MXN.

## Historial completo: guardado y bajo demanda

El SW-PML entrega como máximo 7 días y 20 nodos por consulta, y el MTR existe desde el 27-ene-2017: el historial de un nodo son unas 510 consultas (hasta 20 nodos comparten las mismas consultas). Para no repetir ese trabajo:

- **Nodos de las listas guardadas** (`data/listas_nodos.json`): la GitHub Action `Actualizar historial de PML (MTR)` guarda todo su historial en la rama **`datos`** (un Parquet por nodo e `indice.json`) y suma los días nuevos dos veces al día. Su informe sale de inmediato, con todo el historial. En el buscador llevan 💾.
- **Cualquier otro nodo:** el primer informe descarga todo su historial del CENACE (unas 460-510 consultas en paralelo; el 24-nov-2026 tardó ~1-2 min desde Streamlit Cloud). Si la estimación pasa de 4 minutos, baja primero los últimos 12 meses. Tras muchas consultas seguidas el CENACE empieza a demorar las respuestas: la descarga se corta a los 2.5 min con lo más reciente, y las semanas que falten se completan con **⏪ Un año más** y **⏮ Completar desde 2017**. Lo descargado queda en memoria y lo comparten todas las sesiones mientras la app está despierta; Streamlit Cloud borra la memoria cuando la app se duerme.
- Para guardar un nodo para siempre, agrégalo a una lista en `data/listas_nodos.json`: el push dispara la Action, que descarga su historial completo.

La rama `datos` se publica como un único commit sin historial, así el repositorio no crece. La Action se puede correr a mano en *Actions → Actualizar historial de PML (MTR) → Run workflow*, con claves extra opcionales.

Comportamiento del SW-PML verificado el 24-nov-2026: lapsos de más de 7 días o más de 20 nodos dan HTTP 400; fechas antes del inicio del mercado, también. Un nodo sin datos en el lapso (porque aún no existía) llega vacío sin afectar a los demás. Un lapso que pasa del último día publicado devuelve solo los días publicados. Los días de cambio de horario (hasta 2022) tienen 23 o 25 horas.

## Catálogo automático

El CENACE publica el catálogo de NodosP cada mes en https://www.cenace.gob.mx/Paginas/SIM/NodosP.aspx. La GitHub Action `Actualizar catálogo de NodosP` corre cada lunes:

1. Busca la versión más nueva publicada.
2. Si es posterior a la del repositorio, la descarga y la valida (versión, número de nodos, campos).
3. Escribe `data/catalogo_nodos.xlsx` y `data/catalogo_cambios.json` (altas, bajas, cambios y «Control de Cambios» del CENACE).
4. Hace commit. Streamlit vuelve a publicar solo.

La app muestra las novedades de cada versión, con un botón para agregar los nodos nuevos al buscador.

## Otras funciones

- Filtros que acotan el buscador (sistema, estado, municipio, zona, CCR, kV), importación de claves pegadas desde Excel y listas guardadas.
- **Enlaces:** cada informe deja su enlace (`?nodos=…&moneda=…&periodo=…`), que abre la app y genera el mismo informe. Los enlaces de la versión anterior (`desde`/`hasta`) siguen funcionando.
- **Ubicación aproximada** de los nodos en un mapa (subestaciones de OpenStreetMap con nombre parecido en el mismo estado).

## Diferencias con el original

| Original (v9) | Este fork |
|---|---|
| BESS scoring por caso de uso, Excel BESS, modo «Mapa» y KMZ | Eliminados |
| «Modo de uso» (Análisis / Datos) y fechas a elegir antes de descargar | Un solo flujo: nodos → informe estándar con todo el historial; el período se ajusta después |
| MTR y MDA | Solo MTR |
| Cada consulta descarga todo de nuevo | Historial guardado (rama `datos`) y memoria compartida; solo se descarga lo que falta |
| Dashboard de PML sin componentes | Congestión, pérdidas y diferencial nodal; perfiles para solar y baterías; lecturas automáticas |
| Caja de claves y catálogo actualizado a mano | Buscador desplegable, filtros, importación, listas, enlaces y catálogo automático |
| USD solo con token de Banxico; Excel de datos en USD con valores en MXN | Banxico o DOF, FIX de cada día, error visible si no hay tipo de cambio |
| Plotly `Scattermapbox`, `use_container_width` | `Scattermap` (Plotly 7), `width="stretch"` (Streamlit 1.67) |

## Estructura

```
.
├── app.py                                 # App Streamlit (interfaz)
├── datos_pml.py                           # SW-PML: consultas, cobertura, almacén en memoria y rama datos
├── informe.py                             # Cálculos del informe y lecturas automáticas
├── graficas.py                            # Gráficas Plotly
├── excel_informe.py                       # Excel del informe y datos horarios (Excel, CSV)
├── ubicacion.py                           # Subestaciones de OpenStreetMap y emparejamiento
├── catalogo.py                            # Lectura, etiquetas y diferencias del catálogo
├── scripts/actualizar_datos.py            # Historial de los nodos de las listas (rama datos)
├── scripts/actualizar_catalogo.py         # Descarga y valida el catálogo nuevo
├── .github/workflows/actualizar-datos.yml
├── .github/workflows/actualizar-catalogo.yml
├── data/catalogo_nodos.xlsx               # Catálogo oficial de NodosP del CENACE
├── data/catalogo_cambios.json             # Novedades de la última versión (lo escribe la Action)
├── data/listas_nodos.json                 # Listas guardadas {nombre: [claves]}; sus nodos se guardan en la rama datos
└── tests/                                 # Pruebas con CENACE, Banxico, DOF, GitHub y OSM simulados
```

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opcional, para usar Banxico directo en lugar del DOF: crea `.streamlit/secrets.toml` (no se sube al repositorio) con `BANXICO_TOKEN = "tu-token"`. El token es gratuito: https://www.banxico.org.mx/SieAPIRest/service/v1/token

## Pruebas

```bash
pip install -r requirements.txt pytest
python -m pytest -q tests
```

Las pruebas simulan el CENACE (con sus límites reales), Banxico, el DOF, la rama `datos` y OpenStreetMap, así que no necesitan red.

## Publicación

Streamlit Community Cloud, desde la rama `main`. Cada push a `main` vuelve a publicar la app. El token de Banxico, si se usa, va en *Settings → Secrets* de la app.

## Datos

- PML: CENACE, servicio web SW-PML (`https://ws01.cenace.gob.mx:8082/SWPML/SIM`), proceso MTR (ex post), publicado 7 días después del día de operación.
- Tipo de cambio: FIX de Banxico (SF43718, fecha de determinación) o FIX publicado en el DOF (SIDOF, indicador 158), que sale un día hábil después de su determinación.
- Catálogo de NodosP: CENACE, publicación mensual.
- Ubicación de nodos: OpenStreetMap (Overpass API), por coincidencia de nombres. Es aproximada, no es la georreferencia oficial.
