# ⚡ PML CENACE

Aplicación web para descargar y analizar los Precios Marginales Locales (PML) del Mercado Eléctrico Mayorista, a partir del servicio web SW-PML del CENACE.

**Recurrent Energy / Canadian Solar — México · Inteligencia de Mercado**

Es un fork de [CENACE PML Analyzer](https://github.com/sebasrf44/cenace-pml-analyzer) de Sebastian Roldan (SRF), reducido a lo que trata de PML.

## Funcionalidades

- Descarga paralela de PML horarios (MTR y MDA) del SW-PML del CENACE, con el sistema (SIN, BCA o BCS) detectado de cada nodo en el catálogo de NodosP.
- Conversión a USD con el FIX de Banxico (serie SF43718) del día de operación de cada hora; los fines de semana y feriados usan el último FIX publicado.
- Búsqueda de nodos por sistema, estado, municipio, zona de carga, CCR y tensión.
- **Modo Análisis**: KPIs, mapa de PML por nodo (ubicación aproximada con OpenStreetMap), PML por CCR, tabla resumen por nodo, evolución temporal, heatmap hora × mes, promedio mensual por año, spread por CCR y rankings (PML, volatilidad, horas negativas).
- **Modo Datos**: Excel de datos (una hoja por nodo con PML y componentes de energía, pérdidas y congestión) y Excel de análisis con secciones a elegir (resumen, rankings, multi-año y spread por CCR).

## Diferencias con el original

| Original (v9) | Este fork |
|---|---|
| BESS scoring por caso de uso (arbitraje, SSAA, firming) | Eliminado |
| «Excel de Análisis» BESS y opciones BESS del Excel a la medida | Un solo Excel de análisis de PML, con secciones a elegir |
| Modo «Mapa» (solo geocodificación) y KMZ | Eliminados; el mapa queda en el dashboard, coloreado por PML |
| Registros en el orden en que terminan las consultas paralelas | Ordenados por fecha y hora |
| Excel de datos en USD con los valores en MXN | Cada hora dividida entre el FIX de su día (fines de semana y feriados: último FIX publicado), con columna del TC aplicado |
| Caché que podía mostrar la consulta anterior si la nueva tenía el mismo largo | Caché con huella del contenido y de los tipos de cambio |
| Aviso genérico de rezago de publicación | Aviso según el proceso: MTR a día de operación +7, MDA disponible un día antes |
| Catálogo de NodosP v2026-04-23 | Catálogo v2026-07-16, con la versión visible en la app |
| `go.Scattermapbox` y `use_container_width` | `go.Scattermap` (Plotly 7) y `width="stretch"` (Streamlit 1.67) |

## Estructura

```
.
├── app.py                    # App Streamlit
├── requirements.txt          # Librerías (versiones fijas, probadas)
├── .streamlit/config.toml    # Tema visual RE
├── data/catalogo_nodos.xlsx  # Catálogo oficial de NodosP del CENACE
└── tests/                    # Pruebas con CENACE, Banxico y OSM simulados
```

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre http://localhost:8501

Para la conversión a USD, crea `.streamlit/secrets.toml` (no se sube al repositorio) con:

```toml
BANXICO_TOKEN = "tu-token"
```

El token es gratuito: https://www.banxico.org.mx/SieAPIRest/service/v1/token

## Pruebas

```bash
pip install -r requirements.txt pytest
pytest -q tests
```

La prueba simula las respuestas del CENACE y de OpenStreetMap, así que no necesita red.

## Publicación

Streamlit Community Cloud, desde la rama `main` de este repositorio. Cada push a `main` vuelve a publicar la app. El token de Banxico va en *Settings → Secrets* de la app.

## Datos

- PML: CENACE, servicio web SW-PML (`https://ws01.cenace.gob.mx:8082/SWPML/SIM`). Hasta 20 NodosP y 7 días de operación por consulta; el MTR se publica 7 días después del día de operación y el MDA un día antes.
- Tipo de cambio: Banxico, FIX (SF43718).
- Ubicación de nodos: OpenStreetMap (Overpass API), por coincidencia de nombres. Es aproximada, no es la georreferencia oficial.
