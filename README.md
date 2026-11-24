# ⚡ PML CENACE

Aplicación web para descargar y analizar los Precios Marginales Locales (PML) del Mercado Eléctrico Mayorista, a partir del servicio web SW-PML del CENACE.

**Recurrent Energy / Canadian Solar — México · Inteligencia de Mercado**

Es un fork de [CENACE PML Analyzer](https://github.com/sebasrf44/cenace-pml-analyzer) de Sebastian Roldan (SRF), reducido a lo que trata de PML.

## Funcionalidades

- **Buscador de nodos:** lista desplegable con todo el catálogo. Busca por nombre, clave, zona o estado: escribir «Tast» muestra Tastiota y Maniobras Tastiota, y «Select N matches» agrega todas las coincidencias. Los filtros (sistema, estado, municipio, zona, CCR, kV) acotan la búsqueda.
- **Importar lista de claves** pegadas desde Excel y **listas guardadas** (`data/listas_nodos.json`).
- **Enlaces para compartir:** cada consulta ejecutada deja su enlace (`?nodos=…&proceso=…&desde=…&hasta=…&moneda=…&modo=…`), que abre la app con todo precargado.
- **Novedades del catálogo:** NodosP nuevos, bajas y cambios de la última versión, con botón para agregar los nuevos al buscador.
- Descarga paralela de PML horarios (MTR y MDA) del SW-PML del CENACE, con el sistema (SIN, BCA o BCS) detectado de cada nodo en el catálogo.
- **USD con el FIX del día:** cada hora se divide entre el FIX de su día de operación; los fines de semana y feriados usan el último FIX publicado. Fuente: Banxico SIE (serie SF43718) si hay token; si no, el FIX publicado en el DOF (SIDOF), asignado a su fecha de determinación. Si ninguna responde, la app lo dice con un error y deja todo en MXN.
- **Modo Análisis:** KPIs, mapa de PML por nodo (ubicación aproximada con OpenStreetMap), PML por CCR, tabla resumen por nodo, evolución temporal, heatmap hora × mes, promedio mensual por año, spread por CCR y rankings (PML, volatilidad, horas negativas).
- **Modo Datos:** Excel de datos (una hoja por nodo con PML y componentes de energía, pérdidas y congestión, más el TC aplicado en USD) y Excel de análisis con secciones a elegir (resumen, rankings, multi-año y spread por CCR).

## Catálogo automático

El CENACE publica el catálogo de NodosP cada mes en https://www.cenace.gob.mx/Paginas/SIM/NodosP.aspx. La GitHub Action `Actualizar catálogo de NodosP` (`.github/workflows/actualizar-catalogo.yml`) corre cada lunes y hace lo siguiente:

1. Busca la versión más nueva publicada.
2. Si es posterior a la del repositorio, la descarga y la valida (versión, número de nodos, campos).
3. Escribe `data/catalogo_nodos.xlsx` y `data/catalogo_cambios.json` (altas, bajas, cambios y «Control de Cambios» del CENACE).
4. Hace commit. Streamlit vuelve a publicar solo.

Se puede correr a mano desde *Actions → Actualizar catálogo de NodosP → Run workflow*, o en local:

```bash
python scripts/actualizar_catalogo.py            # revisar y actualizar
python scripts/actualizar_catalogo.py --revisar  # solo informar
```

La subida de un xlsx desde la barra lateral sigue disponible para una sesión.

## Diferencias con el original

| Original (v9) | Este fork |
|---|---|
| BESS scoring por caso de uso (arbitraje, SSAA, firming) | Eliminado |
| «Excel de Análisis» BESS y opciones BESS del Excel a la medida | Un solo Excel de análisis de PML, con secciones a elegir |
| Modo «Mapa» (solo geocodificación) y KMZ | Eliminados; el mapa queda en el dashboard, coloreado por PML |
| Caja de claves y búsqueda por subcadena que reemplazaba la lista | Buscador desplegable por nombre o clave, filtros que acotan, importación, listas guardadas y enlaces |
| Catálogo fijo, actualizado a mano | GitHub Action semanal con novedades de cada versión |
| USD solo con token de Banxico; sin él, MXN con un aviso | Banxico o DOF; error visible si no hay tipo de cambio |
| Excel de datos en USD con los valores en MXN | Cada hora dividida entre el FIX de su día, con columna del TC aplicado |
| Caché que podía mostrar la consulta anterior si la nueva tenía el mismo largo | Caché con huella del contenido y de los tipos de cambio |
| Registros en el orden en que terminan las consultas paralelas | Ordenados por fecha y hora |
| Aviso genérico de rezago de publicación | Aviso según el proceso: MTR a día de operación +7, MDA disponible un día antes |
| `go.Scattermapbox` y `use_container_width` | `go.Scattermap` (Plotly 7) y `width="stretch"` (Streamlit 1.67) |

## Estructura

```
.
├── app.py                                 # App Streamlit
├── catalogo.py                            # Lectura, etiquetas y diferencias del catálogo
├── scripts/actualizar_catalogo.py         # Descarga y valida el catálogo nuevo
├── .github/workflows/actualizar-catalogo.yml
├── requirements.txt                       # Librerías (versiones fijas, probadas)
├── .streamlit/config.toml                 # Tema visual RE
├── data/catalogo_nodos.xlsx               # Catálogo oficial de NodosP del CENACE
├── data/catalogo_cambios.json             # Novedades de la última versión (lo escribe la Action)
├── data/listas_nodos.json                 # Listas guardadas {nombre: [claves]}
└── tests/                                 # Pruebas con CENACE, Banxico, DOF y OSM simulados
```

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre http://localhost:8501

Opcional, para usar Banxico directo en lugar del DOF: crea `.streamlit/secrets.toml` (no se sube al repositorio) con:

```toml
BANXICO_TOKEN = "tu-token"
```

El token es gratuito: https://www.banxico.org.mx/SieAPIRest/service/v1/token

## Pruebas

```bash
pip install -r requirements.txt pytest
pytest -q tests
```

Las pruebas simulan el CENACE, Banxico, el DOF y OpenStreetMap, así que no necesitan red.

## Publicación

Streamlit Community Cloud, desde la rama `main` de este repositorio. Cada push a `main` vuelve a publicar la app. El token de Banxico, si se usa, va en *Settings → Secrets* de la app.

## Datos

- PML: CENACE, servicio web SW-PML (`https://ws01.cenace.gob.mx:8082/SWPML/SIM`). Hasta 20 NodosP y 7 días de operación por consulta; el MTR se publica 7 días después del día de operación y el MDA un día antes.
- Tipo de cambio: FIX de Banxico (SF43718, fecha de determinación) o FIX publicado en el DOF (SIDOF, indicador 158), que sale un día hábil después de su determinación.
- Catálogo de NodosP: CENACE, publicación mensual.
- Ubicación de nodos: OpenStreetMap (Overpass API), por coincidencia de nombres. Es aproximada, no es la georreferencia oficial.
