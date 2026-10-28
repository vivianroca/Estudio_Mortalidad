# Estudio de mortalidad en Guatemala

Investigación descriptiva de las defunciones en Guatemala durante **2009–2024**. El proyecto descarga cuadros oficiales, audita su estructura y transforma indicadores a formato largo, conservando etiquetas y validando totales. Los cuadros son agregados; no son registros individuales de personas.

## Fuentes

Instituto Nacional de Estadística de Guatemala ([INE](https://www.ine.gob.gt/)), sección de estadísticas vitales, «Principales resultados: Defunciones». El descargador utiliza el [endpoint del INE](https://www.ine.gob.gt/bdatos_cargar.php) con categoría `82`, período `1` y referencia al [sistema del INE](https://www.ine.gob.gt/sistema/). La disponibilidad depende del servidor y puede verse limitada por CAPTCHA. Se conservan los originales localmente sin modificarlos.

## Estructura

- `src/`: descarga, auditorías y transformación de nueve indicadores.
- `dashboards/preparar_ds_analisis1.py`: preparación y validación de datos para el dashboard local.
- `docu/`: definición de indicadores y propuesta de modelo de datos en Markdown, Quarto y Mermaid.
- `catalog/catalogo_cie10.csv`: catálogo de códigos y descripciones, sin conteos de defunciones.
- `reports/`: inventarios pequeños de hojas y comparación estructural.
- `data/manifiesto_defunciones.csv`: metadatos de procedencia y descarga.
- `data/raw_data/`: originales del INE, excluidos de Git.
- `data/processed/`: resultados derivados, excluidos de Git.

También se excluyen libros de trabajo de revisión, reportes de validación con conteos, el PDF de referencia y las exportaciones HTML. El dashboard `.qmd` contiene una instantánea de datos procesados y permanece local.

## Entorno y ejecución

El entorno local utiliza Python 3.14. `requirements.txt` fija las versiones instaladas de las dependencias directas (pandas, requests, Beautiful Soup, openpyxl y xlrd) y sus dependencias de entorno, incluido lxml.

Desde la raíz del proyecto:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
mkdir -p data/raw_data data/processed reports/estructura
python src/descargar_defunciones.py 2009
```

Repetir la descarga para cada año hasta 2024; los archivos esperados son `2009.xls`/`.xlsx` hasta `2024.xls`/`.xlsx`. Si el INE bloquea la descarga, obtener los cuadros desde su portal y conservar esos nombres en `data/raw_data/`.

Con los 16 libros locales disponibles:

```bash
python src/auditar_hojas.py
python src/auditar_estructura.py
python src/transformar_ind002.py
python src/transformar_ind004.py
python src/transformar_ind005.py
python src/transformar_ind006.py
python src/transformar_ind009.py
python src/transformar_ind010.py
python src/transformar_ind013.py
python src/transformar_ind014.py
```

Las auditorías regeneran metadatos en `reports/estructura/`, necesarios para los transformadores. Revisar los informes y códigos de salida antes de utilizar los resultados.

IND-007 requiere además los libros locales de revisión aprobada `reports/estructura/propuesta_cat_cie10.xlsx` (hoja `rev_final`) y `catalog/posible_cat.xlsx` (hoja `Revision_IND007`). No se distribuyen en este repositorio. Una vez disponibles, ejecutar `python src/transformar_ind007.py`. Se excluye 2011 solo de este indicador por una inconsistencia del original; su ausencia no significa cero defunciones.

El dashboard requiere los CSV de IND-002, IND-004 e IND-006, el catálogo y el archivo local `dashboards/ds_analisis1.qmd`. Con esos insumos, `python dashboards/preparar_ds_analisis1.py` actualiza la instantánea incrustada. Para renderizar los documentos `.qmd` se necesita Quarto, instalado por separado.

## Estado actual y reproducibilidad

Proyecto de investigación en desarrollo: nueve transformadores, auditorías, validaciones locales, propuesta de modelo y dashboard exploratorio. Consultar `docu/indicadores_definicion.md` para cobertura y advertencias. Los resultados no constituyen inferencias causales.

Git distribuye código, documentación y metadatos; los datos y ciertos insumos de revisión permanecen locales. Por ello, clonar el repositorio por sí solo no reproduce IND-007 ni el dashboard completo. Los scripts de procesamiento se conservan sin cambios.

## Preparación para DVC

DVC **todavía no está inicializado** ni se incluye como dependencia. Las rutas de originales y procesados ya están separadas y excluidas de Git. En una etapa posterior se podrá inicializar DVC, acordar el almacenamiento remoto y versionar datos e insumos de revisión mediante archivos `.dvc` y un flujo reproducible. La caché y la configuración privada de DVC están previstas en `.gitignore`; las futuras referencias `.dvc` podrán incluirse en Git. No guardar credenciales en el repositorio.
