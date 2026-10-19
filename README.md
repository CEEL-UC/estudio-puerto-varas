# Estudio Puerto Varas

Estudio del Centro UC Encuestas y Estudios Longitudinales (CEEL UC) sobre educación y migración estudiantil en el entorno de Puerto Varas: Puerto Varas, Puerto Montt, Frutillar, Osorno y el resto de la Región de Los Lagos, 2015-2025.

La pregunta de partida es si la llegada de hogares de mayores ingresos durante la pandemia cambió la oferta y la demanda de educación particular pagada en la zona, y quiénes son las familias que llegaron.

**Sitio publicado:** https://ceel-uc.github.io/estudio-puerto-varas/

## Cómo está organizado el proyecto

La estructura de carpetas del repositorio es un espejo de la de Dropbox, con los mismos nombres, para que cualquier ruta se pueda ubicar en los dos lados.

```
estudio-puerto-varas/
├── index.html                  Sitio (GitHub Pages lo publica desde la raíz de main)
├── datos_manzanas.js           Datos del mapa del Censo 2024 por manzana
├── datos_cohortes.js           Datos del gráfico de año de llegada
├── logo_ceel.png
├── Análisis y metadata/
│   ├── Dofiles/                Código en Stata y Python (ver orden abajo)
│   ├── Outputs/                Resultados agregados: tablas, Excel de respaldo, informes
│   │   ├── Indice socioeconomico manzanas/   Índice por manzana: documento metodológico y Excel con fórmulas
│   │   └── prototipo_manzanas/               Primer prototipo del mapa por manzana
│   ├── graficado/              Datos consolidados para gráficos
│   └── Resumen de datos.md     Primer mapeo de fuentes (septiembre 2026)
└── Datos/
    └── README.md               Catálogo de las fuentes: qué hay en Dropbox y de dónde se descarga
```

## Flujo de datos

```
Datos fuente (Dropbox)          Código (Dofiles/)                   Productos
──────────────────────          ─────────────────                   ─────────
Mineduc: directorio,      ──►   00_master.do (paneles)        ──►   Outputs/ (tablas Stata y Excel)
resumen y matrícula             01_analisis.do
por estudiante                  02_migracion_destino.do
                                03_migracion_regional.py      ──►   Secciones 01 a 03 del sitio
                                04_cohortes_llegada.py        ──►   datos_cohortes.js
                                05_matricula_niveles.py
Censo 2024: base manzana  ──►   prep_manzanas_censo.py        ──►   datos_manzanas.js (sección 04)
y cartografía Región 10         indice_socioeconomico_*.py    ──►   Outputs/Indice socioeconomico manzanas/
Todo lo anterior          ──►   06_respaldo_sitio_excel.py    ──►   Excel de respaldo de las cifras del sitio
                                07 y 08 (ingreso a 1° básico) ──►   Informe Word
```

Orden de ejecución:

1. `00_master.do`: filtra los archivos nacionales de Mineduc a las comunas del estudio y arma los paneles (directorio, resumen de matrícula, matrícula por estudiante).
2. `01_analisis.do`: indicadores de matrícula por dependencia y curso, establecimientos y perfil por comuna.
3. `02_migracion_destino.do`: destino de los estudiantes que salen de cada comuna.
4. `03_migracion_regional.py` a `05_matricula_niveles.py`: datos de las secciones 01 a 03 del sitio, con la Región de Los Lagos completa.
5. `prep_manzanas_censo.py` e `indice_socioeconomico_manzanas*.py`: mapa por manzana e índice socioeconómico (documento metodológico en `indice_socioeconomico_manzanas_documento.md`).
6. `06_respaldo_sitio_excel.py`: respaldo en Excel de todas las cifras del sitio.
7. `07_analisis_ingreso_1basico.py` y `08_informe_ingreso_1basico_word.py`: informe sobre ingreso a 1° básico y año de llegada.

Requisitos: Stata y Python 3.12 con `polars`, `pyarrow`, `numpy`, `shapely`, `openpyxl` y `python-docx`.

**Rutas.** Los scripts tienen la ruta de la carpeta del proyecto escrita al inicio (`global raiz` en Stata, constantes en mayúscula como `SALIDA` o `REPO` en Python). Para correrlos en otro computador hay que cambiar esas líneas para que apunten a la carpeta de Dropbox y a la copia local de este repositorio.

## Acceso a datos

Los datos del proyecto están en la [carpeta de Dropbox del proyecto](https://www.dropbox.com/scl/fo/wlzpsmzhn1is6qe3lccqw/ALvTera0h52qn6F2zkhiHlk?rlkey=vsv7vprk6z2273aacrk0yiffz&st=4nvf4bxe&dl=0).

## Fuentes

- Ministerio de Educación, Centro de Estudios, [Datos Abiertos](https://datosabiertos.mineduc.cl/): Directorio de Establecimientos, Resumen de Matrícula por Establecimiento y Matrícula por Estudiante.
- Agencia de Calidad de la Educación: SIMCE 2025.
- Instituto Nacional de Estadísticas: Censo de Población y Vivienda 2024 (microdatos y cartografía) y matrices de migración interna.
- Ministerio de Desarrollo Social y Familia: Encuesta CASEN 2017, 2022 y 2024.

## Contacto

Centro UC Encuestas y Estudios Longitudinales, Pontificia Universidad Católica de Chile. https://encuestas.uc.cl
