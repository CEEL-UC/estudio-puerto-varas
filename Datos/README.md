# Datos

Esta carpeta está vacía en GitHub a propósito. Los datos fuente están en la carpeta `Datos\` del proyecto en Dropbox (alrededor de 16 GB). Todos son públicos y se pueden volver a descargar desde los enlaces de abajo.

| Carpeta en Dropbox | Contenido | Tamaño aprox. | Fuente |
|---|---|---|---|
| `Mineduc/Directorio establecimientos/` | Directorio oficial de establecimientos, una carpeta por año (2015-2025): RBD, comuna, dependencia, coordenadas | | [Datos Abiertos Mineduc](https://datosabiertos.mineduc.cl/) |
| `Mineduc/Resumen matricula establecimiento/` | Matrícula por establecimiento y curso, 2004-2025 | | [Datos Abiertos Mineduc](https://datosabiertos.mineduc.cl/) |
| `Mineduc/Matricula por estudiante/` | Matrícula nacional con un registro por estudiante y año, 2004-2025 (identificador MRUN enmascarado; vacío en 2004-2008) | 12 GB (toda la carpeta Mineduc) | [Datos Abiertos Mineduc](https://datosabiertos.mineduc.cl/) |
| `Mineduc/excel-matricula-longuitudinal-2010-2025_no_modificable.xlsx` | Matrícula agregada por comuna, dependencia y año | | Centro de Estudios Mineduc |
| `SIMCE 2025/` | Resultados por establecimiento de 4° básico, 8° básico y 2° medio, con grupo socioeconómico | 6 MB | [Agencia de Calidad de la Educación](https://informacionestadistica.agenciaeducacion.cl/) |
| `Censo 2024/` | Microdatos de personas, hogares y viviendas (parquet), base a nivel de manzana y entidad, cartografía de la Región de Los Lagos y su documentación | 790 MB | [INE, Censo 2024](https://censo2024.ine.gob.cl/) |
| `CASEN/` | CASEN 2017, 2022 y 2024 (Stata) y cuestionario | 2,6 GB | [Observatorio Social, MDSF](https://observatorio.ministeriodesarrollosocial.gob.cl/) |
| `BIDAT Base Residentes/` | Matrices de origen y destino de migración interna por comuna, sexo y tramo de edad, 2018-2024 | 200 MB | INE |
| `[merge] Panel Puerto Varas/` | Paneles construidos por `00_master.do` y caché liviano de los archivos nacionales para los scripts de Python | 620 MB | Elaboración propia |

La carpeta `[merge] Panel Puerto Varas/` no se descarga: se reconstruye corriendo `Análisis y metadata/Dofiles/00_master.do` sobre los datos de Mineduc. La carpeta tiene corchetes en el nombre; al leer los parquet desde Python hay que desactivar los comodines (por ejemplo, `glob=False` en polars).

Para reproducir el estudio en otro computador basta con recrear esta estructura de carpetas con los archivos descargados y ajustar la ruta del proyecto al inicio de cada script.
