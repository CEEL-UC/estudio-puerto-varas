# Resumen de datos - Estudio Puerto Varas

Actualizado: 16 sep 2026

## Datos en carpeta (`Datos\`)

**Mineduc, matrícula longitudinal** (`Datos\Mineduc\`)
Matrícula agregada por comuna, dependencia y año, 2010-2025. No es microdato individual, no permite seguir al mismo alumno en el tiempo. Sirve para el volumen de matrícula particular pagada por comuna y año.

**SIMCE 2025** (`Datos\SIMCE 2025\`)
4° básico, 8° básico y 2° medio. Nivel establecimiento (RBD): comuna, dependencia, grupo socioeconómico, alumnos evaluados, resultados. No trae matrícula total, sí identifica cada colegio particular pagado con su GSE.

**Censo 2024** (`Datos\Censo 2024\`)
Microdato nacional de personas, hogares y viviendas (formato parquet). Incluye lugar de nacimiento, residencia hace 5 años y período de llegada a la vivienda actual (migración pre/post pandemia), escolaridad, ocupación, hijos nacidos vivos. Llega hasta comuna, no a manzana.

**CASEN 2017, 2022 y 2024** (`Datos\CASEN\`)
No traen variable `comuna` plana, la ubicación va dentro de `estrato` (comuna-área-NSE), falta el libro de códigos para saber si Puerto Varas viene identificada o agrupada en "resto de la región". Migración interna (comuna hace 5 años) presente en 2017 y 2022, ausente en 2024.

## Fuentes mapeadas, no descargadas

- Mineduc, matrícula por estudiante y directorio de establecimientos (público, datosabiertos.mineduc.cl)
- Mineduc, acceso restringido con RUN estable para panel de alumnos (trámite pendiente, memo dejado)
- DEIS/Minsal, estadísticas vitales (natalidad por comuna)
- INE, estimaciones de población, Censo 2017 (manzana vía REDATAM), permisos de edificación
- SII, estadísticas de renta por comuna
- Registro Social de Hogares
- BCN, portal SIIT (catálogo completo revisado, sin descargas por ahora)
