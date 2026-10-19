# Por qué la categoría "Ingresó a 1° básico en la comuna" es tan chica en el gráfico de año de llegada (sección 03).
# Reproduce la clasificación de 04_cohortes_llegada.py para la matrícula de Puerto Varas 2025 y sigue a cada
# cohorte de 1° básico 2015-2025 hasta 2025, para ver en qué barra y categoría del gráfico quedó.
# Escribe Outputs\analisis_ingreso_1basico.json (lo leen 06_respaldo_sitio_excel.py y el informe Word).
# Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 07_analisis_ingreso_1basico.py
import json, os
import polars as pl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
CACHE = os.path.join(BASE, 'Datos', '[merge] Panel Puerto Varas', 'nacional_liviano')
SALIDA = os.path.join(BASE, 'Análisis y metadata', 'Outputs', 'analisis_ingreso_1basico.json')
ANIOS = list(range(2014, 2026))
PV = 10109
c = pl.col
CAT = {'ce': 'En la comuna desde 2014 o antes', 'ip': 'Comenzó en parvularia o escuela de lenguaje',
       'ib': 'Ingresó a 1° básico en la comuna', 'sr': 'Sin registro previo, otros cursos',
       'll': 'Llegó desde otra comuna de Los Lagos', 'or': 'Llegó desde otra región'}

d = {a: pl.read_parquet(os.path.join(CACHE, f'mat_{a}.parquet'), glob=False).filter(c('MRUN').is_not_null()) for a in ANIOS}


def una_fila(df):
    # una matrícula por alumno y año; si tiene varias se prefiere el tramo regular (igual que 04_cohortes_llegada.py)
    return df.sort(['MRUN', 'regular'], descending=[False, True]).unique('MRUN', keep='first')


# ---- 1. clasificación de la matrícula regular de Puerto Varas 2025 (mismo criterio que el sitio)
u = d[2025].filter((c('COD_COM_RBD') == PV) & c('regular')).sort(['MRUN', 'RBD']).unique('MRUN', keep='first').select('MRUN', 'curso')
loc = pl.concat([una_fila(d[a].filter(c('MRUN').is_in(u['MRUN'].implode())))
                 .select('MRUN', pl.lit(a, pl.Int32).alias('anio'), c('COD_COM_RBD').alias('com'),
                         pl.when(c('COD_ENSE').is_in([10, 214])).then(pl.lit('par')).when(c('curso') == 1).then(pl.lit('1b')).otherwise(pl.lit('otro')).alias('niv'))
                 for a in ANIOS])
filas = []
for r in loc.pivot(on='anio', index='MRUN', values='com').iter_rows(named=True):
    a = 2025
    while a - 1 >= ANIOS[0] and r.get(str(a - 1)) == PV:  # inicio del tramo continuo más reciente en la comuna
        a -= 1
    filas.append((r['MRUN'], a, r.get(str(a - 1)) if a > ANIOS[0] else None))
s = (pl.DataFrame(filas, schema=['MRUN', 'inicio', 'com_prev'], orient='row').with_columns(c('inicio').cast(pl.Int32))
     .join(loc.select('MRUN', c('anio').alias('inicio'), 'niv'), on=['MRUN', 'inicio'], how='left').join(u, on='MRUN'))
s = s.with_columns(origen=pl.when(c('inicio') == ANIOS[0]).then(pl.lit('ce'))
                   .when(c('com_prev').is_null() & (c('niv') == 'par')).then(pl.lit('ip'))
                   .when(c('com_prev').is_null() & (c('niv') == '1b')).then(pl.lit('ib'))
                   .when(c('com_prev').is_null()).then(pl.lit('sr'))
                   .when(c('com_prev') // 1000 == 10).then(pl.lit('ll')).otherwise(pl.lit('or')))
assert s.height == u.height == 9961

grafico = [[ini] + [s.filter((c('inicio') == ini) & (c('origen') == k)).height for k in CAT] for ini in ANIOS]
llegada_curso = [[ini] + [s.filter((c('inicio') == ini) & (c('curso') == k)).height for k in range(1, 13)] for ini in ANIOS]

# ---- 2. de dónde viene cada año quien está en 1° básico en Puerto Varas (situación el año anterior)
procedencia = []
for t in ANIOS[1:]:
    f = d[t].filter((c('COD_COM_RBD') == PV) & (c('curso') == 1)).select('MRUN').unique()
    p = una_fila(d[t - 1]).select('MRUN', c('COD_COM_RBD').alias('cp'), c('COD_ENSE').alias('ep'), c('curso').alias('kp'))
    j = f.join(p, on='MRUN', how='left')
    n = lambda e: j.filter(e).height
    nac = d[t].filter(c('curso') == 1).select('MRUN').unique()
    nac_sr = nac.join(d[t - 1].select('MRUN').unique(), on='MRUN', how='anti').height
    procedencia.append([t, f.height,
                        n(c('ep').is_in([10, 214]) & (c('cp') == PV)),
                        n(c('ep').is_in([10, 214]) & (c('cp') != PV)),
                        n(c('kp') == 1),
                        n(c('cp').is_null()),
                        n(c('cp').is_not_null() & ~c('ep').is_in([10, 214]) & (c('kp').fill_null(0) != 1)),
                        round(100 * nac_sr / nac.height, 1)])

# ---- 3. seguimiento de cada cohorte de 1° básico de Puerto Varas hasta 2025
ult = una_fila(d[2025]).select('MRUN', c('COD_COM_RBD').alias('com25'), 'regular')
seguimiento = []
for t in ANIOS[1:]:
    f = d[t].filter((c('COD_COM_RBD') == PV) & (c('curso') == 1)).select('MRUN').unique()
    j = f.join(ult, on='MRUN', how='left')
    k = f.join(s, on='MRUN', how='inner')
    seguimiento.append([t, f.height, j.filter(c('com25') == PV).height, k.height,
                        j.filter(c('com25').is_not_null() & (c('com25') != PV)).height, j.filter(c('com25').is_null()).height]
                       + [k.filter(c('origen') == o).height for o in CAT]
                       + [k.filter(c('inicio') == t).height])

# ---- 4. detalle de la cohorte 1° básico 2018: en qué barra del gráfico 2025 quedó
f18 = d[2018].filter((c('COD_COM_RBD') == PV) & (c('curso') == 1)).select('MRUN').unique()
det18 = (f18.join(s, on='MRUN').group_by('inicio', 'origen').len().sort('inicio', 'origen')
         .with_columns(c('origen').replace_strict(CAT)).rows())

out = {
    'categorias': list(CAT.values()),
    'grafico_pv_2025': grafico,
    'llegada_por_curso_2025': llegada_curso,
    'procedencia_1basico': procedencia,
    'seguimiento_cohortes_1basico': seguimiento,
    'detalle_cohorte_2018': det18,
}
json.dump(out, open(SALIDA, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
print('JSON:', SALIDA)
for k in ['grafico_pv_2025', 'procedencia_1basico', 'seguimiento_cohortes_1basico', 'detalle_cohorte_2018']:
    print(k)
    for r in out[k]:
        print('  ', r)
