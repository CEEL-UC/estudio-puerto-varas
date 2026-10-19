# Año de llegada a la comuna de los alumnos matriculados cada año (Región de Los Lagos, 2015-2025).
# Para cada alumno del tramo regular matriculado en la comuna c el año t, el año de llegada es el primer año del
# período continuo más reciente en que estuvo matriculado (en cualquier nivel, incluida parvularia) en un colegio
# de la comuna c. Un hueco de hasta HUECO años sin registro en ningún colegio de Chile, con el alumno en la misma
# comuna antes y después, no corta el período (casi siempre es un colegio que no informó a todos sus alumnos un año).
# El origen es el último lugar donde aparece antes de su llegada, aunque sea varios años antes:
#   otra comuna de Los Lagos / otra región / sin registro previo (no aparece antes en ningún colegio de Chile, o
#   vuelve a la misma comuna tras más de HUECO años fuera del sistema). Quienes no tienen registro previo se
#   separan según el curso con que aparecen por primera vez:
#   comenzó en parvularia o en una escuela de lenguaje (educación especial 214, niños en edad preescolar) en la
#   comuna, ingresó a 1° básico en la comuna, u otro curso (llegó del extranjero, volvió al sistema o su colegio
#   recién empezó a informar a Mineduc).
# Los datos parten en 2009 (el MRUN viene vacío en 2004-2008): quien está en la comuna sin interrupción desde 2009
# queda como "2009 o antes" (origen desconocido). Hasta el 19 oct 2026 partían en 2014.
# Para el total regional se hace lo mismo con la permanencia en la región.
# Usa el caché liviano de los archivos nacionales que arma 03_migracion_regional.py. Escribe datos_cohortes.js
# para el sitio. Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 04_cohortes_llegada.py
import json, os
import polars as pl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
CACHE = os.path.join(BASE, 'Datos', '[merge] Panel Puerto Varas', 'nacional_liviano')
SALIDA = r'C:\Users\gonza\repos\ceel-estudio-puerto-varas\datos_cohortes.js'
NOMBRES = os.path.join(BASE, 'Análisis y metadata', 'Outputs', 'migracion_regional_sitio.json')
ANIOS = list(range(2009, 2026))
REPORTE = list(range(2015, 2026))
PRIMERO = ANIOS[0]
REGION = 10
HUECO = 2  # años sin registro que se rellenan (desde el 19 oct 2026; antes 0)
DEPK = {1: 'M', 5: 'M', 2: 'S', 3: 'P', 4: 'C'}
c = pl.col

COLS = ['MRUN', 'RBD', 'COD_REG_RBD', 'COD_COM_RBD', 'COD_DEPE2', 'COD_ENSE', 'curso', 'regular', 'anio']  # sin NOM_RBD: 17 años nacionales en memoria
datos = {a: pl.read_parquet(os.path.join(CACHE, f'mat_{a}.parquet'), glob=False, columns=COLS) for a in ANIOS}

# alumnos del tramo regular en la región cada año t (universo del gráfico), con su comuna y dependencia
univ = pl.concat([
    datos[t].filter((c('COD_REG_RBD') == REGION) & c('regular') & c('MRUN').is_not_null())
    .sort(['MRUN', 'RBD']).unique('MRUN', keep='first')
    .select('MRUN', c('anio').cast(pl.Int32), c('COD_COM_RBD').alias('com'), c('COD_DEPE2').replace_strict(DEPK, default='C', return_dtype=pl.Utf8).alias('dep'))
    for t in REPORTE])
mruns = univ['MRUN'].unique()

# ubicación de esos alumnos cada año 2009-2025 (una comuna por año; si tiene varias matrículas se prefiere el tramo regular)
loc = pl.concat([
    datos[a].filter(c('MRUN').is_in(mruns)).sort(['MRUN', 'regular'], descending=[False, True]).unique('MRUN', keep='first')
    .select('MRUN', pl.lit(a, pl.Int32).alias('anio'), c('COD_COM_RBD').alias('com'),
            pl.when(c('COD_ENSE').is_in([10, 214])).then(pl.lit('par')).when(c('curso') == 1).then(pl.lit('1b')).otherwise(pl.lit('otro')).alias('niv'))
    for a in ANIOS]).sort(['MRUN', 'anio'])


def llegadas(loc, clave):
    # tramos por alumno según 'clave' (la comuna, o la pertenencia a la región), rellenando huecos de hasta HUECO
    # años; cada fila recibe el año de inicio de su tramo, la comuna del último registro anterior a ese inicio y
    # el curso con que llegó. loc viene ordenado por MRUN y año, así que la fila anterior es el último registro
    l = loc.with_columns(k=clave)
    mismo = c('MRUN') == c('MRUN').shift(1)
    nuevo = ~mismo | (c('anio') - c('anio').shift(1) > HUECO + 1) | (c('k') != c('k').shift(1))
    l = l.with_columns(tramo=nuevo.fill_null(True).cum_sum(), com_ant=pl.when(mismo).then(c('com').shift(1)))
    l = l.with_columns(inicio=c('anio').first().over('tramo'), com_prev=c('com_ant').first().over('tramo'),
                       niv_ini=c('niv').first().over('tramo'))
    return l.select('MRUN', 'anio', 'inicio', 'com_prev', 'niv_ini')


def clasifica(df, regional):
    # volver a la misma comuna (o a la región) tras un hueco largo no es venir de otra parte: queda sin registro previo
    vuelve = (c('com_prev') // 1000 == REGION) if regional else (c('com_prev') == c('com'))
    df = df.with_columns(com_prev=pl.when(vuelve).then(None).otherwise(c('com_prev')))
    origen = (pl.when(c('inicio') == PRIMERO).then(pl.lit('ce'))
              .when(c('com_prev').is_null() & (c('niv_ini') == 'par')).then(pl.lit('ip'))
              .when(c('com_prev').is_null() & (c('niv_ini') == '1b')).then(pl.lit('ib'))
              .when(c('com_prev').is_null()).then(pl.lit('sr'))
              .when(c('com_prev') // 1000 == REGION).then(pl.lit('ll'))
              .otherwise(pl.lit('or')))
    return df.with_columns(origen=origen)


# por comuna
com = clasifica(univ.join(llegadas(loc, c('com')), on=['MRUN', 'anio'], how='left'), False)
# total regional: tramos de permanencia en la región (quien llegó desde otra comuna de Los Lagos no "llega" a la región)
reg = clasifica(univ.join(llegadas(loc, (c('com') // 1000 == REGION)), on=['MRUN', 'anio'], how='left'), True)
assert com['inicio'].null_count() == 0 and reg['inicio'].null_count() == 0
assert reg.filter(c('origen') == 'll').height == 0

nombres = json.load(open(NOMBRES, encoding='utf8'))
NOM = {x['cod']: x['n'] for x in nombres['comunas']}
comunasOrigen = nombres['comunasOrigen']

ORD = ['ll', 'or', 'sr', 'ce', 'ib', 'ip']
coh, orig = {}, {}


def agrega(df, etiqueta_de):
    g = df.group_by(['etq', 'anio', 'dep', 'inicio', 'origen']).len()
    for (etq, anio, dep), sub in g.group_by(['etq', 'anio', 'dep']):
        filas = {}
        for _, _, _, ini, o, n in sub.rows():
            filas.setdefault(ini, [ini, 0, 0, 0, 0, 0, 0])[1 + ORD.index(o)] += n
        coh.setdefault(etq, {}).setdefault(str(anio), {})[dep] = sorted(filas.values())
    o = df.filter(c('origen').is_in(['ll', 'or'])).group_by(['etq', 'anio', 'dep', 'inicio', 'com_prev']).len()
    for (etq, anio, dep, ini), sub in o.group_by(['etq', 'anio', 'dep', 'inicio']):
        lista = sorted(([x, n] for *_, x, n in sub.rows()), key=lambda p: -p[1])
        orig.setdefault(etq, {}).setdefault(str(anio), {}).setdefault(dep, {})[str(ini)] = lista


agrega(com.with_columns(etq=c('com').replace_strict(NOM, default=None, return_dtype=pl.Utf8)), None)
agrega(reg.with_columns(etq=pl.lit('Región de Los Lagos')), None)

# chequeo: por comuna y año, la suma de las cohortes debe ser la matrícula del tramo regular
flujo = nombres['flujo']
for nom in flujo:
    for i, t in enumerate(REPORTE):
        tot = sum(sum(r[1:]) for dep in coh[nom][str(t)].values() for r in dep)
        mat = sum(flujo[nom][k][i] for k in ['Se mantiene', 'Entrante de otra comuna de Los Lagos', 'Entrante de otra región', 'Sin registro el año anterior', 'Ingreso 1° básico'])
        assert tot == mat, (nom, t, tot, mat)

with open(SALIDA, 'w', encoding='utf8') as f:
    f.write('var COHORTES=' + json.dumps(coh, ensure_ascii=False, separators=(',', ':')) + ';\n')
    f.write('var COHORTES_ORIGEN=' + json.dumps(orig, ensure_ascii=False, separators=(',', ':')) + ';\n')
print('JSON:', SALIDA, round(os.path.getsize(SALIDA) / 1e6, 2), 'MB')

for etq in ['Puerto Varas', 'Región de Los Lagos']:
    filas = {}
    for dep in coh[etq]['2025'].values():
        for r in dep:
            filas.setdefault(r[0], [0, 0, 0, 0, 0, 0])
            for j in range(6):
                filas[r[0]][j] += r[1 + j]
    tot = sum(sum(v) for v in filas.values())
    print(etq, '2025, total', tot)
    for ini in sorted(filas):
        v = filas[ini]
        print(f'   {ini}: {sum(v):5d} ({100 * sum(v) / tot:4.1f}%)  LL {v[0]:4d}  otra reg {v[1]:4d}  parvularia {v[5]:4d}  1°básico {v[4]:4d}  otros s/reg {v[2]:4d}  ≤{PRIMERO} {v[3]:4d}')
