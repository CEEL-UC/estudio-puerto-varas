# Región Metropolitana para el sitio (16 nov 2026): mismas series que la Región de Los Lagos, con las mismas
# definiciones de 03_migracion_regional.py (procedencia y salidas), 04_cohortes_llegada.py (año de llegada),
# 05_matricula_niveles.py (matrícula y colegios por nivel) y del mapa de colegios (pines, balance, top de colegios
# y comunas de origen y destino de cada colegio). La "región" de las categorías pasa a ser la RM: "Entrante de
# otra comuna de la RM", "Salida a otra comuna de la RM", "Llegó desde otra comuna de la RM".
#
# Chicureo no es comuna: es el distrito censal 2 de Colina (INE, cartografía del Censo 2024). Se trata como una
# unidad más, sin sacarlo de Colina: un colegio es de Chicureo si su ubicación (mediana de las coordenadas del
# Directorio 2015-2025) cae dentro del polígono del distrito. Para Chicureo, "se mantiene" es seguir en un colegio de
# Chicureo, y quien llega desde el resto de Colina cuenta como entrante de otra comuna de la RM (en las tablas de
# origen aparece como Colina).
#
# Escribe en el repo del sitio:
#   datos_rm.js            matrícula y colegios por nivel (secciones 01 y 02) y procedencia y salidas (sección 03)
#   rm/cohortes_<CUT>.js   año de llegada (sección 03) de cada comuna (total regional: rm/cohortes_region.js)
#   rm/colegios_<CUT>.js   mapa de colegios de cada comuna (Chicureo: rm/colegios_chicureo.js), se carga al elegirla
#
# Validación: con --validar 10 corre todo para la Región de Los Lagos sin escribir nada y compara contra los datos
# publicados (index.html y datos_cohortes.js). Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 11_region_metropolitana.py
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 11_region_metropolitana.py --validar 10
import glob, json, math, os, re, sys, unicodedata
import polars as pl
import pyarrow.parquet as pq
import shapely
from shapely import wkb
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding='utf8')
VALIDAR = '--validar' in sys.argv
REGION = int(sys.argv[sys.argv.index('--validar') + 1]) if VALIDAR else 13

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
CACHE = os.path.join(BASE, 'Datos', '[merge] Panel Puerto Varas', 'nacional_liviano')
ORIGEN = os.path.join(BASE, 'Datos', 'Mineduc', 'Matricula por estudiante')
DIRECTORIO = os.path.join(BASE, 'Datos', 'Mineduc', 'Directorio establecimientos')
DPA = os.path.join(BASE, 'Datos', 'Censo 2024', 'documentacion_manzana', 'diccionario_variables_glosas_censo2024.xlsx')
DISTRITOS_RM = os.path.join(BASE, '..', 'Agenda Pública Consultores', 'Mapa Censo RM', 'Datos', 'cartografia',
                            'Cartografía_censo2024_R13_Distrital.parquet')
REPO = r'C:\Users\gonza\repos\ceel-estudio-puerto-varas'
SITIO_JSON = os.path.join(BASE, 'Análisis y metadata', 'Outputs', 'migracion_regional_sitio.json')

ANIOS_COH = list(range(2009, 2026))   # año de llegada: desde 2009 (antes el MRUN viene vacío)
ANIOS = list(range(2014, 2026))       # procedencia: 2014 sirve para el año anterior de 2015
REPORTE = list(range(2015, 2026))
VAL_SAL = set(REPORTE[:-1])
PRIMERO = ANIOS_COH[0]
HUECO = 2
REG_NOM = {10: 'Región de Los Lagos', 13: 'Región Metropolitana'}[REGION]
REG_CORTO = {10: 'Los Lagos', 13: 'la RM'}[REGION]
CHI_COD, CHI_NOM = 13399, 'Chicureo'   # código ficticio, solo para separar a Chicureo del resto de Colina
c = pl.col

DEPE = {1: 'Municipal', 2: 'Part. Subvencionado', 3: 'Part. Pagado', 4: 'Corp. Adm. Delegada', 5: 'Serv. Local Educ.'}
DEPK = {1: 'M', 5: 'M', 2: 'S', 3: 'P', 4: 'C'}
DEP_MAPA = {1: 'Municipal / SLEP', 5: 'Municipal / SLEP', 2: 'Part. Subvencionado', 3: 'Part. Pagado', 4: 'Adm. delegada'}

ORIG = ['Se mantiene', f'Entrante de otra comuna de {REG_CORTO}', 'Entrante de otra región', 'Sin registro el año anterior', 'Ingreso 1° básico']
DEST = [f'Salida a otra comuna de {REG_CORTO}', 'Salida a otra región', 'Sin registro el año siguiente']

# ---------------------------------------------------------------- nombres de comunas (igual que 03)
wb = load_workbook(DPA, read_only=True)
NOM = {}
for r in list(wb['DPA'].iter_rows(values_only=True))[1:]:
    if r[4]:
        nombre = ' '.join(p if p in ('de', 'del', 'la', 'las', 'los', 'y') else p.capitalize() for p in str(r[5]).strip().lower().split())
        NOM[int(r[4])] = nombre[0].upper() + nombre[1:]
NOM.update({10109: 'Puerto Varas', 10101: 'Puerto Montt', 10105: 'Frutillar', 10301: 'Osorno'})
NOM[CHI_COD] = CHI_NOM
REGNOM = {1: 'Tarapacá', 2: 'Antofagasta', 3: 'Atacama', 4: 'Coquimbo', 5: 'Valparaíso', 6: "O'Higgins", 7: 'Maule',
          8: 'Biobío', 9: 'La Araucanía', 10: 'Los Lagos', 11: 'Aysén', 12: 'Magallanes', 13: 'Metropolitana',
          14: 'Los Ríos', 15: 'Arica y Parinacota', 16: 'Ñuble'}
SITIO = json.load(open(SITIO_JSON, encoding='utf8'))
ORIGEN_PUBLICADO = SITIO['comunasOrigen']   # incluye los códigos antiguos de Ñuble (antes de 2018) ya rotulados


REG_VIEJO = {}


def nombre_origen(x):
    if str(x) in ORIGEN_PUBLICADO:
        return ORIGEN_PUBLICADO[str(x)]
    return [NOM.get(x, f'Comuna {x}'), REG_VIEJO.get(x) or REGNOM.get(x // 1000, f'Región {x // 1000}')]


# ---------------------------------------------------------------- caché nacional (lo arma 03)
COLS = ['MRUN', 'RBD', 'NOM_RBD', 'COD_REG_RBD', 'COD_COM_RBD', 'COD_DEPE2', 'COD_ENSE', 'curso', 'regular', 'anio']
datos = {a: pl.read_parquet(os.path.join(CACHE, f'mat_{a}.parquet'), glob=False, columns=COLS) for a in ANIOS_COH}
print('caché nacional cargado:', {a: datos[a].height for a in (2009, 2025)})

# códigos de comuna que ya no existen: los de Ñuble antes de 2018 (84xx) y los de Valparaíso antes de crearse la
# provincia de Marga Marga en 2010 (Quilpué 5106, Villa Alemana 5108, Limache 5505, Olmué 5507). Se rotulan con el
# nombre y la región actuales de la misma comuna, buscando su nombre en los CSV de Mineduc de 2009 y 2017.
viejos = set()
for a in ANIOS_COH:
    viejos |= set(datos[a]['COD_COM_RBD'].drop_nulls().unique().to_list())
viejos = {x for x in viejos if x not in NOM}
plano = lambda t: unicodedata.normalize('NFKD', t).encode('ascii', 'ignore').decode().lower().strip()
actuales = {}
for k, v in NOM.items():
    actuales.setdefault(plano(v), []).append(k)
for anio in (2009, 2017):
    if not viejos - set(REG_VIEJO):
        break
    f = [x for x in glob.glob(os.path.join(ORIGEN, str(anio), '*')) if x.lower().endswith('.csv')][0]
    hdr = [x.strip().upper() for x in open(f, encoding='latin-1').readline().strip().split(';')]
    n = pl.read_csv(f, separator=';', encoding='utf8-lossy', infer_schema=False, quote_char=None,
                    columns=[hdr.index('COD_COM_RBD'), hdr.index('NOM_COM_RBD')], new_columns=['cod', 'nom']).unique()
    for cod, nom in n.rows():
        cod = int(cod)
        if cod in viejos and cod not in REG_VIEJO:
            # misma comuna con su código actual; entre homónimas, la de la región que tenía entonces (Ñuble para 84xx)
            cand = sorted(actuales.get(plano(nom), []), key=lambda k: (k // 1000 != (16 if cod // 100 == 84 else cod // 1000), k))
            if cand:
                NOM[cod], REG_VIEJO[cod] = NOM[cand[0]], REGNOM[cand[0] // 1000]
            else:
                NOM[cod], REG_VIEJO[cod] = nom.strip().title(), REGNOM.get(16 if cod // 100 == 84 else cod // 1000)
print('códigos antiguos rotulados:', {k: (NOM[k], REG_VIEJO[k]) for k in sorted(REG_VIEJO)}, '| sin rotular:', viejos - set(REG_VIEJO))


# ---------------------------------------------------------------- coordenadas del Directorio, por colegio y año
def leer_directorio(anio):
    carpeta = [x for x in glob.glob(os.path.join(DIRECTORIO, '*')) if x.endswith(str(anio)) or f'-{anio}-' in x or x.endswith(f'{anio}-')][0]
    f = [x for x in glob.glob(os.path.join(carpeta, '**', '*'), recursive=True) if x.lower().endswith('.csv')][0]
    linea = open(f, encoding='latin-1').readline()
    sep = ';' if linea.count(';') > linea.count(',') else ','
    d = pl.read_csv(f, separator=sep, encoding='utf8-lossy', infer_schema=False, quote_char='"')
    d.columns = [re.sub(r'[^A-Z_0-9]', '', x.strip().upper()) for x in d.columns]
    if not d.columns[0].endswith('AGNO') and 'AGNO' not in d.columns:
        pass
    num = lambda x: c(x).str.strip_chars().str.replace(',', '.').cast(pl.Float64, strict=False)
    d = d.select(c('RBD').str.strip_chars().cast(pl.Int64, strict=False).alias('RBD'),
                 c('COD_COM_RBD').str.strip_chars().cast(pl.Int64, strict=False).alias('com'),
                 num('LATITUD').alias('lat'), num('LONGITUD').alias('lon'))
    # fuera de Chile continental = sin coordenada
    ok = c('lat').is_between(-56, -17) & c('lon').is_between(-76, -66)
    return d.with_columns(lat=pl.when(ok).then(c('lat')), lon=pl.when(ok).then(c('lon')), anio=pl.lit(anio))


coords = pl.concat([leer_directorio(a) for a in REPORTE]).filter(c('com') // 1000 == REGION)


def km(a, b):
    return 6371 * math.hypot(math.radians(a[1] - b[1]) * math.cos(math.radians(a[0])), math.radians(a[0] - b[0]))


# corrección de coordenadas erróneas de un año (mismo criterio del 21 sep 2026): un salto aislado de un año (el año
# anterior y el siguiente existen, son parecidos entre sí, y ese año se aleja de ambos) o, en 2025, un salto
# respecto de 3 años previos estables. Se reemplaza por la coordenada del año anterior.
COORD = {}
corregidas = []
for (rbd,), g in coords.filter(c('lat').is_not_null()).group_by(['RBD']):
    p = {a: (la, lo) for a, la, lo in g.select('anio', 'lat', 'lon').rows()}
    for a in sorted(p):
        if a - 1 in p and a + 1 in p and km(p[a - 1], p[a + 1]) < 0.3 and km(p[a], p[a - 1]) > 0.6 and km(p[a], p[a + 1]) > 0.6:
            p[a] = p[a - 1]; corregidas.append((rbd, a))
        elif a == 2025 and all(x in p for x in (2022, 2023, 2024)) and km(p[2022], p[2023]) < 0.3 and km(p[2023], p[2024]) < 0.3 \
                and km(p[2025], p[2024]) > 0.6:
            p[a] = p[2024]; corregidas.append((rbd, a))
    # años sin coordenada en el Directorio: la del año más cercano que la tenga (igual que el mapa de Los Lagos)
    for a in REPORTE:
        if a not in p:
            p[a] = p[min(p, key=lambda x: (abs(x - a), -x))]
    COORD[rbd] = p
print('colegios con coordenadas:', len(COORD), '| coordenadas corregidas:', len(corregidas))

# ---------------------------------------------------------------- Chicureo: colegios dentro del distrito censal
CHI = set()
if REGION == 13:
    dist = [r for r in pq.read_table(DISTRITOS_RM).to_pylist() if str(r['CUT']) == '13301' and r['DISTRITO'] == 'CHICUREO']
    pol = wkb.loads(dist[0]['SHAPE'])
    colina = set(coords.filter(c('com') == 13301)['RBD'].to_list()) | set(datos[2025].filter(c('COD_COM_RBD') == 13301)['RBD'].to_list())
    for rbd in colina:
        if rbd in COORD:
            la = sorted(x[0] for x in COORD[rbd].values())[len(COORD[rbd]) // 2]
            lo = sorted(x[1] for x in COORD[rbd].values())[len(COORD[rbd]) // 2]
            if pol.contains(shapely.Point(lo, la)):
                CHI.add(rbd)
    nom25 = dict(datos[2025].filter(c('RBD').is_in(list(CHI))).select('RBD', 'NOM_RBD').unique('RBD').rows())
    print('Chicureo:', len(CHI), 'colegios ->', sorted(nom25.get(x, f'RBD {x} (sin matrícula 2025)') for x in CHI))
CHI_L = list(CHI)


def unidad(recode):
    # comuna del colegio, o Chicureo si el colegio está en el distrito (solo en la pasada de Chicureo)
    return pl.when(c('RBD').is_in(CHI_L)).then(pl.lit(CHI_COD, pl.Int64)).otherwise(c('COD_COM_RBD')) if recode else c('COD_COM_RBD')


# ================================================================ 05: matrícula y colegios por nivel
NIVEL = (pl.when(c('COD_ENSE') == 10).then(pl.lit('par'))
         .when((c('COD_ENSE') == 110) | c('COD_ENSE').is_in([310, 410, 510, 610, 710, 810, 910])).then(pl.lit('reg'))
         .when(c('COD_ENSE').is_between(200, 299)).then(pl.lit('esp'))
         .otherwise(pl.lit('adu')))
DEP = c('COD_DEPE2').replace_strict(DEPK, default='C', return_dtype=pl.Utf8)
NIVS, DEPS = ['reg', 'par', 'esp', 'adu'], ['M', 'S', 'P', 'C']
BIT = {'reg': 1, 'par': 2, 'esp': 4, 'adu': 8}


def rural(anio):
    f = [x for x in glob.glob(os.path.join(ORIGEN, str(anio), '*')) if x.lower().endswith('.csv')][0]
    hdr = [x.strip().upper() for x in open(f, encoding='latin-1').readline().strip().split(';')]
    hdr[0] = 'AGNO' if hdr[0].endswith('AGNO') else hdr[0]
    r = pl.read_csv(f, separator=';', encoding='utf8-lossy', infer_schema=False, quote_char=None,
                    columns=[hdr.index('RBD'), hdr.index('COD_REG_RBD'), hdr.index('RURAL_RBD')], new_columns=['RBD', 'reg', 'rural'])
    r = r.with_columns([c(x).str.strip_chars().cast(pl.Int64, strict=False) for x in r.columns])
    return r.filter(c('reg') == REGION).group_by('RBD').agg(c('rural').max())


mat, est, RURAL = {}, {}, {}
for i, a in enumerate(REPORTE):
    d = datos[a].filter(c('COD_REG_RBD') == REGION).with_columns(niv=NIVEL, dep=DEP)
    for cod, chi, dep, niv, n in d.with_columns(chi=c('RBD').is_in(CHI_L)).group_by(['COD_COM_RBD', 'chi', 'dep', 'niv']).len().rows():
        for etq in [NOM[cod], REG_NOM] + ([CHI_NOM] if chi else []):
            mat.setdefault(etq, {}).setdefault(dep, {}).setdefault(niv, [0] * len(REPORTE))[i] += n
    ru = rural(a)
    RURAL[a] = dict(ru.rows())
    colegios = d.group_by(['RBD', 'COD_COM_RBD', 'dep']).agg(c('niv').unique()).join(ru, on='RBD', how='left')
    for rbd, cod, dep, nivs, rur in colegios.rows():
        fila = [dep, int(rur or 0), sum(BIT[x] for x in nivs)]
        for etq in [NOM[cod], REG_NOM] + ([CHI_NOM] if rbd in CHI else []):
            est.setdefault(etq, {}).setdefault(str(a), []).append(fila)
    print(a, 'matrículas región', d.height, '| colegios', colegios.height, '| sin ruralidad', colegios['rural'].null_count())
for etq in mat:
    for dep in DEPS:
        for niv in NIVS:
            mat[etq].setdefault(dep, {}).setdefault(niv, [0] * len(REPORTE))
for etq in est:
    for a in est[etq]:
        cnt = {}
        for f in est[etq][a]:
            cnt[tuple(f)] = cnt.get(tuple(f), 0) + 1
        est[etq][a] = [list(k) + [n] for k, n in sorted(cnt.items())]


# ================================================================ 03: procedencia y salidas
def ubicacion(a, recode):
    # una unidad por alumno en el año a: si tiene varias matrículas se prefiere la del tramo regular
    return (datos[a].filter(c('MRUN').is_not_null()).sort(['MRUN', 'regular'], descending=[False, True])
            .unique('MRUN', keep='first').select('MRUN', unidad(recode).alias('u')))


def alumnos(recode):
    filas = []
    for t in REPORTE:
        reg = (datos[t].filter((c('COD_REG_RBD') == REGION) & c('regular') & c('MRUN').is_not_null())
               .sort(['MRUN', 'RBD']).unique('MRUN', keep='first').with_columns(u=unidad(recode)))
        if recode:
            reg = reg.filter(c('u') == CHI_COD)
        reg = reg.join(ubicacion(t - 1, recode).rename({'u': 'u_ant'}), on='MRUN', how='left')
        reg = reg.join(ubicacion(t + 1, recode).rename({'u': 'u_sig'}), on='MRUN', how='left') if t + 1 in datos \
            else reg.with_columns(u_sig=pl.lit(None, pl.Int64))
        filas.append(reg)
    al = pl.concat(filas, how='diagonal_relaxed')
    return al.with_columns(
        origen=pl.when(c('curso') == 1).then(pl.lit(ORIG[4]))
        .when(c('u_ant').is_null()).then(pl.lit(ORIG[3]))
        .when(c('u_ant') == c('u')).then(pl.lit(ORIG[0]))
        .when(c('u_ant') // 1000 == REGION).then(pl.lit(ORIG[1]))
        .otherwise(pl.lit(ORIG[2])),
        destino=pl.when((c('anio') == max(REPORTE)) | (c('curso') == 12)).then(None)
        .when(c('u_sig').is_null()).then(pl.lit(DEST[2]))
        .when(c('u_sig') == c('u')).then(pl.lit('Sigue en la comuna'))
        .when(c('u_sig') // 1000 == REGION).then(pl.lit(DEST[0]))
        .otherwise(pl.lit(DEST[1])),
        depe=c('COD_DEPE2').replace_strict(DEPE, default=None, return_dtype=pl.Utf8),
        dk=c('COD_DEPE2').replace_strict(DEPK, default='C', return_dtype=pl.Utf8))


def serie(df, col, cats, anios_validos=None):
    t = df.group_by(['anio', col]).len()
    out = {}
    for k in cats:
        v = {r[0]: r[2] for r in t.filter(c(col) == k).rows()}
        out[k] = [(v.get(a, 0) if (anios_validos is None or a in anios_validos) else None) for a in REPORTE]
    return out


def suma_sal(f):
    return [None if f[DEST[0]][i] is None else sum(f[k][i] for k in DEST) for i in range(len(REPORTE))]


al_com = alumnos(False)
al_chi = alumnos(True) if CHI else None
print('alumnos tramo regular región (filas alumno-año):', al_com.height, '| Chicureo:', al_chi.height if al_chi is not None else 0)

migFlow, migFlowDepe = {}, {}
for sub_al, unidades in [(al_com, sorted(al_com['u'].unique().to_list())), (al_chi, [CHI_COD] if CHI else [])]:
    for u in unidades:
        nom = NOM.get(u, str(u))
        sub = sub_al.filter(c('u') == u)
        f = serie(sub, 'origen', ORIG)
        f.update(serie(sub, 'destino', DEST, VAL_SAL))
        f['Salidas'] = suma_sal(f)
        migFlow[nom] = f
        migFlowDepe[nom] = {}
        for dk in DEPS:
            g = serie(sub.filter(c('dk') == dk), 'origen', ORIG)
            g.update(serie(sub.filter(c('dk') == dk), 'destino', DEST, VAL_SAL))
            g['Salidas'] = suma_sal(g)
            for k in DEST:
                del g[k]
            migFlowDepe[nom][dk] = g

# ================================================================ mapa de colegios
MAPA_UNIDADES = sorted(al_com['u'].unique().to_list()) + ([CHI_COD] if CHI else [])
if VALIDAR:
    MAPA_UNIDADES = [10109, 10101, 10105, 10301]


def top5(df, col, filtro):
    out = {}
    t = df.filter(filtro).group_by(['RBD', 'anio', col]).len().sort(['RBD', 'anio', 'len', col], descending=[False, False, True, False])
    for (rbd, anio), g in t.group_by(['RBD', 'anio'], maintain_order=True):
        out.setdefault(str(rbd), {})[str(anio)] = [{'c': NOM.get(x, str(x)), 'n': n} for x, n in g.select(col, 'len').rows()[:5]]
    return out


# como en el mapa de Los Lagos: un colegio va como "regular" (r = 1, visible por defecto) si tiene básica regular o
# cualquier enseñanza media (incluida la de adultos); si no, r = 0 y su tipo es especial, parvularia o adultos
# (básica de adultos). Solo la educación especial tiene botón para mostrarse en el mapa.


MAPA = {}
for u in MAPA_UNIDADES:
    nom = NOM[u]
    al = al_chi if u == CHI_COD else al_com
    sub = al.filter(c('u') == u)
    cnt = {(r['RBD'], r['anio']): r for r in sub.group_by(['RBD', 'anio']).agg(
        e=c('origen').is_in(ORIG[1:4]).sum(), i=(c('origen') == ORIG[4]).sum(),
        k=(c('origen') == ORIG[0]).sum(), s=c('destino').is_in(DEST).sum()).rows(named=True)}
    geo, topE, topS = {}, {}, {}
    for a in REPORTE:
        d = datos[a].filter(c('COD_REG_RBD') == REGION)
        d = d.filter(c('RBD').is_in(CHI_L)) if u == CHI_COD else d.filter(c('COD_COM_RBD') == u)
        lista = []
        for rbd, nombre, dep, m, regular, nivs in d.with_columns(niv=NIVEL).group_by('RBD').agg(
                c('NOM_RBD').first(), c('COD_DEPE2').mode().first(), pl.len(), ((c('COD_ENSE') == 110) | (c('COD_ENSE') >= 300)).any(),
                c('niv').unique()).sort('RBD').rows():
            p = COORD.get(rbd, {}).get(a)
            x = cnt.get((rbd, a), {})
            fila = {'rbd': rbd, 'n': (nombre or '').strip().title(), 'd': DEP_MAPA.get(dep, 'Adm. delegada'),
                    'lat': round(p[0], 5) if p else None, 'lon': round(p[1], 5) if p else None, 'm': m,
                    'r': 1 if regular else 0}
            if not regular:
                fila['t'] = 'Especial' if 'esp' in nivs else ('Parvularia' if 'par' in nivs else 'Adultos')
            fila.update(e=x.get('e', 0), i=x.get('i', 0), k=x.get('k', 0), s=(x.get('s', 0) if a in VAL_SAL else None))
            lista.append(fila)
        geo[str(a)] = lista
        con = [f for f in lista if f['r'] == 1]
        topE[str(a)] = [{'rbd': f['rbd'], 'n': f['n'], 'cnt': f['e'], 'mat': f['m'], 'pct': round(100 * f['e'] / f['m'], 1)}
                        for f in sorted(con, key=lambda f: (-f['e'], f['rbd'])) if f['e'] > 0][:5]
        topS[str(a)] = [] if a not in VAL_SAL else [{'rbd': f['rbd'], 'n': f['n'], 'cnt': f['s'], 'mat': f['m'], 'pct': round(100 * f['s'] / f['m'], 1)}
                            for f in sorted(con, key=lambda f: (-f['s'], f['rbd'])) if f['s'] > 0][:5]
    fd = migFlowDepe[nom]
    porDep = {}
    for dk, etq in [('M', 'Municipal / SLEP'), ('S', 'Part. Subvencionado'), ('P', 'Part. Pagado'), ('C', 'Adm. delegada')]:
        g = fd[dk]
        porDep[etq] = {'Mantiene': g[ORIG[0]], 'Entrante': [g[ORIG[1]][i] + g[ORIG[2]][i] + g[ORIG[3]][i] for i in range(len(REPORTE))],
                       'Ingreso 1° básico': g[ORIG[4]], 'Saliente': g['Salidas']}
    MAPA[nom] = {'geo': geo, 'topE': topE, 'topS': topS, 'porDep': porDep,
                 'orig': top5(sub, 'u_ant', c('origen').is_in(ORIG[1:3])),
                 'dest': top5(sub, 'u_sig', c('destino').is_in(DEST[:2]))}
    print('mapa', nom, '| colegios 2025:', len(geo['2025']), '| con coordenadas:', sum(1 for f in geo['2025'] if f['lat'] is not None))

# ================================================================ 04: año de llegada
COH_ORD = ['ll', 'or', 'sr', 'ce', 'ib', 'ip']


def cohortes(recode):
    univ = pl.concat([
        datos[t].filter((c('COD_REG_RBD') == REGION) & c('regular') & c('MRUN').is_not_null())
        .sort(['MRUN', 'RBD']).unique('MRUN', keep='first')
        .select('MRUN', c('anio').cast(pl.Int32), unidad(recode).alias('com'), c('COD_DEPE2').replace_strict(DEPK, default='C', return_dtype=pl.Utf8).alias('dep'))
        for t in REPORTE])
    if recode:
        univ = univ.filter(c('com') == CHI_COD)
    mruns = univ['MRUN'].unique()
    loc = pl.concat([
        datos[a].filter(c('MRUN').is_in(mruns)).sort(['MRUN', 'regular'], descending=[False, True]).unique('MRUN', keep='first')
        .select('MRUN', pl.lit(a, pl.Int32).alias('anio'), unidad(recode).alias('com'),
                pl.when(c('COD_ENSE').is_in([10, 214])).then(pl.lit('par')).when(c('curso') == 1).then(pl.lit('1b')).otherwise(pl.lit('otro')).alias('niv'))
        for a in ANIOS_COH]).sort(['MRUN', 'anio'])

    def llegadas(clave):
        l = loc.with_columns(k=clave)
        mismo = c('MRUN') == c('MRUN').shift(1)
        nuevo = ~mismo | (c('anio') - c('anio').shift(1) > HUECO + 1) | (c('k') != c('k').shift(1))
        l = l.with_columns(tramo=nuevo.fill_null(True).cum_sum(), com_ant=pl.when(mismo).then(c('com').shift(1)))
        l = l.with_columns(inicio=c('anio').first().over('tramo'), com_prev=c('com_ant').first().over('tramo'),
                           niv_ini=c('niv').first().over('tramo'))
        return l.select('MRUN', 'anio', 'inicio', 'com_prev', 'niv_ini')

    def clasifica(df, regional):
        vuelve = (c('com_prev') // 1000 == REGION) if regional else (c('com_prev') == c('com'))
        df = df.with_columns(com_prev=pl.when(vuelve).then(None).otherwise(c('com_prev')))
        origen = (pl.when(c('inicio') == PRIMERO).then(pl.lit('ce'))
                  .when(c('com_prev').is_null() & (c('niv_ini') == 'par')).then(pl.lit('ip'))
                  .when(c('com_prev').is_null() & (c('niv_ini') == '1b')).then(pl.lit('ib'))
                  .when(c('com_prev').is_null()).then(pl.lit('sr'))
                  .when(c('com_prev') // 1000 == REGION).then(pl.lit('ll'))
                  .otherwise(pl.lit('or')))
        return df.with_columns(origen=origen)

    com = clasifica(univ.join(llegadas(c('com')), on=['MRUN', 'anio'], how='left'), False)
    reg = None if recode else clasifica(univ.join(llegadas(c('com') // 1000 == REGION), on=['MRUN', 'anio'], how='left'), True)
    assert com['inicio'].null_count() == 0
    return com, reg


coh, orig = {}, {}


def agrega(df):
    g = df.group_by(['etq', 'anio', 'dep', 'inicio', 'origen']).len()
    for (etq, anio, dep), sub in g.group_by(['etq', 'anio', 'dep']):
        filas = {}
        for _, _, _, ini, o, n in sub.rows():
            filas.setdefault(ini, [ini, 0, 0, 0, 0, 0, 0])[1 + COH_ORD.index(o)] += n
        coh.setdefault(etq, {}).setdefault(str(anio), {})[dep] = sorted(filas.values())
    o = df.filter(c('origen').is_in(['ll', 'or'])).group_by(['etq', 'anio', 'dep', 'inicio', 'com_prev']).len()
    for (etq, anio, dep, ini), sub in o.group_by(['etq', 'anio', 'dep', 'inicio']):
        orig.setdefault(etq, {}).setdefault(str(anio), {}).setdefault(dep, {})[str(ini)] = sorted(([x, n] for *_, x, n in sub.rows()), key=lambda p: (-p[1], p[0]))


com_c, reg_c = cohortes(False)
agrega(com_c.with_columns(etq=c('com').replace_strict(NOM, default=None, return_dtype=pl.Utf8)))
agrega(reg_c.with_columns(etq=pl.lit(REG_NOM)))
if CHI:
    chi_c, _ = cohortes(True)
    agrega(chi_c.with_columns(etq=pl.lit(CHI_NOM)))
# chequeo: la suma de las cohortes es la matrícula del tramo regular de la procedencia
for nom in migFlow:
    for i, t in enumerate(REPORTE):
        tot = sum(sum(r[1:]) for dep in coh[nom][str(t)].values() for r in dep)
        assert tot == sum(migFlow[nom][k][i] for k in ORIG), (nom, t)
print('cohortes: la suma calza con la procedencia en', len(migFlow), 'unidades')

# nombres de todas las comunas que aparecen como origen o destino
cods = set()
for v in orig.values():
    for a in v.values():
        for dd in a.values():
            for lst in dd.values():
                cods.update(x for x, _ in lst)
# (más todos los códigos antiguos: el sitio junta esta lista con la de Los Lagos, a la que le faltaban los de Valparaíso)
comunasOrigen = {str(x): nombre_origen(x) for x in sorted(cods | set(REG_VIEJO))}
sin = [x for x in cods if str(x) not in ORIGEN_PUBLICADO and x not in NOM]
print('comunas de origen:', len(cods), '| sin nombre:', sin)

# ================================================================ validación contra lo publicado (Los Lagos)
if VALIDAR:
    s = open(os.path.join(REPO, 'index.html'), encoding='utf8').read()
    pub = lambda v: json.loads(re.search(r'\nvar ' + v + r' = (.*?);?\n', s).group(1).rstrip(';'))
    fallas = 0

    def igual(nombre, a, b):
        global fallas
        ok = a == b
        fallas += not ok
        print(('OK   ' if ok else 'DIF  ') + nombre)
        if not ok and isinstance(a, dict) and isinstance(b, dict):
            for k in sorted(set(a) | set(b), key=str)[:200]:
                if a.get(k) != b.get(k):
                    print('     ', k, str(a.get(k))[:300], '|', str(b.get(k))[:300]); break

    MATN, ESTN, MF, MFD = pub('MATN'), pub('ESTN'), pub('migFlow'), pub('migFlowDepe')
    igual('MATN', mat, MATN)
    igual('ESTN', est, ESTN)
    igual('migFlow', migFlow, MF)
    igual('migFlowDepe', migFlowDepe, MFD)
    G, TE, TS, OT, DT, FB = pub('colegiosGeo'), pub('topEntrantesEstab'), pub('topSalientesEstab'), pub('estabOrigenTop5'), pub('estabDestinoTop5'), pub('migFlowByDepe')
    for nom in MAPA:
        for a in map(str, REPORTE):
            mine = {f['rbd']: f for f in MAPA[nom]['geo'][a]}
            theirs = {f['rbd']: f for f in G[nom][a]}
            for k in ['m', 'r', 'e', 'i', 'k', 's', 'd', 'lat', 'lon', 'n']:
                dif = [(r, mine[r].get(k), theirs[r].get(k)) for r in theirs if r in mine and mine[r].get(k) != theirs[r].get(k)]
                if dif and k not in ('n',):
                    print(f'DIF  colegiosGeo {nom} {a} {k}: {len(dif)} ej. {dif[:3]}'); fallas += 1
            if set(mine) != set(theirs):
                print(f'DIF  colegiosGeo {nom} {a} rbd: sobran {set(mine) - set(theirs)} faltan {set(theirs) - set(mine)}'); fallas += 1
        igual(f'topEntrantes {nom}', {a: [(x['rbd'], x['cnt'], x['pct']) for x in v] for a, v in MAPA[nom]['topE'].items()},
              {a: [(x['rbd'], x['cnt'], x['pct']) for x in v] for a, v in TE[nom].items()})
        igual(f'topSalientes {nom}', {a: [(x['rbd'], x['cnt'], x['pct']) for x in v] for a, v in MAPA[nom]['topS'].items()},
              {a: [(x['rbd'], x['cnt'], x['pct']) for x in v] for a, v in TS[nom].items()})
        fb = FB[nom]
        suma = {k: [(fb['Municipal'][k][i] or 0) + (fb['Serv. Local Educ.'][k][i] or 0) if fb['Municipal'][k][i] is not None or fb['Serv. Local Educ.'][k][i] is not None else None
                    for i in range(11)] for k in fb['Municipal']}
        igual(f'migFlowByDepe {nom} Municipal/SLEP', MAPA[nom]['porDep']['Municipal / SLEP'], suma)
        igual(f'migFlowByDepe {nom} P', MAPA[nom]['porDep']['Part. Pagado'], fb['Part. Pagado'])
        # top 5 de comunas de origen y destino: se compara el conjunto de cifras (los empates pueden ordenarse distinto)
        for etq, mio, suyo in [('origen', MAPA[nom]['orig'], OT), ('destino', MAPA[nom]['dest'], DT)]:
            d0 = sum(1 for rbd in mio for a in mio[rbd] if sorted(x['n'] for x in mio[rbd][a]) != sorted(x['n'] for x in suyo.get(rbd, {}).get(a, [])))
            print(('OK   ' if d0 == 0 else 'DIF  ') + f'top5 {etq} por colegio {nom}: {d0} colegio-año distintos'); fallas += d0 > 0
    t = open(os.path.join(REPO, 'datos_cohortes.js'), encoding='utf8').read()
    COH = json.loads(re.search(r'var COHORTES=(.*?);\n', t).group(1))
    COHO = json.loads(re.search(r'var COHORTES_ORIGEN=(.*?);\n', t).group(1))
    igual('COHORTES', coh, COH)
    norm = lambda o: {e: {a: {d: {i: sorted(map(tuple, l)) for i, l in v3.items()} for d, v3 in v2.items()} for a, v2 in v1.items()} for e, v1 in o.items()}
    igual('COHORTES_ORIGEN', norm(orig), norm(COHO))
    print('VALIDACIÓN:', 'sin diferencias' if not fallas else f'{fallas} diferencias')
    sys.exit(0)

# ================================================================ salida
UNIDADES = sorted([NOM[u] for u in al_com['u'].unique().to_list()], key=lambda x: unicodedata.normalize('NFKD', x).lower()) + [CHI_NOM]
CUT = {NOM[u]: u for u in al_com['u'].unique().to_list()}
CUT[CHI_NOM] = 'chicureo'
rm = {'region': REG_NOM, 'comunas': UNIDADES, 'cut': CUT, 'MATN': mat, 'ESTN': est, 'migFlow': migFlow,
      'migFlowDepe': migFlowDepe, 'comunasOrigen': comunasOrigen,
      'chicureo': sorted(CHI)}
with open(os.path.join(REPO, 'datos_rm.js'), 'w', encoding='utf8') as f:
    f.write('var RM=' + json.dumps(rm, ensure_ascii=False, separators=(',', ':')) + ';\n')
os.makedirs(os.path.join(REPO, 'rm'), exist_ok=True)
# año de llegada: un archivo por comuna (y uno para el total regional), que el sitio carga al elegirla
CUT[REG_NOM] = 'region'
for nom in coh:
    with open(os.path.join(REPO, 'rm', f'cohortes_{CUT[nom]}.js'), 'w', encoding='utf8') as f:
        f.write('pvCohortesRM(' + json.dumps(nom, ensure_ascii=False) + ',' +
                json.dumps({'coh': coh[nom], 'orig': orig.get(nom, {})}, ensure_ascii=False, separators=(',', ':')) + ');\n')
for nom, v in MAPA.items():
    with open(os.path.join(REPO, 'rm', f'colegios_{CUT[nom]}.js'), 'w', encoding='utf8') as f:
        f.write('pvColegiosRM(' + json.dumps(nom, ensure_ascii=False) + ',' + json.dumps(v, ensure_ascii=False, separators=(',', ':')) + ');\n')
print('datos_rm.js', round(os.path.getsize(os.path.join(REPO, 'datos_rm.js')) / 1e6, 2), 'MB')
for x in ['colegios', 'cohortes']:
    fs = glob.glob(os.path.join(REPO, 'rm', x + '_*.js'))
    print(f'rm/{x}_*.js:', len(fs), 'archivos,', round(sum(map(os.path.getsize, fs)) / 1e6, 1), 'MB, el más pesado',
          round(max(map(os.path.getsize, fs)) / 1e6, 2), 'MB')

# resumen de control
for nom in ['Colina', CHI_NOM, 'Lampa', 'San Bernardo', 'Calera de Tango', 'Padre Hurtado', 'Peñaflor', 'Talagante', REG_NOM]:
    tot = [sum(mat[nom][dp]['reg'][i] for dp in DEPS) for i in range(len(REPORTE))]
    pp = mat[nom]['P']['reg']
    print(f'{nom:20s} tramo regular 2015 {tot[0]:7d} 2025 {tot[-1]:7d} | part. pagado 2015 {pp[0]:6d} 2025 {pp[-1]:6d}')
