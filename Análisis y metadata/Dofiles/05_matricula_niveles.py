# Matrícula y colegios por nivel de enseñanza, dependencia y comuna (Región de Los Lagos, 2015-2025), para que el
# sitio muestre por defecto solo el tramo regular (1° básico a 4° medio) y permita sumar con un botón la educación
# parvularia, la especial y la de adultos.
# Fuente: caché liviano de los archivos nacionales de Matrícula por estudiante (lo arma 03_migracion_regional.py),
# más la ruralidad de cada colegio leída de los mismos CSV. Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 05_matricula_niveles.py
import glob, json, os
import polars as pl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
CACHE = os.path.join(BASE, 'Datos', '[merge] Panel Puerto Varas', 'nacional_liviano')
ORIGEN = os.path.join(BASE, 'Datos', 'Mineduc', 'Matricula por estudiante')
NOMBRES = os.path.join(BASE, 'Análisis y metadata', 'Outputs', 'migracion_regional_sitio.json')
SALIDA = os.path.join(BASE, 'Análisis y metadata', 'Outputs', 'matricula_niveles_sitio.json')
ANIOS = list(range(2015, 2026))
REGION = 10
REG = 'Región de Los Lagos'
c = pl.col

# nivel según código de enseñanza (Mineduc):
#   reg = básica regular (110) y media regular jóvenes, científico-humanista y técnico-profesional (310 a 910)
#   par = educación parvularia (10)
#   esp = educación especial (200 a 299)
#   adu = educación de adultos (básica 160-167, media 360-363 y técnico-profesional 460-963)
NIVEL = (pl.when(c('COD_ENSE') == 10).then(pl.lit('par'))
         .when((c('COD_ENSE') == 110) | c('COD_ENSE').is_in([310, 410, 510, 610, 710, 810, 910])).then(pl.lit('reg'))
         .when(c('COD_ENSE').is_between(200, 299)).then(pl.lit('esp'))
         .otherwise(pl.lit('adu')))
DEP = c('COD_DEPE2').replace_strict({1: 'M', 5: 'M', 2: 'S', 3: 'P', 4: 'C'}, default='C', return_dtype=pl.Utf8)
NIVS, DEPS = ['reg', 'par', 'esp', 'adu'], ['M', 'S', 'P', 'C']
BIT = {'reg': 1, 'par': 2, 'esp': 4, 'adu': 8}

nombres = json.load(open(NOMBRES, encoding='utf8'))
NOM = {x['cod']: x['n'] for x in nombres['comunas']}


def rural(anio):
    # RURAL_RBD por colegio: no está en el caché liviano, se lee directo del CSV del año
    f = [x for x in glob.glob(os.path.join(ORIGEN, str(anio), '*')) if x.lower().endswith('.csv')][0]
    hdr = [x.strip().upper() for x in open(f, encoding='latin-1').readline().strip().split(';')]
    hdr[0] = 'AGNO' if hdr[0].endswith('AGNO') else hdr[0]
    r = pl.read_csv(f, separator=';', encoding='utf8-lossy', infer_schema=False, quote_char=None,
                    columns=[hdr.index('RBD'), hdr.index('COD_REG_RBD'), hdr.index('RURAL_RBD')], new_columns=['RBD', 'reg', 'rural'])
    r = r.with_columns([c(x).str.strip_chars().cast(pl.Int64, strict=False) for x in r.columns])
    return r.filter(c('reg') == REGION).group_by('RBD').agg(c('rural').max())


mat = {}          # mat[comuna][dep][nivel] = serie 2015-2025 (matrículas)
est = {}          # est[comuna][año] = [[dep, rural, máscara de niveles con matrícula], ...] un elemento por colegio
for i, a in enumerate(ANIOS):
    d = pl.read_parquet(os.path.join(CACHE, f'mat_{a}.parquet'), glob=False).filter(c('COD_REG_RBD') == REGION)
    d = d.with_columns(niv=NIVEL, dep=DEP)
    for (cod, dep, niv), n in d.group_by(['COD_COM_RBD', 'dep', 'niv']).len().rows() and [((r[0], r[1], r[2]), r[3]) for r in d.group_by(['COD_COM_RBD', 'dep', 'niv']).len().rows()]:
        for etq in (NOM[cod], REG):
            s = mat.setdefault(etq, {}).setdefault(dep, {}).setdefault(niv, [0] * len(ANIOS))
            s[i] += n
    ru = rural(a)
    colegios = d.group_by(['RBD', 'COD_COM_RBD', 'dep']).agg(c('niv').unique()).join(ru, on='RBD', how='left')
    for rbd, cod, dep, nivs, rur in colegios.rows():
        mask = sum(BIT[x] for x in nivs)
        fila = [dep, int(rur or 0), mask]
        est.setdefault(NOM[cod], {}).setdefault(str(a), []).append(fila)
        est.setdefault(REG, {}).setdefault(str(a), []).append(fila)
    print(a, 'matrículas región', d.height, '| colegios', colegios.height, '| sin dato de ruralidad', colegios['rural'].null_count())

# completa ceros para que todas las combinaciones existan
for etq in mat:
    for dep in DEPS:
        for niv in NIVS:
            mat[etq].setdefault(dep, {}).setdefault(niv, [0] * len(ANIOS))
# compacta los colegios: [dep, rural, máscara, cantidad]
for etq in est:
    for a in est[etq]:
        cnt = {}
        for f in est[etq][a]:
            k = tuple(f)
            cnt[k] = cnt.get(k, 0) + 1
        est[etq][a] = [list(k) + [n] for k, n in sorted(cnt.items())]

json.dump({'anios': ANIOS, 'mat': mat, 'est': est}, open(SALIDA, 'w', encoding='utf8'), ensure_ascii=False, separators=(',', ':'))
print('JSON:', SALIDA, round(os.path.getsize(SALIDA) / 1e3), 'KB')
pv = mat['Puerto Varas']
print('PV total todas las categorías:', [sum(pv[dp][nv][i] for dp in DEPS for nv in NIVS) for i in range(len(ANIOS))])
print('PV tramo regular:', [sum(pv[dp]['reg'][i] for dp in DEPS) for i in range(len(ANIOS))])
print('PV parvularia:', [sum(pv[dp]['par'][i] for dp in DEPS) for i in range(len(ANIOS))])
print('PV pagado reg / par:', pv['P']['reg'][-1], pv['P']['par'][-1])
