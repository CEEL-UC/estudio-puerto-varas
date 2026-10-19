# Respaldo en Excel de todos los datos que muestra el sitio (ceel-uc.github.io/estudio-puerto-varas).
# Lee los datos tal como están publicados (index.html, datos_cohortes.js y datos_manzanas.js del repo del CEEL), así
# que cada cifra del Excel es la misma que ve quien abre el sitio. Agrega al final las tablas del análisis de
# ingreso a 1° básico (07_analisis_ingreso_1basico.py, que hay que correr antes).
# Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 06_respaldo_sitio_excel.py
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REPO = r'C:\Users\gonza\repos\ceel-estudio-puerto-varas'
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
OUT = os.path.join(BASE, 'Análisis y metadata', 'Outputs')
SALIDA = os.path.join(OUT, 'respaldo_datos_sitio_puerto_varas.xlsx')
ANIOS = list(range(2015, 2026))
DEP = {'M': 'Municipal / SLEP', 'S': 'Part. Subvencionado', 'P': 'Part. Pagado', 'C': 'Adm. delegada'}
NIVEL = {'reg': 'Tramo regular (1° básico a 4° medio)', 'par': 'Parvularia', 'esp': 'Especial', 'adu': 'Adultos'}
PRIMERO = 2009  # primer año del panel de cohortes (04_cohortes_llegada.py)
COH = ['Llegó desde otra comuna de Los Lagos', 'Llegó desde otra región', 'Sin registro previo, otros cursos',
       f'En la comuna desde {PRIMERO} o antes', 'Inició su escolaridad en la comuna (parvularia, escuela de lenguaje o 1° básico)']
ini_txt = lambda x: f'{PRIMERO} o antes' if int(x) == PRIMERO else int(x)


def leer():
    d = {}
    for ln in open(os.path.join(REPO, 'index.html'), encoding='utf8').read().split('\n'):
        m = re.match(r'var (MATN|ESTN|migFlowDepe|comunasOrigen|migFlowByDepe|topEntrantesEstab|topSalientesEstab|'
                     r'estabOrigenTop5|estabDestinoTop5|colegiosGeo) = (.*);\s*$', ln)
        if m:
            d[m.group(1)] = json.loads(m.group(2))
    for f in ['datos_cohortes.js', 'datos_manzanas.js']:
        for trozo in open(os.path.join(REPO, f), encoding='utf8').read().split(';\n'):
            m = re.match(r'\s*var (\w+)=(.*)$', trozo, re.S)
            if m:
                d[m.group(1)] = json.loads(m.group(2).rstrip().rstrip(';'))
    return d


d = leer()
NOMC = {int(k): v for k, v in d['comunasOrigen'].items()}
nomc = lambda cod: NOMC.get(cod, ['Comuna ' + str(cod), ''])
nombre_rbd = {}
for com, porAnio in d['colegiosGeo'].items():
    for filas in porAnio.values():
        for r in filas:
            nombre_rbd[r['rbd']] = (r['n'], com)

hojas = []  # (nombre, sección del sitio, descripción, columnas, filas, notas)


def hoja(nombre, seccion, desc, cols, filas, nota=''):
    hojas.append((nombre, seccion, desc, cols, filas, nota))


# ---- 01 Matrícula general
filas = []
for com, deps in d['MATN'].items():
    for dk, niveles in deps.items():
        for nk, serie in niveles.items():
            for a, v in zip(ANIOS, serie):
                filas.append([com, DEP[dk], NIVEL[nk], a, v])
hoja('01_matricula', '01 Matrícula general',
     'Matrícula por comuna, dependencia, nivel y año. El sitio muestra por defecto solo el tramo regular; los botones '
     '"Incluir también" suman parvularia, especial y adultos.',
     ['Comuna', 'Dependencia', 'Nivel', 'Año', 'Matrícula'], filas,
     'Municipal / SLEP agrupa municipales, corporaciones municipales y Servicios Locales de Educación Pública.')

# ---- 02 Establecimientos
filas = []
for com, porAnio in d['ESTN'].items():
    for a, regs in porAnio.items():
        for dk, rural, mask, n in regs:
            filas.append([com, int(a), DEP[dk], 'Rural' if rural else 'Urbano',
                          *['Sí' if mask & b else 'No' for b in (1, 2, 4, 8)], n])
hoja('02_establecimientos', '02 Establecimientos',
     'Número de colegios con matrícula mayor que cero, por comuna, año, dependencia, ruralidad y combinación de niveles '
     'que imparten. Un colegio se cuenta en el sitio si imparte al menos uno de los niveles elegidos.',
     ['Comuna', 'Año', 'Dependencia', 'Ruralidad', 'Imparte tramo regular', 'Imparte parvularia', 'Imparte especial',
      'Imparte adultos', 'N° colegios'], filas)

# ---- 03 Procedencia y salidas
filas = []
for com, deps in d['migFlowDepe'].items():
    for dk, cats in deps.items():
        for cat, serie in cats.items():
            for a, v in zip(ANIOS, serie):
                filas.append([com, DEP[dk], cat, a, v])
hoja('03_procedencia_salidas', '03 Migración estudiantil',
     'Matrícula del tramo regular de cada año según dónde estaba el alumno el año anterior, y salidas (alumnos que no '
     'aparecen en la comuna al año siguiente).',
     ['Comuna', 'Dependencia', 'Categoría', 'Año', 'Alumnos'], filas,
     'Salidas 2025 vacías: requieren la matrícula 2026. "Salidas" junta salida a otra comuna de Los Lagos, a otra región '
     'y sin registro el año siguiente.')

# ---- 03 Año de llegada (cohortes)
filas = []
for com, porAnio in d['COHORTES'].items():
    for a, deps in porAnio.items():
        for dk, regs in deps.items():
            for r in regs:
                filas.append([com, int(a), DEP[dk], ini_txt(r[0]), *r[1:5], r[5] + r[6], sum(r[1:])])
filas.sort(key=lambda f: (f[0], f[1], f[2], str(f[3])))
hoja('03_anio_llegada', '03 Migración estudiantil',
     'Matrícula del tramo regular de cada año (stock) según el año de llegada a la comuna y la procedencia. Es el gráfico '
     '"Matrícula total (stock) según año de llegada a la comuna".',
     ['Comuna', 'Año matrícula', 'Dependencia', 'Año de llegada'] + COH + ['Total'], filas,
     'Año de llegada: primer año del período continuo más reciente en que el alumno estuvo matriculado en la comuna, en '
     f'cualquier nivel (también parvularia). Para "Región de Los Lagos" se usa la permanencia en la región. Datos desde {PRIMERO}.')

filas = []
for com, porAnio in d['COHORTES_ORIGEN'].items():
    for a, deps in porAnio.items():
        for dk, porIni in deps.items():
            for ini, lista in porIni.items():
                for cod, n in lista:
                    nm, rg = nomc(cod)
                    filas.append([com, int(a), DEP[dk], ini_txt(ini), cod, nm, rg, n])
filas.sort(key=lambda f: (f[0], f[1], f[2], str(f[3]), -f[7]))
hoja('03_anio_llegada_origen', '03 Migración estudiantil',
     'Comuna de origen de quienes llegaron desde otra comuna de Los Lagos o desde otra región, por año de llegada '
     '(tabla "Ver comunas de origen" bajo el gráfico de año de llegada).',
     ['Comuna', 'Año matrícula', 'Dependencia', 'Año de llegada', 'Código comuna origen', 'Comuna origen', 'Región origen',
      'Alumnos'], filas,
     'Los códigos antiguos de Ñuble (84xx, antes de 2018) se rotulan con el nombre y la región actuales.')

# ---- 04 Mapa de colegios
filas = []
for com, porAnio in d['colegiosGeo'].items():
    for a, regs in porAnio.items():
        for r in regs:
            filas.append([com, int(a), r['rbd'], r['n'], r['d'], 'Sí' if r['r'] == 1 else 'No', r.get('t', ''), r['lat'],
                          r['lon'], r['m'], r['e'], r['i'], r['k'], r['s']])
filas.sort(key=lambda f: (f[0], f[1], -(f[9] or 0)))
hoja('04_colegios_mapa', '04 Mapas: colegios y migración',
     'Colegios del mapa con coordenadas, matrícula y balance de entradas, ingreso a 1° básico, permanencia y salidas.',
     ['Comuna', 'Año', 'RBD', 'Colegio', 'Dependencia', 'Tramo regular', 'Tipo (si no es regular)', 'Latitud', 'Longitud',
      'Matrícula', 'Entradas', 'Ingreso 1° básico', 'Se mantienen', 'Salidas'], filas,
     'Salidas vacías en 2025 (requieren 2026) y en colegios que no imparten tramo regular.')

filas = []
for com, deps in d['migFlowByDepe'].items():
    for dep, mets in deps.items():
        for met, serie in mets.items():
            for a, v in zip(ANIOS, serie):
                filas.append([com, dep, met, a, v])
hoja('04_balance_mapa', '04 Mapas: colegios y migración',
     'Balance por comuna y dependencia del cuadro superior izquierdo del mapa de colegios.',
     ['Comuna', 'Dependencia', 'Indicador', 'Año', 'Alumnos'], filas,
     'Desde 2021 los colegios municipales de Puerto Varas y Frutillar pasaron al SLEP (Serv. Local Educ.).')

for clave, nom, txt in [('topEntrantesEstab', '04_top_colegios_entradas', 'entradas'),
                        ('topSalientesEstab', '04_top_colegios_salidas', 'salidas')]:
    filas = []
    for com, porAnio in d[clave].items():
        for a, regs in porAnio.items():
            for pos, r in enumerate(regs, 1):
                filas.append([com, int(a), pos, r['rbd'], r['n'], r['cnt'], r['mat'], r['pct']])
    hoja(nom, '04 Mapas: colegios y migración', f'Colegios con más {txt} cada año (cuadro derecho del mapa).',
         ['Comuna', 'Año', 'Posición', 'RBD', 'Colegio', 'Alumnos', 'Matrícula del colegio', '% sobre la matrícula'], filas)

for clave, nom, txt in [('estabOrigenTop5', '04_colegio_comunas_origen', 'origen de sus entrantes'),
                        ('estabDestinoTop5', '04_colegio_comunas_destino', 'destino de sus salientes')]:
    filas = []
    for rbd, porAnio in d[clave].items():
        nm, com = nombre_rbd.get(int(rbd), ('', ''))
        for a, regs in porAnio.items():
            for pos, r in enumerate(regs, 1):
                filas.append([com, int(rbd), nm, int(a), pos, r['c'], r['n']])
    filas.sort(key=lambda f: (f[0], f[2], f[3], f[4]))
    hoja(nom, '04 Mapas: colegios y migración', f'Al hacer click en un colegio: top 5 comunas de {txt}.',
         ['Comuna del colegio', 'RBD', 'Colegio', 'Año', 'Posición', 'Comuna', 'Alumnos'], filas)

# ---- 04 Censo 2024
VAR_CENSO = {
    'edad': 'Edad promedio', 'men18': '% menores de 18', 'men6': '% 0 a 5 años', 'may60': '% 60 o más',
    'inm': '% inmigrantes', 'n05': 'NNA 0 a 5 (n)', 'n613': 'NNA 6 a 13 (n)', 'n1417': 'NNA 14 a 17 (n)',
    'perhog': 'Personas por hogar', 'prop': '% hogares propietarios', 'arr': '% hogares arrendatarios',
    'desoc': '% viviendas desocupadas', 'depto': '% departamentos', 'dorm4': '% 4 dormitorios o más',
    'hacin': '% viviendas hacinadas', 'alleg': 'Hogares allegados por 100 viviendas', 'irrec': '% viviendas irrecuperables',
    'precar': '% paredes precarias', 'alcant': '% con alcantarillado', 'lena': '% calefacción a leña',
    'esc': 'Escolaridad promedio (18+)', 'sup': '% educación superior (18+)', 'prof': '% directivos y profesionales',
    'elem': '% ocupaciones elementales', 'desocup': '% desocupación', 'compu': '% hogares con computador',
    'inet': '% hogares con internet fija', 'casado': '% casados (15+)', 'conviv': '% convivientes sin acuerdo (15+)',
    'auc': '% AUC (15+)', 'sepdiv': '% separados o divorciados (15+)', 'viudo': '% viudos (15+)',
    'soltero': '% solteros (15+)', 'unip': '% hogares unipersonales', 'hog60': '% hogares solo 60+',
    'hogmen': '% hogares con menores de 15', 'jefm': '% jefatura femenina', 'ha': 'Superficie (ha)',
    'ise': 'Índice socioeconómico (proxy)', 'w': 'Personas 18+ (ponderador)'}
filas = [[r['c'], 'Manzana urbana' if r['t'] == 'M' else 'Entidad rural', r['id'], r['nom'], *r['v'][:2 + len(d['MZK'])]]
         for r in d['MZ']]
filas.sort(key=lambda f: (f[0], f[1], f[2]))
hoja('04_censo_manzanas', '04 Mapas: caracterización Censo 2024',
     'Indicadores por manzana urbana y entidad rural del Censo 2024 (sin las geometrías de los polígonos).',
     ['Comuna', 'Tipo', 'ID manzana o entidad', 'Nombre', 'Personas', 'Viviendas'] + [VAR_CENSO.get(k, k) for k in d['MZK']],
     filas, 'Celdas vacías: el INE suprime el dato por confidencialidad o no aplica. Fuente: INE, Censo de Población y '
            'Vivienda 2024, resultados por manzana.')

# ---- análisis de ingreso a 1° básico (07_analisis_ingreso_1basico.py)
a = json.load(open(os.path.join(OUT, 'analisis_ingreso_1basico.json'), encoding='utf8'))
cat = a['categorias']
hoja('A1_grafico_PV_2025', 'Análisis 1° básico',
     'Gráfico de año de llegada de Puerto Varas 2025 en su versión anterior (datos desde 2014, parvularia y 1° básico separadas), la que motivó el análisis.',
     ['Año de llegada'] + cat + ['Total'], [[('2014 o antes' if r[0] == 2014 else r[0])] + r[1:] + [sum(r[1:])]
                                            for r in a['grafico_pv_2025']])
hoja('A2_llegada_por_curso', 'Análisis 1° básico',
     'Matrícula de Puerto Varas 2025 por año de llegada (filas) y curso que cursa en 2025 (columnas). Cada barra del '
     'gráfico mezcla varios cursos.',
     ['Año de llegada'] + [f'{k}° básico' for k in range(1, 9)] + [f'{k}° medio' for k in range(1, 5)] + ['Total'],
     [[('2014 o antes' if r[0] == 2014 else r[0])] + r[1:] + [sum(r[1:])] for r in a['llegada_por_curso_2025']])
hoja('A3_procedencia_1basico', 'Análisis 1° básico',
     'Alumnos de 1° básico en Puerto Varas cada año, según su situación el año anterior.',
     ['Año', 'Alumnos en 1° básico', 'Kínder o esc. de lenguaje en Puerto Varas', 'Kínder o esc. de lenguaje en otra comuna',
      'Repiten 1° básico', 'Sin registro en ningún colegio de Chile', 'Otra situación',
      '% sin registro el año anterior, total país'], a['procedencia_1basico'])
hoja('A4_seguimiento_1basico', 'Análisis 1° básico',
     'Cada cohorte de 1° básico de Puerto Varas seguida hasta 2025: cuántos siguen en la comuna y en qué categoría del '
     'gráfico de año de llegada quedaron.',
     ['Año en 1° básico', 'Alumnos en 1° básico', 'En Puerto Varas en 2025 (cualquier nivel)',
      'En el gráfico 2025 (tramo regular)', 'En otra comuna en 2025', 'Sin registro en 2025'] + cat
     + ['Quedan en la barra de su año de 1° básico'], a['seguimiento_cohortes_1basico'],
     'Las columnas de categoría suman "En el gráfico 2025".')
hoja('A5_cohorte_2018', 'Análisis 1° básico',
     'Los alumnos que estaban en 1° básico en Puerto Varas en 2018 y siguen en la comuna en 2025: barra y categoría '
     'en que aparecen en el gráfico.',
     ['Año de llegada (barra)', 'Categoría', 'Alumnos'], [[('2014 o antes' if r[0] == 2014 else r[0]), r[1], r[2]]
                                                         for r in a['detalle_cohorte_2018']])

# ---- escribir
F = 'Arial'
fh = Font(name=F, bold=True, color='FFFFFF')
fill = PatternFill('solid', fgColor='1F4E79')
wb = Workbook()
ind = wb.active
ind.title = 'Indice'
ind['A1'] = 'Respaldo de datos del sitio Estudio Puerto Varas (CEEL UC)'
ind['A1'].font = Font(name=F, bold=True, size=14)
ind['A2'] = ('Datos leídos directamente de la versión publicada en ceel-uc.github.io/estudio-puerto-varas. '
             'Todas las hojas están en formato largo (una fila por combinación), listas para filtrar o hacer tablas dinámicas.')
ind['A3'] = ('Fuentes: Mineduc, Datos Abiertos, Matrícula por estudiante 2014-2025 (archivos nacionales) y Directorio de '
             'establecimientos; INE, Censo 2024. Años 2015-2025; 2014 se usa solo como año previo.')
for r in (2, 3):
    ind.cell(r, 1).font = Font(name=F, size=10)
for j, h in enumerate(['Hoja', 'Sección del sitio', 'Contenido', 'Filas', 'Notas'], 1):
    x = ind.cell(5, j, h)
    x.font, x.fill = fh, fill
for i, (nom, sec, desc, cols, filas, nota) in enumerate(hojas, 6):
    ind.cell(i, 1, nom).hyperlink = f"#'{nom}'!A1"
    ind.cell(i, 1).font = Font(name=F, color='0563C1', underline='single')
    for j, v in enumerate([sec, desc, len(filas), nota], 2):
        ind.cell(i, j, v).font = Font(name=F)
        ind.cell(i, j).alignment = Alignment(wrap_text=True, vertical='top')
    ind.cell(i, 1).alignment = Alignment(vertical='top')
fila_g = 6 + len(hojas) + 1
ind.cell(fila_g, 1, 'Glosario').font = Font(name=F, bold=True, size=12)
GLOS = [
    ('Tramo regular', 'Enseñanza básica y media para niños y jóvenes (1° básico a 4° medio, códigos de enseñanza 110 y 310 a '
                      '810). Excluye parvularia, educación especial y adultos.'),
    ('Se mantiene', 'Matriculado en la misma comuna el año anterior.'),
    ('Entrante de otra comuna de Los Lagos / de otra región', 'Matriculado el año anterior en un colegio de otra comuna de '
                                                              'la región o de otra región.'),
    ('Sin registro el año anterior', 'No aparece en ningún colegio de Chile el año anterior y no está en 1° básico: llegó del '
                                     'extranjero, volvió al sistema o su colegio recién empezó a informar a Mineduc.'),
    ('Ingreso 1° básico (procedencia y mapa)', 'Todos los alumnos que cursan 1° básico ese año en la comuna (comienzan '
                                               'el tramo regular), vengan de donde vengan.'),
    ('Año de llegada', 'Primer año del período continuo más reciente en que el alumno estuvo matriculado en la comuna, en '
                       'cualquier nivel. Quien hizo prekínder o kínder en la comuna llega el año de su primer registro en '
                       'parvularia, no el año en que entra a 1° básico. Datos desde 2009.'),
    ('Inició su escolaridad en la comuna (año de llegada)', 'Alumnos cuyo primer registro en el sistema escolar es en la '
                                                             'comuna, en parvularia, escuela de lenguaje o 1° básico.'),
    ('Salida', 'No aparece en la comuna al año siguiente (a otra comuna de Los Lagos, a otra región o sin registro).'),
]
for i, (t, x) in enumerate(GLOS, fila_g + 1):
    ind.cell(i, 1, t).font = Font(name=F, bold=True)
    ind.cell(i, 2, x).font = Font(name=F)
    ind.merge_cells(start_row=i, start_column=2, end_row=i, end_column=5)
    ind.cell(i, 2).alignment = Alignment(wrap_text=True, vertical='top')
    ind.cell(i, 1).alignment = Alignment(wrap_text=True, vertical='top')
    ind.row_dimensions[i].height = 32
for col, w in zip('ABCDE', [30, 30, 80, 10, 60]):
    ind.column_dimensions[col].width = w

for nom, sec, desc, cols, filas, nota in hojas:
    ws = wb.create_sheet(nom)
    ws['A1'] = desc
    ws['A1'].font = Font(name=F, italic=True, size=9)
    if nota:
        ws['A2'] = 'Nota: ' + nota
        ws['A2'].font = Font(name=F, italic=True, size=9)
    for j, h in enumerate(cols, 1):
        x = ws.cell(4, j, h)
        x.font, x.fill = fh, fill
        x.alignment = Alignment(wrap_text=True, vertical='center')
    for i, f in enumerate(filas, 5):
        for j, v in enumerate(f, 1):
            ws.cell(i, j, v).font = Font(name=F, size=10)
    ws.freeze_panes = 'A5'
    ws.auto_filter.ref = f'A4:{get_column_letter(len(cols))}{4 + len(filas)}'
    ws.row_dimensions[4].height = 45
    for j, h in enumerate(cols, 1):
        largo = max([len(str(h)) // 2] + [len(str(f[j - 1])) for f in filas[:300]])
        ws.column_dimensions[get_column_letter(j)].width = min(max(10, largo + 2), 45)
wb.save(SALIDA)
print('Excel:', SALIDA)
for nom, _, _, _, filas, _ in hojas:
    print(f'  {nom}: {len(filas)} filas')
