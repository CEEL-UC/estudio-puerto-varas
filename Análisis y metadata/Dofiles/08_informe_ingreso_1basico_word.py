# Informe Word: por qué la categoría "Ingresó a 1° básico en la comuna" es tan chica en el gráfico de año de llegada.
# Usa las tablas de 07_analisis_ingreso_1basico.py (hay que correrlo antes).
# Correr con:
#   C:\Users\gonza\AppData\Local\Programs\Python\Python312\python.exe 08_informe_ingreso_1basico_word.py
import json, os
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')
OUT = os.path.join(BASE, 'Análisis y metadata', 'Outputs')
SALIDA = os.path.join(OUT, 'informe_ingreso_1basico_anio_llegada.docx')
a = json.load(open(os.path.join(OUT, 'analisis_ingreso_1basico.json'), encoding='utf8'))
fmt = lambda x: f'{x:,}'.replace(',', '.') if isinstance(x, int) and not 2000 <= x <= 2030 else str(x)
pct = lambda x, t: f'{100 * x / t:.1f}'.replace('.', ',') + '%'

doc = Document()
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2.3)
st = doc.styles['Normal']
st.font.name = 'Arial'
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for n, tam in [('Title', 20), ('Heading 1', 14), ('Heading 2', 12)]:
    s = doc.styles[n]
    s.font.name = 'Arial'
    s.font.size = Pt(tam)
    s.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)


def p(texto, negrita_inicio=None, estilo=None):
    par = doc.add_paragraph(style=estilo)
    if negrita_inicio:
        par.add_run(negrita_inicio).bold = True
    par.add_run(texto)
    return par


def vineta(texto, negrita_inicio=None):
    return p(texto, negrita_inicio, 'List Bullet')


def sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def tabla(cols, filas, anchos, titulo, fuente=None, destacar=None):
    t = doc.add_paragraph()
    t.add_run(titulo).bold = True
    t.paragraph_format.space_after = Pt(3)
    t.paragraph_format.keep_with_next = True
    tb = doc.add_table(rows=1, cols=len(cols))
    tb.style = 'Table Grid'
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    for j, h in enumerate(cols):
        c = tb.rows[0].cells[j]
        c.text = ''
        r = c.paragraphs[0].add_run(h)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), RGBColor(0xFF, 0xFF, 0xFF)
        sombrear(c, '1F4E79')
    for i, f in enumerate(filas):
        cells = tb.add_row().cells
        for j, v in enumerate(f):
            cells[j].text = ''
            r = cells[j].paragraphs[0].add_run(fmt(v))
            r.font.size = Pt(8.5)
            if j > 0 and not isinstance(v, str) or (isinstance(v, str) and v.endswith('%')):
                cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if destacar is not None and i == destacar:
                r.bold = True
                sombrear(cells[j], 'FFF2CC')
    for fila in tb.rows:
        for j, w in enumerate(anchos):
            fila.cells[j].width = Cm(w)
    if fuente:
        f = doc.add_paragraph()
        r = f.add_run(fuente)
        r.italic, r.font.size = True, Pt(8)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(2)


g = {r[0]: r[1:] for r in a['grafico_pv_2025']}           # ce, ip, ib, sr, ll, or
seg = {r[0]: r for r in a['seguimiento_cohortes_1basico']}
proc = {r[0]: r for r in a['procedencia_1basico']}
cur = {r[0]: r[1:] for r in a['llegada_por_curso_2025']}
s18 = seg[2018]
d18 = a['detalle_cohorte_2018']
parv18 = sum(n for ini, cat, n in d18 if cat.startswith('Comenzó'))
barra18 = sum(n for ini, cat, n in d18 if ini == 2018)
otros18 = sum(n for ini, cat, n in d18 if cat.startswith('Sin registro'))
ib = [seg[t][8] for t in range(2015, 2026)]
FUENTE = 'Fuente: Mineduc, Matrícula por estudiante 2014-2025. Elaboración propia.'

doc.add_paragraph('Ingreso a 1° básico en el gráfico de año de llegada', style='Title')
p('Estudio Puerto Varas, CEEL UC. Nota técnica, octubre de 2026.').runs[0].italic = True

doc.add_heading('La pregunta', 1)
p('En el gráfico "Matrícula total 2025 (stock) de Puerto Varas según año de llegada a la comuna", la categoría '
  '"Ingresó a 1° básico en la comuna" es muy chica: 347 alumnos en total (3,5% de la matrícula), y solo 14 en la barra '
  'de 2018. Si cada año entran unos 700 niños a 1° básico en Puerto Varas, uno esperaría ver cientos en cada barra. '
  'La hipótesis inicial fue que muchos ya habían terminado sus 12 años de escolaridad, pero eso no calza: quien '
  'entró a 1° básico en 2018 está hoy en 8° básico.')

doc.add_heading('Respuesta corta', 1)
p('No es un error en los datos ni una pérdida de alumnos. Es la definición del gráfico:', None)
vineta(' el gráfico fecha la llegada a la comuna, no la entrada a 1° básico. El año de llegada es el primer año del '
       'período continuo más reciente en que el alumno estuvo matriculado en un colegio de la comuna, en cualquier nivel, '
       'incluida la parvularia.', 'El año de llegada cuenta desde la parvularia:')
vineta(' un niño que hizo prekínder en Puerto Varas en 2016, kínder en 2017 y 1° básico en 2018 "llegó" en 2016, y '
       'aparece en esa barra como "Comenzó en parvularia o escuela de lenguaje". No aparece en la barra de 2018.',
       'Por eso casi todos los que entran a 1° básico quedan en barras anteriores:')
vineta(' solo reúne a los niños cuyo primer registro en todo el sistema escolar chileno es 1° básico, es decir, que no '
       'tienen prekínder, kínder ni escuela de lenguaje registrados en ningún colegio del país. Ese grupo es chico en '
       'todo Chile: entre 5% y 7% de los alumnos de 1° básico cada año.', 'La categoría "Ingresó a 1° básico" es residual:')
vineta(' los que entraron a 1° básico en 2018 siguen en el colegio. Egresan de 4° medio 12 años después de entrar a '
       '1° básico, así que solo las cohortes de 2013 o antes han egresado, y esas están fuera del rango del gráfico.',
       'La hipótesis del egreso no aplica:')

doc.add_heading('Dónde está la cohorte de 1° básico de 2018', 1)
p(f'En 2018 había {fmt(s18[1])} alumnos en 1° básico en Puerto Varas. En 2025, {fmt(s18[2])} siguen matriculados en la '
  f'comuna ({fmt(s18[3])} en el tramo regular, que es lo que muestra el gráfico), {fmt(s18[4])} están en un colegio de '
  f'otra comuna y {fmt(s18[5])} no aparecen en ningún colegio de Chile. De los {fmt(s18[3])} que siguen en el gráfico, '
  f'casi todos están en 8° básico (o en 7° si repitieron), pero se reparten así entre las barras:')
suma = lambda f: sum(n for ini, cat, n in d18 if f(ini, cat))
otra = lambda cat: cat.startswith('Llegó')
tabla(['Barra del gráfico (año de llegada)', 'Categoría', 'Alumnos'],
      [['2014 o antes', 'En la comuna desde 2014 o antes', suma(lambda i, k: i == 2014)],
       ['2015 a 2017', 'Comenzó en parvularia o escuela de lenguaje', parv18],
       ['2015 a 2017', 'Llegó desde otra comuna de Los Lagos u otra región', suma(lambda i, k: 2015 <= i <= 2017 and otra(k))],
       ['2018', 'Ingresó a 1° básico en la comuna', suma(lambda i, k: i == 2018 and k.startswith('Ingresó'))],
       ['2018', 'Llegó desde otra comuna de Los Lagos u otra región', suma(lambda i, k: i == 2018 and otra(k))],
       ['2020 a 2025', 'Sin registro previo, otros cursos (registro interrumpido)', otros18],
       ['2020 a 2025', 'Llegó desde otra comuna de Los Lagos u otra región', suma(lambda i, k: i >= 2019 and otra(k))],
       ['Total', '', s18[3]]],
      [4.2, 9.5, 2.3], 'Tabla 1. Alumnos de 1° básico 2018 que siguen en Puerto Varas en 2025, según barra y categoría',
      FUENTE)
p(f'{fmt(parv18)} de ellos ({pct(parv18, s18[3])}) están en las barras de 2015, 2016 y 2017 como "Comenzó en parvularia", '
  f'porque esos fueron sus años de prekínder, kínder o escuela de lenguaje en la comuna. Otros {fmt(d18[0][2])} ya estaban en la '
  f'comuna en 2014 (en prekínder o antes) y quedan en "2014 o antes". En la barra de 2018 quedan solo {fmt(barra18)}: los 14 que no tenían '
  f'ningún registro previo y los que llegaron ese año desde otra comuna o región.')

doc.add_heading('Qué hay entonces en la barra de 2018', 1)
c18 = cur[2018]
p(f'La barra de 2018 ({fmt(sum(c18))} alumnos en 2025) no es la cohorte de 1° básico de 2018, sino quienes empezaron su '
  f'período actual en la comuna ese año, en cualquier curso. El grupo principal ({fmt(g[2018][1])} alumnos) son niños que '
  f'entraron a prekínder, kínder o una escuela de lenguaje en 2018 y que hoy cursan sobre todo 5° y 6° básico. Cada barra '
  f'mezcla varios cursos, como muestra la tabla 2.')
tabla(['Año de llegada', '1°-3° B', '4° B', '5° B', '6° B', '7° B', '8° B', '1°-4° M', 'Total'],
      [[('2014 o antes' if k == 2014 else k), sum(v[0:3]), v[3], v[4], v[5], v[6], v[7], sum(v[8:12]), sum(v)]
       for k, v in sorted(cur.items())],
      [3.2, 1.6, 1.4, 1.4, 1.4, 1.4, 1.4, 1.7, 1.7],
      'Tabla 2. Matrícula de Puerto Varas 2025 por año de llegada y curso que cursa en 2025', FUENTE,
      destacar=4)

doc.add_heading('De dónde vienen los niños que entran a 1° básico', 1)
p('La tabla 3 mira el problema desde el otro lado: cada año, dónde estaban el año anterior los alumnos de 1° básico de '
  'Puerto Varas. Entre 80% y 88% hizo kínder (o escuela de lenguaje) en la misma comuna, por lo que su año de llegada es '
  'anterior a 1° básico. Solo entre 22 y 82 niños por año no tenían ningún registro previo; ese es el universo de la '
  'categoría "Ingresó a 1° básico".')
tabla(['Año', 'En 1° básico', 'Kínder en P. Varas', 'Kínder en otra comuna', 'Repiten', 'Sin registro previo',
       '% sin registro', '% sin registro, Chile'],
      [[r[0], r[1], r[2], r[3], r[4], r[5], pct(r[5], r[1]), str(r[7]).replace('.', ',') + '%'] for r in a['procedencia_1basico']],
      [1.3, 1.8, 2.2, 2.2, 1.6, 2.2, 1.9, 2.2],
      'Tabla 3. Alumnos de 1° básico en Puerto Varas según su situación el año anterior', FUENTE, destacar=3)
p('El porcentaje sin registro previo en Puerto Varas es parecido al del país hasta 2017 (3% a 4% contra 5% a 6%) y sube '
  'desde 2018 (6% a 11%). El alza coincide con colegios particulares pagados que abren o empiezan a informar a Mineduc '
  'en esos años (Terravida en 2018, Patagonia en 2020), cuyos alumnos no traen kínder registrado. Esto no está '
  'verificado colegio por colegio.')

doc.add_heading('Por qué la categoría crece en las barras recientes', 1)
p(f'En el gráfico la categoría pasa de entre 6 y 14 alumnos en las barras 2015-2018 a entre 45 y 50 desde 2020 '
  f'({", ".join(fmt(x) for x in ib)} para 2015-2025). Hay dos razones:')
vineta(' desde 2018 hay más niños de 1° básico sin registro previo (tabla 3).', 'Más niños sin kínder registrado:')
vineta(' una cohorte antigua tuvo más años para irse de la comuna o interrumpir su registro. Si un alumno deja de '
       'aparecer un año y vuelve, su año de llegada se reinicia y pasa a otra barra.', 'Más tiempo de exposición:')

doc.add_heading('Un hallazgo menor: registros interrumpidos', 1)
p(f'{fmt(otros18)} alumnos de la cohorte 2018 aparecen como "Sin registro previo, otros cursos" en barras posteriores, '
  'casi todos en 2020. Estaban en Puerto Varas en 2018, no aparecen en ningún colegio de Chile en 2019 y vuelven en '
  '2020. La mayoría (24) estaba en el Colegio Germania del Verbo Divino en 2018 y en 2020, que informa 834 alumnos en '
  '2019 contra 926 en 2018 y 1.014 en 2020. Parece un reporte incompleto de ese colegio en 2019, no una salida real. El '
  'efecto en el gráfico es chico pero conviene tenerlo presente: lleva alumnos antiguos a barras recientes.')

doc.add_heading('Qué falta definir', 1)
p('El gráfico está bien calculado; el problema es de rotulado. Quedan abiertas estas decisiones:')
vineta(' si "Comenzó en parvularia" e "Ingresó a 1° básico" se presentan juntas, por ejemplo como "Inició su '
       'escolaridad en la comuna", ya que ambas son niños que empiezan su trayectoria escolar en Puerto Varas.',
       'Unir categorías:')
vineta(' si se agrega una nota bajo el gráfico que aclare que el año de llegada cuenta desde la parvularia y que quien '
       'hizo kínder en la comuna no aparece en el año en que entró a 1° básico.', 'Nota aclaratoria:')
vineta(' si la pregunta de interés es cuántos alumnos entran a 1° básico cada año, ese dato ya está en el gráfico de '
       'procedencia y salidas de la misma sección ("Ingreso 1° básico", 737 en 2025), que cuenta a todos los alumnos de '
       '1° básico.', 'Remitir al otro gráfico:')
vineta(' si se corrigen los registros interrumpidos de un año (tratar como continuo a quien falta un solo año y vuelve '
       'al mismo colegio).', 'Huecos de un año:')

doc.add_heading('Respaldo', 1)
p('Las tablas de esta nota están en las hojas A1 a A5 de respaldo_datos_sitio_puerto_varas.xlsx, junto con todos los datos '
  'del sitio. Se calculan con 07_analisis_ingreso_1basico.py (carpeta Dofiles), que reproduce exactamente las cifras '
  'publicadas del gráfico para Puerto Varas 2025 (9.961 alumnos).')

doc.save(SALIDA)
print('Word:', SALIDA)
