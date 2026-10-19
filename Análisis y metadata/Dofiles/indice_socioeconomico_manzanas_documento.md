---
title: "Índice socioeconómico por manzana"
subtitle: "Censo de Población y Vivienda 2024. Puerto Varas, Puerto Montt, Frutillar y Osorno"
date: "Octubre de 2026"
lang: es-CL
---

# 1. Propósito

El índice socioeconómico por manzana resume en un solo número la situación educacional, ocupacional y material de cada manzana urbana y entidad rural de las cuatro comunas del estudio. Se construye con la base manzana-entidad del Censo 2024 y es el indicador que muestra el mapa "Caracterización Censo 2024" del sitio del estudio, en la pestaña Socioeconómico.

Es un **proxy**: el Censo no mide ingresos, de modo que el índice no reproduce el grupo socioeconómico (GSE) de la Asociación de Investigadores de Mercado (AIM), cuya clasificación actual se basa principalmente en el ingreso del hogar. Tampoco clasifica hogares: describe el perfil que predomina en cada manzana.

# 2. Fuente y unidad de análisis

- **Fuente:** base manzana-entidad y cartografía censal del Censo de Población y Vivienda 2024 (INE), capas `Manzanas` y `Entidades` de la Región de Los Lagos. Se usan solo los polígonos con datos censales (`MZ_BASE_CENSO = 1`).
- **Unidad de análisis:** el polígono, que puede ser una manzana urbana o una entidad rural. Las entidades rurales son polígonos de Thiessen trazados por el INE alrededor de las viviendas censadas; son una referencia estadística, no límites reales en terreno.
- **Cobertura:** 6.369 polígonos de las comunas de Puerto Varas (10109), Puerto Montt (10101), Frutillar (10105) y Osorno (10301): 5.371 manzanas urbanas y 998 entidades rurales. De ellos, 6.230 tienen índice; los 139 restantes no cumplen el mínimo de componentes con dato (sección 5).
- **Supresión del INE:** el INE oculta las cifras de personas en celdas con menos de 4 casos. Los datos de hogares y viviendas no se suprimen. Las manzanas con 1 a 3 viviendas ocupadas fueron fusionadas por el INE con una vecina.

# 3. Componentes

El índice combina siete componentes: cuatro que indican mejor situación (signo positivo) y tres que indican carencia (signo negativo). Para la manzana $j$:

| Componente | Variable $x_{jk}$ | Signo $s_k$ |
|---|---|:-:|
| Escolaridad | Años de escolaridad promedio de las personas de 18 años o más | $+1$ |
| Educación superior | % de personas de 18 años o más con educación terciaria, maestría o doctorado | $+1$ |
| Directivos y profesionales | % de ocupados en los grandes grupos 1 y 2 de la CIUO-08 | $+1$ |
| Computador | % de hogares con computador | $+1$ |
| Hacinamiento | % de viviendas ocupadas hacinadas | $-1$ |
| Paredes precarias | % de viviendas ocupadas con paredes de material precario | $-1$ |
| Allegamiento | Hogares allegados por cada 100 viviendas ocupadas | $-1$ |

Las fórmulas usan los nombres de las variables de la base manzana-entidad del INE.

**Escolaridad.** Se usa directamente el promedio que publica el INE:

$$x_{j,\text{esc}} = \texttt{prom\_escolaridad18}_j$$

**Educación superior.** El denominador son las personas de 18 años o más, que se obtienen restando a la población total los tramos de 0 a 17 años:

$$x_{j,\text{sup}} = 100 \cdot \frac{\texttt{n\_cine\_terciaria\_maestria\_doctorado}_j}{\texttt{n\_per}_j - \left(\texttt{n\_edad\_0\_5}_j + \texttt{n\_edad\_6\_13}_j + \texttt{n\_edad\_14\_17}_j\right)}$$

**Directivos y profesionales.** CIUO 1 corresponde a directores, gerentes y administradores; CIUO 2, a profesionales, científicos e intelectuales:

$$x_{j,\text{prof}} = 100 \cdot \frac{\texttt{n\_ciuo\_1}_j + \texttt{n\_ciuo\_2}_j}{\texttt{n\_ocupado}_j}$$

**Computador:**

$$x_{j,\text{comp}} = 100 \cdot \frac{\texttt{n\_serv\_compu}_j}{\texttt{n\_hog}_j}$$

**Hacinamiento y allegamiento.** El universo del INE para estas variables son las viviendas particulares ocupadas con moradores presentes. Ese total no viene como columna propia, por lo que se reconstruye sumando los tipos de vivienda:

$$V_j = \sum_{t \in T} \texttt{n\_tipo\_viv\_}t_j, \qquad T = \{\text{casa, depto, indígena, pieza, mediagua, móvil, otro}\}$$

$$x_{j,\text{hac}} = 100 \cdot \frac{\texttt{n\_viv\_hacinadas}_j}{V_j} \qquad\qquad x_{j,\text{all}} = 100 \cdot \frac{\texttt{n\_hog\_allegados}_j}{V_j}$$

El hacinamiento y el allegamiento siguen las definiciones del INE. Como una vivienda puede tener más de un hogar allegado, $x_{j,\text{all}}$ es una tasa por cada 100 viviendas y no un porcentaje acotado en 100.

**Paredes precarias.** Se consideran precarias las paredes de tabique sin forro interior, las de adobe, barro, pirca, quincha u otros materiales artesanales, y las de materiales precarios o de desecho. El denominador es la suma de todas las categorías de materialidad de paredes:

$$x_{j,\text{prec}} = 100 \cdot \frac{\texttt{tabique\_sin\_forro}_j + \texttt{artesanal}_j + \texttt{precarios}_j}{\sum_{m \in M} \texttt{n\_mat\_paredes\_}m_j}$$

donde $M$ = {hormigón, albañilería, tabique forrado, tabique sin forro, artesanal, precarios}.

**Reglas para datos faltantes:**

- Si el numerador fue suprimido por el INE, el componente queda sin dato. No se reemplaza por cero.
- En los numeradores compuestos (directivos y profesionales, paredes precarias), el componente queda sin dato solo si todas sus partes vienen suprimidas.
- En los denominadores que se arman sumando categorías, una categoría suprimida cuenta como cero.
- Si el denominador es cero, el componente queda sin dato.
- Todos los componentes se redondean a un decimal antes de estandarizarlos, con el redondeo habitual (los empates, como 6,25, suben a 6,3), igual que la función REDONDEAR de Excel.

# 4. Estandarización

Cada componente se lleva a puntaje $z$ con la media y la desviación estándar de **todos los polígonos de las cuatro comunas** que tienen dato en ese componente, y se multiplica por su signo. Así, en los siete puntajes un valor más alto siempre significa mejor situación:

$$z_{jk} = s_k \cdot \frac{x_{jk} - \mu_k}{\sigma_k}$$

$$\mu_k = \frac{1}{n_k}\sum_{j=1}^{n_k} x_{jk} \qquad\qquad \sigma_k = \sqrt{\frac{1}{n_k - 1}\sum_{j=1}^{n_k}\left(x_{jk} - \mu_k\right)^2}$$

donde $n_k$ es el número de polígonos con dato en el componente $k$. Cada polígono pesa lo mismo, sin ponderar por población.

Parámetros obtenidos:

| Componente | Signo | Media $\mu_k$ | Desv. estándar $\sigma_k$ | Polígonos con dato $n_k$ |
|--------------------------------------|--------:|--------------:|---------------:|---------------:|
| Escolaridad (años) | $+1$ | 11,99 | 2,06 | 6.234 |
| Educación superior (%) | $+1$ | 29,45 | 20,99 | 6.234 |
| Directivos y profesionales (%) | $+1$ | 19,98 | 19,48 | 6.232 |
| Computador (%) | $+1$ | 53,63 | 20,87 | 6.230 |
| Hacinamiento (%) | $-1$ | 4,29 | 5,43 | 6.365 |
| Paredes precarias (%) | $-1$ | 4,06 | 7,67 | 6.230 |
| Allegamiento (por 100 viviendas) | $-1$ | 1,26 | 3,48 | 6.365 |

# 5. Cálculo del índice

El índice de la manzana $j$ es el promedio simple de los puntajes $z$ disponibles, todos con el mismo peso:

$$I_j = \frac{1}{|K_j|}\sum_{k \in K_j} z_{jk} \qquad \text{si } |K_j| \geq 4$$

donde $K_j$ es el conjunto de componentes con dato en la manzana $j$. Si la manzana tiene menos de 4 componentes con dato, el índice queda sin dato. El resultado se redondea a dos decimales.

# 6. Interpretación

- $I_j = 0$ corresponde a una manzana promedio de las cuatro comunas. Un valor positivo indica mejor situación que ese promedio y uno negativo, peor.
- La escala está en unidades de desviación estándar, pero **no está acotada** entre $-1$ y $1$. En los datos va de $-4{,}15$ a $1{,}89$, con mediana $-0{,}09$.
- El rango es asimétrico. Hacinamiento, paredes precarias y allegamiento son cero en la mayoría de las manzanas y muy altos en unas pocas, por lo que una manzana con carencias concentradas puede bajar varias desviaciones estándar. En la parte alta, escolaridad y educación superior tienen techo y no se alejan tanto del promedio. Por eso la cola larga va hacia valores negativos.
- Como los parámetros se calculan con las cuatro comunas juntas, los valores son comparables entre comunas.

# 7. Clasificaciones derivadas

**Percentil.** Ubica cada manzana respecto de todas las manzanas con índice de las cuatro comunas ($N = 6.230$):

$$P_j = 100 \cdot \frac{\#\{\, i : I_i < I_j \,\}}{N}$$

Por ejemplo, $P_j = 80$ significa que el 80% de las manzanas de las cuatro comunas tiene un índice menor.

**Quintiles por comuna (los que muestra el mapa).** El mapa colorea las manzanas en cinco grupos dentro de la comuna elegida. Si $v_{(1)} \leq \dots \leq v_{(n)}$ son los índices ordenados de la comuna, los cortes son:

$$c_q = v_{(\lfloor q n / 5 \rfloor + 1)}, \qquad q = 1, 2, 3, 4$$

y cada manzana se asigna al grupo $1 + \#\{\, q : I_j \geq c_q \,\}$. Si un corte se repite o es igual al mínimo, se omite y quedan menos de cinco grupos (no ocurre con este índice, porque es continuo).

Los quintiles dependen de qué polígonos entran al cálculo. El mapa, por defecto, muestra solo manzanas urbanas; al encender las entidades rurales, los cortes se recalculan con todos los polígonos. El Excel trae ambas versiones. Cortes obtenidos:

| Comuna | Alcance | Corte 1 | Corte 2 | Corte 3 | Corte 4 |
|---|---|--:|--:|--:|--:|
| Puerto Varas | Solo manzanas urbanas | −0,53 | −0,19 | 0,25 | 0,97 |
| Puerto Varas | Todos los polígonos | −0,45 | −0,09 | 0,43 | 1,04 |
| Puerto Montt | Solo manzanas urbanas | −0,48 | −0,19 | 0,14 | 0,60 |
| Puerto Montt | Todos los polígonos | −0,49 | −0,20 | 0,12 | 0,58 |
| Frutillar | Solo manzanas urbanas | −0,53 | −0,31 | −0,10 | 0,26 |
| Frutillar | Todos los polígonos | −0,53 | −0,29 | −0,04 | 0,51 |
| Osorno | Solo manzanas urbanas | −0,62 | −0,32 | −0,03 | 0,45 |
| Osorno | Todos los polígonos | −0,61 | −0,31 | −0,02 | 0,45 |

Como los quintiles son relativos a cada comuna, el quintil 5 de Frutillar y el de Puerto Varas no representan el mismo nivel absoluto. Para comparar entre comunas hay que usar el índice o el percentil.

# 8. Resumen por comuna

Estadísticas del índice entre los polígonos con dato, sin ponderar por población. N.º = polígonos con índice.

| Comuna y tipo de polígono | N.º | Personas | Media | Mediana | Mínimo | Máximo |
|------------------------|------:|----------:|-------:|---------:|--------:|--------:|
| Puerto Varas, manzanas urbanas | 422 | 34.787 | 0,13 | 0,02 | −1,97 | 1,83 |
| Puerto Varas, entidades rurales | 193 | 16.841 | 0,43 | 0,42 | −1,05 | 1,82 |
| Puerto Montt, manzanas urbanas | 2.788 | 233.141 | 0,03 | −0,04 | −4,15 | 1,83 |
| Puerto Montt, entidades rurales | 210 | 35.615 | −0,12 | −0,18 | −1,49 | 1,16 |
| Frutillar, manzanas urbanas | 223 | 15.981 | −0,13 | −0,23 | −1,72 | 1,61 |
| Frutillar, entidades rurales | 166 | 6.354 | 0,08 | −0,09 | −1,49 | 1,78 |
| Osorno, manzanas urbanas | 1.902 | 146.624 | −0,10 | −0,18 | −2,99 | 1,89 |
| Osorno, entidades rurales | 326 | 15.810 | −0,07 | −0,14 | −1,80 | 1,86 |

# 9. Relación con el ISMT

El índice sigue la lógica del Índice Socio Material Territorial (ISMT) del Observatorio de Ciudades UC, construido con el Censo 2017 por manzana, que combina escolaridad del jefe de hogar, materialidad de la vivienda, hacinamiento y allegamiento. No es una réplica exacta:

- La base 2024 por manzana no trae la escolaridad del jefe de hogar. Se reemplaza por la escolaridad promedio de todas las personas de 18 años o más y por el porcentaje con educación superior.
- Se agregan dos componentes que el ISMT no usa y que discriminan mejor en la parte alta de la distribución: directivos y profesionales, y hogares con computador.
- El ISMT asigna categorías tipo ABC1, C2, C3, D y E. Este índice no lo hace, porque sin datos de ingreso esas etiquetas sugerirían una equivalencia con el GSE de la AIM que no existe. Si se requieren, podrían asignarse calibrando los cortes con la distribución de GSE publicada para la región, siempre presentándolas como aproximación.

# 10. Limitaciones

- **Es un indicador de la manzana, no del hogar.** Una manzana con índice alto es una donde predomina ese perfil; puede tener hogares de situación muy distinta (falacia ecológica).
- **No mide ingresos.** Es un proxy educacional, ocupacional y material.
- **Pesos iguales y sin ponderar por población.** Cada componente pesa lo mismo en el índice y cada polígono pesa lo mismo en la media y la desviación estándar, aunque tenga 5 o 500 habitantes.
- **Manzanas pequeñas.** En manzanas con pocos habitantes un porcentaje cambia mucho con una o dos personas.
- **Entidades rurales.** Un mismo polígono de Thiessen puede reunir parcelas de agrado y viviendas campesinas, y el índice promedia ambas realidades.
- **Supresión y fusión del INE.** Las cifras de personas con menos de 4 casos vienen suprimidas, y las manzanas muy pequeñas fueron fusionadas con vecinas.
- **Personas sin manzana.** Las personas en situación de calle y las censadas en operativos especiales solo están a nivel comunal y no forman parte del índice.

# 11. Archivos y reproducibilidad

- `indice_socioeconomico_manzanas_censo2024_formulas.xlsx` (esta carpeta). Los datos del INE van como insumo y todo lo demás son fórmulas de Excel. Contiene las hojas:
  - **Manzanas:** una fila por polígono, con identificación, conteos, los siete componentes, sus puntajes $z$, el número de componentes con dato, el índice, el percentil y los dos quintiles por comuna.
  - **Parámetros:** medias y desviaciones estándar de la sección 4.
  - **Cortes de quintiles:** cortes de la sección 7.
  - **Resumen por comuna:** tabla de la sección 8.
  - **Diccionario:** descripción de cada columna.
- El código `MANZENT` es la llave para cruzar el Excel con la cartografía censal del INE.
- Cálculo: `Análisis y metadata\Dofiles\indice_socioeconomico_manzanas.py` (genera el Excel) y `prep_manzanas_censo.py` (genera los datos del mapa). Ambos usan la misma definición. Se verificó que el índice del Excel coincide en los 6.369 polígonos con el que muestra el sitio.
