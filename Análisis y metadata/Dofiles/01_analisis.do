global raiz "C:\Users\gonza\Dropbox\Proyectos personales\Estudio Puerto Varas"
local panel "$raiz/Datos/[merge] Panel Puerto Varas"
local out "$raiz/Análisis y metadata/Outputs"
local xlsx "`out'/indicadores_puerto_varas.xlsx"

* descripcion general de colegios y matricula (todos los niveles), desde el resumen por establecimiento
* la base ya trae 2004-2025, pero reportamos desde 2015 (antes faltan variables clave y no es el foco del encargo)
use "`panel'/panel_resumen_matricula.dta", clear
keep if anio >= 2015
label define depe2 1 "Municipal" 2 "Particular Subvencionado" 3 "Particular Pagado" 4 "Corp. Adm. Delegada" 5 "Servicio Local de Educacion", replace
label values cod_depe2 depe2
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno", replace
label values cod_com_rbd comuna

tab cod_com_rbd cod_depe2 if anio == 2025 & mat_total > 0
table cod_com_rbd cod_depe2 if anio == 2025 & mat_total > 0, statistic(sum mat_total)
tab cod_com_rbd rural_rbd if anio == 2025 & mat_total > 0
table cod_com_rbd rural_rbd anio if mat_total > 0, statistic(frequency)

preserve
collapse (sum) matricula = mat_total, by(cod_com_rbd anio)
label values cod_com_rbd comuna
table cod_com_rbd anio, statistic(sum matricula)
restore

preserve
keep if mat_total > 0
collapse (count) n_colegios = rbd (sum) matricula = mat_total, by(cod_com_rbd nom_com_rbd cod_depe2 anio)
label values cod_com_rbd comuna
label values cod_depe2 depe2
export excel using "`xlsx'", sheet("colegios_matricula") firstrow(variables) replace
save "`out'/01_colegios_matricula_por_comuna.dta", replace
restore




* base de alumnos: tipos corregidos, todo etiquetado, curso 1B-4M armado
* se mantiene 2014 como colchon para que los indicadores de entrada/salida del primer anio reportado (2015) calculen bien el rezago; se filtra a 2015+ mas abajo, despues de calcular los rezagos
use "`panel'/panel_matricula_estudiante.dta", clear
keep if anio >= 2014
capture destring edad_alu, replace
capture destring cod_grado, replace
capture destring cod_com_alu, replace

label define depe1 1 "Corporacion Municipal" 2 "Municipal DAEM" 3 "Particular Subvencionado" 4 "Particular Pagado" 5 "Corp. Adm. Delegada" 6 "Servicio Local de Educacion", replace
label values cod_depe depe1
label define depe2 1 "Municipal" 2 "Particular Subvencionado" 3 "Particular Pagado" 4 "Corp. Adm. Delegada" 5 "Servicio Local de Educacion", replace
label values cod_depe2 depe2
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno", replace
label values cod_com_rbd comuna
label define sexo1 1 "Hombre" 2 "Mujer", replace
label values gen_alu sexo1
label define rural 0 "Urbano" 1 "Rural", replace
label values rural_rbd rural
label define estado 1 "Funcionando" 2 "En receso" 3 "Cerrado" 4 "Autorizado sin matricula", replace
label values estado_estab estado

gen curso = .
replace curso = cod_grado if cod_ense == 110
replace curso = cod_grado + 8 if inlist(cod_ense, 310, 410, 510, 610, 710, 810)
label define curso_lbl 1 "1 Basico" 2 "2 Basico" 3 "3 Basico" 4 "4 Basico" 5 "5 Basico" 6 "6 Basico" 7 "7 Basico" 8 "8 Basico" 9 "1 Medio" 10 "2 Medio" 11 "3 Medio" 12 "4 Medio", replace
label values curso curso_lbl

gen mismo_comuna = cod_com_rbd == cod_com_alu if !missing(cod_com_alu) & cod_com_alu != 0
label variable mismo_comuna "Alumno vive en la misma comuna del colegio"

save "`out'/panel_matricula_estudiante_etiquetado.dta", replace

tab curso cod_com_rbd if anio == 2025
tab curso cod_depe2 if anio == 2025

tab anio cod_com_rbd if cod_depe == 2 & anio >= 2015
tab anio cod_com_rbd if cod_depe == 3 & anio >= 2015
tab anio cod_com_rbd if cod_depe == 4 & anio >= 2015
tab anio cod_com_rbd if cod_depe == 6 & anio >= 2015

* indicador: matricula por curso y establecimiento, con desglose por sexo
preserve
keep if anio >= 2015
keep if !missing(curso)
gen nina = gen_alu == 2
gen nino = gen_alu == 1
collapse (count) matricula = mrun (sum) ninas = nina ninos = nino, by(rbd nom_rbd cod_com_rbd nom_com_rbd cod_depe2 anio curso)
label values cod_com_rbd comuna
label values cod_depe2 depe2
label values curso curso_lbl
table cod_com_rbd curso, statistic(sum matricula)
restore

* indicador: perfil demografico por comuna, dependencia y año
preserve
keep if anio >= 2015
gen es_hombre = gen_alu == 1 if gen_alu != 0
gen edad = edad_alu if edad_alu < 90
collapse (count) n_alumnos = mrun (mean) pct_mismo_comuna = mismo_comuna pct_hombres = es_hombre edad_prom = edad, by(cod_com_rbd nom_com_rbd cod_depe2 anio)
label values cod_com_rbd comuna
label values cod_depe2 depe2
list cod_com_rbd nom_com_rbd cod_depe2 anio n_alumnos pct_mismo_comuna pct_hombres edad_prom, sepby(cod_com_rbd) noobs
restore


preserve
keep mrun anio cod_com_rbd curso
duplicates drop mrun anio, force
fillin mrun anio
xtset mrun anio

gen presente = !missing(curso)
gen presente_t1 = L.presente
gen presente_t2 = F.presente

gen mantiene = (presente == 1 & presente_t1 == 1)
gen entrante = (presente == 1 & presente_t1 == 0 & curso != 1)
gen ingreso_1basico = (presente == 1 & presente_t1 == 0 & curso == 1)
gen saliente = (presente == 1 & presente_t2 == 0 & curso != 12) if !missing(presente_t2)

keep if presente == 1
keep if anio >= 2015
collapse (sum) mantiene entrante ingreso_1basico saliente (count) total = mrun, by(cod_com_rbd anio)
label values cod_com_rbd comuna

list cod_com_rbd anio total mantiene entrante ingreso_1basico saliente, sepby(cod_com_rbd) noobs
restore








**DESCRIPCIÓN MIGRATORIA

* origen: alumnos que llegaron de otra comuna, por año y por dependencia de entrada
preserve
keep mrun anio cod_com_rbd cod_depe2 curso
duplicates drop mrun anio, force
fillin mrun anio
xtset mrun anio

gen presente = !missing(curso)
gen presente_t1 = L.presente
gen comuna_t1 = L.cod_com_rbd
gen curso_t1 = L.curso
gen depe_t1 = L.cod_depe2

gen llego = (presente == 1 & presente_t1 == 0 & curso != 1) if !missing(presente_t1)
gen se_fue = (presente == 0 & presente_t1 == 1 & curso_t1 != 12) if !missing(presente_t1)

label values cod_com_rbd comuna
label values comuna_t1 comuna
label values cod_depe2 depe2
label values depe_t1 depe2

tab cod_depe2 cod_com_rbd if llego == 1 & anio >= 2015
tab depe_t1 comuna_t1 if se_fue == 1 & anio >= 2015
table cod_com_rbd cod_depe2 anio if llego == 1 & anio >= 2015, statistic(frequency)
table comuna_t1 depe_t1 anio if se_fue == 1 & anio >= 2015, statistic(frequency)
restore

**comunas de ingreso
preserve
keep mrun anio cod_com_rbd cod_com_alu curso
duplicates drop mrun anio, force
fillin mrun anio
xtset mrun anio

gen presente = !missing(curso)
gen presente_t1 = L.presente
gen llego = (presente == 1 & presente_t1 == 0 & curso != 1) if !missing(presente_t1)

gen origen = .
replace origen = 1 if cod_com_alu == cod_com_rbd & !missing(cod_com_alu)
replace origen = 2 if inlist(cod_com_alu, 10101, 10105, 10109, 10301) & cod_com_alu != cod_com_rbd
replace origen = 3 if !missing(cod_com_alu) & cod_com_alu != 0 & !inlist(cod_com_alu, 10101, 10105, 10109, 10301)
replace origen = 4 if missing(cod_com_alu) | cod_com_alu == 0

label define origen_lbl 1 "Misma comuna" 2 "Otra de las 4 comunas" 3 "Fuera de las 4 comunas" 4 "Sin informacion", replace
label values origen origen_lbl
label values cod_com_rbd comuna

table cod_com_rbd origen anio if llego == 1 & anio >= 2015, statistic(frequency)
restore

**top 5 comunas 
preserve
keep mrun anio cod_com_rbd cod_com_alu nom_com_alu curso
duplicates drop mrun anio, force
fillin mrun anio
xtset mrun anio

gen presente = !missing(curso)
gen presente_t1 = L.presente
gen llego = (presente == 1 & presente_t1 == 0 & curso != 1) if !missing(presente_t1)

keep if llego == 1 & anio >= 2015 & !missing(cod_com_alu) & cod_com_alu != 0 & cod_com_alu != cod_com_rbd
label values cod_com_rbd comuna
contract cod_com_rbd anio nom_com_alu, freq(n)
gsort cod_com_rbd anio -n
bysort cod_com_rbd anio: gen rank = _n
keep if rank <= 5
list cod_com_rbd anio rank nom_com_alu n, sepby(cod_com_rbd) noobs
restore

** establecimientos
use "`panel'/panel_directorio.dta", clear
keep if anio >= 2015
label values cod_com_rbd comuna
label values cod_depe2 depe2
table cod_com_rbd cod_depe2 anio if matricula == 1, statistic(frequency)

label define rural 0 "Urbano" 1 "Rural", replace
label values rural_rbd rural
table cod_com_rbd rural_rbd anio if matricula == 1, statistic(frequency)


use "`panel'/panel_resumen_matricula.dta", clear
keep if anio >= 2015
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno", replace
label values cod_com_rbd comuna
collapse (sum) matricula = mat_total, by(cod_com_rbd anio)
label values cod_com_rbd comuna
table cod_com_rbd anio, statistic(sum matricula)


