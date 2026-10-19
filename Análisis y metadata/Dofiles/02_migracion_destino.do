* 02_migracion_destino.do
* Busca a que comuna (a nivel nacional) llegan los alumnos "salientes" de las 4 comunas del estudio.
* Requiere haber corrido antes 00_master.do y 01_analisis.do (usa panel_matricula_estudiante_etiquetado.dta).
* No descarga nada nuevo: los CSV de "Matricula por estudiante" ya son nacionales, el pipeline anterior
* solo los filtraba a las 4 comunas al importar. Aca se usan esos mismos CSV sin filtrar, pero cruzados
* contra una lista chica (los salientes de cada anio), no se arma un panel nacional completo.

global raiz "C:\Users\gonza\Dropbox\Proyectos personales\Estudio Puerto Varas"
local out "$raiz/Análisis y metadata/Outputs"
local nacional "$raiz/Datos/Mineduc/Matricula por estudiante"

* -------- 1. identificar salientes, mismo criterio que el grafico de flujo de 01_analisis.do --------
* saliente en el anio Y = presente en Y, ausente en Y+1, y no egreso de 4to medio en Y (curso==12)
use "`out'/panel_matricula_estudiante_etiquetado.dta", clear
keep mrun anio cod_com_rbd nom_com_rbd cod_depe2 curso
duplicates drop mrun anio, force
fillin mrun anio
xtset mrun anio

gen presente = !missing(curso)
gen presente_t2 = F.presente
gen saliente = (presente == 1 & presente_t2 == 0 & curso != 12) if !missing(presente_t2)

keep if saliente == 1 & anio >= 2015
gen anio_busqueda = anio + 1
rename cod_com_rbd cod_com_origen
rename nom_com_rbd nom_com_origen
rename cod_depe2 cod_depe2_origen
keep mrun anio anio_busqueda cod_com_origen nom_com_origen cod_depe2_origen
duplicates drop
save "`out'/_salientes_pendientes.dta", replace

* -------- 2. para cada anio de busqueda, importar el CSV nacional SIN filtrar y cruzar por mrun --------
capture erase "`out'/panel_destino_salientes.dta"
levelsof anio_busqueda, local(anios)
foreach anio of local anios {
	use "`out'/_salientes_pendientes.dta", clear
	keep if anio_busqueda == `anio'
	count
	if r(N) == 0 continue

	tempfile buscar
	save "`buscar'"

	cd "`nacional'/`anio'"
	local archivo : dir "." files "*.csv"
	if "`archivo'" == "" local archivo : dir "." files "*.CSV"
	import delimited `archivo', clear
	ds
	local vlist `r(varlist)'
	local hasagno 0
	foreach v of local vlist {
		if lower("`v'") == "agno" local hasagno 1
	}
	if !`hasagno' {
		local v1 : word 1 of `vlist'
		rename `v1' agno
	}
	rename *, lower
	capture destring mrun, replace force
	keep mrun rbd nom_rbd cod_com_rbd nom_com_rbd cod_depe2
	bysort mrun (rbd): keep if _n == 1
	tempfile nacional_anio
	save "`nacional_anio'"

	use "`buscar'", clear
	merge 1:1 mrun using "`nacional_anio'", keep(master match) keepusing(cod_com_rbd nom_com_rbd cod_depe2)
	gen encontrado = _merge == 3
	rename cod_com_rbd cod_com_destino
	rename nom_com_rbd nom_com_destino
	rename cod_depe2 cod_depe2_destino
	drop _merge

	capture confirm file "`out'/panel_destino_salientes.dta"
	if _rc == 0 append using "`out'/panel_destino_salientes.dta"
	save "`out'/panel_destino_salientes.dta", replace
}
capture erase "`out'/_salientes_pendientes.dta"

* -------- 3. resumen: top 5 comunas de destino por comuna de origen y anio --------
use "`out'/panel_destino_salientes.dta", clear
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno", replace
label values cod_com_origen comuna
label values cod_com_destino comuna

tab cod_com_origen encontrado, row
* "encontrado==0" = no aparece matriculado en NINGUN colegio del pais al anio siguiente
* (probable salida del sistema escolar subvencionado/publico: colegio particular pagado fuera
* de las 4 comunas sin reporte a Mineduc con este identificador, cambio de pais, o abandono escolar)

preserve
keep if encontrado == 1
contract cod_com_origen anio nom_com_destino, freq(n)
gsort cod_com_origen anio -n
bysort cod_com_origen anio: gen rank = _n
keep if rank <= 5
list cod_com_origen anio rank nom_com_destino n, sepby(cod_com_origen) noobs
restore

save "`out'/panel_destino_salientes.dta", replace
