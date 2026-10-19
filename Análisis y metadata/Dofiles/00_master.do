global raiz "C:\Users\gonza\Dropbox\Proyectos personales\Estudio Puerto Varas"
local comunas "10101, 10105, 10109, 10301"

local base1 "$raiz/Datos/Mineduc/Directorio establecimientos"
local salida1 "$raiz/Datos/[merge] Panel Puerto Varas/panel_directorio.dta"
capture erase "`salida1'"
cd "`base1'"
local carpetas1 : dir "." dirs "*"
foreach carpeta of local carpetas1 {
	cd "`base1'"
	cd `carpeta'
	local archivo : dir "." files "*.csv"
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
	* mineduc siempre llama "agno" a esta columna en sus csv; la usamos como "anio" de aqui en adelante
	capture rename agno anio
	keep if inlist(cod_com_rbd, `comunas')
	* latitud/longitud: la mayoria de los anios usa coma decimal (2017 en adelante) mas blancos; 2015-2016 vienen limpios en numerico;
	* 2020 es el unico anio con coma en CERO filas (100% punto decimal) + blancos. "dpcomma" trata el punto como separador de miles
	* SIEMPRE (no solo cuando hay coma), lo que corrompia por completo un anio 100% en formato punto como 2020. Por eso NO se usa
	* dpcomma: se reemplaza la coma por punto a mano (no hace nada si no hay coma) y se destringea normal.
	capture confirm string variable latitud
	if !_rc {
		replace latitud = subinstr(latitud, ",", ".", .)
		destring latitud, replace force
	}
	capture confirm string variable longitud
	if !_rc {
		replace longitud = subinstr(longitud, ",", ".", .)
		destring longitud, replace force
	}
	keep anio rbd nom_rbd cod_com_rbd nom_com_rbd cod_depe cod_depe2 rural_rbd matricula estado_estab pago_matricula pago_mensual latitud longitud
	capture confirm file "`salida1'"
	if _rc == 0 append using "`salida1'"
	save "`salida1'", replace
}
use "`salida1'", clear
label define depe2 1 "Municipal" 2 "Particular Subvencionado" 3 "Particular Pagado" 4 "Corp. Adm. Delegada" 5 "Servicio Local de Educacion"
label values cod_depe2 depe2
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno"
label values cod_com_rbd comuna
save "`salida1'", replace

local base2 "$raiz/Datos/Mineduc/Resumen matricula establecimiento"
local salida2 "$raiz/Datos/[merge] Panel Puerto Varas/panel_resumen_matricula.dta"
capture erase "`salida2'"
cd "`base2'"
local carpetas2 : dir "." dirs "*"
foreach carpeta of local carpetas2 {
	cd "`base2'"
	cd `carpeta'
	local archivo : dir "." files "*.csv"
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
	capture rename agno anio
	keep if inlist(cod_com_rbd, `comunas')
	* 2004-2009 no trae cod_depe2 y 2004-2014 no trae estado_estab (aparecen recien desde 2010 y 2015 resp.)
	capture confirm variable cod_depe2
	if _rc gen cod_depe2 = .
	capture confirm variable estado_estab
	if _rc gen estado_estab = .
	keep anio rbd nom_rbd cod_com_rbd nom_com_rbd cod_depe cod_depe2 rural_rbd estado_estab mat_ens_1 mat_ens_2 mat_ens_3 mat_ens_4 mat_ens_5 mat_ens_6 mat_ens_7 mat_ens_8 mat_hom_tot mat_muj_tot mat_total
	capture confirm file "`salida2'"
	if _rc == 0 append using "`salida2'"
	save "`salida2'", replace
}
use "`salida2'", clear
label define depe2 1 "Municipal" 2 "Particular Subvencionado" 3 "Particular Pagado" 4 "Corp. Adm. Delegada" 5 "Servicio Local de Educacion"
label values cod_depe2 depe2
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno"
label values cod_com_rbd comuna
label variable mat_ens_1 "Matricula Parvularia"
label variable mat_ens_2 "Matricula Basica Regular"
label variable mat_ens_3 "Matricula Basica Adulto"
label variable mat_ens_4 "Matricula Especial"
label variable mat_ens_5 "Matricula Media HC Jovenes"
label variable mat_ens_6 "Matricula Media HC Adultos"
label variable mat_ens_7 "Matricula Media TP Jovenes"
label variable mat_ens_8 "Matricula Media TP Adultos"
save "`salida2'", replace

local base3 "$raiz/Datos/Mineduc/Matricula por estudiante"
local salida3 "$raiz/Datos/[merge] Panel Puerto Varas/panel_matricula_estudiante.dta"
capture erase "`salida3'"
cd "`base3'"
local anios : dir "." dirs "*"
foreach anio of local anios {
	cd "`base3'"
	cd `anio'
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
	capture rename agno anio
	keep if inlist(cod_com_rbd, `comunas')
	* 2004-2014 no trae estado_estab (aparece recien desde 2015)
	capture confirm variable estado_estab
	if _rc gen estado_estab = .
	keep anio rbd nom_rbd cod_com_rbd nom_com_rbd cod_depe cod_depe2 rural_rbd estado_estab cod_ense cod_grado cod_grado2 mrun gen_alu fec_nac_alu edad_alu cod_com_alu nom_com_alu
	capture destring cod_com_alu, replace
	capture destring edad_alu, replace
	capture destring cod_grado, replace
	* 2004-2008 traen mrun en blanco para un puñado de registros a nivel nacional, lo que hace que Stata importe la columna como string
	capture destring mrun, replace force
	capture confirm file "`salida3'"
	if _rc == 0 append using "`salida3'"
	save "`salida3'", replace
}
use "`salida3'", clear
label define depe2 1 "Municipal" 2 "Particular Subvencionado" 3 "Particular Pagado" 4 "Corp. Adm. Delegada" 5 "Servicio Local de Educacion"
label values cod_depe2 depe2
label define comuna 10101 "Puerto Montt" 10105 "Frutillar" 10109 "Puerto Varas" 10301 "Osorno"
label values cod_com_rbd comuna
label define sexo 1 "Hombre" 2 "Mujer"
label values gen_alu sexo
save "`salida3'", replace
