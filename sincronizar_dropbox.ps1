# Copia al repo lo que cambió en la carpeta del proyecto en Dropbox.
# Trae "Análisis y metadata" completo; lo que no se publica (microdatos,
# archivos temporales) lo filtra .gitignore. La carpeta Datos no se copia.
#
# Uso, desde la raíz del repo:
#   .\sincronizar_dropbox.ps1 -Dropbox "C:\Users\<usuario>\Dropbox\Proyectos personales\Estudio Puerto Varas"
# Después revisar con `git status` antes de hacer commit.

param(
    [Parameter(Mandatory = $true)]
    [string]$Dropbox
)

$repo = $PSScriptRoot
$carpeta = 'Análisis y metadata'

if (-not (Test-Path (Join-Path $Dropbox $carpeta))) {
    Write-Error "No encuentro '$carpeta' dentro de $Dropbox"
    exit 1
}

# /E incluye subcarpetas; /XO no pisa archivos más nuevos del repo;
# /XF excluye los archivos de bloqueo de Office y los microdatos.
robocopy (Join-Path $Dropbox $carpeta) (Join-Path $repo $carpeta) /E /XO /NJH /NJS /NP `
    /XF '~$*' 'panel_matricula_estudiante_etiquetado.dta' 'panel_destino_salientes.dta' '_*.dta'

# robocopy devuelve 0-7 cuando no hay errores
if ($LASTEXITCODE -ge 8) { exit $LASTEXITCODE } else { exit 0 }
