# RUA EZ-Fac: instala o actualiza la app desde GitHub con una sola linea. En PowerShell:
#
#   irm https://raw.githubusercontent.com/tk3piernas/rua-ez-fac/main/instalar.ps1 | iex
#
# Baja la ultima version, la copia en Documentos\RUA EZ-Fac (nunca adentro de OneDrive), le saca la marca de
# "bajado de internet" y corre "Instalar o actualizar RUA EZ-Fac.bat". Si la app ya estaba, conserva
# config.toml, datos\ y .venv\ (no vienen en la descarga, asi que no se pisan).
#
# Para probarlo sin tocar nada real: $env:RUA_EZFAC_DESTINO (otra carpeta), $env:RUA_EZFAC_ZIP (un zip local en vez
# de bajarlo) y $env:RUA_EZFAC_SIN_INSTALAR = "1" (copia los archivos pero no corre el instalador).

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'          # Invoke-WebRequest is much slower with the progress bar
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$Repo = 'https://github.com/tk3piernas/rua-ez-fac/archive/refs/heads/main.zip'
$Nombre = 'RUA EZ-Fac'

function Salir-ConError([string]$texto) {
    Write-Host ''
    Write-Host "  $texto" -ForegroundColor Red
    Write-Host ''
    Read-Host '  Apreta Enter para cerrar' | Out-Null
    throw $texto
}

function Elegir-Destino {
    if ($env:RUA_EZFAC_DESTINO) { return $env:RUA_EZFAC_DESTINO }
    $docs = [Environment]::GetFolderPath('MyDocuments')
    # Documentos puede estar adentro de OneDrive (copia de seguridad de carpetas): el OneDrive de RUA no deja
    # guardar programas, asi que en ese caso va a la carpeta local del usuario.
    if (-not $docs -or $docs -match 'OneDrive') {
        $docs = Join-Path $env:USERPROFILE 'Documents'
        if ((-not (Test-Path $docs)) -or ((Get-Item $docs).Target -match 'OneDrive')) {
            return (Join-Path $env:LOCALAPPDATA $Nombre)
        }
    }
    # una instalacion anterior en "Documentos\App Facturas" se sigue usando
    $vieja = Join-Path $docs 'App Facturas'
    if (Test-Path (Join-Path $vieja 'Instalar o actualizar RUA EZ-Fac.bat')) { return $vieja }
    return (Join-Path $docs $Nombre)
}

Write-Host ''
Write-Host '  ==================================================' -ForegroundColor Cyan
Write-Host '    RUA EZ-Fac - instalar o actualizar desde GitHub' -ForegroundColor Cyan
Write-Host '  ==================================================' -ForegroundColor Cyan

$destino = Elegir-Destino
if ($destino -match 'OneDrive') {
    Salir-ConError "La carpeta $destino esta en OneDrive y el OneDrive de RUA no deja guardar programas."
}
if (Test-Path (Join-Path $destino '.git')) {
    Salir-ConError "$destino es una copia de Git (la de desarrollo): actualizala con git pull, no con este script."
}
Write-Host "  Carpeta: $destino"

$tmp = Join-Path $env:TEMP ("rua-ez-fac-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $zip = Join-Path $tmp 'app.zip'
    if ($env:RUA_EZFAC_ZIP) {
        Copy-Item $env:RUA_EZFAC_ZIP $zip
    } else {
        Write-Host '  [1/3] Bajando la ultima version ...'
        try {
            Invoke-WebRequest -Uri $Repo -OutFile $zip -UseBasicParsing
        } catch {
            Salir-ConError "No pude bajar la app de GitHub ($($_.Exception.Message)). Revisa internet, o baja el zip a mano: github.com/tk3piernas/rua-ez-fac -> Code -> Download ZIP."
        }
    }
    Write-Host '  [2/3] Copiando los archivos ...'
    Expand-Archive -Path $zip -DestinationPath $tmp -Force
    $origen = Get-ChildItem -Path $tmp -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'Instalar o actualizar RUA EZ-Fac.bat') } | Select-Object -First 1
    if (-not $origen) { Salir-ConError 'La descarga no tiene el instalador: avisale a Matheo.' }
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
    # robocopy /E copies new and changed files and leaves everything else (config.toml, datos, .venv) as is
    robocopy $origen.FullName $destino /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
    if ($LASTEXITCODE -ge 8) { Salir-ConError "No pude copiar los archivos a $destino (robocopy $LASTEXITCODE)." }
    $global:LASTEXITCODE = 0                   # 1-7 is robocopy's "copied fine", not an error
    Get-ChildItem -Path $destino -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
} finally {
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($env:RUA_EZFAC_SIN_INSTALAR -eq '1') {
    Write-Host '  Listo (sin correr el instalador).'
    return
}
Write-Host '  [3/3] Instalando (la primera vez tarda unos minutos) ...'
Write-Host ''
$bat = Join-Path $destino 'Instalar o actualizar RUA EZ-Fac.bat'
Push-Location $destino
try {
    & cmd.exe /c "`"$bat`""
} finally {
    Pop-Location
}
