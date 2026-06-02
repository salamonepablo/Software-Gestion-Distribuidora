#requires -Version 7.0
<#
.SYNOPSIS
Diagnostica y repara problemas de VB6 MSComCtl2.MonthView (MSCOMCT2.OCX) tras cambios regionales.
.DESCRIPTION
Script seguro, idempotente y verboso para:
- Verificar privilegios de administrador.
- Verificar existencia de MSCOMCT2.OCX en SysWOW64 y System32.
- Registrar (o re-registrar) el OCX con regsvr32 de SysWOW64.
- Verificar claves COM esperadas.
- Mostrar configuración regional/cultural y formato internacional.
- Opcionalmente corregir formato corto de fecha para el usuario actual.
.PARAMETER ReRegister
Desregistra y vuelve a registrar el OCX (usar si registro actual está dañado).
.PARAMETER FixRegional
Configura para el usuario actual:
- sShortDate = dd/MM/yyyy
- sDate      = /
Puede requerir cerrar sesión y volver a entrar.
.EXAMPLE
pwsh -File .\Diagnostico-MonthView.ps1
.EXAMPLE
pwsh -File .\Diagnostico-MonthView.ps1 -ReRegister
.EXAMPLE
pwsh -File .\Diagnostico-MonthView.ps1 -FixRegional
.EXAMPLE
pwsh -File .\Diagnostico-MonthView.ps1 -ReRegister -FixRegional
#>
[CmdletBinding()]
param(
    [switch]$ReRegister,
    [switch]$FixRegional
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
function Write-Section {
    param([string]$Text)
    Write-Host ""
    Write-Host "==== $Text ====" -ForegroundColor Cyan
}
function Write-Result {
    param(
        [ValidateSet('PASS','WARN','FAIL')]
        [string]$Level,
        [string]$Message
    )
    $color = switch ($Level) {
        'PASS' { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
    }
    Write-Host ("[{0}] {1}" -f $Level, $Message) -ForegroundColor $color
}
function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}
function Invoke-Regsvr32 {
    param(
        [ValidateSet('Register','Unregister')]
        [string]$Action,
        [Parameter(Mandatory)]
        [string]$Regsvr32Path,
        [Parameter(Mandatory)]
        [string]$OcxPath
    )
    $args = if ($Action -eq 'Unregister') { '/s /u "{0}"' -f $OcxPath } else { '/s "{0}"' -f $OcxPath }
    try {
        $proc = Start-Process -FilePath $Regsvr32Path -ArgumentList $args -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            Write-Result PASS ("regsvr32 {0} correcto. ExitCode=0" -f $Action)
            return $true
        } else {
            Write-Result FAIL ("regsvr32 {0} falló. ExitCode={1}" -f $Action, $proc.ExitCode)
            return $false
        }
    } catch {
        Write-Result FAIL ("Error ejecutando regsvr32 ({0}): {1}" -f $Action, $_.Exception.Message)
        return $false
    }
}
$results = [ordered]@{
    Admin               = 'WARN'
    OcxSysWOW64         = 'WARN'
    OcxSystem32         = 'WARN'
    RegisterAction      = 'WARN'
    ComProgId           = 'WARN'
    ComClsid            = 'WARN'
    RegionalRead        = 'WARN'
    RegionalFix         = 'WARN'
}
Write-Section "Inicio del diagnóstico MSCOMCT2.MonthView"
Write-Host "Fecha/Hora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "PowerShell: $($PSVersionTable.PSVersion)"
Write-Host "Usuario: $env:USERNAME"
Write-Host "Equipo: $env:COMPUTERNAME"
Write-Section "1) Verificar elevación (Administrador)"
$isAdmin = Test-IsAdmin
if ($isAdmin) {
    Write-Result PASS "La sesión está elevada (Administrador)."
    $results.Admin = 'PASS'
} else {
    Write-Result FAIL "La sesión NO está elevada. Regsvr32 puede fallar por permisos."
    Write-Host "Sugerencia: abrir PowerShell 7 como Administrador y volver a ejecutar."
    $results.Admin = 'FAIL'
}
Write-Section "2) Verificar existencia de MSCOMCT2.OCX"
$ocxSysWOW64 = Join-Path $env:WINDIR 'SysWOW64\MSCOMCT2.OCX'
$ocxSystem32 = Join-Path $env:WINDIR 'System32\MSCOMCT2.OCX'
$regsvr32SysWOW64 = Join-Path $env:WINDIR 'SysWOW64\regsvr32.exe'
if (Test-Path -LiteralPath $ocxSysWOW64) {
    Write-Result PASS "Encontrado: $ocxSysWOW64"
    $results.OcxSysWOW64 = 'PASS'
} else {
    Write-Result FAIL "No encontrado: $ocxSysWOW64"
    $results.OcxSysWOW64 = 'FAIL'
}
if (Test-Path -LiteralPath $ocxSystem32) {
    Write-Result WARN "Encontrado en System32: $ocxSystem32 (para VB6 32-bit importa SysWOW64)."
    $results.OcxSystem32 = 'WARN'
} else {
    Write-Result WARN "No encontrado en System32: $ocxSystem32 (puede ser normal en algunos equipos)."
    $results.OcxSystem32 = 'WARN'
}
if (-not (Test-Path -LiteralPath $regsvr32SysWOW64)) {
    Write-Result FAIL "No existe regsvr32 esperado: $regsvr32SysWOW64"
    $results.RegisterAction = 'FAIL'
}
Write-Section "3) Registrar OCX (SysWOW64 regsvr32)"
$didRegister = $false
if ((Test-Path -LiteralPath $ocxSysWOW64) -and (Test-Path -LiteralPath $regsvr32SysWOW64)) {
    if ($ReRegister) {
        Write-Host "Modo -ReRegister activo: primero desregistrar, luego registrar."
        $okUnreg = Invoke-Regsvr32 -Action Unregister -Regsvr32Path $regsvr32SysWOW64 -OcxPath $ocxSysWOW64
        $okReg   = Invoke-Regsvr32 -Action Register   -Regsvr32Path $regsvr32SysWOW64 -OcxPath $ocxSysWOW64
        $didRegister = $okUnreg -and $okReg
    } else {
        Write-Host "Registro estándar: registrar sin desregistrar (idempotente)."
        $didRegister = Invoke-Regsvr32 -Action Register -Regsvr32Path $regsvr32SysWOW64 -OcxPath $ocxSysWOW64
    }
    $results.RegisterAction = if ($didRegister) { 'PASS' } else { 'FAIL' }
} elseif ($results.RegisterAction -ne 'FAIL') {
    Write-Result FAIL "No se puede registrar: falta OCX en SysWOW64 o regsvr32 en SysWOW64."
    $results.RegisterAction = 'FAIL'
}
Write-Section "4) Verificar claves COM requeridas"
$progIdPath = 'Registry::HKEY_CLASSES_ROOT\MSComCtl2.MonthView.2'
try {
    if (Test-Path -LiteralPath $progIdPath) {
        Write-Result PASS "Existe ProgID: HKCR\MSComCtl2.MonthView.2"
        $results.ComProgId = 'PASS'
        # Resolver CLSID REAL desde el ProgID (sin hardcodear)
        $monthViewClsid = (Get-ItemProperty -Path "$progIdPath\CLSID" -ErrorAction Stop).'(default)'
        Write-Host "CLSID resuelto desde ProgID: $monthViewClsid"
        # Validar InprocServer32 en vista 32-bit (clave para VB6)
        $inproc32 = & reg query "HKCR\CLSID\$monthViewClsid\InprocServer32" /reg:32 /ve 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Result PASS "Existe InprocServer32 (vista 32-bit) para $monthViewClsid"
            $results.ComClsid = 'PASS'
            if ($inproc32 -match 'MSCOMCT2\.OCX') {
                Write-Result PASS "InprocServer32 apunta a MSCOMCT2.OCX"
            } else {
                Write-Result WARN "InprocServer32 existe pero no referencia claramente MSCOMCT2.OCX"
            }
        } else {
            Write-Result FAIL "No existe InprocServer32 (vista 32-bit) para $monthViewClsid"
            $results.ComClsid = 'FAIL'
        }
    } else {
        Write-Result FAIL "No existe ProgID: HKCR\MSComCtl2.MonthView.2"
        $results.ComProgId = 'FAIL'
        $results.ComClsid = 'FAIL'
    }
} catch {
    Write-Result FAIL ("Error verificando ProgID/CLSID dinámico: {0}" -f $_.Exception.Message)
    $results.ComProgId = 'FAIL'
    $results.ComClsid = 'FAIL'
}
Write-Section "5) Diagnóstico regional/cultural"
try {
    $culture = Get-Culture
    $uiCulture = Get-UICulture
    $intlPath = 'HKCU:\Control Panel\International'
    $intl = Get-ItemProperty -Path $intlPath
    Write-Host "Culture actual         : $($culture.Name) - $($culture.DisplayName)"
    Write-Host "UI Culture actual      : $($uiCulture.Name) - $($uiCulture.DisplayName)"
    Write-Host "sShortDate (HKCU)      : $($intl.sShortDate)"
    Write-Host "sDate separador (HKCU) : $($intl.sDate)"
    Write-Host "sDecimal (HKCU)        : $($intl.sDecimal)"
    Write-Host "sThousand (HKCU)       : $($intl.sThousand)"
    $results.RegionalRead = 'PASS'
} catch {
    Write-Result FAIL ("No se pudo leer configuración regional: {0}" -f $_.Exception.Message)
    $results.RegionalRead = 'FAIL'
}
Write-Section "6) Aplicar corrección regional opcional (-FixRegional)"
if ($FixRegional) {
    try {
        $intlPath = 'HKCU:\Control Panel\International'
        $current = Get-ItemProperty -Path $intlPath
        $needShort = $current.sShortDate -ne 'dd/MM/yyyy'
        $needSep   = $current.sDate -ne '/'
        if (-not $needShort -and -not $needSep) {
            Write-Result PASS "La configuración regional ya está en los valores deseados. No se hicieron cambios."
            $results.RegionalFix = 'PASS'
        } else {
            if ($needShort) {
                Set-ItemProperty -Path $intlPath -Name sShortDate -Value 'dd/MM/yyyy'
                Write-Result PASS "Actualizado sShortDate -> dd/MM/yyyy"
            } else {
                Write-Host "sShortDate ya estaba correcto."
            }
            if ($needSep) {
                Set-ItemProperty -Path $intlPath -Name sDate -Value '/'
                Write-Result PASS "Actualizado sDate -> /"
            } else {
                Write-Host "sDate ya estaba correcto."
            }
            Write-Result WARN "Cambios regionales aplicados. Puede requerir cerrar sesión y volver a entrar."
            $results.RegionalFix = 'WARN'
        }
    } catch {
        Write-Result FAIL ("No se pudo aplicar -FixRegional: {0}" -f $_.Exception.Message)
        $results.RegionalFix = 'FAIL'
    }
} else {
    Write-Host "No se solicitó -FixRegional. Se omite modificación regional."
    $results.RegionalFix = 'WARN'
}
Write-Section "7) Resumen final y próximos pasos"
$hasFail = $results.Values -contains 'FAIL'
$hasWarn = $results.Values -contains 'WARN'
$globalStatus = if ($hasFail) { 'FAIL' } elseif ($hasWarn) { 'WARN' } else { 'PASS' }
Write-Result $globalStatus "Estado global: $globalStatus"
Write-Host ""
Write-Host "Detalle:"
$results.GetEnumerator() | ForEach-Object {
    Write-Host (" - {0}: {1}" -f $_.Key, $_.Value)
}
Write-Host ""
if ($globalStatus -eq 'PASS') {
    Write-Host "Próximas acciones recomendadas:"
    Write-Host "  1) Abrir SPC-Core y validar el control MonthView en la pantalla afectada."
    Write-Host "  2) Si persiste error, reiniciar Windows y revalidar."
} elseif ($globalStatus -eq 'WARN') {
    Write-Host "Próximas acciones recomendadas:"
    Write-Host "  1) Revisar advertencias (especialmente regionales) y cerrar sesión si aplicaste -FixRegional."
    Write-Host "  2) Ejecutar nuevamente con -ReRegister si aún hay fallas de inicialización."
    Write-Host "  3) Validar en SPC-Core el formulario que usa MonthView."
} else {
    Write-Host "Próximas acciones recomendadas:"
    Write-Host "  1) Ejecutar PowerShell 7 como Administrador y repetir."
    Write-Host "  2) Confirmar que existe '$ocxSysWOW64'."
    Write-Host "  3) Si falta el OCX, reinstalarlo desde instalador confiable y volver a registrar."
    Write-Host "  4) Verificar antivirus/políticas que bloqueen regsvr32."
}
Write-Host ""
Write-Host "DISCLAIMER: Si falta MSCOMCT2.OCX, descargalo SOLO de fuentes confiables y verificadas (instalador oficial o repositorio corporativo). Evitá sitios no oficiales."