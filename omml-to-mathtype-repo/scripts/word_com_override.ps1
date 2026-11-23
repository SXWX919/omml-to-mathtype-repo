param(
    [ValidateSet('status', 'apply', 'revert')][string]$Action = 'status'
)
# Diagnose / temporarily redirect the per-user Word.Application COM class.
#
# Problem: on machines with WPS Office, HKLM 32-bit CLSID {000209FF-...} points to
# wps.exe /Automation, so every `New-Object -ComObject Word.Application` reaches WPS
# instead of Microsoft Word. While a WPS automation server is running it owns the COM
# class object and a registry override is ignored - close WPS completely first.
#
# apply  : write HKCU override -> real WINWORD.EXE, then create the object and print
#          which executable answered.
# revert : remove the HKCU override.
# status : show registrations in both registry views and the current live object.

$key = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$candidates = @(
    'C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE',
    'C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE',
    'C:\Program Files\Microsoft Office\Office16\WINWORD.EXE'
)

function Show-Status {
    "--- HKLM 64-bit ---"
    $k = 'HKLM:\SOFTWARE\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
    if (Test-Path $k) { (Get-ItemProperty $k).'(default)' } else { '(none)' }
    "--- HKLM 32-bit ---"
    $k = 'HKLM:\SOFTWARE\WOW6432Node\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
    if (Test-Path $k) { (Get-ItemProperty $k).'(default)' } else { '(none)' }
    "--- HKCU override ---"
    if (Test-Path $key) { (Get-ItemProperty $key).'(default)' } else { '(none)' }
    "--- live object ---"
    try {
        $app = New-Object -ComObject Word.Application
        "Name=$($app.Name) Path=$($app.Path) Version=$($app.Version)"
    }
    catch { "cannot create: $($_.Exception.Message)" }
}

switch ($Action) {
    'status' { Show-Status }
    'apply' {
        $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
        if (-not $exe) { "no WINWORD.EXE found in the usual Office paths"; exit 1 }
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name '(default)' -Value ('"' + $exe + '" /Automation')
        "override -> $exe"
        $wps = Get-Process wps -ErrorAction SilentlyContinue
        if ($wps) { "WARNING: WPS is still running; close it or the override is ignored." }
        Show-Status
    }
    'revert' {
        if (Test-Path $key) {
            Remove-Item $key -Force
            "removed $key"
        }
        else { "override not present" }
    }
}
