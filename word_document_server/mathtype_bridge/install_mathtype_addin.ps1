# Registers the MathType Word add-in (per-user, no admin rights). Run once,
# then restart Word. Undo with uninstall_mathtype_addin.ps1.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$assemblyPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\MathTypeAddIn.dll'
$classId = '{C4B5A18E-3427-49D9-94D4-36C8AF8B5F61}'
$progId = 'WordMcpLive.MathTypeAddIn'
$className = 'WordMcpLive.MathTypeBridge.MathTypeAddIn'
$assemblyName = 'MathTypeAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'

if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
    throw "MathTypeAddIn.dll is missing. Run build_mathtype_bridge.ps1 (next to this script) first."
}

$codeBase = ([Uri](Resolve-Path -LiteralPath $assemblyPath).Path).AbsoluteUri
$classes = 'Registry::HKEY_CURRENT_USER\Software\Classes'
$progIdKey = Join-Path $classes $progId
$classKey = Join-Path $classes "CLSID\$classId"
$inprocKey = Join-Path $classKey 'InprocServer32'
$versionKey = Join-Path $inprocKey '1.0.0.0'

New-Item -Path $progIdKey -Force | Out-Null
Set-Item -Path $progIdKey -Value $className
New-Item -Path (Join-Path $progIdKey 'CLSID') -Force | Out-Null
Set-Item -Path (Join-Path $progIdKey 'CLSID') -Value $classId

New-Item -Path $classKey -Force | Out-Null
Set-Item -Path $classKey -Value $className
New-Item -Path $inprocKey -Force | Out-Null
Set-Item -Path $inprocKey -Value 'mscoree.dll'
Set-ItemProperty -Path $inprocKey -Name ThreadingModel -Value Both
Set-ItemProperty -Path $inprocKey -Name Class -Value $className
Set-ItemProperty -Path $inprocKey -Name Assembly -Value $assemblyName
Set-ItemProperty -Path $inprocKey -Name RuntimeVersion -Value 'v4.0.30319'
Set-ItemProperty -Path $inprocKey -Name CodeBase -Value $codeBase

New-Item -Path $versionKey -Force | Out-Null
Set-ItemProperty -Path $versionKey -Name Class -Value $className
Set-ItemProperty -Path $versionKey -Name Assembly -Value $assemblyName
Set-ItemProperty -Path $versionKey -Name RuntimeVersion -Value 'v4.0.30319'
Set-ItemProperty -Path $versionKey -Name CodeBase -Value $codeBase

New-Item -Path (Join-Path $classKey 'ProgId') -Force | Out-Null
Set-Item -Path (Join-Path $classKey 'ProgId') -Value $progId
New-Item -Path (Join-Path $classKey 'Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}') -Force | Out-Null

$officeKey = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\$progId"
New-Item -Path $officeKey -Force | Out-Null
Set-ItemProperty -Path $officeKey -Name FriendlyName -Value 'Word MCP Live MathType Bridge'
Set-ItemProperty -Path $officeKey -Name Description -Value 'In-process semantic MathML access for MathType OLE equations.'
Set-ItemProperty -Path $officeKey -Name LoadBehavior -Type DWord -Value 3
Set-ItemProperty -Path $officeKey -Name CommandLineSafe -Type DWord -Value 0

Write-Output 'MathType Word add-in registered for the current user. Restart Word to load it.'
