# Removes the per-user registration created by install_mathtype_addin.ps1.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$classId = '{C4B5A18E-3427-49D9-94D4-36C8AF8B5F61}'
$progId = 'WordMcpLive.MathTypeAddIn'

$keys = @(
    "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\$progId",
    "Registry::HKEY_CURRENT_USER\Software\Classes\$progId",
    "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$classId"
)
foreach ($key in $keys) {
    if (Test-Path -LiteralPath $key) {
        Remove-Item -LiteralPath $key -Recurse -Force
    }
}
Write-Output 'MathType Word add-in registration removed. Restart Word to unload it.'
