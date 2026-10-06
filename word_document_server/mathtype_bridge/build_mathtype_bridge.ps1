# Compiles MathTypeBridge.exe (stdio helper) and MathTypeAddIn.dll (Word add-in)
# from the single C# source next to this script. Quit Word first: the DLL is
# locked while the add-in is loaded.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot 'MathTypeBridge.cs'
$outputDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin'
$output = Join-Path $outputDirectory 'MathTypeBridge.exe'
$addInOutput = Join-Path $outputDirectory 'MathTypeAddIn.dll'

$windowsDirectory = Split-Path -Parent ([Environment]::SystemDirectory)
$framework = $null
foreach ($candidate in @('Framework64', 'Framework')) {
    $path = Join-Path $windowsDirectory "Microsoft.NET\$candidate\v4.0.30319"
    if (Test-Path -LiteralPath (Join-Path $path 'csc.exe') -PathType Leaf) {
        $framework = $path
        break
    }
}
if (-not $framework) {
    throw '.NET Framework 4.x C# compiler (csc.exe) was not found.'
}
$compiler = Join-Path $framework 'csc.exe'

# Office Extensibility assembly lives in the GAC under a version-specific path.
$extensibility = Get-ChildItem -Path (Join-Path $windowsDirectory 'assembly'), (Join-Path $windowsDirectory 'Microsoft.NET\assembly') `
        -Recurse -Filter 'Extensibility.dll' -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $extensibility) {
    throw 'Microsoft Office Extensibility assembly was not found in the GAC.'
}

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

# anycpu: the EXE follows the OS, the DLL follows the hosting Word process
# (works for both 32- and 64-bit Word installs).
foreach ($target in @(
        @{ Kind = 'exe'; Out = $output },
        @{ Kind = 'library'; Out = $addInOutput })) {
    & $compiler /nologo /target:$($target.Kind) /platform:anycpu /optimize+ /utf8output `
        /out:$($target.Out) `
        /reference:"$framework\System.dll" `
        /reference:"$framework\System.Core.dll" `
        /reference:"$framework\System.Security.dll" `
        /reference:"$framework\System.Web.Extensions.dll" `
        /reference:"$framework\System.Xml.dll" `
        /reference:"$framework\Microsoft.CSharp.dll" `
        /reference:$extensibility `
        $source
    if ($LASTEXITCODE -ne 0) {
        throw "C# compiler exited with code $LASTEXITCODE"
    }
    Write-Output $target.Out
}
