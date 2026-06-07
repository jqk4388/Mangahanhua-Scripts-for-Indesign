<#
.SYNOPSIS
    Launch an InDesign JSX script from PowerShell.

.DESCRIPTION
    Usage:
      .\run.ps1                    # default: run scriptRun.jsx in same folder as run.ps1
      .\run.ps1 "C:\path\to\script.jsx"  # absolute path to jsx file
      .\run.ps1 "myscript.jsx"     # relative path (relative to current working directory)
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string]$ScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-ScriptPath {
    param(
        [string]$InputPath
    )

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        return $null
    }

    if ($InputPath.Length -ge 2 -and $InputPath[1] -eq ':' -and $InputPath[0] -match '[A-Za-z]') {
        return $InputPath
    }

    if ($InputPath.StartsWith('\\')) {
        return $InputPath
    }

    if ($InputPath.StartsWith('/')) {
        return $InputPath
    }

    return Join-Path -Path (Get-Location) -ChildPath $InputPath
}

function Get-DefaultScriptPath {
    $scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
    return Join-Path -Path $scriptFolder -ChildPath 'scriptRun.jsx'
}

function Invoke-InDesignScript {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $indesign = New-Object -ComObject 'InDesign.Application'
    $jsType = 1246973031
    $indesign.DoScript($Path, $jsType)
}

if ($PSBoundParameters.ContainsKey('ScriptPath') -and -not [string]::IsNullOrWhiteSpace($ScriptPath)) {
    $resolvedPath = Resolve-ScriptPath -InputPath $ScriptPath
} else {
    $resolvedPath = Get-DefaultScriptPath
}

if (-not (Test-Path -Path $resolvedPath -PathType Leaf)) {
    Write-Error "Error: script file not found: $resolvedPath"
    exit 1
}

try {
    Invoke-InDesignScript -Path $resolvedPath
} catch {
    Write-Error "Error executing InDesign script: $($_.Exception.Message)"
    exit 1
}
