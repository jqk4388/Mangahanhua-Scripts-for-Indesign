<#
.SYNOPSIS
    Manga Layout Automation - PowerShell Launcher

.DESCRIPTION
    Launches Adobe InDesign, loads the manga_layout.jsx script, and logs execution details.

.NOTES
    Version: 1.0.0
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$logPath = Join-Path $scriptPath 'manga_layout_vbs.log'

function Log-Message {
    param(
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $logPath -Value $line
    Write-Host $Message
}

function Get-InDesignApplication {
    $versions = @('2026', '2025', '2024', '2023', '2022', '2021', '2020', 'CC 2019', 'CC 2018')

    foreach ($version in $versions) {
        try {
            $progId = "InDesign.Application.$version"
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            if ($app) {
                Log-Message "Connected to InDesign $version"
                return $app
            }
        } catch {
            # Ignore and continue
        }
    }

    foreach ($version in $versions) {
        try {
            $progId = "InDesign.Application.$version"
            $app = New-Object -ComObject $progId
            if ($app) {
                Log-Message "Created InDesign $version instance"
                return $app
            }
        } catch {
            # Ignore and continue
        }
    }

    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('InDesign.Application')
        if ($app) {
            Log-Message 'Connected to InDesign (default version)'
            return $app
        }
    } catch {
        # Ignore
    }

    try {
        $app = New-Object -ComObject 'InDesign.Application'
        if ($app) {
            Log-Message 'Created InDesign (default version) instance'
            return $app
        }
    } catch {
        # Ignore
    }

    return $null
}

function Is-InDesignRunning {
    try {
        $processes = Get-Process -Name 'InDesign' -ErrorAction SilentlyContinue
        return ($processes -ne $null -and $processes.Count -gt 0)
    } catch {
        return $false
    }
}

function Execute-JSX {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string]$JsxPath
    )

    $idJavascript = 1246973031

    try {
        $Application.DoScript($JsxPath, $idJavascript)
        return $true
    } catch {
        Log-Message "DoScript error: $($_.Exception.Message)"
        return $false
    }
}

function Main {
    $startTime = Get-Date

    Log-Message '===== Manga Layout PS1 Launcher Started ====='
    Log-Message "Script path: $scriptPath"
    Log-Message "Start time: $startTime"

    $isRunning = Is-InDesignRunning
    Log-Message "Is InDesign already running: $isRunning"

    Log-Message 'Connecting to InDesign application...'
    $app = Get-InDesignApplication

    if (-not $app) {
        Log-Message 'Error: Unable to connect to InDesign - ensure InDesign is installed and available.'
        exit 1
    }

    Log-Message 'Successfully connected to InDesign'
    Start-Sleep -Milliseconds 1000

    $jsxPath = Join-Path $scriptPath 'manga_layout.jsx'
    if (-not (Test-Path -Path $jsxPath -PathType Leaf)) {
        Log-Message "Error: Script file does not exist - $jsxPath"
        exit 1
    }

    Log-Message "Script file: $jsxPath"

    $configPath = Join-Path $scriptPath 'manga_layout_config.json'
    if (-not (Test-Path -Path $configPath -PathType Leaf)) {
        Log-Message "Warning: Config file does not exist - $configPath"
        Log-Message 'Will run with default configuration'
    } else {
        Log-Message "Config file: $configPath"
    }

    Log-Message 'Starting JSX script execution...'
    $result = Execute-JSX -Application $app -JsxPath $jsxPath

    $endTime = Get-Date
    $duration = New-TimeSpan -Start $startTime -End $endTime

    if (-not $result) {
        Log-Message "Error: Script execution failed"
        Log-Message "Execution duration: $($duration.TotalSeconds) seconds"
        exit 1
    }

    Log-Message 'Script execution completed'
    Log-Message "Execution duration: $([math]::Round($duration.TotalSeconds, 2)) seconds"
    Log-Message '===== Execution successful ====='

    return 0
}

Main
