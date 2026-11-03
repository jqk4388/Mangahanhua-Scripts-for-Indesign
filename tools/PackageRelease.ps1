<#
  打包发布 zip

  规则与历史 release 保持一致：
    1. 只打包 git 已跟踪的文件
    2. 排除开发用目录与文件（img、Indesign-script-for-claw、_build 之外的工具等）
    3. 有对应 .jsxbin 的脚本只发布 jsxbin，不发布源码
    4. 把 _build 中的 .jsxbin 合并到对应脚本目录

  用法：
    pwsh -File tools/PackageRelease.ps1 -Tag 261103
#>
param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [string]$OutputDir
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
if (-not $OutputDir) { $OutputDir = Join-Path $root 'dist' }

$pkgName = "Mangahanhua-Scripts-for-Indesign_$Tag"
$zipPath = Join-Path $OutputDir "$pkgName.zip"
$stageRoot = Join-Path $env:TEMP "release_$Tag"

# 不参与分发的路径（正则匹配）
$excludePatterns = @(
    '^\.git',                          # .gitattributes / .gitignore
    '^img/',
    '^Indesign-script-for-claw/',
    '^\.github/',
    '^tools/',
    '^dist/',
    '^_build/',
    '^BTjsx/Readme\.md$',
    '^BTjsx/final_comic_only_text\.py$',
    '^BTjsx/bt.*\.json$'               # BT 导出的方框例子
)

# _build 中的 jsxbin -> 发布目录中的位置
$bins = [ordered]@{
    '页面调整/1.开局新建图层图框.jsxbin'                = '1.开局新建图层图框.jsxbin'
    '导入导出/2.导入样式和复合字体.jsxbin'              = '2.导入样式和复合字体.jsxbin'
    '导入导出/3.LabelPlus-script-id-UI.jsxbin'         = '3.LabelPlus-script-id-UI.jsxbin'
    '文本修改-格式化/3.样式匹配.jsxbin'                 = '3.样式匹配.jsxbin'
    'BTjsx/3.BT-JSON-to-ID.jsxbin'                     = '3.BT-JSON-to-ID.jsxbin'
    '样式修改-文本特效/AI特效-粗糙化.jsxbin'            = 'AI特效-粗糙化.jsxbin'
    '样式修改-文本特效/AI特效-黑底雕.jsxbin'            = 'AI特效-黑底雕.jsxbin'
    '样式修改-文本特效/ID特效-白底雕.jsxbin'            = 'ID特效-白底雕.jsxbin'
    '样式修改-文本特效/ID特效-白底雕斜.jsxbin'          = 'ID特效-白底雕斜.jsxbin'
    '样式修改-文本特效/PS特效-背景字.jsxbin'            = 'PS特效-背景字.jsxbin'
    '页码链接/[集英社]显隐文字图层 ネーム.jsxbin'       = '[集英社]显隐文字图层 ネーム.jsxbin'
}

# 已被 jsxbin 取代的源码，不重复发布
$replacedSources = @{}
foreach ($dest in $bins.Keys) {
    $replacedSources[$dest -replace '\.jsxbin$', '.jsx'] = $true
}

# 取 git 跟踪的文件列表（-z 避免中文路径被转义）
$tracked = (git -C $root ls-files -z) -split [char]0 | Where-Object { $_ -ne '' }

if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force }
$stagePkg = Join-Path $stageRoot 'Mangahanhua-Scripts-for-Indesign'
New-Item -ItemType Directory -Path $stagePkg -Force | Out-Null

$copied = 0
foreach ($rel in $tracked) {
    if ($replacedSources.ContainsKey($rel)) { continue }
    $skip = $false
    foreach ($p in $excludePatterns) {
        if ($rel -match $p) { $skip = $true; break }
    }
    if ($skip) { continue }

    $dest = Join-Path $stagePkg $rel
    $destDir = Split-Path -Parent $dest
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    Copy-Item -LiteralPath (Join-Path $root $rel) -Destination $dest -Force
    $copied++
}

# 合并 _build 中的 jsxbin
$binCopied = 0
foreach ($destRel in $bins.Keys) {
    $src = Join-Path $root "_build/$($bins[$destRel])"
    if (-not (Test-Path -LiteralPath $src)) {
        throw "缺少编译产物：_build/$($bins[$destRel])"
    }
    $dest = Join-Path $stagePkg $destRel
    $destDir = Split-Path -Parent $dest
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $src -Destination $dest -Force
    $binCopied++
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

[System.IO.Compression.ZipFile]::CreateFromDirectory(
    $stagePkg, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $true, [System.Text.Encoding]::UTF8)

$entries = ([System.IO.Compression.ZipFile]::OpenRead($zipPath)).Entries.Count
Write-Host "已打包 $entries 个条目（脚本 $copied，jsxbin $binCopied）：$zipPath"
$zipPath
