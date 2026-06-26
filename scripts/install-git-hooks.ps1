# Instala hooks versionados deste repositório (core.hooksPath local).
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot
git config core.hooksPath .githooks
Write-Host "Hooks instalados: core.hooksPath=.githooks"
