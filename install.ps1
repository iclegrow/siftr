<#
.SYNOPSIS
    Set up Siftr for first-time use — adds the module to your PowerShell profile.

.DESCRIPTION
    This script:
    1. Detects the Siftr repo root
    2. Adds a dot-source line to your PowerShell profile
    3. Sources the module into the current session

    After running, just open Copilot CLI and say "siftr" to begin setup.

.EXAMPLE
    .\install.ps1
#>

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSCommandPath
$modulePath = Join-Path $repoRoot 'modules\Siftr-Inbox.ps1'

if (-not (Test-Path $modulePath)) {
    Write-Error "Cannot find Siftr module at: $modulePath"
    return
}

Write-Host "🔧 Siftr Install" -ForegroundColor Cyan
Write-Host ""

# Check prerequisites
$prereqs = @(
    @{ Name = 'Outlook'; Check = { try { New-Object -ComObject Outlook.Application | Out-Null; $true } catch { $false } } }
    @{ Name = 'Node.js'; Check = { try { node --version | Out-Null; $true } catch { $false } } }
)

foreach ($p in $prereqs) {
    $ok = & $p.Check
    if ($ok) {
        Write-Host "  ✅ $($p.Name) found" -ForegroundColor Green
    } else {
        Write-Host "  ⚠️  $($p.Name) not found — some features may not work" -ForegroundColor Yellow
    }
}

Write-Host ""

# Add to PS profile
$profilePath = $PROFILE
$sourceLine = "if (Test-Path `"$modulePath`") { . `"$modulePath`" }"

if (-not (Test-Path $profilePath)) {
    New-Item -Path $profilePath -ItemType File -Force | Out-Null
    Write-Host "  📄 Created PowerShell profile: $profilePath"
}

$profileContent = Get-Content $profilePath -Raw -ErrorAction SilentlyContinue
if ($profileContent -and $profileContent.Contains($modulePath)) {
    Write-Host "  ✅ Siftr module already in profile" -ForegroundColor Green
} else {
    Add-Content -Path $profilePath -Value "`n# Siftr — email triage skill`n$sourceLine"
    Write-Host "  ✅ Added Siftr module to profile: $profilePath" -ForegroundColor Green
}

# Source into current session
. $modulePath
Write-Host ""
Write-Host "✅ Siftr installed! Open Copilot CLI and say 'siftr' to begin." -ForegroundColor Green
Write-Host "   (First run will walk you through configuration)" -ForegroundColor DarkGray
