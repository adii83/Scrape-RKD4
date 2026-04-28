param(
    [string]$ConfigPath = "scraper.settings.json",
    [string]$CdpUrl,
    [Nullable[int]]$StartRow,
    [Nullable[int]]$EndRow,
    [Nullable[int]]$BatchSize,
    [Nullable[double]]$SearchDelayMin,
    [Nullable[double]]$SearchDelayMax,
    [Nullable[double]]$ProfileDelayMin,
    [Nullable[double]]$ProfileDelayMax,
    [Nullable[double]]$BatchCooldownMin,
    [Nullable[double]]$BatchCooldownMax,
    [Nullable[double]]$BlockRecheckMin,
    [Nullable[double]]$BlockRecheckMax
)

$ErrorActionPreference = "Stop"

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host "[OK]   $msg" -ForegroundColor Green }
function Write-Err($msg) { Write-Host "[ERR]  $msg" -ForegroundColor Red }

try {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Config tidak ditemukan: $ConfigPath"
    }

    $cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json

    if ($PSBoundParameters.ContainsKey("CdpUrl")) { $cfg.cdp_url = $CdpUrl }
    if ($PSBoundParameters.ContainsKey("StartRow")) { $cfg.start_row = $StartRow.Value }
    if ($PSBoundParameters.ContainsKey("EndRow")) { $cfg.end_row = $EndRow.Value }
    if ($PSBoundParameters.ContainsKey("BatchSize")) { $cfg.batch_size = $BatchSize.Value }
    if ($PSBoundParameters.ContainsKey("SearchDelayMin")) { $cfg.search_delay_min = $SearchDelayMin.Value }
    if ($PSBoundParameters.ContainsKey("SearchDelayMax")) { $cfg.search_delay_max = $SearchDelayMax.Value }
    if ($PSBoundParameters.ContainsKey("ProfileDelayMin")) { $cfg.profile_delay_min = $ProfileDelayMin.Value }
    if ($PSBoundParameters.ContainsKey("ProfileDelayMax")) { $cfg.profile_delay_max = $ProfileDelayMax.Value }
    if ($PSBoundParameters.ContainsKey("BatchCooldownMin")) { $cfg.batch_cooldown_min = $BatchCooldownMin.Value }
    if ($PSBoundParameters.ContainsKey("BatchCooldownMax")) { $cfg.batch_cooldown_max = $BatchCooldownMax.Value }
    if ($PSBoundParameters.ContainsKey("BlockRecheckMin")) { $cfg.block_recheck_min = $BlockRecheckMin.Value }
    if ($PSBoundParameters.ContainsKey("BlockRecheckMax")) { $cfg.block_recheck_max = $BlockRecheckMax.Value }

    if ([int]$cfg.end_row -lt [int]$cfg.start_row) {
        throw "end_row harus >= start_row"
    }
    if ([int]$cfg.batch_size -lt 1) {
        throw "batch_size minimal 1"
    }
    if ([double]$cfg.search_delay_max -lt [double]$cfg.search_delay_min) {
        throw "search_delay_max harus >= search_delay_min"
    }
    if ([double]$cfg.batch_cooldown_max -lt [double]$cfg.batch_cooldown_min) {
        throw "batch_cooldown_max harus >= batch_cooldown_min"
    }

    ($cfg | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
    Write-Ok "Config diperbarui: $ConfigPath"
    Write-Info "Nilai saat ini:"
    Get-Content -LiteralPath $ConfigPath
}
catch {
    Write-Err $_.Exception.Message
    exit 1
}
