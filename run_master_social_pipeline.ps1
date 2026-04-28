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
    [Nullable[double]]$BlockRecheckMax,
    [switch]$SkipCdpCheck
)

$ErrorActionPreference = "Stop"

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host "[OK]   $msg" -ForegroundColor Green }
function Write-Err($msg) { Write-Host "[ERR]  $msg" -ForegroundColor Red }

function Run-HarnessScript($scriptPath, $label) {
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "Script tidak ditemukan: $scriptPath"
    }
    Write-Info "Menjalankan $label ..."
    browser-harness -c "exec(open('$scriptPath', encoding='utf-8').read())"
    Write-Ok "$label selesai"
}

try {
    Write-Info "Memulai master pipeline (LinkedIn -> Instagram -> Facebook -> TikTok)..."

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Config tidak ditemukan: $ConfigPath"
    }
    $cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json

    $CdpUrl = if ($PSBoundParameters.ContainsKey("CdpUrl")) { $CdpUrl } else { [string]$cfg.cdp_url }
    $StartRow = if ($PSBoundParameters.ContainsKey("StartRow")) { $StartRow.Value } else { [int]$cfg.start_row }
    $EndRow = if ($PSBoundParameters.ContainsKey("EndRow")) { $EndRow.Value } else { [int]$cfg.end_row }
    $BatchSize = if ($PSBoundParameters.ContainsKey("BatchSize")) { $BatchSize.Value } else { [int]$cfg.batch_size }
    $SearchDelayMin = if ($PSBoundParameters.ContainsKey("SearchDelayMin")) { $SearchDelayMin.Value } else { [double]$cfg.search_delay_min }
    $SearchDelayMax = if ($PSBoundParameters.ContainsKey("SearchDelayMax")) { $SearchDelayMax.Value } else { [double]$cfg.search_delay_max }
    $ProfileDelayMin = if ($PSBoundParameters.ContainsKey("ProfileDelayMin")) { $ProfileDelayMin.Value } else { [double]$cfg.profile_delay_min }
    $ProfileDelayMax = if ($PSBoundParameters.ContainsKey("ProfileDelayMax")) { $ProfileDelayMax.Value } else { [double]$cfg.profile_delay_max }
    $BatchCooldownMin = if ($PSBoundParameters.ContainsKey("BatchCooldownMin")) { $BatchCooldownMin.Value } else { [double]$cfg.batch_cooldown_min }
    $BatchCooldownMax = if ($PSBoundParameters.ContainsKey("BatchCooldownMax")) { $BatchCooldownMax.Value } else { [double]$cfg.batch_cooldown_max }
    $BlockRecheckMin = if ($PSBoundParameters.ContainsKey("BlockRecheckMin")) { $BlockRecheckMin.Value } else { [double]$cfg.block_recheck_min }
    $BlockRecheckMax = if ($PSBoundParameters.ContainsKey("BlockRecheckMax")) { $BlockRecheckMax.Value } else { [double]$cfg.block_recheck_max }

    if ($EndRow -lt $StartRow) { throw "EndRow harus >= StartRow." }
    if ($BatchSize -lt 1) { throw "BatchSize minimal 1." }

    $bh = Get-Command browser-harness -ErrorAction SilentlyContinue
    if (-not $bh) { throw "Perintah 'browser-harness' tidak ditemukan di PATH." }
    Write-Ok "browser-harness tersedia: $($bh.Source)"

    $env:BU_CDP_URL = $CdpUrl
    $env:BH_START_ROW = "$StartRow"
    $env:BH_END_ROW = "$EndRow"
    $env:BH_BATCH_SIZE = "$BatchSize"
    $env:BH_RESUME_ENABLED = "1"
    $env:BH_SEARCH_DELAY_MIN = "$SearchDelayMin"
    $env:BH_SEARCH_DELAY_MAX = "$SearchDelayMax"
    $env:BH_PROFILE_DELAY_MIN = "$ProfileDelayMin"
    $env:BH_PROFILE_DELAY_MAX = "$ProfileDelayMax"
    $env:BH_BATCH_COOLDOWN_MIN = "$BatchCooldownMin"
    $env:BH_BATCH_COOLDOWN_MAX = "$BatchCooldownMax"
    $env:BH_BLOCK_RECHECK_MIN = "$BlockRecheckMin"
    $env:BH_BLOCK_RECHECK_MAX = "$BlockRecheckMax"
    Write-Info "Parameter pipeline sudah di-set (row $StartRow-$EndRow, batch $BatchSize)."

    if (-not $SkipCdpCheck) {
        Write-Info "Cek koneksi CDP..."
        try {
            $null = Invoke-RestMethod "$CdpUrl/json/version" -TimeoutSec 5
            Write-Ok "CDP aktif."
        } catch {
            throw "CDP tidak bisa diakses di $CdpUrl. Pastikan Brave dibuka dengan --remote-debugging-port."
        }
    } else {
        Write-Info "Lewatkan cek CDP (--SkipCdpCheck)."
    }

    Run-HarnessScript "tools/alumni_linkedin_lookup.py" "LinkedIn"
    Run-HarnessScript "tools/alumni_instagram_lookup.py" "Instagram"
    Run-HarnessScript "tools/alumni_facebook_lookup.py" "Facebook"
    Run-HarnessScript "tools/alumni_tiktok_lookup.py" "TikTok"

    Write-Info "Merging final output..."
    python "tools/merge_social_outputs.py"
    Write-Ok "Master output selesai."
    Write-Ok "Cek outputs/alumni_master_${StartRow}_${EndRow}.json dan .xlsx"
}
catch {
    Write-Err $_.Exception.Message
    exit 1
}
