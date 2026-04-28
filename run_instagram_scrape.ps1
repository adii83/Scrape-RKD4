param(
    [string]$CdpUrl = "http://127.0.0.1:9333",
    [string]$ScriptPath = "tools/alumni_instagram_lookup.py",
    [string]$ExcelPath = "Data Alumni.xlsx",
    [switch]$SkipCdpCheck
)

$ErrorActionPreference = "Stop"

function Write-Info($msg) {
    Write-Host "[INFO] $msg" -ForegroundColor Cyan
}

function Write-Ok($msg) {
    Write-Host "[OK]   $msg" -ForegroundColor Green
}

function Write-Err($msg) {
    Write-Host "[ERR]  $msg" -ForegroundColor Red
}

try {
    Write-Info "Memulai runner alumni Instagram scrape..."

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Script tidak ditemukan: $ScriptPath"
    }
    Write-Ok "Script ditemukan: $ScriptPath"

    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        throw "File Excel tidak ditemukan: $ExcelPath"
    }
    Write-Ok "File Excel ditemukan: $ExcelPath"

    $bh = Get-Command browser-harness -ErrorAction SilentlyContinue
    if (-not $bh) {
        throw "Perintah 'browser-harness' tidak ditemukan di PATH."
    }
    Write-Ok "browser-harness tersedia: $($bh.Source)"

    $env:BU_CDP_URL = $CdpUrl
    Write-Info "BU_CDP_URL di-set ke: $CdpUrl"

    if (-not $SkipCdpCheck) {
        Write-Info "Cek koneksi CDP..."
        try {
            $null = Invoke-RestMethod "$CdpUrl/json/version" -TimeoutSec 5
            Write-Ok "CDP aktif."
        }
        catch {
            throw "CDP tidak bisa diakses di $CdpUrl. Pastikan Brave dibuka dengan --remote-debugging-port."
        }
    }
    else {
        Write-Info "Lewatkan cek CDP (--SkipCdpCheck)."
    }

    Write-Info "Menjalankan scraping Instagram..."
    browser-harness -c "exec(open('$ScriptPath', encoding='utf-8').read())"

    Write-Ok "Selesai. Cek folder outputs untuk hasil batch/all."
}
catch {
    Write-Err $_.Exception.Message
    exit 1
}
