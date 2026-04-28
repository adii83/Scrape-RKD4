
## 1) Prasyarat

- Windows + PowerShell
- Python (sudah aktif di environment Anda)
- Brave Browser
- `browser-harness` command tersedia
- File data: `Data Alumni.xlsx`

## 2) Aktivasi virtual environment (opsional, kalau sudah pakai `.venv`)

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& ".\.venv\Scripts\Activate.ps1")
```

## 3) Jalankan Brave dengan remote debugging

```powershell
Stop-Process -Name brave -Force -ErrorAction SilentlyContinue

$brave = "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
$userData = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"

& $brave `
  --remote-debugging-port=9333 `
  "--user-data-dir=$userData" `
  --profile-directory=Default `
  about:blank
```

Verifikasi:

```powershell
Invoke-RestMethod "http://127.0.0.1:9333/json/version"
```

Jika keluar `webSocketDebuggerUrl`, berarti siap.

## 4) Atur konfigurasi scraping

Edit file:
- `scraper.settings.json`

Contoh:

```json
{
  "cdp_url": "http://127.0.0.1:9333",
  "start_row": 1,
  "end_row": 1000,
  "batch_size": 100,
  "search_delay_min": 5.0,
  "search_delay_max": 15.0,
  "profile_delay_min": 2.0,
  "profile_delay_max": 4.5,
  "batch_cooldown_min": 20.0,
  "batch_cooldown_max": 90.0,
  "block_recheck_min": 120.0,
  "block_recheck_max": 240.0
}
```

Alternatif: ubah config lewat script helper

```powershell
.\set_scraper_settings.ps1 -StartRow 1 -EndRow 1000 -BatchSize 100
```

## 5) Jalankan pipeline utama

```powershell
.\run_master_social_pipeline.ps1
```

## 6) Pause / Resume

- Pause saat proses: `Ctrl + C`
- Resume: jalankan lagi command yang sama

Semua scraper punya progress file sendiri, jadi proses lanjut dari titik terakhir.

## 7) Output

Output ada di folder `outputs/`.

File final gabungan:
- `outputs/alumni_master_<start_row>_<end_row>.json`
- `outputs/alumni_master_<start_row>_<end_row>.xlsx`

Field final utama:
- `alamat_sosial_media_linkedin`
- `alamat_sosial_media_instagram`
- `alamat_sosial_media_facebook`
- `alamat_sosial_media_tiktok`
- `email`
- `no_hp`
- `tempat_bekerja`
- `alamat_bekerja`
- `posisi`
- `kategori_pekerjaan`
- `alamat_sosial_media_tempat_bekerja`

## 8) Catatan

- Jika terdeteksi captcha/rate-limit, scraper auto-wait lalu cek ulang sampai normal.
- Search engine fallback untuk non-LinkedIn: DuckDuckGo -> Yahoo -> Yandex -> Bing.
- Link kampus seperti `ummcampus` difilter agar tidak masuk hasil personal alumni.
