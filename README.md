# FOE Duty Dashboard - Codex Optimized Copy

This `_codex` folder is an isolated optimized copy. The original app one folder up is unchanged.

## What changed

- Flask now builds a compact duty index from the workbook once, then serves names, months, validation, and duty reports from memory/JSON cache.
- Google Sheets refresh runs in the background. Users keep seeing cached data instead of waiting on a slow download.
- `/api/bootstrap`, `/api/healthz`, and `/api/readyz` were added for faster app startup and cloud health checks.
- Gunicorn defaults are tuned for small Railway/Render machines: one worker with several threads.
- The React dashboard loads bootstrap data in one request, shows warming/stale/refreshing states, keeps the old report visible during refresh, supports URL sharing, recent names, and CSV export.

## Useful environment variables

- `DATA_DIR`: directory for `_cached_foe_data.xlsx` and `_duty_index.json`.
- `CACHE_TTL_SECONDS`: workbook freshness window. Default: `3600`.
- `DOWNLOAD_TIMEOUT_SECONDS`: Google Sheets download timeout. Default: `12`.
- `WEB_CONCURRENCY`: Gunicorn workers. Default: `1`.
- `GUNICORN_THREADS`: Gunicorn threads per worker. Default: `4`.

## Run locally

```powershell
cd "C:\Users\blade\Desktop\Duty Checker\_codex"
..\env\Scripts\python.exe server.py
```

Build the dashboard:

```powershell
cd "C:\Users\blade\Desktop\Duty Checker\_codex\dashboard"
npm run build
```
