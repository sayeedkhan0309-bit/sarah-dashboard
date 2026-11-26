# GFS Chartink Dashboard

This repository runs six saved Chartink screeners. Chartink remains the only source for screener membership and market-data fields.

## Configured screeners

| Excel sheet | Chartink URL | Report layout |
|---|---|---|
| ADVANCEGFSCASHSTOCKS | https://chartink.com/screener/advancegfscashstocks | Stock columns, including Volume |
| ADVANCEGFSSTOCKSFUTURE | https://chartink.com/screener/advancegfsstocksfuture | Stock columns, including Volume |
| VLAUEBUYSTOCKSCASH | https://chartink.com/screener/vlauebuystockscash | Stock columns, including Volume |
| GFSSTOCKCASH | https://chartink.com/screener/gfsstockcash | Stock columns, including Volume |
| GFSWITHSECTOR | https://chartink.com/screener/gfswithsector | Stock columns, including Volume |
| SECTORINDICES | https://chartink.com/screener/gfs-60610539 | Index columns; Volume omitted |

The `VLAUEBUYSTOCKSCASH` label intentionally follows the exact spelling in the supplied Chartink URL.

## Output

Each run creates:

- `gfs_dashboard_latest.xlsx` — a Dashboard sheet, one formatted result sheet per screener, and one raw-response sheet per screener.
- `gfs_dashboard_latest.json` — timestamp, status, original Chartink fields, and records for all six screeners.

The formatted stock columns are: Sr, Stock Name, Daily RSI, Weekly RSI, Monthly RSI, Close, %_change, Volume.

The formatted index columns are: Sr, Index Name, Daily RSI, Weekly RSI, Monthly RSI, Close, %_change. The index sheet intentionally has no Volume column.

## Run in GitHub

1. Open the repository's **Actions** tab.
2. Select **Refresh GFS Chartink Dashboard**.
3. Select **Run workflow** to test manually.
4. Open the completed run and download the `gfs-dashboard-latest` artifact.

The workflow is scheduled on weekdays during the Indian market session. GitHub may delay scheduled runs. Chartink can also change its page or response behavior, so check the run log if a screener fails.

## Technical notes

- `gfs_dashboard.py` opens each saved Chartink page and captures the screener's response. It does not construct or modify scan clauses.
- The existing conditions and universe for each screener remain controlled by the saved Chartink query itself.
- The script does not use Yahoo Finance, NSE feeds, broker feeds, or other market-data sources, and it never places orders.
- Raw response sheets and JSON are retained so actual Chartink field names can be checked if a formatted RSI column is blank.
- `requirements.txt` contains the Python dependencies.
- `.github/workflows/gfs-refresh.yml` defines the scheduled and manual GitHub Actions runs.
