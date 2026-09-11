# Qualys Fetch Asset Tags

Fetches **all** Asset Management (AM) tags from a Qualys tenant via the QPS REST API and exports them to CSV and XLSX. Built to pull live tag data straight from the API rather than relying on a possibly stale manual export.

## What it does

- Pages through `POST /qps/rest/2.0/search/am/tag` on the Qualys EU Platform 2 pod (`qualysapi.qg2.apps.qualys.eu`), 100 tags per page, until every tag has been retrieved.
- Resolves each tag's parent tag name by building a complete `id -> name` map across all pages first, so children whose parent appears on a later page still resolve correctly.
- Classifies each tag as **Static** or **Dynamic** based on `ruleType`, and best-effort maps each tag's color to a business meaning per the standard color scheme (business-critical, users/roles, technologies, projects, vulnerability/EOL, etc.).
- Writes:
  - `qualys_tags_<YYYY-MM-DD>.csv`
  - `qualys_tags_<YYYY-MM-DD>.xlsx` — a `Tags` sheet (frozen header row, autofilter, wrapped rule-value column, sorted by parent then name) and a `Summary` sheet (counts by tag type and rule type).
- Prints a console summary: total tags, static/dynamic counts, tags with no color, root tags (no parent), and any tag whose `parentTagId` didn't resolve to a name.

## Requirements

- Python 3
- `requests`, `openpyxl` (`pip install requests openpyxl`)

## Credentials

Never hardcoded. The script resolves credentials in this order:

1. Environment variables `QUALYS_USERNAME` / `QUALYS_PASSWORD`
2. A `qualys_creds.txt` file next to the script, formatted as:

   ```
   user:	<username>
   pass:	<password>
   ```

`qualys_creds.txt` is git-ignored and must never be committed.

## Usage

```bash
python fetch_qualys_tags.py
```

Re-running is safe — it always does a full fetch and overwrites/creates that day's `qualys_tags_<date>.csv`/`.xlsx`.

## Reliability

- Sleeps `INITIAL_DELAY` (1s) between successful page requests.
- On HTTP 409/429 (rate limiting), honors `X-RateLimit-ToWait-Sec` / `Retry-After` and backs off exponentially, capped at `MAX_RATE_LIMIT_SLEEP` (300s).
- Retries up to `MAX_RETRIES` (5) attempts per page before failing loudly with a clear error.
