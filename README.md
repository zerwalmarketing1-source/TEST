# STOURS DMC Hotel Catalogue

Static premium hotel catalogue generated from the provided Excel and STOURS DMC brand assets.

## Local catalogue admin

The local admin app lets you update hotel content, replace or switch image slots, download photos from URLs into `assets/official/...`, and import hotel names/classifications/official links from an Excel file.

```bash
python3 -m pip install -r requirements-admin.txt
python3 admin_app.py
```

Then open `http://127.0.0.1:8000/admin`.

Excel import expects a header row with a hotel-name column such as `HOTEL NAME`; optional columns include `CITY`, `CLASSIFICATION`, and a website/link column. It updates matching hotel names and creates missing hotels inside matching city sections when possible. Every save creates an `.admin_backups/` copy of `index.html`.
