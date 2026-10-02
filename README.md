# WDS Inventory

Private, login-protected inventory dashboard for Western Door Solutions. Imports the 1,420 products from the WDS physical inventory workbook; purchase price and QuickBooks sales price are imported when provided in the source. Missing prices remain blank and can be entered in the app. The workbook is uploaded only through the authenticated app and is not stored in the public GitHub repository.

## Features
- Admin-only login restricted to `ADMIN_EMAIL` (default `wds@telus.net`) with password held in `ADMIN_PASSWORD`.
- Product search by part number, name, brand, category, and location.
- Browser camera or image upload with OCR-based part-number search.
- Quantity controls with an inventory movement history.
- Purchase and list price editing; dashboard quantity, inventory cost, and list value totals.
- PostgreSQL persistence and a health endpoint.

## Local setup

```sh
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
export DATABASE_URL='postgresql://...'
export ADMIN_EMAIL='wds@telus.net'
export ADMIN_PASSWORD='set-a-strong-password'
export COOKIE_SECURE='false'
export SESSION_SECRET='a-long-random-secret'
python app.py
```

After first sign-in, choose **Import workbook** and select the WDS inventory workbook. The import is accepted only while the product catalog is empty. Subsequent starts preserve the database and do not overwrite changes.

## Render

The included `render.yaml` creates a web service and a free PostgreSQL database. Set `ADMIN_PASSWORD` as a secret before first deployment. Render free PostgreSQL instances expire after 30 days; use a paid database before storing business inventory for long-term use. Free web services may spin down while idle.

## Security and deployment notes

Never commit passwords or session secrets. `ADMIN_PASSWORD` must be set in the service environment. Change it from the temporary password previously supplied in chat before sharing the app. Login sessions use secure, HTTP-only cookies; production must use HTTPS.
