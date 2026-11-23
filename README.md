# data-reconciliation

## Deploy full stack on Render

Create a Render Web Service from this repository with:

- Root Directory: `.`
- Build Command: `npm ci && npm run build`
- Start Command: `npm start`

Render provides `PORT`. NestJS serves the frontend from `public/` and handles
the reconciliation API on the same domain. The scheduled job remains inactive
until its source API and reconciliation settings are configured.

`GET /reconciliation/codes` returns the available reconciliation codes and
names used to populate the payment-channel selector.

## MongoDB logs

Set `MONGODB_URI` to enable MongoDB logging. Optionally set
`MONGODB_DATABASE`; the default database is `data-reconciliation`. Configure
the connection string as an environment variable or in the project-root `.env`
file; `.env` is loaded at startup. Existing process environment values take
precedence. Do not commit credentials.

The app writes user actions to `user_logs` and system events to `system_logs`.
The temporary `userKey` is the request IP (not an authenticated identity).
Logs expire after 10 days. When `MONGODB_URI` is unset, the app continues
without MongoDB persistence and writes a startup warning.

## Upload input files

`POST /reconciliation/upload` accepts one multipart file per request:

- `file`: an `.xlsx` workbook with exactly one worksheet, up to 70 MB.
- `type`: `partner` or `afc`.

Files are stored locally under `uploads/` using their original filename; the
`type` is recorded in `uploads/.metadata/`. Uploading a filename that already
exists returns HTTP 409, and existing files are never overwritten.

## Scheduled reconciliation

The application checks for source files every 10 minutes. Configure these
environment variables to enable the job:

- `RECONCILIATION_SOURCE_URL`: HTTPS GET endpoint returning JSON with
  `partnerFileUrl` and `afcFileUrl`, both HTTPS URLs to `.xlsx` files.
- `RECONCILIATION_CODE`: one code returned by `GET /reconciliation/codes`.
- `RECONCILIATION_FROM_DATE` and `RECONCILIATION_TO_DATE`: fixed dates in
  `YYYY-MM-DD` format.
- `RECONCILIATION_OUTPUT_DIR`: output directory; defaults to `output/`.

When `RECONCILIATION_SOURCE_URL` is empty, the scheduled job skips that run.
Downloaded files are limited to 20 MB each, and generated workbooks are saved
with a timestamped filename in the output directory.