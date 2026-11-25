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
Successful B2 uploads are also recorded in the persistent `uploaded_files`
collection with `name`, `links`, and `status` fields, plus file metadata.
`links` contains the B2 object URL; private buckets still require authorization
to open the URL.

## Upload input files

`POST /reconciliation/upload` accepts one multipart file per request and uploads
it to Backblaze B2:

- `file`: an `.xlsx` workbook with exactly one worksheet, up to 70 MB.
- `type`: `partner` or `afc`.

Set `B2_BUCKET_NAME`, `B2_ENDPOINT`, `B2_KEY_ID`, and `B2_APPLICATION_KEY` in
`.env` or the process environment. `B2_REGION` is optional when it can be
derived from a standard Backblaze S3 endpoint. The API response includes the
B2 object `key` and MongoDB `id`. Download the uploaded workbook with
`GET /reconciliation/download/:id`; the API streams it from B2. The upload uses
an in-memory buffer and does not write the input workbook to the local
`uploads/` directory.

`GET /reconciliation/list-uploaded-files` lists the current user's uploaded
files. Use `page` (default `1`) and `limit` (default `10`, maximum `100`) query
parameters for pagination. The response includes `data`, `page`, `limit`,
`total`, and `totalPages`.

During reconciliation, Excel date/time values are normalized to
`Asia/Ho_Chi_Minh` (UTC+7) before the date-range filter is applied.

## Scheduled reconciliation

The application checks for source files every 10 minutes. The job starts by
default; set `RECONCILIATION_CRON_ENABLED=false` to disable its timer without
removing `CronJobService` from the module. The service also exposes `stop()` to
stop the timer programmatically. Configure these environment variables to use
the job:

- `RECONCILIATION_SOURCE_URL`: HTTPS GET endpoint returning JSON with
  `partnerFileUrl` and `afcFileUrl`, both HTTPS URLs to `.xlsx` files.
- `RECONCILIATION_CODE`: one code returned by `GET /reconciliation/codes`.
- `RECONCILIATION_FROM_DATE` and `RECONCILIATION_TO_DATE`: fixed dates in
  `YYYY-MM-DD` format.
- `RECONCILIATION_OUTPUT_DIR`: output directory; defaults to `output/`.

When `RECONCILIATION_SOURCE_URL` is empty, the scheduled job skips that run.
Downloaded files are limited to 20 MB each, and generated workbooks are saved
with a timestamped filename in the output directory.