# data-reconciliation

## Deploy full stack on Render

Create a Render Web Service from this repository with:

- Root Directory: `.`
- Build Command: `npm ci && npm run build`
- Start Command: `npm start`

No environment variables are required; Render provides `PORT`. NestJS serves the
frontend from `public/` and handles the reconciliation API on the same domain.

`GET /reconciliation/codes` returns the available reconciliation codes and
names used to populate the payment-channel selector.