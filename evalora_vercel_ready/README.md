# Evalora AI — Vercel-ready starter

## Features
- FastAPI backend exposed through Vercel's root `app.py` entrypoint.
- Three-stage Gemini review with configurable model fallback.
- Text extraction from PDF, DOCX, PPTX, TXT, MD, and CSV.
- Appends evaluation records to a Google Sheet when credentials are configured.

## 1. Deploy
1. Upload this folder to a GitHub repository.
2. Import the repository into Vercel. Leave the build command and output directory at their defaults.
3. In **Project → Settings → Environment Variables**, add the variables listed below.
4. Redeploy after adding or changing environment variables.
5. Open `/api/health` on the deployed domain. It should return JSON.

## 2. Environment variables
Required for evaluation:
- `GEMINI_API_KEY`: API key from Google AI Studio / Gemini API.

Optional model overrides (defaults are shown):
- `GEMINI_MODEL_1=gemini-3.6-flash`
- `GEMINI_MODEL_2=gemini-3.1-flash-lite`
- `GEMINI_MODEL_3=gemini-3-flash`

For Google Sheets saving:
- `GOOGLE_SERVICE_ACCOUNT_JSON`: the full service-account JSON object as a single-line JSON value.
- `GOOGLE_SHEET_ID`: the ID between `/d/` and `/edit` in the spreadsheet URL.
- `GOOGLE_WORKSHEET=Evaluations` (optional).

Enable the Google Sheets API in the Google Cloud project, create a service account key, and share the target spreadsheet with the service account's `client_email` as an Editor. Do not commit credentials to GitHub or put them in frontend code.

## 3. Test
- `https://YOUR-DEPLOYMENT.vercel.app/api/health`
- Submit a small text-based PDF or DOCX first.
- If the model returns quota or unavailable errors, check Google AI Studio usage and model access. You can change the fallback IDs in Vercel environment variables.
- If the response succeeds but the sheet status says it was not saved, check the Google Sheets environment variables and service-account sharing.

## Important limits
- Each uploaded file is limited to 8 MB; scanned PDFs without embedded text need OCR before upload.
- Three sequential model requests can take longer than short serverless limits. `maxDuration: 60` is set for broad plan compatibility. Three sequential model requests can still exceed this limit; if that happens, use a plan that supports longer functions or redesign the evaluation as an asynchronous job.
- Model IDs and availability can change. Confirm your API key can access the configured IDs before production use.
- This starter does not include user authentication. Add access control before exposing evaluation data publicly.
