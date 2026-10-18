# Evalora AI — Vercel Web UI + Python API

This deployment replaces the Gradio interface with a polished responsive HTML/CSS/JavaScript frontend. The evaluation prompt, four phase question sets, rating mapping, Gemini fallback, compliance analysis, and Google Sheets nine-column append all remain in `api/index.py`.

## Deploy
1. Extract ZIP and push the folder contents to GitHub (do not upload secrets).
2. Import the repo at https://vercel.com/new.
3. Add environment variables in Vercel → Project Settings → Environment Variables, then deploy/redeploy.
4. Test `/api/health` on your deployment.

## Required environment variables
- `GEMINI_API_KEY`: Gemini API key.
- `GOOGLE_SERVICE_ACCOUNT_JSON`: full service-account JSON pasted as a single-line JSON value.
- `SPREADSHEET_ID`: defaults to the ID from the original app.
- `GOOGLE_WORKSHEET`: defaults to `Evaluations`.
Optional model overrides: `GEMINI_MODEL_1`, `GEMINI_MODEL_2`, `GEMINI_MODEL_3`. The original model IDs are retained as defaults; verify they are currently available to your API key. If one fails, the app tries the next configured model.

## Google Sheets
Enable Google Sheets API and Google Drive API in Google Cloud. Create a service account and paste its JSON into Vercel environment variables. Share the spreadsheet with the service account `client_email` as Editor. Row 1 must have exactly these nine headers: `Name`, `Project ID`, `Project Name`, `Phase Type`, `Questions`, `AI Confidence Score`, `Feedback`, `Suggestions`, `Overall Score`. Each successful evaluation appends one row of exactly nine values. Compliance details are displayed in the report but not written to Sheets, matching the original behavior.

## Files
TXT, PDF, DOCX, PPTX, RTF, MD, CSV; max 8 MB per file. Legacy DOC/PPT conversion is not included because LibreOffice is unavailable in the Vercel serverless runtime. Convert those to DOCX/PPTX first. Scanned PDFs need OCR.

## Important
Vercel execution time limits depend on plan. The API performs one Gemini request, with sequential model fallback; long requests may time out on plans with short limits. There is no authentication in this sample, so restrict access before handling private student data. Never commit API keys or service-account credentials to GitHub.
