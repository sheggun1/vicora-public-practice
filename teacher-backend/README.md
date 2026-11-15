# VICORA TEACHER — Foundation

This is the isolated foundation for VICORA TEACHER, a Nigerian academic workspace for Primary 1–6, JSS 1–3 and SS 1–3.

## Project decisions
- Separate teacher registration and sign-in.
- First registrant creating a school becomes School Admin; additional teachers join with the school join code.
- Dedicated Google Apps Script backend and Google Sheets database.
- Central Google Drive storage, organised by school, plus a global curriculum library.
- AI provider router/model IDs aligned with VICORA Public Practice.
- Print-ready A4 output through browser print and editable DOCX export.
- Responsive frontend uses the VICORA Practice navy/blue/teal visual language.

## Important separation
All new code is on branch `vicora-teacher-foundation` and under new paths. The existing Practice frontend, API, Apps Script project and database are not changed. Do not merge this branch into the live Practice branch until the separate Teacher deployment has been tested.

## Files
- `public/teacher/index.html` — Teacher workspace UI.
- `api/teacher/[...path].js` — server-side proxy to Apps Script.
- `teacher-backend/Code.gs` — dedicated Apps Script backend.
- `teacher-backend/appsscript.json` — Apps Script manifest.
- `docs/VICORA-TEACHER-TECHNICAL-BLUEPRINT.md` — architecture and delivery plan.

## Apps Script setup
1. Open Google Apps Script and create a standalone project named **VICORA TEACHER Backend**.
2. Copy `teacher-backend/Code.gs` into the Apps Script editor.
3. Open Project Settings, enable manifest file visibility, and replace the manifest with `teacher-backend/appsscript.json`.
4. Save. Select `setupVICORATeacher` and run it once. Approve the Google Sheets and Drive permissions.
5. Copy the one-time Platform Admin key from the execution log and store it securely. Do not commit it to GitHub or put it in frontend code.
6. Deploy → New deployment → Web app. Execute as **Me** and choose **Anyone** for access. Copy the URL ending in `/exec`.
7. In a separate Vercel project for this Teacher deployment, set `VICORA_TEACHER_APPS_SCRIPT_URL` to that /exec URL. Do not change the existing VICORA Practice project's environment variables.
8. Deploy this branch/root to the separate Teacher Vercel project. The interface is at `/teacher/`.

## AI setup
The provider IDs, model defaults and fallback order follow VICORA Public Practice. However, Apps Script Script Properties are isolated per project. Therefore, add the same API keys, model IDs and provider order in the Teacher Platform Admin → AI settings. The keys are stored server-side in the Teacher Apps Script project; they are never sent to the browser.

Supported provider IDs: `gemini, groq, openrouter, mistral, cloudflare, nvidia, cohere, sambanova, openai, meta, claude, grok`.

## Initial capabilities in this foundation
- Teacher registration and sign-in.
- School workspace creation and join-code registration.
- Platform Admin sign-in.
- Lesson note, examination, test, scheme, syllabus and report generation.
- Saved document list, document opening and editing.
- Browser print / Save as PDF.
- DOCX export through the browser document converter.
- PDF, DOCX and TXT text extraction in the browser and original file storage in Drive.
- School and global reference-library access.
- AI teaching assistant.
- Platform Admin school listing and AI provider settings.

## Current implementation notes
- Uploaded files are capped at 3 MB.
- The first document editor stores content in Google Sheets; keep generated document HTML/text below the Apps Script cell-size limit.
- PDF output uses the browser print dialog; choose Save as PDF.
- DOCX export depends on the external html-docx-js library loading successfully.
- PDF/DOCX text extraction is performed in the browser before the original file and extracted text are sent to the backend.
- The first retrieval implementation uses keyword matching over extracted text. Semantic/vector retrieval is a later enhancement.
- The first pass has no automated password-reset email flow, school-specific template builder, document deletion, full examination-format controls or curriculum approval workflow yet.
- Review all generated material before official classroom use. The AI must not be treated as an authority for official curriculum approval.

## Smoke-test checklist
1. Apps Script setup creates the database tabs and Drive folder.
2. Teacher can create a new school and sign in.
3. Second account can join using the school code.
4. Teacher cannot read another school's documents or resources.
5. School Admin can upload a curriculum file; teacher can see its reference metadata.
6. AI generation succeeds using a configured provider and saves a draft.
7. Document edits save and persist after re-login.
8. Print preview is A4 and has no application navigation.
9. DOCX opens in Microsoft Word or LibreOffice.
10. Platform Admin can view schools and save AI provider settings.
