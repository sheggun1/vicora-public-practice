# VICORA Public Practice

A separate free public CBT practice portal for WAEC, NECO, JAMB and general subjects. Student self-registration is open. Answers and explanations are revealed after submission.

## Pages
- `/practice/` — student sign-in and practice dashboard
- `/practice/register.html` — open student registration
- `/practice/admin/` — administrator question-bank dashboard

## Separate Apps Script backend

The complete standalone Google Apps Script backend is in `apps-script/Code.gs`; its V8 manifest is `apps-script/appsscript.json`.

1. Open [Google Apps Script](https://script.google.com/) and create a **new standalone project** named **VICORA Public Practice Backend**.
2. Replace the default `Code.gs` with the contents of [apps-script/Code.gs](apps-script/Code.gs). Set the project manifest from [apps-script/appsscript.json](apps-script/appsscript.json), or leave the default V8 runtime.
3. Save, select `setupVICORAPublicPractice` in the function dropdown, and click **Run**. Approve the Google permissions.
4. Open the execution log. The setup function creates the dedicated **VICORA Public Practice Database** spreadsheet and its five tabs, generates the administrator key, and stores it in Script Properties. **Copy and save the administrator key securely.** It is shown in the setup result/log; it is not committed to this repository.
5. In Apps Script, choose **Deploy → New deployment → Web app**. Set **Execute as: Me** and **Who has access: Anyone**, then deploy. Copy the URL ending in `/exec`.
6. In the Vercel project settings, add/update `VICORA_APPS_SCRIPT_URL` for Production, Preview and Development with that `/exec` URL, then redeploy.
7. Open `/practice/admin/` and enter the generated administrator key. Add questions before students start a test.

The setup is safe to run again: it reuses the existing database and administrator key and does not clear records. If you lose the key, generate a replacement in Apps Script Script Properties and use the new value in the admin dashboard.

## Vercel deployment

Import this repository into Vercel with the repository root as project root. No build command is required. Set `VICORA_APPS_SCRIPT_URL` for Production, Preview and Development, then deploy.

The public practice backend and database are separate from the VICORA School CBT backend. Never put the administrator key in GitHub source code or in the Vercel environment variables; it is entered into the protected admin page and checked by Apps Script.
