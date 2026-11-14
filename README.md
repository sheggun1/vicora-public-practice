VICORA Public Practice

A separate free public CBT practice portal for WAEC, NECO, JAMB and general subjects. Student self-registration is open. Answers and explanations are revealed after submission.

Pages:
- /practice/ — student sign-in and practice dashboard
- /practice/register.html — open student registration
- /practice/admin/ — administrator question-bank dashboard

Backend: Google Apps Script + Google Sheets. Configure the Vercel environment variable VICORA_APPS_SCRIPT_URL with the deployed Apps Script /exec URL. The Apps Script project must contain the practice.* actions and the VICORA_PRACTICE_ADMIN_KEY Script Property. Never commit the administrator key.

Deployment: Import this repository into Vercel with the repository root as project root. No build command is required. Set VICORA_APPS_SCRIPT_URL for Production, Preview and Development, then deploy. Existing VICORA school CBT remains independent and untouched.
