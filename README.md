# Lead-Capture-Automation-Google-Workspace
Free lead capture system using Google Forms, Sheets and Apps Script
# Lead Capture Automation (Google Workspace)

A free, no-code-tool lead capture system built with **Google Forms, Google Sheets and Google Apps Script**.

When someone submits the "Request a Quote" form:
- ✅ Their details are saved to Google Sheets automatically
- 📧 They get an instant, branded thank-you email
- 🔔 The business owner gets an email alert with a link to the lead
- 🏷️ The lead is marked **New** in a colour-coded Status column (New → Contacted → Quote sent → Won / Lost)

No Zapier, no paid tools, no monthly fees. It runs on a free Google account.


## Tech
Google Apps Script (JavaScript) · Google Forms · Google Sheets · Gmail (MailApp) · Installable triggers

## Setup
1. Create a Google Form and link it to a Google Sheet.
2. In the Sheet, open **Extensions → Apps Script** and paste in `LeadCapture.gs`.
3. Edit the `CONFIG` block (your email, business name, tab name).
4. Run `setup` once and approve the permissions.
5. Submit the form to test.
