# Connect the shared system (about 5 minutes)

1. Create a new Google Sheet (name it e.g. "MBR2 Tracker data").
2. **Extensions → Apps Script**. Delete the sample code, paste all of `Code.gs`, then **Save**.
3. Choose the function `setup` in the toolbar and press **Run**. Approve the permission prompt.
4. Back in the Sheet, open the **Users** tab. It now lists each person with a 4-digit PIN. Share each PIN privately. Edit them if you like.
5. In Apps Script: **Deploy → New deployment → Web app**. Execute as: **Me**. Who has access: **Anyone**. Press **Deploy** and copy the **Web app URL**.
6. In `index.html`, find `const API=''` near the top of the script and paste the URL between the quotes, e.g. `const API='https://script.google.com/macros/s/XXXX/exec'`.
7. Upload the changed `index.html` to GitHub.

If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → New version** so the same URL keeps working.
Staff can only enter data for their own lines. "Manager" can enter for everyone.
