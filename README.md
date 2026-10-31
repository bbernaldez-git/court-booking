# Basketball court booking

## 1. Google Sheet + Apps Script
1. Create a new Google Sheet. Open **Extensions > Apps Script**, paste `apps-script/Code.gs`.
2. Edit `CONFIG` at the top (court name, hours, rate, GCash/Maya details).
3. **Project Settings > Script properties**: add `ADMIN_PASSWORD` = a strong password.
4. Select `setup` in the toolbar and click **Run** (grants Sheets/Drive/Mail permissions). It creates the `Bookings` tab and a Drive folder for screenshots.
5. **Deploy > New deployment > Web app**: Execute as **Me**, Who has access **Anyone**. Copy the `/exec` URL.
6. After any code change, use **Deploy > Manage deployments > Edit > New version**.

## 2. Frontend on Vercel
1. Paste the `/exec` URL into `config.js`.
2. Push this folder (without `apps-script/` if you like) to GitHub and import it in Vercel (framework preset: **Other**, no build command), or run `npx vercel --prod` in this folder.
3. Customers use `your-app.vercel.app`; you use `your-app.vercel.app/admin.html`.

## Flow
Customer books and uploads proof, row is saved as **Pending** and the slot is held. Admin opens the screenshot, taps **Confirm** (or **Reject** to free the slot). Confirmed bookings email the customer if they gave an email.

## Notes
- Screenshots stay private in your Drive; open them while signed in to the Google account that owns the script.
- Set the script time zone (Project Settings) to Asia/Manila so "past slot" checks are right.
- Everything also lives in the Sheet, so you can edit the `Status` column manually.
