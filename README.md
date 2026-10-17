# Supply Day Book

Delivery and collection app for supplymen. Opens on the phone with no network, saves every entry,
photo and GPS fix locally, and sends them to the office Google Sheet when signal returns.

- App (install from phone): https://netpill24x7-byte.github.io/supply-app/
- Office backend: `backend/Code.gs` (Google Apps Script bound to the "Supply Day Book" sheet)
- Server URL is set in `config.js`

## Sheet tabs
- **Supplymen** — id, name, pin, role (`supplyman` / `admin`), active
- **Bills** — date, supplyman_id, bill_no, party, amount, mobile, otp (optional). Paste the morning Marg list here.
- **Entries** — written by the phones. One row per entry; photos go to the Drive folder "Supply Day Book Photos".
