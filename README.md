# REC 2001 - Batchmates: Names, Roll Nos, Branches & Comments

Public GitHub Pages site backed live by a private Google Sheet
("REC 2001 - Names, Roll Nos & Comments", Entries tab) through a Google
Apps Script web app (`Code.gs`, deployed as "REC 2001 Comments Backend",
access: Anyone, execute as owner).

- Anyone with the link can view, add and edit rows; changes appear
  immediately. There is no delete on the page or in the backend.
- `data.js` is only an offline fallback snapshot; the page live-reads the
  sheet via JSONP and renders whatever the sheet holds.
- Every change lands in the sheet, whose native version history is the
  recoverable backup.

To change the backend: edit the "REC 2001 Comments Backend" Apps Script
project and redeploy (new version), then update `WEB_APP_URL` in
`index.html` if the URL changes.
