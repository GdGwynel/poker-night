# Poker Night

Same idea as the wishlist: a page on GitHub Pages, with a Google Sheet storing the data.

## 1. Google Sheet
1. Create a new Google Sheet named "Poker Night".
2. Open Extensions → Apps Script, delete what's there, and paste in all of `Code.gs`.
3. On the line `const ADMIN_PIN = '0000';`, change `0000` to your own host PIN.
4. Optional: pick `importPastGames` in the function dropdown and click Run to load the 6 earlier games from ChatGPT. Approve the permission prompt.
5. Deploy → New deployment → type "Web app". Set "Execute as" to Me and "Who has access" to Anyone. Click Deploy and copy the URL that ends in `/exec`.

## 2. Page
1. In `index.html`, paste that URL into `const API_URL = '';`.
2. Create a GitHub repo (for example `poker-night`), upload `index.html`, and turn on Settings → Pages (branch `main`, root).
3. Send the link to the group chat.

If you change `Code.gs` later, use Deploy → Manage deployments → Edit → New version, so the URL stays the same.
