# Team Target Run

A live, game-style running leaderboard for Team Chawin.

## Data source
Google Sheet: `Team Records (Few X PS)` → `TPSep2026`

The Apps Script filters rows where column A (`Team`) starts with `Chawin`, so it includes Chawin (SM), Chawin (Pik), Chawin (Little), etc. Junior's team is excluded.

## Make it live
1. Open the Google Sheet.
2. Extensions → Apps Script.
3. Paste `google-apps-script.gs`.
4. Deploy → New deployment → Web app.
5. Execute as: Me. Who has access: Anyone.
6. Copy the `/exec` URL.
7. Paste it into `API_URL` at the top of `script.js`.

The website polls every 10 seconds and keeps the last successful data if a refresh fails.

## Local preview
Open `index.html` directly, or run a static server:

```bash
python3 -m http.server 8080
```

The project ships in demo mode using a snapshot of current TPSep2026 values until API_URL is set.
