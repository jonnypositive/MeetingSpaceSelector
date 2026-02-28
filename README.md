# Hotel Polaris Space Suggester (Phase 1)

Phase 1 includes:
- Capacity-chart knowledge base loaded from `capacity-chart.pdf` into `room_catalog.json`
- Cvent RFP PDF upload UI
- Header extraction for:
  - Account Name
  - Group Name / RFP Title
  - Arrival Date
  - Departure Date
  - RFP Contact Name
  - RFP Contact Company
- Best-fit room recommendations by attendee count + requested setup style
- Recommendation-only output (no exclusion list)
- Salesforce (Delphi) integration scaffold UI (no credential storage, no live auth yet)

## Run

```bash
cd "/Users/jonathanmiller/Documents/App Development Sandbox/MeetingSpaceSelector"
python3 app.py
```

Then open: `http://127.0.0.1:8080`

## Notes

- The parser is currently optimized for Cvent-style RFP PDFs.
- Salesforce/Delphi in Phase 2 should use OAuth 2.0 + MFA (Authenticator) rather than username/password form auth.

## Deploy (Render + Custom Subdomain)

This app is server-rendered and has backend API routes, so it cannot run on GitHub Pages by itself.

1. Push this repo to GitHub.
2. In Render, create a new `Web Service` from this repo.
3. Render will detect `render.yaml` and use:
   - build: `pip install -r requirements.txt`
   - start: `python app.py`
4. After deploy, open the Render service URL and confirm the app loads.
5. In Render service settings, add custom domain: `spaceselector.reunited.us`.
6. In your DNS provider for `reunited.us`, create/update the `CNAME` for `spaceselector` to the Render target shown in Render.
7. In GitHub repo settings, disable GitHub Pages for this project (to avoid conflicting DNS/hosting behavior).

After DNS propagation, `https://spaceselector.reunited.us` should serve the live app and backend endpoints from the same host.
