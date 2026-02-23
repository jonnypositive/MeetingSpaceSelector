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
