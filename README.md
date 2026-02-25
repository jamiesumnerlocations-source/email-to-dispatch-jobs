# email-to-dispatch-jobs
# Email → Dispatch Jobs (Google Apps Script)

A Google Apps Script pipeline that scans Gmail from authorised senders, extracts dispatch job details from unstructured emails, and writes structured rows to a Google Sheet.

## What it does
- Searches recent Gmail threads from an allowlist of senders
- Parses email body text to extract:
  - Date context (weekday headers or bare dates)
  - Time (24h or am/pm)
  - Origin / destination (labels or inline “from/to” text)
  - Vehicle counts (truck/van/car)
- Writes one row per dispatchable “move” into a `Jobs` sheet
- Optionally calculates distance and duration via Google Maps Directions
- Dedupe key prevents duplicate job creation:
  `SourceEmailId + Date + Time + Origin + Destination`

## Why it exists
Dispatch requests often arrive as semi-structured emails. This script converts them into a consistent job table automatically, reducing manual admin and improving reliability.

## Sheet requirements
Create a Google Sheet with a `Jobs` tab and headers including (example):
- JobID, CreatedAt, SourceEmailId, Requester
- VehicleType, Date, Time, Origin, Destination
- DistanceKm, DurationMins, MapsUrl
- Status, Notes

## Setup
1) Copy `src/code.gs` into an Apps Script project attached to your Sheet.
2) Set:
   - `CONFIG.spreadsheetId`
   - `CONFIG.authorisedSenders`
3) Reload the Sheet; use the `Dispatch` menu → “Scan authorised Gmail → Jobs”.

## Notes / Future improvements
- Add support for attachments (PDF OCR) if needed
- Expand parsing patterns for more email formats
- Add structured logging and alerting for parsing failures
