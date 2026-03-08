# Implementation Document

## Project Scope
- Backend: `votabase-backend`
- Frontend: `Votabase-mobile-ui`
- Runtime target DB: `Survey_Production (13.233.40.235:5432/surveydb)`
- Date: March 9, 2026

## Requested Outcomes
1. Backend and mobile app must connect correctly.
2. APIs should use Survey_Production data, especially `public.assembly`, `public.wards`, `public.booths`, `public.voters`.
3. Resolve auth and visibility issues for booth/snapshot APIs.
4. Fix large snapshot response problem (5+ lakh voters) causing Postman and iOS failures.
5. Keep response structures compatible with app usage.
6. Ensure local backend logs clearly show incoming API hits and DB target.

## Final Design Implemented
1. Keep snapshot API as link-based response to avoid huge direct payload.
2. Add lightweight snapshot mode for booth listing:
   - `GET /votebase/v1/api/voters/snapshot?assemblyCode=...&includeVoters=false`
   - Returns all wards/booths + `voterStats`, with `voters: []`.
3. Add booth-level lazy voter load:
   - `GET /votebase/v1/api/voters/by-booth?boothId=...`
   - Returns one booth with full voters list.
4. Update mobile to use lightweight snapshot for Search Booth and fetch booth voters on tap.
5. Cache only lite snapshot for booth search use-case.

## Backend Changes
File: `votabase-backend/app/main.py`

1. Environment loading
- `.env` values now override inherited shell envs for deterministic runtime.

2. Startup/diagnostic logs
- Added DB target startup log:
  - `[DB_TARGET] host:port/db`
- Fixed DB target regex parsing.

3. Request logging middleware
- Added request logs for key APIs with auth presence, status, duration.
- Added masked body logging for selected endpoints (phone/token/password masked).

4. Booth API source change
- `/votebase/v1/api/booth` now reads from `public.booths`.
- Returns all booths for authenticated users.

5. Snapshot link mode
- `/voters/snapshot` returns URL in `data.result`.
- `/voters/snapshot/content/{snapshot_id}` serves cached snapshot payload.
- Added short-lived in-memory cache and cleanup.

6. Public table snapshot mapper
- Added dynamic schema column detection via `information_schema`.
- Added `_build_public_snapshot(..., include_voters=bool)` mapping:
  - `public.wards` + `public.booths` + `public.voters`
  - Compatible output shape: `assembly -> wards -> booths -> voters`.

7. Lightweight snapshot mode
- `includeVoters=false` support:
  - omits large voter arrays in snapshot
  - adds per-booth stats:
    - `voterStats.total`
    - `voterStats.male`
    - `voterStats.female`

8. New per-booth voters endpoint
- Added `/votebase/v1/api/voters/by-booth?boothId=...`
- Uses `public.booths/public.voters/public.wards`
- Returns booth metadata + full voters + `voterStats`.

## Frontend Changes
Files:
- `Votabase-mobile-ui/src/apis/Api.js`
- `Votabase-mobile-ui/src/screens/VotersManagement/SearchBooth.tsx`
- `Votabase-mobile-ui/src/screens/LoginManagement/LoadData.tsx`

1. New API wrappers
- `loadDataLite()` calls snapshot with `includeVoters=false`.
- `fetchBoothVoters(boothId)` calls new `by-booth` endpoint.

2. SearchBooth flow updated
- Uses lightweight snapshot endpoint.
- Stores compact snapshot in AsyncStorage key: `boothSnapshotLite`.
- Displays stats from `voterStats` without loading full voters.
- On booth tap: fetches full voters for selected booth only, then navigates.

3. LoadData behavior updated
- Loads/stores lite snapshot (`boothSnapshotLite`) instead of huge full payload for this flow.

## Root Cause of "String length exceeds limit"
1. Mobile app attempted to parse/store very large JSON snapshot containing all voters.
2. JS engine / bridge / AsyncStorage limits triggered on large string/object operations.
3. Postman also hit maximum response size when requesting huge content links.

## Fix Strategy for 5-Lakh Response Issue
1. Do not return full dataset in one request for UI bootstrap.
2. Return compact structure + counts for listing/search pages.
3. Fetch full voters only for selected booth (lazy load).
4. Optional next step: paginate `by-booth` results if some booths are very large.

## Validation Performed
1. Verified backend points to Survey_Production from startup log:
   - `[DB_TARGET] 13.233.40.235:5432/surveydb`
2. Verified compact snapshot on live backend:
   - `WARDS = 27`
   - `BOOTHS = 508`
   - First booth has `voterStats`
   - First booth `voters` length is `0` when `includeVoters=false`
3. Verified per-booth endpoint returns full voters and correct stats.
4. Verified API hit logs appear in backend terminal for requested endpoints.

## Operational Notes
1. If old behavior appears, backend was running stale process.
2. Must restart backend on `8082` after code changes.
3. For Postman:
   - Use snapshot lite endpoint first.
   - Then open returned content link.
   - Do not test old heavy content links.

## Current Recommended API Usage Pattern
1. App startup / booth search:
   - `GET /voters/snapshot?...&includeVoters=false`
2. Booth open:
   - `GET /voters/by-booth?boothId=...`
3. Optional voter-global search pages:
   - keep scoped/filter-based loading, avoid full assembly-wide voter blob.
