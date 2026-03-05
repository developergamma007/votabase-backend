# Implementation Document

## Project
- Repository: `votabase-backend`
- Date: March 6, 2026

## User Requests Captured
1. Run the app.
2. Create one implementation document containing every point requested with proposed solutions.

## Work Completed
1. Inspected project structure and startup docs.
2. Identified runtime stack: FastAPI + Uvicorn + PostgreSQL.
3. Created local virtual environment: `.venv`.
4. Installed dependencies from `requirements.txt`.
5. Attempted to start app:
   - Command: `source .venv/bin/activate && uvicorn app.main:app --host 0.0.0.0 --port 8081`

## Observed Issue
- App startup fails during DB-dependent startup logic (super admin seed path).
- Error summary:
  - Could not connect to PostgreSQL host `65.0.75.0:5432`
  - Failure mode observed: `Operation not permitted` / `Operation timed out`
- Effect:
  - API never finishes startup.
  - Port `8081` is not bound.

## Root Cause
- The application requires active PostgreSQL connectivity during startup.
- Current configured database endpoint is unreachable from this run environment.

## Proposed Solution
### Option A (Recommended): Use a reachable dev/local PostgreSQL
1. Update `.env` `DATABASE_URL` to a reachable DB (local docker or accessible dev DB).
2. Verify connectivity with a DB check (`psql` or SQLAlchemy test connect).
3. Start app again on `8081`.
4. Validate API responds at:
   - `http://127.0.0.1:8081/votebase/v1`

### Option B: Make startup DB seed optional
1. Guard startup seeding with an env flag (example: `SEED_ON_STARTUP=true/false`).
2. Skip seed flow when DB is unavailable or when flag is disabled.
3. Allow API process boot for non-DB smoke checks.
4. Re-enable seeding in connected environments.

### Option C: Add resilience + diagnostics
1. Add retry with timeout/backoff for initial DB connect.
2. Improve startup log message with explicit DB host/port and env hint.
3. Fail fast with clear actionable error if DB remains unreachable.

## Implementation Plan (Pragmatic)
1. Confirm preferred approach (A/B/C or combo).
2. If A:
   - Update `.env` and rerun app.
3. If B:
   - Patch startup logic in `app/main.py` to gate seeding by env.
4. Validate with:
   - Process starts cleanly
   - Port `8081` listening
   - Base route responds

## Acceptance Criteria
1. App process starts without crashing.
2. `uvicorn` binds `0.0.0.0:8081`.
3. Endpoint under `/votebase/v1` responds (even if auth-protected).
4. Startup logs are actionable for DB failure scenarios.

## Notes
- No destructive changes were made.
- Dependencies are installed in `.venv`.
