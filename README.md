# VoteBase Python Backend

This is a Python/FastAPI migration of the original Java Spring Boot backend.

## What is preserved
- Same base path: `/votebase/v1`
- Same API routes (auth, user, tenant, voter, family, booth, association, volunteer stats/dropdown, excel upload, snapshot)
- Same JWT claims/format and role checks
- Same PostgreSQL schemas/tables (`metastore`, `data`, `snapshot`, `error`)
- Same response envelope where Java used `ApiResponse`

## Setup

### If you are at repo root (`votabase-backend`)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r python_backend/requirements.txt
```

### If you are inside `python_backend`
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run Python API

### From repo root
```bash
source .venv/bin/activate
cp python_backend/.env.example python_backend/.env
uvicorn python_backend.app.main:app --host 0.0.0.0 --port 8081
```

### From `python_backend`
```bash
source .venv/bin/activate
cp .env.example .env
uvicorn app.main:app --host 0.0.0.0 --port 8081
```

API base URL:
- `http://localhost:8081/votebase/v1`

## Endpoint Parity Test Suite (Java vs Python)

Tests compare status/body across both servers endpoint-by-endpoint.

1. Start Java API (recommended on `8082`) and Python API (recommended on `8081`).
2. Run parity tests from repo root:

```bash
source .venv/bin/activate
JAVA_BASE_URL=http://127.0.0.1:8082/votebase/v1 \
PY_BASE_URL=http://127.0.0.1:8081/votebase/v1 \
pytest python_backend/tests/test_parity.py
```

Or:
```bash
source .venv/bin/activate
./python_backend/run_parity_tests.sh
```

Notes:
- Tests use existing DB users (`SUPER_ADMIN`, first `ADMIN`, first `USER`) for login.
- Tests focus on non-destructive endpoints.
- Dynamic fields (token/timestamps/presigned URLs) are normalized before comparison.

## Config

Python config is fully separate now and uses only `.env` / environment variables.

Required:
- `DATABASE_URL`

Optional:
- `AWS_REGION`
- `AWS_S3_BUCKET`
- `AWS_ACCESS_KEY_ID`
- `AWS_SECRET_ACCESS_KEY`
- `AWS_S3_PROFILE_PICS`
- `CONTEXT_PATH`
- `PORT`
- `SUPERADMIN_USERNAME`

## Notes
- This migration keeps DB unchanged and operates against existing tables.
- Super admin seeding is preserved at startup (`role=SUPER_ADMIN`).
