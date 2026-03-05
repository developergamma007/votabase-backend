#!/usr/bin/env bash
set -euo pipefail

JAVA_BASE_URL="${JAVA_BASE_URL:-http://127.0.0.1:8082/votebase/v1}"
PY_BASE_URL="${PY_BASE_URL:-http://127.0.0.1:8081/votebase/v1}"
export JAVA_BASE_URL
export PY_BASE_URL

pytest python_backend/tests/test_parity.py \
  --maxfail=1 \
  -q
