import json
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

import psycopg2
import pytest
import requests


JAVA_BASE = os.getenv("JAVA_BASE_URL", "http://127.0.0.1:8082/votebase/v1")
PY_BASE = os.getenv("PY_BASE_URL", "http://127.0.0.1:8081/votebase/v1")
TIMEOUT = int(os.getenv("PARITY_TIMEOUT", "20"))


def _sqla_to_conninfo(sqla_url: str) -> str:
    m = re.match(r"postgresql(?:\+psycopg2)?://([^:]+):([^@]+)@([^:/]+)(?::(\d+))?/([^?]+)", sqla_url)
    if not m:
        raise RuntimeError(
            "Invalid DATABASE_URL. Expected: postgresql+psycopg2://user:pass@host:5432/dbname"
        )
    user, password, host, port, db = m.group(1), m.group(2), m.group(3), m.group(4) or "5432", m.group(5)
    return f"host={host} port={port} dbname={db} user={user} password={password}"


@dataclass
class Principal:
    first_name: str
    phone: str
    tenant_id: Optional[str] = None
    assignment_type: Optional[str] = None
    assignment_id: Optional[int] = None


@pytest.fixture(scope="session")
def db_conn():
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        pytest.skip("DATABASE_URL is required for parity tests")
    conninfo = _sqla_to_conninfo(db_url)
    conn = psycopg2.connect(conninfo)
    conn.autocommit = True
    try:
        yield conn
    finally:
        conn.close()


@pytest.fixture(scope="session")
def principals(db_conn):
    cur = db_conn.cursor()

    superadmin_username = os.getenv("SUPERADMIN_USERNAME", "admin@iswot.io")
    superadmin = Principal(first_name=superadmin_username, phone="8867038709")

    cur.execute(
        """
        SELECT u.first_name, u.phone, t.tenant_id, u.assignment_type, u.assignment_id
        FROM metastore.users u
        JOIN metastore.tenant t ON u.tenant_id = t.id
        WHERE u.role = 'ADMIN' AND COALESCE(u.blocked, false) = false AND COALESCE(u.deleted, false) = false
        ORDER BY u.id ASC
        LIMIT 1
        """
    )
    row = cur.fetchone()
    if row is None:
        pytest.skip("No ADMIN user found in DB for parity tests")
    admin = Principal(*row)

    cur.execute(
        """
        SELECT u.first_name, u.phone, t.tenant_id, u.assignment_type, u.assignment_id
        FROM metastore.users u
        JOIN metastore.tenant t ON u.tenant_id = t.id
        WHERE u.role = 'USER' AND COALESCE(u.blocked, false) = false AND COALESCE(u.deleted, false) = false
        ORDER BY u.id ASC
        LIMIT 1
        """
    )
    row = cur.fetchone()
    user = Principal(*row) if row else None

    cur.execute("SELECT tenant_id FROM metastore.tenant ORDER BY id ASC LIMIT 1")
    tenant_row = cur.fetchone()
    sample_tenant_id = tenant_row[0] if tenant_row else None

    cur.execute(
        """
        SELECT a.assembly_code
        FROM data.assembly_details a
        WHERE a.tenant_id = %s
        ORDER BY a.assembly_id ASC
        LIMIT 1
        """,
        (admin.tenant_id,),
    )
    r = cur.fetchone()
    sample_assembly_code = r[0] if r else None

    cur.execute(
        """
        SELECT b.booth_id
        FROM data.booth_details b
        WHERE b.tenant_id = %s
        ORDER BY b.booth_id ASC
        LIMIT 1
        """,
        (admin.tenant_id,),
    )
    r = cur.fetchone()
    sample_booth_id = r[0] if r else None

    cur.execute(
        """
        SELECT family_id
        FROM data.family
        WHERE tenant_id = %s AND COALESCE(deleted,false)=false
        ORDER BY family_id ASC
        LIMIT 1
        """,
        (admin.tenant_id,),
    )
    r = cur.fetchone()
    sample_family_id = r[0] if r else None

    return {
        "superadmin": superadmin,
        "admin": admin,
        "user": user,
        "sample_tenant_id": sample_tenant_id,
        "sample_assembly_code": sample_assembly_code,
        "sample_booth_id": sample_booth_id,
        "sample_family_id": sample_family_id,
    }


def _login(base: str, p: Principal) -> str:
    resp = requests.post(
        f"{base}/api/auth/login",
        json={"firstName": p.first_name, "phone": p.phone},
        timeout=TIMEOUT,
    )
    assert resp.status_code == 200, f"Login failed on {base}: {resp.status_code} {resp.text}"
    return resp.json()["data"]["result"]["token"]


@pytest.fixture(scope="session")
def tokens(principals):
    out = {
        "superadmin": _login(PY_BASE, principals["superadmin"]),
        "admin": _login(PY_BASE, principals["admin"]),
    }
    if principals["user"]:
        out["user"] = _login(PY_BASE, principals["user"])

    # Ensure Java login also works (sanity).
    _ = _login(JAVA_BASE, principals["superadmin"])
    _ = _login(JAVA_BASE, principals["admin"])
    if principals["user"]:
        _ = _login(JAVA_BASE, principals["user"])

    return out


def _normalize(value: Any) -> Any:
    if isinstance(value, dict):
        out = {}
        for k, v in value.items():
            if k in {
                "token",
                "timestamp",
                "lastUpdated",
                "createdAt",
                "updatedAt",
                "profilePicUrl",
                "s3Url",
            }:
                continue
            out[k] = _normalize(v)
        return out
    if isinstance(value, list):
        return [_normalize(v) for v in value]
    return value


def _request_pair(
    method: str,
    path: str,
    token: Optional[str] = None,
    params: Optional[Dict[str, Any]] = None,
    json_body: Optional[Dict[str, Any]] = None,
) -> Tuple[requests.Response, requests.Response]:
    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    j = requests.request(method, f"{JAVA_BASE}{path}", headers=headers, params=params, json=json_body, timeout=TIMEOUT)
    p = requests.request(method, f"{PY_BASE}{path}", headers=headers, params=params, json=json_body, timeout=TIMEOUT)
    return j, p


def assert_parity(j: requests.Response, p: requests.Response):
    assert j.status_code == p.status_code, f"Status mismatch\nJava: {j.status_code} {j.text}\nPy: {p.status_code} {p.text}"

    if not j.text and not p.text:
        return

    try:
        j_data = _normalize(j.json())
        p_data = _normalize(p.json())
    except Exception:
        assert j.text == p.text, f"Body mismatch\nJava: {j.text}\nPy: {p.text}"
        return

    assert j_data == p_data, "JSON mismatch\nJava:\n%s\n\nPy:\n%s" % (
        json.dumps(j_data, indent=2, sort_keys=True),
        json.dumps(p_data, indent=2, sort_keys=True),
    )


def test_auth_login_parity(principals):
    j, p = _request_pair(
        "POST",
        "/api/auth/login",
        json_body={"firstName": principals["admin"].first_name, "phone": principals["admin"].phone},
    )
    assert_parity(j, p)


def test_tenant_list_parity(tokens):
    j, p = _request_pair("GET", "/api/tenant", token=tokens["superadmin"], params={"page": 0, "size": 5})
    assert_parity(j, p)


def test_tenant_get_parity(tokens, principals):
    tenant_id = principals["sample_tenant_id"]
    if not tenant_id:
        pytest.skip("No tenant available")
    j, p = _request_pair("GET", f"/api/tenant/{tenant_id}", token=tokens["superadmin"])
    assert_parity(j, p)


def test_user_list_parity(tokens):
    j, p = _request_pair(
        "GET",
        "/api/user",
        token=tokens["admin"],
        params={"page": 0, "size": 10, "sortBy": "firstName", "direction": "asc"},
    )
    assert_parity(j, p)


def test_profile_parity(tokens):
    j, p = _request_pair("GET", "/api/user/profile", token=tokens["admin"])
    assert_parity(j, p)


def test_assignments_parity(tokens):
    for kind in ["ASSEMBLY", "WARD", "BOOTH"]:
        j, p = _request_pair("GET", "/api/assignments", token=tokens["admin"], params={"type": kind})
        assert_parity(j, p)


def test_booth_parity(tokens):
    j, p = _request_pair("GET", "/api/booth", token=tokens["admin"])
    assert_parity(j, p)


def test_association_list_parity(tokens, principals):
    booth_id = principals["sample_booth_id"]
    if not booth_id:
        pytest.skip("No booth available")
    j, p = _request_pair("GET", "/api/association", token=tokens["admin"], params={"boothId": booth_id})
    assert_parity(j, p)


def test_family_list_parity(tokens, principals):
    booth_id = principals["sample_booth_id"]
    if not booth_id:
        pytest.skip("No booth available")
    j, p = _request_pair("GET", "/api/family", token=tokens["admin"], params={"boothId": booth_id, "page": 0, "size": 5})
    assert_parity(j, p)


def test_family_get_parity(tokens, principals):
    family_id = principals["sample_family_id"]
    if not family_id:
        pytest.skip("No family available")
    j, p = _request_pair("GET", f"/api/family/{family_id}", token=tokens["admin"])
    assert_parity(j, p)


def test_volunteer_dropdown_parity(tokens):
    j, p = _request_pair("GET", "/api/volunteers/dropdown", token=tokens["admin"], params={"level": "ASSEMBLY"})
    assert_parity(j, p)


def test_voter_gender_stats_parity(tokens):
    j, p = _request_pair("GET", "/api/voters/stats/gender", token=tokens["admin"])
    assert_parity(j, p)


def test_volunteer_stats_parity(tokens, principals):
    params = {}
    admin = principals["admin"]
    if admin.assignment_type and admin.assignment_id is not None:
        params = {"level": admin.assignment_type, "id": admin.assignment_id}
    j, p = _request_pair("GET", "/api/volunteers/stats", token=tokens["admin"], params=params)
    assert_parity(j, p)


def test_voter_snapshot_parity(tokens, principals):
    assembly_code = principals["sample_assembly_code"]
    if not assembly_code:
        pytest.skip("No assembly found")
    j, p = _request_pair("GET", "/api/voters/snapshot", token=tokens["admin"], params={"assemblyCode": assembly_code})
    assert_parity(j, p)
