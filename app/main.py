import json
import traceback
import os
import re
import tempfile
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import boto3
import jwt
from botocore.exceptions import NoCredentialsError
from fastapi import Depends, FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from openpyxl import load_workbook
from pydantic import BaseModel, Field
from sqlalchemy import Boolean, DateTime, Double, ForeignKey, Integer, String, and_, asc, case, create_engine, desc, func, or_, text
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker


# ---------------------------
# Config
# ---------------------------
def _load_dotenv_from_python_backend() -> None:
    env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            # Keep backend deterministic: values in votabase-backend/.env
            # should override inherited shell variables.
            os.environ[key.strip()] = value.strip().strip("\"'")


_load_dotenv_from_python_backend()

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL is required. Example: postgresql+psycopg2://user:pass@host:5432/dbname")

CONTEXT_PATH = os.getenv("CONTEXT_PATH", "/votebase/v1")
AWS_REGION = os.getenv("AWS_REGION", "ap-south-1")
AWS_BUCKET = os.getenv("AWS_S3_BUCKET", "")
AWS_ACCESS_KEY = os.getenv("AWS_ACCESS_KEY_ID", "")
AWS_SECRET_KEY = os.getenv("AWS_SECRET_ACCESS_KEY", "")
PROFILE_UPLOAD_DIR = os.getenv("AWS_S3_PROFILE_PICS", "profile_pics")
SUPERADMIN_USERNAME = os.getenv("SUPERADMIN_USERNAME", "admin@iswot.io")

JWT_SECRET = "supersecretkeysupersecretkeysupersecretkey"
JWT_EXPIRATION_SECONDS = 60 * 60 * 24 * 365
SNAPSHOT_CACHE_TTL_SECONDS = 60 * 15


# In-memory cache for large snapshot payloads served via short-lived links.
_snapshot_cache: Dict[str, Dict[str, Any]] = {}


def _safe_db_target(db_url: str) -> str:
    # Print DB target without exposing credentials.
    m = re.match(r"postgresql(?:\+psycopg2)?://[^@]+@([^:/]+)(?::(\d+))?/([^?]+)", db_url or "")
    if not m:
        return "unknown"
    host = m.group(1)
    port = m.group(2) or "5432"
    db = m.group(3)
    return f"{host}:{port}/{db}"


def _cleanup_snapshot_cache() -> None:
    now_ts = int(datetime.now(timezone.utc).timestamp())
    expired = [
        k
        for k, v in _snapshot_cache.items()
        if now_ts - int(v.get("created_ts", 0)) > SNAPSHOT_CACHE_TTL_SECONDS
    ]
    for k in expired:
        _snapshot_cache.pop(k, None)


def _cache_snapshot(snapshot: Dict[str, Any]) -> str:
    _cleanup_snapshot_cache()
    snapshot_id = uuid.uuid4().hex
    _snapshot_cache[snapshot_id] = {
        "created_ts": int(datetime.now(timezone.utc).timestamp()),
        "payload": snapshot,
    }
    return snapshot_id


# ---------------------------
# DB models
# ---------------------------
class Base(DeclarativeBase):
    pass


class Tenant(Base):
    __tablename__ = "tenant"
    __table_args__ = {"schema": "metastore"}

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20), unique=True)
    name: Mapped[str] = mapped_column(String(255))
    description: Mapped[Optional[str]] = mapped_column(String)
    contact_email: Mapped[str] = mapped_column(String(255), unique=True)
    contact_phone: Mapped[Optional[str]] = mapped_column(String(50))
    active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[Optional[datetime]] = mapped_column(DateTime)


class User(Base):
    __tablename__ = "users"
    __table_args__ = {"schema": "metastore"}

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_ref: Mapped[Optional[int]] = mapped_column("tenant_id", ForeignKey("metastore.tenant.id"))
    role: Mapped[str] = mapped_column(String(30))
    assignment_type: Mapped[Optional[str]] = mapped_column(String(30))
    assignment_id: Mapped[Optional[int]] = mapped_column(Integer)
    first_name: Mapped[str] = mapped_column(String(100), unique=True)
    phone: Mapped[str] = mapped_column(String(10), unique=True)
    profile_pic_url: Mapped[Optional[str]] = mapped_column(String)
    blocked: Mapped[bool] = mapped_column(Boolean, default=False)
    deleted: Mapped[bool] = mapped_column(Boolean, default=False)

    tenant: Mapped[Optional[Tenant]] = relationship(Tenant)


class VolunteerUser(Base):
    __tablename__ = "volunteer_users"
    __table_args__ = {"schema": "metastore"}

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    role: Mapped[str] = mapped_column(String(30), default="USER")
    working_level: Mapped[Optional[str]] = mapped_column(String(30))
    assignment_type: Mapped[Optional[str]] = mapped_column(String(30))
    assignment_id: Mapped[Optional[str]] = mapped_column(String)
    assembly_ids: Mapped[Optional[str]] = mapped_column(String)
    ward_ids: Mapped[Optional[str]] = mapped_column(String)
    booth_ids: Mapped[Optional[str]] = mapped_column(String)
    first_name: Mapped[str] = mapped_column(String(100))
    phone: Mapped[str] = mapped_column(String(10), unique=True)
    profile_pic_url: Mapped[Optional[str]] = mapped_column(String)
    blocked: Mapped[bool] = mapped_column(Boolean, default=False)
    deleted: Mapped[bool] = mapped_column(Boolean, default=False)

class Assembly(Base):
    __tablename__ = "assembly"
    __table_args__ = {"schema": "data"}

    assembly_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    assembly_name_en: Mapped[Optional[str]] = mapped_column(String(255))
    assembly_name_local: Mapped[Optional[str]] = mapped_column(String(255))
    assembly_code: Mapped[str] = mapped_column(String(12), unique=True)


class Ward(Base):
    __tablename__ = "wards"
    __table_args__ = {"schema": "data"}

    ward_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    assembly_id: Mapped[int] = mapped_column(ForeignKey("data.assembly.assembly_id"))
    tenant_id: Mapped[str] = mapped_column(String(20))
    ward_name_en: Mapped[Optional[str]] = mapped_column(String(255))
    ward_name_local: Mapped[Optional[str]] = mapped_column(String(255))
    ward_code: Mapped[str] = mapped_column(String(20), unique=True)


class Booth(Base):
    __tablename__ = "booths"
    __table_args__ = {"schema": "data"}

    booth_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    ward_id: Mapped[int] = mapped_column(ForeignKey("data.wards.ward_id"))
    tenant_id: Mapped[str] = mapped_column(String(20))
    polling_station_adr_en: Mapped[Optional[str]] = mapped_column(String(255))
    polling_station_adr_local: Mapped[Optional[str]] = mapped_column(String(255))


class Association(Base):
    __tablename__ = "association"
    __table_args__ = {"schema": "data"}

    association_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    association_name: Mapped[str] = mapped_column(String(100))
    booth_id: Mapped[Optional[int]] = mapped_column(ForeignKey("data.booths.booth_id"))
    association_address: Mapped[Optional[str]] = mapped_column(String)
    association_head_name: Mapped[Optional[str]] = mapped_column(String)
    phone: Mapped[Optional[str]] = mapped_column(String)
    latitude: Mapped[Optional[float]] = mapped_column(Double)
    longitude: Mapped[Optional[float]] = mapped_column(Double)
    tenant_id: Mapped[str] = mapped_column(String(20))


class Voter(Base):
    __tablename__ = "voters"
    __table_args__ = {"schema": "data"}

    voter_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    booth_id: Mapped[int] = mapped_column("booth_id", ForeignKey("data.booths.booth_id"))
    sr_no: Mapped[Optional[int]] = mapped_column(Integer)
    epic_no: Mapped[Optional[str]] = mapped_column(String(20), unique=True)
    first_middle_name_en: Mapped[Optional[str]] = mapped_column(String(150))
    last_name_en: Mapped[Optional[str]] = mapped_column(String(100))
    first_middle_name_local: Mapped[Optional[str]] = mapped_column(String(150))
    last_name_local: Mapped[Optional[str]] = mapped_column(String(100))
    relation_type: Mapped[Optional[str]] = mapped_column(String(20))
    relation_first_middle_name_en: Mapped[Optional[str]] = mapped_column(String(150))
    relation_last_name_en: Mapped[Optional[str]] = mapped_column(String(100))
    relation_first_middle_name_local: Mapped[Optional[str]] = mapped_column(String(150))
    relation_last_name_local: Mapped[Optional[str]] = mapped_column(String(100))
    house_no_en: Mapped[Optional[str]] = mapped_column(String(50))
    house_no_local: Mapped[Optional[str]] = mapped_column(String(50))
    gender: Mapped[Optional[str]] = mapped_column(String(10))
    age: Mapped[Optional[int]] = mapped_column(Integer)
    dob: Mapped[Optional[datetime]] = mapped_column(DateTime)
    mobile: Mapped[Optional[str]] = mapped_column(String(15))
    address_en: Mapped[Optional[str]] = mapped_column(String(255))
    address_local: Mapped[Optional[str]] = mapped_column(String(255))
    status: Mapped[Optional[str]] = mapped_column(String(20))
    community: Mapped[Optional[str]] = mapped_column(String(100))
    caste: Mapped[Optional[str]] = mapped_column(String(100))
    residence_type: Mapped[Optional[str]] = mapped_column(String(100))
    civic_issue: Mapped[Optional[str]] = mapped_column(String(255))
    mother_tongue: Mapped[Optional[str]] = mapped_column(String(100))
    team: Mapped[Optional[str]] = mapped_column(String(100))
    ownership: Mapped[Optional[str]] = mapped_column(String(20))
    education: Mapped[Optional[str]] = mapped_column(String(20))
    nature_of_voter: Mapped[Optional[str]] = mapped_column(String(100))
    latitude: Mapped[Optional[float]] = mapped_column(Double)
    longitude: Mapped[Optional[float]] = mapped_column(Double)


class Family(Base):
    __tablename__ = "family"
    __table_args__ = {"schema": "data"}

    familyId: Mapped[int] = mapped_column("family_id", Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    family_name: Mapped[str] = mapped_column(String(30))
    family_address: Mapped[Optional[str]] = mapped_column(String(555))
    head_voter_id: Mapped[Optional[int]] = mapped_column(ForeignKey("data.family_members.member_id"))
    phone: Mapped[Optional[str]] = mapped_column(String(15))
    points: Mapped[Optional[int]] = mapped_column(Integer)
    points_provided: Mapped[Optional[int]] = mapped_column(Integer)
    latitude: Mapped[Optional[float]] = mapped_column(Double)
    longitude: Mapped[Optional[float]] = mapped_column(Double)
    booth_id: Mapped[int] = mapped_column(ForeignKey("data.booths.booth_id"))
    association_id: Mapped[Optional[int]] = mapped_column(ForeignKey("data.association.association_id"))
    economic_status: Mapped[Optional[str]] = mapped_column(String(50))
    family_nature: Mapped[Optional[str]] = mapped_column(String(50))
    deleted: Mapped[bool] = mapped_column(Boolean, default=False)


class FamilyMember(Base):
    __tablename__ = "family_members"
    __table_args__ = {"schema": "data"}

    member_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    family_id: Mapped[int] = mapped_column(ForeignKey("data.family.family_id"))
    voter_id: Mapped[int] = mapped_column(ForeignKey("data.voters.voter_id"))
    is_head: Mapped[bool] = mapped_column(Boolean, default=False)


class VoterSnapshot(Base):
    __tablename__ = "voter_snapshot"
    __table_args__ = {"schema": "snapshot"}

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    assembly_code: Mapped[str] = mapped_column(String(12))
    ward_code: Mapped[Optional[str]] = mapped_column(String(20))
    booth_id: Mapped[Optional[int]] = mapped_column(Integer)
    s3_url: Mapped[str] = mapped_column(String)
    snapshot_level: Mapped[str] = mapped_column(String(20))
    version: Mapped[int] = mapped_column(Integer, default=1)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[Optional[datetime]] = mapped_column(DateTime)


class VoterChangeLog(Base):
    __tablename__ = "voter_changelog"
    __table_args__ = {"schema": "data"}

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(20))
    voter_id: Mapped[int] = mapped_column(ForeignKey("data.voters.voter_id"))
    updated_by: Mapped[int] = mapped_column(ForeignKey("metastore.users.id"))
    field_name: Mapped[str] = mapped_column(String(100))
    old_value: Mapped[Optional[str]] = mapped_column(String)
    new_value: Mapped[Optional[str]] = mapped_column(String)
    updated_at: Mapped[datetime] = mapped_column(DateTime)
    update_latitude: Mapped[Optional[float]] = mapped_column(Double)
    update_longitude: Mapped[Optional[float]] = mapped_column(Double)


class VoterEnrichment(Base):
    __tablename__ = "voter_enrichment"
    __table_args__ = {"schema": "public"}

    epic: Mapped[str] = mapped_column(String(20), primary_key=True)
    ward_code: Mapped[Optional[str]] = mapped_column(String(20))
    booth_no: Mapped[Optional[str]] = mapped_column(String(20))
    first_middle_name_en: Mapped[Optional[str]] = mapped_column(String)
    last_name_en: Mapped[Optional[str]] = mapped_column(String)
    first_middle_name_local: Mapped[Optional[str]] = mapped_column(String)
    last_name_local: Mapped[Optional[str]] = mapped_column(String)
    relation_type: Mapped[Optional[str]] = mapped_column(String)
    relation_first_middle_name_en: Mapped[Optional[str]] = mapped_column(String)
    relation_last_name_en: Mapped[Optional[str]] = mapped_column(String)
    relation_first_middle_name_local: Mapped[Optional[str]] = mapped_column(String)
    relation_last_name_local: Mapped[Optional[str]] = mapped_column(String)
    house_no_en: Mapped[Optional[str]] = mapped_column(String)
    house_no_local: Mapped[Optional[str]] = mapped_column(String)
    gender: Mapped[Optional[str]] = mapped_column(String(20))
    age: Mapped[Optional[int]] = mapped_column(Integer)
    dob: Mapped[Optional[str]] = mapped_column(String(50))
    mobile: Mapped[Optional[str]] = mapped_column(String(30))
    address_en: Mapped[Optional[str]] = mapped_column(String)
    address_local: Mapped[Optional[str]] = mapped_column(String)
    status: Mapped[Optional[str]] = mapped_column(String(100))
    community: Mapped[Optional[str]] = mapped_column(String(100))
    caste: Mapped[Optional[str]] = mapped_column(String(100))
    residence_type: Mapped[Optional[str]] = mapped_column(String(100))
    civic_issue: Mapped[Optional[str]] = mapped_column(String)
    mother_tongue: Mapped[Optional[str]] = mapped_column(String(100))
    team: Mapped[Optional[str]] = mapped_column(String(100))
    ownership: Mapped[Optional[str]] = mapped_column(String(100))
    education: Mapped[Optional[str]] = mapped_column(String(100))
    nature_of_voter: Mapped[Optional[str]] = mapped_column(String(100))
    voter_points: Mapped[Optional[str]] = mapped_column(String(100))
    govt_scheme_tracking: Mapped[Optional[str]] = mapped_column(String)
    engagement_potential: Mapped[Optional[str]] = mapped_column(String(100))
    if_shifted: Mapped[Optional[str]] = mapped_column(String(100))
    notes: Mapped[Optional[str]] = mapped_column(String)
    present_address: Mapped[Optional[str]] = mapped_column(String)
    new_ward: Mapped[Optional[str]] = mapped_column(String(100))
    new_booth_no: Mapped[Optional[str]] = mapped_column(String(100))
    new_serial_no: Mapped[Optional[str]] = mapped_column(String(100))
    not_available_reason: Mapped[Optional[str]] = mapped_column(String)
    latitude: Mapped[Optional[float]] = mapped_column(Double)
    longitude: Mapped[Optional[float]] = mapped_column(Double)
    updated_fields: Mapped[str] = mapped_column(String, default="[]")
    updated_by: Mapped[Optional[int]] = mapped_column(ForeignKey("metastore.users.id"))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


engine = create_engine(DATABASE_URL, future=True, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# ---------------------------
# API response helpers
# ---------------------------
def api_success(message: str, result: Any) -> Dict[str, Any]:
    return {"success": True, "message": message, "data": {"result": result}}


def api_error(message: str, error: Any) -> Dict[str, Any]:
    return {"success": False, "message": message, "data": {"error": error}}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


class InvalidCredentialsException(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.code = "INVALID_CREDENTIALS"
        self.details = message


class ResourceAlreadyExistsException(Exception):
    def __init__(self, resource_name: str, field_name: str, field_value: Any):
        super().__init__(f"{resource_name} already exists with {field_name}: '{field_value}'")
        self.code = "RESOURCE_ALREADY_EXISTS"
        self.resource_name = resource_name
        self.field_name = field_name
        self.field_value = field_value
        self.details = f"{resource_name} with {field_name} '{field_value}' already exists."


class ResourceNotFoundException(Exception):
    def __init__(self, resource_name: str, field_name: str, field_value: Any):
        super().__init__(f"{resource_name} not found with {field_name}: '{field_value}'")
        self.code = "RESOURCE_NOT_FOUND"
        self.resource_name = resource_name
        self.field_name = field_name
        self.field_value = field_value
        self.details = f"{resource_name} with {field_name} '{field_value}' was not found."


# ---------------------------
# Auth
# ---------------------------
@dataclass
class JwtUserDetails:
    phone: str
    firstName: str
    role: str
    tenantId: Optional[str]
    assignmentType: Optional[str]
    assignmentId: Optional[str]


def _generate_token(first_name: str, role: str, tenant_id: Optional[str], assignment_type: Optional[str], assignment_id: Optional[str], phone: str) -> str:
    payload = {
        "sub": first_name,
        "role": role,
        "tenantId": tenant_id,
        "firstName": first_name,
        "assignmentType": assignment_type,
        "assignmentId": assignment_id,
        "phone": phone,
        "iat": int(datetime.now(timezone.utc).timestamp()),
        "exp": int((datetime.now(timezone.utc) + timedelta(seconds=JWT_EXPIRATION_SECONDS)).timestamp()),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")


def _parse_token(token: str) -> Dict[str, Any]:
    return jwt.decode(token, JWT_SECRET, algorithms=["HS256"])


def _external_base_url(request: Request) -> str:
    forwarded_proto = request.headers.get("x-forwarded-proto")
    forwarded_host = request.headers.get("x-forwarded-host")
    host = forwarded_host or request.headers.get("host")
    scheme = forwarded_proto or request.url.scheme
    if host:
        return f"{scheme}://{host}"
    return str(request.base_url).rstrip("/")


def _resolve_tenant_id(user: Optional[User] = None, payload: Optional[Dict[str, Any]] = None) -> Optional[str]:
    if user is not None and getattr(user, "tenant", None) is not None:
        tenant_value = getattr(user.tenant, "tenant_id", None)
        if tenant_value:
            return str(tenant_value)
    if payload is not None:
        tenant_value = payload.get("tenantId")
        if tenant_value:
            return str(tenant_value)
    return None


def _resolve_tenant_id_for_entity(entity: Any) -> Optional[str]:
    if entity is None:
        return None
    if isinstance(entity, VolunteerUser):
        return entity.tenant_id
    return _resolve_tenant_id(user=entity)


def _auth_user(request: Request, db: Session) -> JwtUserDetails:
    auth = request.headers.get("Authorization")
    if not auth or not auth.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid Authorization header")

    token = auth[7:]
    try:
        payload = _parse_token(token)
    except Exception:
        raise HTTPException(status_code=401, detail="Invalid token")

    base_query = (
        db.query(User)
        .join(Tenant, User.tenant_ref == Tenant.id, isouter=True)
        .filter(User.first_name == payload.get("firstName"), User.phone == payload.get("phone"))
    )
    user = base_query.first()
    volunteer = None
    if not user:
        volunteer = (
            db.query(VolunteerUser)
            .filter(VolunteerUser.first_name == payload.get("firstName"), VolunteerUser.phone == payload.get("phone"))
            .first()
        )

    target = user or volunteer
    if target and (target.blocked or target.deleted):
        reason = "blocked" if target.blocked else "deleted"
        raise HTTPException(status_code=403, detail=f"User is {reason}")

    return JwtUserDetails(
        phone=payload.get("phone"),
        firstName=payload.get("firstName"),
        role=payload.get("role"),
        tenantId=_resolve_tenant_id_for_entity(target) or payload.get("tenantId"),
        assignmentType=payload.get("assignmentType"),
        assignmentId=payload.get("assignmentId"),
    )


def get_current_user(request: Request, db: Session = Depends(get_db)) -> JwtUserDetails:
    return _auth_user(request, db)


def require_roles(*roles: str):
    def dep(user: JwtUserDetails = Depends(get_current_user)) -> JwtUserDetails:
        role = (user.role or "").replace("ROLE_", "")
        if role not in roles:
            if role in {"ASSEMBLY", "WARD", "BOOTH"} and "USER" in roles:
                return user
            raise HTTPException(status_code=401, detail="Access denied")
        return user

    return dep


# ---------------------------
# Schemas
# ---------------------------
class LoginRequest(BaseModel):
    firstName: str
    phone: str


class UserDetailsIn(BaseModel):
    role: str
    tenantId: Optional[str] = None
    assignmentType: Optional[str] = None
    assignmentId: Optional[str] = None
    firstName: str
    phone: Optional[str] = None
    profilePicUrl: Optional[str] = None
    blocked: Optional[bool] = None
    deleted: Optional[bool] = None


class UserBlockRequest(BaseModel):
    firstName: Optional[str] = None
    phone: Optional[str] = None
    userEmail: Optional[str] = None
    block: bool


class UserDeleteRequest(BaseModel):
    firstName: Optional[str] = None
    phone: Optional[str] = None
    userEmail: Optional[str] = None
    delete: bool


class UserBulkActionRequest(BaseModel):
    userFirstNames: Optional[List[str]] = None
    userEmails: Optional[List[str]] = None
    action: bool


class TenantDtoIn(BaseModel):
    id: Optional[int] = None
    tenantId: Optional[str] = None
    name: str
    description: Optional[str] = None
    contactEmail: str
    contactPhone: Optional[str] = None
    active: Optional[bool] = True


class VoterUpdatePayload(BaseModel):
    updateLocationLat: Optional[float] = None
    updateLocationLng: Optional[float] = None
    updateRequest: Dict[str, Any]


class PublicVoterUpdatePayload(BaseModel):
    wardCode: Optional[str] = None
    boothNo: Optional[str] = None
    updateRequest: Dict[str, Any]


class CreateAssociationRequest(BaseModel):
    associationName: str
    associationAddress: Optional[str] = None
    associationHeadName: Optional[str] = None
    phone: Optional[str] = None
    longitude: Optional[float] = None
    latitude: Optional[float] = None
    boothId: int


class CreateFamilyRequest(BaseModel):
    familyName: str
    familyAddress: Optional[str] = None
    phone: Optional[str] = None
    points: Optional[int] = None
    pointsProvided: Optional[int] = None
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    boothId: int
    associationId: Optional[int] = None
    headEpicNo: str
    memberEpicNos: List[str]
    economicStatus: Optional[str] = None
    familyNature: Optional[str] = None


class UpdateFamilyRequest(CreateFamilyRequest):
    pass


class UserProfileDto(BaseModel):
    firstName: str
    phone: str
    profilePicUrl: Optional[str] = None
    tenantId: Optional[str] = None
    role: Optional[str] = None


class VolunteerCreateRequest(BaseModel):
    firstName: str
    phone: str
    workingLevel: str
    assemblyIds: Optional[List[int]] = None
    wardIds: Optional[List[int]] = None
    boothIds: Optional[List[int]] = None


# ---------------------------
# Utility converters
# ---------------------------
def to_user_details(u: User) -> Dict[str, Any]:
    def parse_ids(value: Optional[str]) -> List[int]:
        if not value:
            return []
        return [int(v) for v in str(value).split(",") if str(v).strip().isdigit()]

    return {
        "role": u.role,
        "tenantId": _resolve_tenant_id_for_entity(u),
        "assignmentType": u.assignment_type,
        "assignmentId": u.assignment_id,
        "workingLevel": getattr(u, "working_level", None) or u.assignment_type,
        "assemblyIds": parse_ids(getattr(u, "assembly_ids", None)),
        "wardIds": parse_ids(getattr(u, "ward_ids", None)),
        "boothIds": parse_ids(getattr(u, "booth_ids", None)),
        "firstName": u.first_name,
        "lastName": "",
        "userName": u.first_name,
        "phone": u.phone,
        "profilePicUrl": u.profile_pic_url,
        "blocked": u.blocked,
        "deleted": u.deleted,
    }


def to_tenant_dto(t: Tenant) -> Dict[str, Any]:
    return {
        "id": t.id,
        "tenantId": t.tenant_id,
        "name": t.name,
        "description": t.description,
        "contactEmail": t.contact_email,
        "contactPhone": t.contact_phone,
        "active": t.active,
    }


def _serialize_id_list(values: Optional[List[int]]) -> Optional[str]:
    if not values:
        return None
    return ",".join(str(int(v)) for v in values if v is not None)


def _first_id(values: Optional[List[int]]) -> Optional[int]:
    if not values:
        return None
    for v in values:
        if v is not None:
            return int(v)
    return None


def build_page(content: List[Any], page: int, size: int, total: int, sort_field: str = "", direction: str = "asc") -> Dict[str, Any]:
    total_pages = (total + size - 1) // size if size else 1
    return {
        "content": content,
        "pageable": {
            "pageNumber": page,
            "pageSize": size,
            "sort": {"sorted": bool(sort_field), "unsorted": not bool(sort_field), "empty": not bool(sort_field)},
            "offset": page * size,
            "paged": True,
            "unpaged": False,
        },
        "last": page >= max(total_pages - 1, 0),
        "totalPages": total_pages,
        "totalElements": total,
        "size": size,
        "number": page,
        "sort": {"sorted": bool(sort_field), "unsorted": not bool(sort_field), "empty": not bool(sort_field)},
        "first": page == 0,
        "numberOfElements": len(content),
        "empty": len(content) == 0,
    }


def _parse_id_list(raw: Optional[str]) -> List[int]:
    if not raw:
        return []
    items = []
    for part in str(raw).split(","):
        part = part.strip()
        if part.isdigit():
            items.append(int(part))
    return items


def _get_access_scope(db: Session, current: JwtUserDetails) -> Optional[Dict[str, Any]]:
    role = (current.role or "").replace("ROLE_", "")
    if role in {"SUPER_ADMIN", "ADMIN"}:
        return None

    volunteer = (
        db.query(VolunteerUser)
        .filter(VolunteerUser.first_name == current.firstName, VolunteerUser.phone == current.phone)
        .first()
    )
    assignment_type = normalize_optional_text(current.assignmentType) or (
        normalize_optional_text(volunteer.assignment_type) if volunteer else None
    ) or (normalize_optional_text(volunteer.working_level) if volunteer else None) or role
    assignment_type = (assignment_type or role).upper()

    assembly_ids = _parse_id_list(volunteer.assembly_ids) if volunteer else []
    ward_ids = _parse_id_list(volunteer.ward_ids) if volunteer else []
    booth_ids = _parse_id_list(volunteer.booth_ids) if volunteer else []

    if not (assembly_ids or ward_ids or booth_ids):
        fallback_ids = _parse_id_list(str(current.assignmentId or ""))
        if assignment_type == "ASSEMBLY":
            assembly_ids = fallback_ids
        elif assignment_type == "WARD":
            ward_ids = fallback_ids
        elif assignment_type == "BOOTH":
            booth_ids = fallback_ids

    return {
        "assignment_type": assignment_type,
        "assembly_ids": assembly_ids,
        "ward_ids": ward_ids,
        "booth_ids": booth_ids,
        "role": role,
    }


def _resolve_access_scope_ids(db: Session, current: JwtUserDetails) -> Optional[Dict[str, Any]]:
    scope = _get_access_scope(db, current)
    if not scope:
        return None

    allowed_assembly_ids = set(scope.get("assembly_ids") or [])
    allowed_ward_ids = set(scope.get("ward_ids") or [])
    allowed_booth_ids = set(scope.get("booth_ids") or [])

    # booth_ids may be stored as booth_no values, so map both booth_id and booth_no to booth_id
    if allowed_booth_ids:
        booth_cols = _get_table_columns(db, "public", "booths")
        booth_pk_col = "id" if "id" in booth_cols else ("booth_id" if "booth_id" in booth_cols else None)
        booth_no_col = "booth_no" if "booth_no" in booth_cols else None
        if booth_pk_col:
            raw_values = [str(v).strip() for v in allowed_booth_ids if v is not None and str(v).strip() != ""]
            ids_int = sorted({int(v) for v in raw_values if v.isdigit()})
            ids_text = sorted({str(v) for v in raw_values if str(v).strip()})
            where_parts = []
            params: Dict[str, Any] = {}
            if ids_int:
                clause_id, params_id = _build_in_clause(booth_pk_col, ids_int, "scope_booth_id")
                where_parts.append(f"({clause_id})")
                params.update(params_id)
            if booth_no_col and ids_text:
                clause_no, params_no = _build_in_clause(f"CAST({booth_no_col} AS TEXT)", ids_text, "scope_booth_no")
                where_parts.append(f"({clause_no})")
                params.update(params_no)
            if not where_parts:
                where_parts.append("1=0")
            where_clause = f"WHERE {' OR '.join(where_parts)}"
            rows = db.execute(
                text(
                    f"""
                    SELECT {booth_pk_col} AS booth_id, {booth_no_col if booth_no_col else booth_pk_col} AS booth_no
                    FROM public.booths
                    {where_clause}
                    """
                ),
                params,
            ).all()
            allowed_booth_ids = {row.booth_id for row in rows if row.booth_id is not None}

    if allowed_assembly_ids:
        ward_rows = db.query(Ward.ward_id).filter(Ward.assembly_id.in_(allowed_assembly_ids)).all()
        allowed_ward_ids.update([row[0] for row in ward_rows])

    if allowed_booth_ids:
        ward_rows = db.query(Booth.ward_id).filter(Booth.booth_id.in_(allowed_booth_ids)).all()
        allowed_ward_ids.update([row[0] for row in ward_rows if row[0] is not None])

    if allowed_ward_ids:
        booth_rows = db.query(Booth.booth_id).filter(Booth.ward_id.in_(allowed_ward_ids)).all()
        allowed_booth_ids.update([row[0] for row in booth_rows])
        assembly_rows = db.query(Ward.assembly_id).filter(Ward.ward_id.in_(allowed_ward_ids)).all()
        allowed_assembly_ids.update([row[0] for row in assembly_rows if row[0] is not None])

    allowed_assembly_ids = {v for v in allowed_assembly_ids if v is not None}
    allowed_ward_ids = {v for v in allowed_ward_ids if v is not None}
    allowed_booth_ids = {v for v in allowed_booth_ids if v is not None}

    return {
        **scope,
        "allowed_assembly_ids": allowed_assembly_ids,
        "allowed_ward_ids": allowed_ward_ids,
        "allowed_booth_ids": allowed_booth_ids,
    }


def _build_in_clause(column_expr: str, values: List[Any], prefix: str) -> tuple[str, Dict[str, Any]]:
    if not values:
        return "1=0", {}
    params: Dict[str, Any] = {}
    placeholders: List[str] = []
    for idx, value in enumerate(values):
        key = f"{prefix}_{idx}"
        params[key] = value
        placeholders.append(f":{key}")
    return f"{column_expr} IN ({', '.join(placeholders)})", params


def _get_s3_client():
    return boto3.client(
        "s3",
        region_name=AWS_REGION,
        aws_access_key_id=AWS_ACCESS_KEY or None,
        aws_secret_access_key=AWS_SECRET_KEY or None,
    )


def s3_extract_key(url: str) -> str:
    idx = url.find(".com/")
    if idx == -1:
        raise ValueError(f"Invalid S3 URL: {url}")
    return url[idx + 5 :]


def s3_presigned_url(key: str, minutes: int, fallback_url: Optional[str] = None) -> str:
    client = _get_s3_client()
    try:
        return client.generate_presigned_url(
            "get_object",
            Params={"Bucket": AWS_BUCKET, "Key": key},
            ExpiresIn=minutes * 60,
        )
    except NoCredentialsError:
        if fallback_url:
            return fallback_url
        if AWS_BUCKET:
            return f"https://{AWS_BUCKET}.s3.amazonaws.com/{key}"
        raise


def s3_upload_bytes(content: bytes, content_type: str, key: str) -> str:
    client = _get_s3_client()
    client.put_object(Bucket=AWS_BUCKET, Key=key, Body=content, ContentType=content_type)
    return f"https://{AWS_BUCKET}.s3.amazonaws.com/{key}"


def normalize_assembly_code(value: Any) -> str:
    s = str(value)
    if len(s) == 12:
        return s
    return f"{int(s):012d}"


def normalize_optional_text(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    trimmed = value.strip()
    return trimmed if trimmed else None


def parse_optional_bool(value: Optional[str | bool]) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    cleaned = normalize_optional_text(value)
    if cleaned is None:
        return None
    lowered = cleaned.lower()
    if lowered in {"true", "1", "yes"}:
        return True
    if lowered in {"false", "0", "no"}:
        return False
    raise ValueError(f"Invalid boolean value: {value}")


def resolve_user_first_name(first_name: Optional[str], fallback_username: Optional[str]) -> Optional[str]:
    return normalize_optional_text(first_name) or normalize_optional_text(fallback_username)


def resolve_bulk_usernames(payload: UserBulkActionRequest) -> List[str]:
    candidates = payload.userFirstNames if payload.userFirstNames is not None else payload.userEmails
    if not candidates:
        raise ValueError("At least one user identifier is required")
    usernames = [name.strip() for name in candidates if name and name.strip()]
    if not usernames:
        raise ValueError("At least one valid user identifier is required")
    return usernames


# ---------------------------
# App + exception handlers
# ---------------------------
app = FastAPI(title="VoteBase Python API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def log_all_requests(request: Request, call_next):
    started_at = datetime.now(timezone.utc)
    query = request.url.query
    path_with_query = f"{request.url.path}?{query}" if query else request.url.path
    auth_present = bool(request.headers.get("Authorization"))
    body_log = None

    target_paths = {
        f"{CONTEXT_PATH}/api/auth/login",
        f"{CONTEXT_PATH}/api/booth",
        f"{CONTEXT_PATH}/api/voters/snapshot",
    }

    def mask_sensitive(value: Any) -> Any:
        if isinstance(value, dict):
            out: Dict[str, Any] = {}
            for k, v in value.items():
                lk = k.lower()
                if lk in {"phone", "token", "password", "authorization"}:
                    out[k] = "***"
                else:
                    out[k] = mask_sensitive(v)
            return out
        if isinstance(value, list):
            return [mask_sensitive(v) for v in value]
        return value

    if request.url.path in target_paths:
        content_type = (request.headers.get("content-type") or "").lower()
        if "application/json" in content_type:
            try:
                raw_body = await request.body()
                if raw_body:
                    parsed = json.loads(raw_body.decode("utf-8"))
                    body_log = json.dumps(mask_sensitive(parsed), ensure_ascii=False)

                    # Re-inject the already-read body so downstream handlers can read it.
                    async def receive():
                        return {"type": "http.request", "body": raw_body, "more_body": False}

                    request = Request(request.scope, receive)
            except Exception:
                body_log = "<unparseable-json>"

    try:
        response = await call_next(request)
        duration_ms = int((datetime.now(timezone.utc) - started_at).total_seconds() * 1000)
        line = (
            f"[API_HIT] method={request.method} path={path_with_query} "
            f"status={response.status_code} auth={'yes' if auth_present else 'no'} "
            f"durationMs={duration_ms}"
        )
        if body_log is not None:
            line += f" body={body_log}"
        print(line)
        return response
    except Exception as ex:
        duration_ms = int((datetime.now(timezone.utc) - started_at).total_seconds() * 1000)
        line = (
            f"[API_HIT] method={request.method} path={path_with_query} "
            f"status=500 auth={'yes' if auth_present else 'no'} durationMs={duration_ms} "
            f"error={type(ex).__name__}"
        )
        if body_log is not None:
            line += f" body={body_log}"
        print(line)
        raise


@app.exception_handler(InvalidCredentialsException)
async def handle_invalid_credentials(_: Request, ex: InvalidCredentialsException):
    return JSONResponse(
        status_code=401,
        content=api_error(ex.args[0], {"timestamp": now_iso(), "code": ex.code, "details": ex.details}),
    )


@app.exception_handler(ResourceAlreadyExistsException)
async def handle_resource_exists(_: Request, ex: ResourceAlreadyExistsException):
    return JSONResponse(
        status_code=409,
        content=api_error(
            ex.args[0],
            {
                "timestamp": now_iso(),
                "code": ex.code,
                "resource": ex.resource_name,
                "field": ex.field_name,
                "value": ex.field_value,
                "details": ex.details,
            },
        ),
    )


@app.exception_handler(ResourceNotFoundException)
async def handle_resource_not_found(_: Request, ex: ResourceNotFoundException):
    return JSONResponse(
        status_code=404,
        content=api_error(
            "Resource not found",
            {
                "timestamp": now_iso(),
                "code": ex.code,
                "resource": ex.resource_name,
                "field": ex.field_name,
                "value": ex.field_value,
                "details": ex.details,
            },
        ),
    )


@app.exception_handler(ValueError)
async def handle_validation(_: Request, ex: ValueError):
    return JSONResponse(
        status_code=400,
        content=api_error("Validation failed", {"code": "INVALID_PARAMETER", "details": str(ex), "timestamp": now_iso()}),
    )


@app.exception_handler(Exception)
async def handle_generic(_: Request, ex: Exception):
    if isinstance(ex, HTTPException):
        return JSONResponse(status_code=ex.status_code, content=api_error(ex.detail, {"timestamp": now_iso(), "code": "HTTP_ERROR", "details": ex.detail}))
    return JSONResponse(
        status_code=500,
        content=api_error("An unexpected error occurred", {"timestamp": now_iso(), "code": "INTERNAL_SERVER_ERROR", "details": str(ex)}),
    )


# ---------------------------
# Startup: super admin seed
# ---------------------------
@app.on_event("startup")
def startup_seed_super_admin() -> None:
    print(f"[DB_TARGET] { _safe_db_target(DATABASE_URL) }")
    db = SessionLocal()
    try:
        existing = db.query(User).filter(User.role == "SUPER_ADMIN").first()
        if not existing:
            super_admin = User(
                first_name=SUPERADMIN_USERNAME,
                phone="8867038709",
                role="SUPER_ADMIN",
                tenant_ref=None,
                assignment_type=None,
                assignment_id=None,
                blocked=False,
                deleted=False,
            )
            db.add(super_admin)
            db.commit()
    finally:
        db.close()


@app.on_event("startup")
def startup_ensure_voter_enrichment() -> None:
    VoterEnrichment.__table__.create(bind=engine, checkfirst=True)
    VolunteerUser.__table__.create(bind=engine, checkfirst=True)
    db = SessionLocal()
    try:
        existing_cols = {
            row[0]
            for row in db.execute(
                text(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_schema = 'metastore'
                      AND table_name = 'volunteer_users'
                    """
                )
            ).all()
        }
        if "working_level" not in existing_cols:
            db.execute(text("ALTER TABLE metastore.volunteer_users ADD COLUMN working_level varchar(30)"))
        if "assembly_ids" not in existing_cols:
            db.execute(text("ALTER TABLE metastore.volunteer_users ADD COLUMN assembly_ids text"))
        if "ward_ids" not in existing_cols:
            db.execute(text("ALTER TABLE metastore.volunteer_users ADD COLUMN ward_ids text"))
        if "booth_ids" not in existing_cols:
            db.execute(text("ALTER TABLE metastore.volunteer_users ADD COLUMN booth_ids text"))

        col = db.execute(
            text(
                """
                SELECT data_type
                FROM information_schema.columns
                WHERE table_schema = 'metastore'
                  AND table_name = 'volunteer_users'
                  AND column_name = 'assignment_id'
                """
            )
        ).scalar()
        if col and col not in {"character varying", "text"}:
            db.execute(text("ALTER TABLE metastore.volunteer_users ALTER COLUMN assignment_id TYPE text USING assignment_id::text"))
        db.commit()
    except Exception:
        db.rollback()
    finally:
        db.close()


# ---------------------------
# Routes
# ---------------------------
@app.post(f"{CONTEXT_PATH}/api/auth/login")
def login(payload: LoginRequest, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.first_name == payload.firstName, User.phone == payload.phone).first()
    volunteer = None
    if not user:
        volunteer = db.query(VolunteerUser).filter(VolunteerUser.first_name == payload.firstName, VolunteerUser.phone == payload.phone).first()
        if not volunteer:
            raise InvalidCredentialsException("Invalid firstname or phone")

    tenant_id = None
    assignment_type = None
    assignment_id = 0
    if user and user.role != "SUPER_ADMIN":
        if not user.tenant:
            raise InvalidCredentialsException("Tenant information missing for user")
        tenant_id = user.tenant.tenant_id
        if user.role != "ADMIN":
            if user.assignment_type is None or user.assignment_id == -1:
                raise InvalidCredentialsException("Assignment information missing for user")
        assignment_type = user.assignment_type
        assignment_id = user.assignment_id
    elif volunteer:
        if volunteer.blocked or volunteer.deleted:
            reason = "blocked" if volunteer.blocked else "deleted"
            raise HTTPException(status_code=403, detail=f"User is {reason}")
        tenant_id = volunteer.tenant_id
        assignment_type = volunteer.assignment_type
        assignment_id = volunteer.assignment_id

    effective = user or volunteer
    token = _generate_token(effective.first_name, effective.role, tenant_id, assignment_type, assignment_id, effective.phone)
    return api_success(
        "Login successful",
        {
            "userName": payload.firstName,
            "token": token,
            "role": effective.role,
            "tenantId": tenant_id,
        },
    )


@app.get(f"{CONTEXT_PATH}/api/admin/dashboard")
def get_dashboard(user: JwtUserDetails = Depends(require_roles("ADMIN"))):
    return api_success(f"Welcome, ['ROLE_{user.role}']", {"user": user.firstName})


@app.post(f"{CONTEXT_PATH}/api/user/register")
def register_user(
    payload: UserDetailsIn,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN")),
):
    if current.role not in {"ADMIN", "SUPER_ADMIN"}:
        raise HTTPException(status_code=403, detail="You are not authorized to create users")

    requested_role = (payload.role or "").strip().upper()
    role_to_create = "USER"
    if requested_role not in {"", "USER"}:
        raise ValueError("Invalid role in register request")

    tenant_id = current.tenantId if current.role == "ADMIN" else (payload.tenantId or "").strip()
    if not tenant_id and current.role == "SUPER_ADMIN":
        tenants = db.query(Tenant).all()
        if len(tenants) == 1:
            tenant_id = tenants[0].tenant_id
        else:
            raise HTTPException(status_code=400, detail="tenantId is required for SUPER_ADMIN")

    tenant = db.query(Tenant).filter(Tenant.tenant_id == tenant_id).first()
    if not tenant:
        raise ResourceNotFoundException("Tenant", "tenantId", tenant_id)

    exists = db.query(VolunteerUser).filter(VolunteerUser.first_name == payload.firstName, VolunteerUser.phone == payload.phone).first()
    if exists:
        raise ResourceAlreadyExistsException("registerUser", "userName", payload.firstName)

    volunteer = VolunteerUser(
        role=role_to_create,
        tenant_id=tenant.tenant_id,
        working_level=payload.assignmentType,
        assignment_type=payload.assignmentType,
        assignment_id=payload.assignmentId,
        first_name=payload.firstName,
        phone=payload.phone,
        blocked=False,
        deleted=False,
    )
    db.add(volunteer)
    db.commit()

    out = payload.model_dump()
    out["role"] = role_to_create
    out["tenantId"] = tenant_id
    out["userName"] = payload.firstName
    return api_success("User registered successfully", out)


@app.post(f"{CONTEXT_PATH}/api/volunteers")
def create_volunteer(
    payload: VolunteerCreateRequest,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD")),
):
    if current.role not in {"ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD"}:
        raise HTTPException(status_code=403, detail="You are not authorized to create volunteers")

    working_level = (payload.workingLevel or "").strip().upper()
    if working_level not in {"ASSEMBLY", "WARD", "BOOTH"}:
        raise ValueError("workingLevel must be ASSEMBLY, WARD, or BOOTH")

    phone = normalize_optional_text(payload.phone)
    if not phone or not phone.isdigit() or len(phone) != 10:
        raise ValueError("phone must be a 10 digit number")

    tenant_id = current.tenantId
    if not tenant_id:
        tenants = db.query(Tenant).all()
        if len(tenants) == 1:
            tenant_id = tenants[0].tenant_id
        elif len(tenants) == 0:
            tenant_id = ""
        else:
            raise HTTPException(status_code=400, detail="tenantId is required for SUPER_ADMIN")

    exists = db.query(VolunteerUser).filter(VolunteerUser.first_name == payload.firstName, VolunteerUser.phone == phone).first()
    if exists:
        raise ResourceAlreadyExistsException("volunteer", "userName", payload.firstName)

    assembly_ids = payload.assemblyIds or []
    ward_ids = payload.wardIds or []
    booth_ids = payload.boothIds or []
    assignment_id = _first_id(assembly_ids) or _first_id(ward_ids) or _first_id(booth_ids)

    volunteer = VolunteerUser(
        role=working_level,
        tenant_id=tenant_id,
        working_level=working_level,
        assignment_type=working_level,
        assignment_id=str(assignment_id) if assignment_id is not None else None,
        assembly_ids=_serialize_id_list(assembly_ids),
        ward_ids=_serialize_id_list(ward_ids),
        booth_ids=_serialize_id_list(booth_ids),
        first_name=payload.firstName,
        phone=phone,
        blocked=False,
        deleted=False,
    )
    db.add(volunteer)
    db.commit()

    return api_success(
        "Volunteer created successfully",
        to_user_details(volunteer),
    )


@app.get(f"{CONTEXT_PATH}/api/volunteers")
def list_volunteers(
    page: int = 0,
    size: int = 10,
    search: Optional[str] = None,
    blocked: Optional[str] = None,
    deleted: Optional[str] = None,
    workingLevel: Optional[str] = None,
    sortBy: str = "firstName",
    direction: str = "asc",
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD")),
):
    q = db.query(VolunteerUser).filter(VolunteerUser.role.in_(["USER", "ASSEMBLY", "WARD", "BOOTH"]))
    if current.role != "SUPER_ADMIN" and current.tenantId:
        q = q.filter(VolunteerUser.tenant_id == current.tenantId)

    blocked_filter = parse_optional_bool(blocked)
    deleted_filter = parse_optional_bool(deleted)
    level_filter = normalize_optional_text(workingLevel)

    if blocked_filter is not None:
        q = q.filter(VolunteerUser.blocked == blocked_filter)
    if deleted_filter is not None:
        q = q.filter(VolunteerUser.deleted == deleted_filter)
    if level_filter is not None:
        q = q.filter(VolunteerUser.working_level == level_filter)
    if search and search.strip():
        s = f"%{search.lower()}%"
        q = q.filter(or_(func.lower(VolunteerUser.first_name).like(s), VolunteerUser.phone.like(f"%{search}%")))

    sort_map = {
        "firstName": VolunteerUser.first_name,
        "phone": VolunteerUser.phone,
        "role": VolunteerUser.role,
        "workingLevel": VolunteerUser.working_level,
        "assignmentType": VolunteerUser.assignment_type,
        "assignmentId": VolunteerUser.assignment_id,
        "blocked": VolunteerUser.blocked,
        "deleted": VolunteerUser.deleted,
    }
    sort_col = sort_map.get(sortBy, VolunteerUser.first_name)
    q = q.order_by(desc(sort_col) if direction.lower() == "desc" else asc(sort_col))

    total = q.count()
    users = q.offset(page * size).limit(size).all()
    content = [to_user_details(u) for u in users]
    return build_page(content, page, size, total, sortBy, direction)


@app.put(f"{CONTEXT_PATH}/api/user/block")
def block_user(payload: UserBlockRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD"))):
    target_first_name = resolve_user_first_name(payload.firstName, payload.userEmail)
    if not target_first_name:
        raise ValueError("firstName or userEmail is required")

    q = db.query(VolunteerUser).filter(VolunteerUser.first_name == target_first_name)
    if current.tenantId:
        q = q.filter(VolunteerUser.tenant_id == current.tenantId)
    user = q.first()
    if not user:
        return api_error("User not found", {"details": f"User not found with FirstName: '{target_first_name}'"})

    user.blocked = payload.block
    db.commit()
    action = "blocked" if payload.block else "unblocked"
    return api_success(f"User {action} successfully", {"firstName": target_first_name})


@app.get(f"{CONTEXT_PATH}/api/user")
def list_users(
    role: Optional[str] = None,
    page: int = 0,
    size: int = 10,
    search: Optional[str] = None,
    blocked: Optional[str] = None,
    deleted: Optional[str] = None,
    assignmentType: Optional[str] = None,
    sortBy: str = "firstName",
    direction: str = "asc",
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN")),
):
    q = db.query(VolunteerUser).filter(VolunteerUser.role == "USER", VolunteerUser.tenant_id == current.tenantId)

    blocked_filter = parse_optional_bool(blocked)
    deleted_filter = parse_optional_bool(deleted)
    assignment_type_filter = normalize_optional_text(assignmentType)

    if blocked_filter is not None:
        q = q.filter(VolunteerUser.blocked == blocked_filter)
    if deleted_filter is not None:
        q = q.filter(VolunteerUser.deleted == deleted_filter)
    if assignment_type_filter is not None:
        q = q.filter(VolunteerUser.assignment_type == assignment_type_filter)
    if search and search.strip():
        s = f"%{search.lower()}%"
        q = q.filter(or_(func.lower(VolunteerUser.first_name).like(s), VolunteerUser.phone.like(f"%{search}%")))

    sort_map = {
        "firstName": VolunteerUser.first_name,
        "phone": VolunteerUser.phone,
        "role": VolunteerUser.role,
        "assignmentType": VolunteerUser.assignment_type,
        "assignmentId": VolunteerUser.assignment_id,
        "blocked": VolunteerUser.blocked,
        "deleted": VolunteerUser.deleted,
    }
    sort_col = sort_map.get(sortBy, VolunteerUser.first_name)
    q = q.order_by(desc(sort_col) if direction.lower() == "desc" else asc(sort_col))

    total = q.count()
    users = q.offset(page * size).limit(size).all()
    content = [to_user_details(u) for u in users]
    return build_page(content, page, size, total, sortBy, direction)


@app.put(f"{CONTEXT_PATH}/api/user/delete")
def delete_user(payload: UserDeleteRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD"))):
    target_first_name = resolve_user_first_name(payload.firstName, payload.userEmail)
    if not target_first_name:
        raise ValueError("firstName or userEmail is required")

    q = db.query(VolunteerUser).filter(VolunteerUser.first_name == target_first_name)
    if current.tenantId:
        q = q.filter(VolunteerUser.tenant_id == current.tenantId)
    user = q.first()
    if not user:
        return api_error("User not found", {"details": f"User not found with FirstName: '{target_first_name}'"})

    user.deleted = payload.delete
    db.commit()
    action = "deleted" if payload.delete else "restored"
    return api_success(f"User {action} successfully", {"email": payload.model_dump()})


@app.put(f"{CONTEXT_PATH}/api/user/block/bulk")
def bulk_block(payload: UserBulkActionRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD"))):
    usernames = resolve_bulk_usernames(payload)
    q = db.query(VolunteerUser).filter(VolunteerUser.first_name.in_(usernames))
    if current.tenantId:
        q = q.filter(VolunteerUser.tenant_id == current.tenantId)
    users = q.all()
    if len(users) != len(usernames):
        return api_error("Some users not found", {"details": "Some users not found"})

    for user in users:
        user.blocked = payload.action
    db.commit()
    action = "blocked" if payload.action else "unblocked"
    return api_success(f"Users {action} successfully", {"emails": usernames})


@app.put(f"{CONTEXT_PATH}/api/user/delete/bulk")
def bulk_delete(payload: UserBulkActionRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "SUPER_ADMIN", "ASSEMBLY", "WARD"))):
    usernames = resolve_bulk_usernames(payload)
    q = db.query(VolunteerUser).filter(VolunteerUser.first_name.in_(usernames))
    if current.tenantId:
        q = q.filter(VolunteerUser.tenant_id == current.tenantId)
    users = q.all()
    if len(users) != len(usernames):
        return api_error("Some users not found", {"details": "Some users not found"})

    for user in users:
        user.deleted = payload.action
    db.commit()
    action = "deleted" if payload.action else "restored"
    return api_success(f"Users {action} successfully", {"emails": usernames})


@app.get(f"{CONTEXT_PATH}/api/user/profile")
def get_profile(db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    user = db.query(User).filter(User.first_name == current.firstName, User.phone == current.phone).first()
    volunteer = None
    if not user:
        volunteer = db.query(VolunteerUser).filter(VolunteerUser.first_name == current.firstName, VolunteerUser.phone == current.phone).first()
        if not volunteer:
            raise ResourceNotFoundException("User", "username", current.firstName)

    target = user or volunteer
    presigned = ""
    if target.profile_pic_url:
        key = s3_extract_key(target.profile_pic_url)
        presigned = s3_presigned_url(key, 15, fallback_url=target.profile_pic_url)

    return {
        "firstName": target.first_name,
        "lastName": "",
        "userName": target.first_name,
        "phone": target.phone,
        "profilePicUrl": presigned,
        "tenantId": _resolve_tenant_id_for_entity(target),
        "role": target.role,
    }


@app.put(f"{CONTEXT_PATH}/api/user/profile")
def update_profile(payload: UserProfileDto, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    user = db.query(User).filter(User.first_name == current.firstName, User.phone == current.phone).first()
    volunteer = None
    if not user:
        volunteer = db.query(VolunteerUser).filter(VolunteerUser.first_name == current.firstName, VolunteerUser.phone == current.phone).first()
        if not volunteer:
            raise ResourceNotFoundException("User", "username", current.firstName)

    target = user or volunteer
    target.first_name = payload.firstName
    target.phone = payload.phone
    db.commit()
    db.refresh(target)
    return get_profile(db, JwtUserDetails(phone=target.phone, firstName=target.first_name, role=target.role, tenantId=current.tenantId, assignmentType=current.assignmentType, assignmentId=current.assignmentId))


@app.post(f"{CONTEXT_PATH}/api/user/profile/upload")
def upload_profile(file: UploadFile = File(...), db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    user = db.query(User).filter(User.first_name == current.firstName, User.phone == current.phone).first()
    volunteer = None
    if not user:
        volunteer = db.query(VolunteerUser).filter(VolunteerUser.first_name == current.firstName, VolunteerUser.phone == current.phone).first()
        if not volunteer:
            raise ResourceNotFoundException("User", "username", current.firstName)

    ext = Path(file.filename or "").suffix
    target = user or volunteer
    tenant_segment = _resolve_tenant_id_for_entity(target) or "global"
    key = f"{PROFILE_UPLOAD_DIR}/{tenant_segment}/{target.first_name}/{uuid.uuid4()}{ext}"
    raw = file.file.read()
    s3_url = s3_upload_bytes(raw, file.content_type or "application/octet-stream", key)
    target.profile_pic_url = s3_url
    db.commit()

    presigned = s3_presigned_url(key, 15, fallback_url=s3_url)
    return {
        "firstName": target.first_name,
        "lastName": "",
        "userName": target.first_name,
        "phone": target.phone,
        "profilePicUrl": presigned,
        "tenantId": _resolve_tenant_id_for_entity(target),
        "role": target.role,
    }


@app.post(f"{CONTEXT_PATH}/api/tenant")
def create_tenant(payload: TenantDtoIn, db: Session = Depends(get_db), _: JwtUserDetails = Depends(require_roles("SUPER_ADMIN"))):
    existing = db.query(Tenant).filter(Tenant.contact_email == payload.contactEmail).first()
    if existing:
        raise ResourceAlreadyExistsException("Tenant", "contactEmail", payload.contactEmail)

    tenant = Tenant(
        name=payload.name,
        description=payload.description,
        contact_email=payload.contactEmail,
        contact_phone=payload.contactPhone,
        active=True,
        created_at=datetime.utcnow(),
        updated_at=datetime.utcnow(),
        tenant_id=f"iswot-{int(datetime.utcnow().timestamp() * 1000) % 100_000_000:08d}",
    )
    db.add(tenant)
    db.flush()

    admin = User(
        role="ADMIN",
        tenant_ref=tenant.id,
        first_name="Tenant",
        phone=payload.contactPhone or "0000000000",
        blocked=False,
        deleted=False,
    )
    db.add(admin)
    db.commit()
    db.refresh(tenant)
    return to_tenant_dto(tenant)


@app.get(f"{CONTEXT_PATH}/api/tenant/{{tenantId}}")
def get_tenant(tenantId: str, db: Session = Depends(get_db), _: JwtUserDetails = Depends(require_roles("SUPER_ADMIN"))):
    tenant = db.query(Tenant).filter(Tenant.tenant_id == tenantId).first()
    if not tenant:
        raise ValueError(f"Tenant not found: {tenantId}")
    return to_tenant_dto(tenant)


@app.put(f"{CONTEXT_PATH}/api/tenant/{{tenantId}}")
def update_tenant(tenantId: str, payload: TenantDtoIn, db: Session = Depends(get_db), _: JwtUserDetails = Depends(require_roles("SUPER_ADMIN"))):
    tenant = db.query(Tenant).filter(Tenant.tenant_id == tenantId).first()
    if not tenant:
        raise ValueError(f"Tenant not found: {tenantId}")

    tenant.name = payload.name
    tenant.description = payload.description
    tenant.contact_phone = payload.contactPhone
    tenant.active = payload.active if payload.active is not None else tenant.active
    tenant.updated_at = datetime.utcnow()
    db.commit()
    return to_tenant_dto(tenant)


@app.get(f"{CONTEXT_PATH}/api/tenant")
def list_tenants(page: int = 0, size: int = 10, db: Session = Depends(get_db), _: JwtUserDetails = Depends(require_roles("SUPER_ADMIN"))):
    q = db.query(Tenant).order_by(Tenant.id)
    total = q.count()
    tenants = q.offset(page * size).limit(size).all()
    return build_page([to_tenant_dto(t) for t in tenants], page, size, total)


@app.get(f"{CONTEXT_PATH}/api/assignments")
def get_assignments(type: str = Query(...), db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "ADMIN", "USER"))):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        return []
    t = type.upper()
    if t == "ASSEMBLY":
        q = db.query(Assembly)
        if current.tenantId is not None:
            q = q.filter(Assembly.tenant_id == current.tenantId)
        if scope and scope.get("allowed_assembly_ids"):
            q = q.filter(Assembly.assembly_id.in_(scope.get("allowed_assembly_ids")))
        rows = q.order_by(Assembly.assembly_name_en.asc()).all()
        return [{"id": r.assembly_id, "name": r.assembly_name_en} for r in rows]
    if t == "WARD":
        q = db.query(Ward)
        if current.tenantId is not None:
            q = q.filter(Ward.tenant_id == current.tenantId)
        if scope and scope.get("allowed_ward_ids"):
            q = q.filter(Ward.ward_id.in_(scope.get("allowed_ward_ids")))
        rows = q.order_by(Ward.ward_name_en.asc()).all()
        return [{"id": r.ward_id, "name": r.ward_name_en} for r in rows]
    if t == "BOOTH":
        q = db.query(Booth)
        if current.tenantId is not None:
            q = q.filter(Booth.tenant_id == current.tenantId)
        if scope and scope.get("allowed_booth_ids"):
            q = q.filter(Booth.booth_id.in_(scope.get("allowed_booth_ids")))
        rows = q.order_by(Booth.polling_station_adr_en.asc()).all()
        return [{"id": r.booth_id, "name": r.polling_station_adr_en} for r in rows]
    raise ValueError("Invalid assignment type")


@app.get(f"{CONTEXT_PATH}/api/booth")
def get_booths(db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "ADMIN", "USER"))):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        return api_success("Booths fetched successfully", [])
    booth_cols = _get_table_columns(db, "public", "booths")
    booth_pk_col = "id" if "id" in booth_cols else ("booth_id" if "booth_id" in booth_cols else "id")
    booth_ward_id_col = "ward_id" if "ward_id" in booth_cols else None

    where_parts: List[str] = []
    params: Dict[str, Any] = {}
    if scope:
        allowed_booth_ids = sorted([v for v in (scope.get("allowed_booth_ids") or []) if v is not None])
        allowed_ward_ids = sorted([v for v in (scope.get("allowed_ward_ids") or []) if v is not None])
        if allowed_booth_ids:
            clause, clause_params = _build_in_clause(f"b.{booth_pk_col}", allowed_booth_ids, "scope_booth")
            where_parts.append(clause)
            params.update(clause_params)
        elif allowed_ward_ids and booth_ward_id_col:
            clause, clause_params = _build_in_clause(f"b.{booth_ward_id_col}", allowed_ward_ids, "scope_ward")
            where_parts.append(clause)
            params.update(clause_params)

    where_clause = f"WHERE {' AND '.join(where_parts)}" if where_parts else ""
    rows = db.execute(
        text(
            f"""
            SELECT
                COALESCE(b.booth_no, b.{booth_pk_col}) AS id,
                b.booth_no AS booth_id,
                COALESCE(NULLIF(b.booth_add_en, ''), 'Booth ' || COALESCE(b.booth_no::text, b.{booth_pk_col}::text)) AS name_en
            FROM public.booths b
            {where_clause}
            ORDER BY
                b.booth_no ASC NULLS LAST,
                b.{booth_pk_col} ASC
            """
        ),
        params,
    ).all()
    dto = [{"id": int(r.id), "boothId": int(r.booth_id), "nameEn": r.name_en} for r in rows]
    return api_success("Booths fetched successfully", dto)


@app.get(f"{CONTEXT_PATH}/api/booths")
def get_booths_plural(
    assemblyCode: Optional[str] = None,
    wardId: Optional[int] = None,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "ADMIN", "USER")),
):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        return []
    q = (
        db.query(Booth)
        .join(Ward, Booth.ward_id == Ward.ward_id)
        .join(Assembly, Ward.assembly_id == Assembly.assembly_id)
    )
    if current.tenantId is not None:
        q = q.filter(Booth.tenant_id == current.tenantId)
    if scope:
        allowed_booth_ids = {v for v in (scope.get("allowed_booth_ids") or set()) if v is not None}
        allowed_ward_ids = {v for v in (scope.get("allowed_ward_ids") or set()) if v is not None}
        if allowed_booth_ids:
            q = q.filter(Booth.booth_id.in_(allowed_booth_ids))
        elif allowed_ward_ids:
            q = q.filter(Booth.ward_id.in_(allowed_ward_ids))
    if assemblyCode:
        q = q.filter(Assembly.assembly_code == assemblyCode)
    if wardId:
        if scope and (scope.get("allowed_ward_ids") or set()) and wardId not in (scope.get("allowed_ward_ids") or set()):
            return []
        q = q.filter(Ward.ward_id == wardId)

    booths = q.order_by(Booth.booth_id.asc()).all()
    if booths:
        return [
            {
                "boothId": b.booth_id,
                "pollingStationAdrEn": b.polling_station_adr_en,
                "pollingStationAdrLocal": b.polling_station_adr_local,
                "wardId": b.ward_id,
                "tenantId": b.tenant_id,
            }
            for b in booths
        ]

    booth_cols = _get_table_columns(db, "public", "booths")
    id_col = "booth_id" if "booth_id" in booth_cols else ("booth_no" if "booth_no" in booth_cols else ("id" if "id" in booth_cols else None))
    name_col = "polling_station_adr_en" if "polling_station_adr_en" in booth_cols else ("booth_add_en" if "booth_add_en" in booth_cols else ("name_en" if "name_en" in booth_cols else None))
    ward_ref = "ward_id" if "ward_id" in booth_cols else ("ward_no" if "ward_no" in booth_cols else None)
    if not id_col or not name_col:
        return []
    where = []
    params: Dict[str, Any] = {}
    if wardId and ward_ref:
        where.append(f"{ward_ref} = :ward_id")
        params["ward_id"] = wardId
    where_clause = f"WHERE {' AND '.join(where)}" if where else ""
    rows = db.execute(
        text(
            f"""
            SELECT {id_col} AS id, {name_col} AS name, {ward_ref if ward_ref else 'NULL'} AS ward_id
            FROM public.booths
            {where_clause}
            ORDER BY {name_col}
            """
        ),
        params,
    ).all()
    if scope:
        allowed_booth_ids = scope.get("allowed_booth_ids") or set()
        allowed_ward_ids = scope.get("allowed_ward_ids") or set()
        rows = [
            r
            for r in rows
            if (not allowed_booth_ids or int(r.id) in allowed_booth_ids)
            and (not allowed_ward_ids or (r.ward_id is not None and int(r.ward_id) in allowed_ward_ids))
        ]
    return [
        {
            "boothId": int(r.id),
            "pollingStationAdrEn": r.name,
            "pollingStationAdrLocal": None,
            "wardId": int(r.ward_id) if r.ward_id is not None else None,
            "tenantId": None,
        }
        for r in rows
    ]


@app.get(f"{CONTEXT_PATH}/api/wards")
def get_wards(
    assemblyId: Optional[int] = None,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "ADMIN", "USER")),
):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        return []
    q = db.query(Ward)
    if current.tenantId is not None:
        q = q.filter(Ward.tenant_id == current.tenantId)
    if assemblyId:
        q = q.filter(Ward.assembly_id == assemblyId)
    if scope:
        allowed_ward_ids = {v for v in (scope.get("allowed_ward_ids") or set()) if v is not None}
        if allowed_ward_ids:
            q = q.filter(Ward.ward_id.in_(allowed_ward_ids))
    wards = q.order_by(Ward.ward_id.asc()).all()
    if wards:
        return [
            {
                "wardId": w.ward_id,
                "wardNameEn": w.ward_name_en,
                "assemblyId": w.assembly_id,
                "tenantId": w.tenant_id,
            }
            for w in wards
        ]

    ward_cols = _get_table_columns(db, "public", "wards")
    id_col = "ward_id" if "ward_id" in ward_cols else ("ward_no" if "ward_no" in ward_cols else ("id" if "id" in ward_cols else None))
    name_col = "ward_name_en" if "ward_name_en" in ward_cols else ("name_en" if "name_en" in ward_cols else ("ward_name_local" if "ward_name_local" in ward_cols else ("name_kannada" if "name_kannada" in ward_cols else None)))
    assembly_ref = "assembly_id" if "assembly_id" in ward_cols else ("assembly_no" if "assembly_no" in ward_cols else None)
    if not id_col or not name_col:
        return []
    where = []
    params: Dict[str, Any] = {}
    if assemblyId and assembly_ref:
        where.append(f"{assembly_ref} = :assembly_id")
        params["assembly_id"] = assemblyId
    where_clause = f"WHERE {' AND '.join(where)}" if where else ""
    rows = db.execute(
        text(
            f"""
            SELECT {id_col} AS id, {name_col} AS name, {assembly_ref if assembly_ref else 'NULL'} AS assembly_id
            FROM public.wards
            {where_clause}
            ORDER BY {id_col}
            """
        ),
        params,
    ).all()
    if scope:
        allowed_ward_ids = scope.get("allowed_ward_ids") or set()
        if allowed_ward_ids:
            rows = [r for r in rows if int(r.id) in allowed_ward_ids]
        else:
            rows = []
    return [
        {
            "wardId": int(r.id),
            "wardNameEn": r.name,
            "assemblyId": int(r.assembly_id) if r.assembly_id is not None else None,
            "tenantId": None,
        }
        for r in rows
    ]


@app.get(f"{CONTEXT_PATH}/api/assemblies")
def get_assemblies_plural(
    assemblyCode: Optional[str] = None,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "USER")),
):
    q = db.query(Assembly).filter(Assembly.tenant_id == current.tenantId)
    if assemblyCode:
        q = q.filter(Assembly.assembly_code == assemblyCode)

    assemblies = q.order_by(Assembly.assembly_id.asc()).all()
    return [
        {
            "assemblyId": a.assembly_id,
            "assemblyCode": a.assembly_code,
            "assemblyNameEn": a.assembly_name_en,
            "assemblyNameLocal": a.assembly_name_local,
            "tenantId": a.tenant_id,
        }
        for a in assemblies
    ]


@app.put(f"{CONTEXT_PATH}/api/voters/{{voterId}}")
def update_voter(voterId: int, payload: VoterUpdatePayload, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    req = dict(payload.updateRequest or {})
    req["voterId"] = voterId

    updated_by = db.query(User).filter(User.first_name == current.firstName, User.phone == current.phone).first()
    if not updated_by:
        raise ValueError(f"User not found with Email: {current.firstName}")

    voter = db.query(Voter).filter(Voter.voter_id == req.get("voterId")).first()
    if not voter:
        raise ValueError(f"Voter not found with ID: {req.get('voterId')}")

    field_map = {
        "houseNoEn": "house_no_en",
        "houseNoLocal": "house_no_local",
        "relationType": "relation_type",
        "firstMiddleNameEn": "first_middle_name_en",
        "lastNameEn": "last_name_en",
        "firstMiddleNameLocal": "first_middle_name_local",
        "lastNameLocal": "last_name_local",
        "gender": "gender",
        "age": "age",
        "dob": "dob",
        "mobile": "mobile",
        "addressEn": "address_en",
        "addressLocal": "address_local",
        "latitude": "latitude",
        "longitude": "longitude",
    }

    logs: List[VoterChangeLog] = []
    for k, v in req.items():
        if k in ("voterId", "updatedByUserId") or v is None:
            continue
        if k not in field_map:
            raise ValueError(f"Invalid field in update request: {k}")

        attr = field_map[k]
        old = getattr(voter, attr)
        if old != v:
            logs.append(
                VoterChangeLog(
                    tenant_id=current.tenantId,
                    voter_id=voter.voter_id,
                    updated_by=updated_by.id,
                    field_name=k,
                    old_value=str(old) if old is not None else None,
                    new_value=str(v),
                    updated_at=datetime.utcnow(),
                    update_latitude=payload.updateLocationLat,
                    update_longitude=payload.updateLocationLng,
                )
            )
            setattr(voter, attr, v)

    db.add(voter)
    for log in logs:
        db.add(log)
    db.commit()

    voter_dict = {c.name: getattr(voter, c.name) for c in Voter.__table__.columns}
    return api_success("Voter updated successfully", voter_dict)


@app.put(f"{CONTEXT_PATH}/api/voters/by-epic/{{epic}}")
def update_public_voter_by_epic(epic: str, payload: PublicVoterUpdatePayload, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    normalized_epic = normalize_optional_text(epic)
    if not normalized_epic:
        raise ValueError("EPIC is required")

    voter_cols = _get_table_columns(db, "public", "voters")
    if not voter_cols:
        raise ValueError("public.voters table not found")

    req = dict(payload.updateRequest or {})
    if not req:
        raise ValueError("updateRequest is required")

    unsupported = sorted(k for k in req.keys() if k not in VOTER_ENRICHMENT_FIELD_MAP)
    if unsupported:
        raise ValueError(f"Invalid field in update request: {unsupported[0]}")

    normalized_values: Dict[str, Any] = {
        api_field: _serialize_enrichment_value(api_field, req.get(api_field))
        for api_field in req.keys()
    }
    if not normalized_values:
        raise ValueError("No supported fields provided in updateRequest")

    ward_code = normalize_optional_text(payload.wardCode)
    booth_no = normalize_optional_text(payload.boothNo)

    where_parts = ["epic = :epic"]
    params: Dict[str, Any] = {"epic": normalized_epic}
    if ward_code and "ward_code" in voter_cols:
        where_parts.append("CAST(ward_code AS TEXT) = :ward_code")
        params["ward_code"] = ward_code
    if booth_no and "booth_no" in voter_cols:
        where_parts.append("CAST(booth_no AS TEXT) = :booth_no")
        params["booth_no"] = booth_no

    public_voter = db.execute(
        text(
            f"""
            SELECT epic, ward_code, booth_no, sl, house, name_en, name_kannada, gender, age, rel_eng, rel_kannada, rel_type, mobile
            FROM public.voters
            WHERE {" AND ".join(where_parts)}
            LIMIT 1
            """
        ),
        params,
    ).first()
    if not public_voter and (ward_code or booth_no):
        public_voter = db.execute(
            text(
                """
                SELECT epic, ward_code, booth_no, sl, house, name_en, name_kannada, gender, age, rel_eng, rel_kannada, rel_type, mobile
                FROM public.voters
                WHERE epic = :epic
                LIMIT 1
                """
            ),
            {"epic": normalized_epic},
        ).first()
    if not public_voter:
        raise ResourceNotFoundException("Public voter", "epic", normalized_epic)

    updated_by = db.query(User).filter(User.first_name == current.firstName, User.phone == current.phone).first()
    enrichment = db.query(VoterEnrichment).filter(VoterEnrichment.epic == normalized_epic).first()
    if not enrichment:
        enrichment = VoterEnrichment(
            epic=normalized_epic,
            ward_code=str(public_voter.ward_code) if public_voter.ward_code is not None else ward_code,
            booth_no=str(public_voter.booth_no) if public_voter.booth_no is not None else booth_no,
        )
        db.add(enrichment)

    updated_fields = _parse_updated_fields(enrichment.updated_fields)
    for api_field, serialized_value in normalized_values.items():
        setattr(enrichment, VOTER_ENRICHMENT_FIELD_MAP[api_field], serialized_value)
        updated_fields.add(api_field)

    enrichment.ward_code = str(public_voter.ward_code) if public_voter.ward_code is not None else enrichment.ward_code
    enrichment.booth_no = str(public_voter.booth_no) if public_voter.booth_no is not None else enrichment.booth_no
    enrichment.updated_fields = json.dumps(sorted(updated_fields))
    enrichment.updated_at = datetime.utcnow()
    enrichment.updated_by = updated_by.id if updated_by else None
    db.commit()
    db.refresh(enrichment)
    return api_success(
        "Public voter updated successfully",
        _merge_voter_payload_with_enrichment(_build_public_voter_result(public_voter), _build_voter_enrichment_payload(enrichment)),
    )


@app.get(f"{CONTEXT_PATH}/api/voters")
def get_voters(
    assemblyCode: str,
    page: int = 0,
    size: int = 500,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "USER")),
):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        return []
    q = (
        db.query(Voter)
        .join(Booth, Voter.booth_id == Booth.booth_id)
        .join(Ward, Booth.ward_id == Ward.ward_id)
        .join(Assembly, Ward.assembly_id == Assembly.assembly_id)
        .filter(
            Voter.tenant_id == current.tenantId,
            Assembly.assembly_code == assemblyCode,
        )
        .order_by(Voter.voter_id.asc())
    )
    if scope:
        if scope.get("allowed_booth_ids"):
            q = q.filter(Booth.booth_id.in_(scope.get("allowed_booth_ids")))
        elif scope.get("allowed_ward_ids"):
            q = q.filter(Ward.ward_id.in_(scope.get("allowed_ward_ids")))

    total = q.count()
    voters = q.offset(page * size).limit(size).all()
    return [_build_voter_map(v) for v in voters]


@app.get(f"{CONTEXT_PATH}/api/voter-search")
@app.get(f"{CONTEXT_PATH}/api/voters/search")
def search_voters(
    assemblyCode: str,
    searchQuery: Optional[str] = None,
    wardId: Optional[int] = None,
    boothNumber: Optional[str] = None,
    mobileNumber: Optional[str] = None,
    epicId: Optional[str] = None,
    relationName: Optional[str] = None,
    houseNumber: Optional[str] = None,
    page: int = 0,
    size: int = 50,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "USER", "ADMIN")),
):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        empty_payload = api_success("Voter search fetched", [])
        empty_payload["data"]["meta"] = {
            "total": 0,
            "male": 0,
            "female": 0,
            "returned": 0,
            "page": page,
            "size": max(1, min(size, 2000)),
            "hasMore": False,
        }
        return empty_payload
    booth_cols = _get_table_columns(db, "public", "booths")
    voter_cols = _get_table_columns(db, "public", "voters")
    ward_cols = _get_table_columns(db, "public", "wards")

    booth_id_col = "id" if "id" in booth_cols else ("booth_id" if "booth_id" in booth_cols else None)
    booth_no_col = "booth_no" if "booth_no" in booth_cols else booth_id_col
    booth_ward_code_col = "ward_code" if "ward_code" in booth_cols else None
    booth_ward_id_col = "ward_id" if "ward_id" in booth_cols else None
    booth_name_en_col = "booth_add_en" if "booth_add_en" in booth_cols else ("polling_station_adr_en" if "polling_station_adr_en" in booth_cols else None)
    booth_name_local_col = "booth_add_local" if "booth_add_local" in booth_cols else ("polling_station_adr_local" if "polling_station_adr_local" in booth_cols else None)

    if not booth_id_col or not booth_no_col:
        return JSONResponse(status_code=404, content=api_error("Search failed", "public.booths missing required columns"))

    ward_id_col = "id" if "id" in ward_cols else ("ward_id" if "ward_id" in ward_cols else None)
    ward_code_col = "ward_code" if "ward_code" in ward_cols else ("ward_no" if "ward_no" in ward_cols else None)
    ward_name_en_col = "ward_name_en" if "ward_name_en" in ward_cols else ("name_en" if "name_en" in ward_cols else None)
    ward_name_local_col = "ward_name_local" if "ward_name_local" in ward_cols else ("name_kannada" if "name_kannada" in ward_cols else None)

    if scope and scope.get("allowed_assembly_ids"):
        assembly_cols = _ensure_public_assembly_code(db)
        assembly_pk_col = "id" if "id" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else None)
        assembly_no_col = "assembly_no" if "assembly_no" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else None)
        assembly_code_expr = "assembly_code" if "assembly_code" in assembly_cols else (
            f"LPAD(CAST({assembly_no_col} AS TEXT), 12, '0')" if assembly_no_col else "NULL"
        )
        assembly_row = db.execute(
            text(
                f"""
                SELECT {assembly_pk_col if assembly_pk_col else 'NULL'} AS assembly_pk,
                       {assembly_no_col if assembly_no_col else 'NULL'} AS assembly_no
                FROM public.assembly
                WHERE {assembly_code_expr} = :assembly_code
                   OR CAST({assembly_no_col if assembly_no_col else assembly_code_expr} AS TEXT) = :assembly_no_text
                LIMIT 1
                """
            ),
            {
                "assembly_code": normalize_assembly_code(assemblyCode),
                "assembly_no_text": str(int(normalize_assembly_code(assemblyCode))),
            },
        ).first()
        if not assembly_row:
            raise HTTPException(status_code=404, detail="Assembly not found")
        allowed_assembly_ids = {v for v in (scope.get("allowed_assembly_ids") or set()) if v is not None}
        assembly_id = assembly_row.assembly_pk if assembly_row.assembly_pk is not None else assembly_row.assembly_no
        if allowed_assembly_ids and assembly_id not in allowed_assembly_ids:
            raise HTTPException(status_code=403, detail="Access denied for requested assembly")

    ward_by_id: Dict[int, Dict[str, Any]] = {}
    if ward_id_col:
        ward_rows = db.execute(
            text(
                f"""
                SELECT
                    {ward_id_col} AS ward_id,
                    {ward_code_col if ward_code_col else 'NULL'} AS ward_code,
                    {ward_name_en_col if ward_name_en_col else 'NULL'} AS ward_name_en,
                    {ward_name_local_col if ward_name_local_col else 'NULL'} AS ward_name_local
                FROM public.wards
                """
            )
        ).all()
        for r in ward_rows:
            if scope and scope.get("allowed_ward_ids") and int(r.ward_id) not in scope.get("allowed_ward_ids"):
                continue
            ward_by_id[int(r.ward_id)] = {
                "wardId": int(r.ward_id),
                "wardCode": str(r.ward_code) if r.ward_code is not None else None,
                "wardNameEn": r.ward_name_en,
                "wardNameLocal": r.ward_name_local,
            }

    ward_code_filter: Optional[str] = None
    if wardId is not None and wardId in ward_by_id:
        ward_code_filter = ward_by_id[wardId].get("wardCode")
    if scope and scope.get("allowed_ward_ids") and wardId is not None and wardId not in scope.get("allowed_ward_ids"):
        empty_payload = api_success("Voter search fetched", [])
        empty_payload["data"]["meta"] = {
            "total": 0,
            "male": 0,
            "female": 0,
            "returned": 0,
            "page": page,
            "size": max(1, min(size, 2000)),
            "hasMore": False,
        }
        return empty_payload

    booth_sql = f"""
        SELECT
            {booth_id_col} AS booth_id,
            {booth_no_col} AS booth_no,
            {booth_ward_code_col if booth_ward_code_col else 'NULL'} AS ward_code,
            {booth_ward_id_col if booth_ward_id_col else 'NULL'} AS ward_id,
            {booth_name_en_col if booth_name_en_col else 'NULL'} AS booth_name_en,
            {booth_name_local_col if booth_name_local_col else 'NULL'} AS booth_name_local
        FROM public.booths
    """
    booth_rows = db.execute(text(booth_sql)).all()
    booth_by_key: Dict[tuple, Dict[str, Any]] = {}
    booths_by_no: Dict[str, List[Dict[str, Any]]] = {}
    for b in booth_rows:
        if scope:
            allowed_booth_ids = {v for v in (scope.get("allowed_booth_ids") or set()) if v is not None}
            allowed_ward_ids = {v for v in (scope.get("allowed_ward_ids") or set()) if v is not None}
            booth_pk = int(b.booth_id)
            ward_pk = int(b.ward_id) if b.ward_id is not None else None
            if allowed_booth_ids and booth_pk not in allowed_booth_ids:
                continue
            if allowed_ward_ids and ward_pk is not None and ward_pk not in allowed_ward_ids:
                continue
        row = {
            "boothId": int(b.booth_id),
            "boothNo": str(b.booth_no if b.booth_no is not None else b.booth_id),
            "wardCode": str(b.ward_code) if b.ward_code is not None else None,
            "wardId": int(b.ward_id) if b.ward_id is not None else None,
            "boothNameEn": b.booth_name_en,
            "boothNameLocal": b.booth_name_local,
        }
        booth_by_key[(row["wardCode"], row["boothNo"])] = row
        booths_by_no.setdefault(row["boothNo"], []).append(row)

    voter_id_col = "voter_id" if "voter_id" in voter_cols else ("id" if "id" in voter_cols else None)
    voter_sr_col = "sl" if "sl" in voter_cols else ("sr_no" if "sr_no" in voter_cols else None)
    voter_epic_col = "epic" if "epic" in voter_cols else ("epic_no" if "epic_no" in voter_cols else None)
    voter_name_en_col = "name_en" if "name_en" in voter_cols else ("first_middle_name_en" if "first_middle_name_en" in voter_cols else None)
    voter_name_local_col = "name_kannada" if "name_kannada" in voter_cols else ("first_middle_name_local" if "first_middle_name_local" in voter_cols else None)
    voter_relation_en_col = "relation_name_en" if "relation_name_en" in voter_cols else ("relation_first_middle_name_en" if "relation_first_middle_name_en" in voter_cols else ("rel_eng" if "rel_eng" in voter_cols else None))
    voter_relation_local_col = "relation_name_local" if "relation_name_local" in voter_cols else ("relation_first_middle_name_local" if "relation_first_middle_name_local" in voter_cols else ("rel_kannada" if "rel_kannada" in voter_cols else None))
    voter_relation_type_col = "relation_type" if "relation_type" in voter_cols else ("rel_type" if "rel_type" in voter_cols else None)
    voter_age_col = "age" if "age" in voter_cols else None
    voter_house_col = "house" if "house" in voter_cols else ("house_no_en" if "house_no_en" in voter_cols else None)
    voter_relation_en_col = "relation_name_en" if "relation_name_en" in voter_cols else ("relation_first_middle_name_en" if "relation_first_middle_name_en" in voter_cols else ("rel_eng" if "rel_eng" in voter_cols else None))
    voter_relation_local_col = "relation_name_local" if "relation_name_local" in voter_cols else ("relation_first_middle_name_local" if "relation_first_middle_name_local" in voter_cols else ("rel_kannada" if "rel_kannada" in voter_cols else None))
    voter_relation_type_col = "relation_type" if "relation_type" in voter_cols else ("rel_type" if "rel_type" in voter_cols else None)
    voter_age_col = "age" if "age" in voter_cols else None
    voter_gender_col = "gender" if "gender" in voter_cols else None
    voter_mobile_col = "mobile" if "mobile" in voter_cols else None
    voter_booth_no_col = "booth_no" if "booth_no" in voter_cols else ("booth_id" if "booth_id" in voter_cols else None)
    voter_ward_code_col = "ward_code" if "ward_code" in voter_cols else None

    if not voter_booth_no_col:
        return JSONResponse(status_code=404, content=api_error("Search failed", "public.voters missing booth mapping column"))

    where_parts = ["1=1"]
    params: Dict[str, Any] = {"limit": max(1, min(size, 2000)), "offset": max(page, 0) * max(1, min(size, 2000))}

    if scope:
        allowed_ward_codes = sorted({str(v.get("wardCode")) for v in ward_by_id.values() if v.get("wardCode")})
        allowed_booth_nos = sorted({row.get("boothNo") for row in booth_by_key.values() if row.get("boothNo")})
        if allowed_ward_codes and voter_ward_code_col:
            clause, clause_params = _build_in_clause(f"CAST({voter_ward_code_col} AS TEXT)", allowed_ward_codes, "scope_ward_code")
            where_parts.append(clause)
            params.update(clause_params)
        if allowed_booth_nos:
            clause, clause_params = _build_in_clause(f"CAST({voter_booth_no_col} AS TEXT)", allowed_booth_nos, "scope_booth_no")
            where_parts.append(clause)
            params.update(clause_params)

    if ward_code_filter and voter_ward_code_col:
        where_parts.append(f"{voter_ward_code_col} = :ward_code")
        params["ward_code"] = ward_code_filter
    if boothNumber:
        where_parts.append(f"CAST({voter_booth_no_col} AS TEXT) ILIKE :booth_number")
        params["booth_number"] = f"%{boothNumber.strip()}%"
    if mobileNumber and voter_mobile_col:
        where_parts.append(f"CAST({voter_mobile_col} AS TEXT) ILIKE :mobile_number")
        params["mobile_number"] = f"%{mobileNumber.strip()}%"
    if epicId and voter_epic_col:
        where_parts.append(f"CAST({voter_epic_col} AS TEXT) ILIKE :epic_id")
        params["epic_id"] = f"%{epicId.strip()}%"
    if houseNumber and voter_house_col:
        where_parts.append(f"CAST({voter_house_col} AS TEXT) ILIKE :house_number")
        params["house_number"] = f"%{houseNumber.strip()}%"
    if relationName and voter_relation_en_col:
        where_parts.append(f"CAST({voter_relation_en_col} AS TEXT) ILIKE :relation_name")
        params["relation_name"] = f"%{relationName.strip()}%"

    search_fields: List[str] = []
    if voter_name_en_col:
        search_fields.append(f"CAST({voter_name_en_col} AS TEXT)")
    if voter_epic_col:
        search_fields.append(f"CAST({voter_epic_col} AS TEXT)")
    if voter_mobile_col:
        search_fields.append(f"CAST({voter_mobile_col} AS TEXT)")
    if voter_sr_col:
        search_fields.append(f"CAST({voter_sr_col} AS TEXT)")
    search_fields.append(f"CAST({voter_booth_no_col} AS TEXT)")
    if voter_house_col:
        search_fields.append(f"CAST({voter_house_col} AS TEXT)")
    if voter_relation_en_col:
        search_fields.append(f"CAST({voter_relation_en_col} AS TEXT)")
    if searchQuery and search_fields:
        where_parts.append("(" + " OR ".join([f"{f} ILIKE :search_query" for f in search_fields]) + ")")
        params["search_query"] = f"%{searchQuery.strip()}%"

    where_sql = " AND ".join(where_parts)

    order_expr = f"CAST({voter_sr_col} AS INT)" if voter_sr_col else (f"CAST({voter_epic_col} AS TEXT) ASC NULLS LAST" if voter_epic_col else voter_booth_no_col)
    voter_id_expr = f"{voter_id_col}" if voter_id_col else (f"COALESCE(CAST({voter_sr_col} AS BIGINT), ROW_NUMBER() OVER ())" if voter_sr_col else "ROW_NUMBER() OVER ()")
    voters_rows = db.execute(
        text(
            f"""
            SELECT
                {voter_id_expr} AS voter_id,
                {voter_sr_col if voter_sr_col else 'NULL'} AS sl,
                {voter_epic_col if voter_epic_col else 'NULL'} AS epic_no,
                {voter_name_en_col if voter_name_en_col else 'NULL'} AS name_en,
                {voter_name_local_col if voter_name_local_col else 'NULL'} AS name_local,
                {voter_relation_en_col if voter_relation_en_col else 'NULL'} AS relation_name_en,
                {voter_relation_local_col if voter_relation_local_col else 'NULL'} AS relation_name_local,
                {voter_relation_type_col if voter_relation_type_col else 'NULL'} AS relation_type,
                {voter_house_col if voter_house_col else 'NULL'} AS house_no_en,
                {voter_gender_col if voter_gender_col else 'NULL'} AS gender,
                {voter_age_col if voter_age_col else 'NULL'} AS age,
                {voter_mobile_col if voter_mobile_col else 'NULL'} AS mobile,
                {voter_booth_no_col} AS booth_no,
                {voter_ward_code_col if voter_ward_code_col else 'NULL'} AS ward_code
            FROM public.voters
            WHERE {where_sql}
            ORDER BY {order_expr}
            LIMIT :limit OFFSET :offset
            """
        ),
        params,
    ).all()

    result: List[Dict[str, Any]] = []
    for v in voters_rows:
        ward_code = str(v.ward_code) if v.ward_code is not None else None
        booth_no = str(v.booth_no)
        booth = booth_by_key.get((ward_code, booth_no))
        if not booth:
            candidates = booths_by_no.get(booth_no, [])
            booth = candidates[0] if candidates else None

        ward_id_mapped = booth.get("wardId") if booth else None
        if wardId is not None and ward_id_mapped is not None and int(ward_id_mapped) != int(wardId):
            continue

        ward_info = ward_by_id.get(int(ward_id_mapped)) if ward_id_mapped is not None else None
        result.append(
            {
                "voterId": int(v.voter_id) if v.voter_id is not None else None,
                "sl": v.sl,
                "epicNo": v.epic_no,
                "firstMiddleNameEn": v.name_en,
                "lastNameEn": "",
                "firstMiddleNameLocal": v.name_local,
                "lastNameLocal": "",
                "relationFirstMiddleNameEn": v.relation_name_en,
                "relationFirstMiddleNameLocal": v.relation_name_local,
                "relationType": v.relation_type,
                "relationLastNameEn": "",
                "houseNoEn": str(v.house_no_en) if v.house_no_en is not None else None,
                "houseNoLocal": None,
                "gender": v.gender,
                "age": v.age,
                "mobile": str(v.mobile) if v.mobile is not None else None,
                "wardCode": ward_code,
                "boothNo": booth_no,
                "boothInfo": {
                    "boothId": booth.get("boothId") if booth else None,
                    "boothNo": booth.get("boothNo") if booth else booth_no,
                    "wardCode": booth.get("wardCode") if booth else ward_code,
                    "boothNameEn": booth.get("boothNameEn") if booth else None,
                    "boothNameLocal": booth.get("boothNameLocal") if booth else None,
                },
                "wardId": ward_id_mapped,
                "wardNameEn": ward_info.get("wardNameEn") if ward_info else None,
                "wardNameLocal": ward_info.get("wardNameLocal") if ward_info else None,
            }
        )
    _merge_voter_payloads_with_enrichment(db, result)

    count_sql = f"""
        SELECT
            COUNT(*)::int AS total_count,
            SUM(CASE WHEN UPPER(COALESCE({voter_gender_col}, '')) LIKE 'M%' THEN 1 ELSE 0 END)::int AS male_count,
            SUM(CASE WHEN UPPER(COALESCE({voter_gender_col}, '')) LIKE 'F%' THEN 1 ELSE 0 END)::int AS female_count
        FROM public.voters
        WHERE {where_sql}
    """ if voter_gender_col else f"""
        SELECT
            COUNT(*)::int AS total_count,
            0::int AS male_count,
            0::int AS female_count
        FROM public.voters
        WHERE {where_sql}
    """
    count_params = {k: v for k, v in params.items() if k not in {"limit", "offset"}}
    counts = db.execute(text(count_sql), count_params).first()
    total_count = int((counts.total_count if counts else 0) or 0)
    male_count = int((counts.male_count if counts else 0) or 0)
    female_count = int((counts.female_count if counts else 0) or 0)

    payload = api_success("Voter search fetched", result)
    payload["data"]["meta"] = {
        "total": total_count,
        "male": male_count,
        "female": female_count,
        "returned": len(result),
        "page": page,
        "size": max(1, min(size, 2000)),
        "hasMore": (max(page, 0) * max(1, min(size, 2000)) + len(result)) < total_count,
    }
    return payload


@app.get(f"{CONTEXT_PATH}/api/voters/by-booth")
def get_voters_by_booth(
    boothId: int,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "USER", "ADMIN")),
):
    scope = _resolve_access_scope_ids(db, current)
    if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
        raise HTTPException(status_code=403, detail="No access scope defined for this user")
    booth_cols = _get_table_columns(db, "public", "booths")
    voter_cols = _get_table_columns(db, "public", "voters")
    ward_cols = _get_table_columns(db, "public", "wards")

    booth_id_col = "id" if "id" in booth_cols else ("booth_id" if "booth_id" in booth_cols else None)
    booth_no_col = "booth_no" if "booth_no" in booth_cols else booth_id_col
    booth_ward_code_col = "ward_code" if "ward_code" in booth_cols else None
    booth_ward_id_col = "ward_id" if "ward_id" in booth_cols else None
    booth_name_en_col = "booth_add_en" if "booth_add_en" in booth_cols else ("polling_station_adr_en" if "polling_station_adr_en" in booth_cols else None)
    booth_name_local_col = "booth_add_local" if "booth_add_local" in booth_cols else ("polling_station_adr_local" if "polling_station_adr_local" in booth_cols else None)

    if not booth_id_col or not booth_no_col:
        return JSONResponse(status_code=404, content=api_error("Booth not found", "public.booths missing required columns"))

    booth_row = db.execute(
        text(
            f"""
            SELECT
                {booth_id_col} AS booth_id,
                {booth_no_col} AS booth_no,
                {booth_ward_code_col if booth_ward_code_col else 'NULL'} AS ward_code,
                {booth_ward_id_col if booth_ward_id_col else 'NULL'} AS ward_id,
                {booth_name_en_col if booth_name_en_col else 'NULL'} AS booth_name_en,
                {booth_name_local_col if booth_name_local_col else 'NULL'} AS booth_name_local
            FROM public.booths
            WHERE {booth_id_col} = :booth_id
            LIMIT 1
            """
        ),
        {"booth_id": boothId},
    ).first()
    if not booth_row:
        return JSONResponse(status_code=404, content=api_error("Booth not found", f"Invalid boothId: {boothId}"))
    if scope:
        allowed_booth_ids = scope.get("allowed_booth_ids") or set()
        allowed_ward_ids = scope.get("allowed_ward_ids") or set()
        if allowed_booth_ids and int(booth_row.booth_id) not in allowed_booth_ids:
            raise HTTPException(status_code=403, detail="Access denied for requested booth")
        if allowed_ward_ids and booth_row.ward_id is not None and int(booth_row.ward_id) not in allowed_ward_ids:
            raise HTTPException(status_code=403, detail="Access denied for requested booth")

    ward_name_en = None
    ward_name_local = None
    if booth_row.ward_id is not None and ("id" in ward_cols or "ward_id" in ward_cols):
        ward_id_col = "id" if "id" in ward_cols else "ward_id"
        ward_name_en_col = "ward_name_en" if "ward_name_en" in ward_cols else ("name_en" if "name_en" in ward_cols else None)
        ward_name_local_col = "ward_name_local" if "ward_name_local" in ward_cols else ("name_kannada" if "name_kannada" in ward_cols else None)
        if ward_name_en_col or ward_name_local_col:
            ward_row = db.execute(
                text(
                    f"""
                    SELECT
                        {ward_name_en_col if ward_name_en_col else 'NULL'} AS ward_name_en,
                        {ward_name_local_col if ward_name_local_col else 'NULL'} AS ward_name_local
                    FROM public.wards
                    WHERE {ward_id_col} = :ward_id
                    LIMIT 1
                    """
                ),
                {"ward_id": booth_row.ward_id},
            ).first()
            if ward_row:
                ward_name_en = ward_row.ward_name_en
                ward_name_local = ward_row.ward_name_local

    voter_id_col = "voter_id" if "voter_id" in voter_cols else ("id" if "id" in voter_cols else None)
    voter_sr_col = "sl" if "sl" in voter_cols else ("sr_no" if "sr_no" in voter_cols else None)
    voter_epic_col = "epic" if "epic" in voter_cols else ("epic_no" if "epic_no" in voter_cols else None)
    voter_name_en_col = "name_en" if "name_en" in voter_cols else ("first_middle_name_en" if "first_middle_name_en" in voter_cols else None)
    voter_name_local_col = "name_kannada" if "name_kannada" in voter_cols else ("first_middle_name_local" if "first_middle_name_local" in voter_cols else None)
    voter_relation_en_col = "relation_name_en" if "relation_name_en" in voter_cols else ("relation_first_middle_name_en" if "relation_first_middle_name_en" in voter_cols else ("rel_eng" if "rel_eng" in voter_cols else None))
    voter_relation_local_col = "relation_name_local" if "relation_name_local" in voter_cols else ("relation_first_middle_name_local" if "relation_first_middle_name_local" in voter_cols else ("rel_kannada" if "rel_kannada" in voter_cols else None))
    voter_relation_type_col = "relation_type" if "relation_type" in voter_cols else ("rel_type" if "rel_type" in voter_cols else None)
    voter_age_col = "age" if "age" in voter_cols else None
    voter_house_col = "house" if "house" in voter_cols else ("house_no_en" if "house_no_en" in voter_cols else None)
    voter_gender_col = "gender" if "gender" in voter_cols else None
    voter_booth_no_col = "booth_no" if "booth_no" in voter_cols else ("booth_id" if "booth_id" in voter_cols else None)
    voter_ward_code_col = "ward_code" if "ward_code" in voter_cols else None

    if not voter_booth_no_col:
        return JSONResponse(status_code=404, content=api_error("Voters not found", "public.voters missing booth mapping column"))

    where_clause = f"{voter_booth_no_col} = :booth_no"
    params: Dict[str, Any] = {"booth_no": booth_row.booth_no}
    if voter_ward_code_col and booth_row.ward_code is not None:
        where_clause += f" AND {voter_ward_code_col} = :ward_code"
        params["ward_code"] = booth_row.ward_code

    voter_id_expr = f"{voter_id_col}" if voter_id_col else (f"COALESCE(CAST({voter_sr_col} AS BIGINT), ROW_NUMBER() OVER ())" if voter_sr_col else "ROW_NUMBER() OVER ()")
    order_expr = (
        f"CAST({voter_sr_col} AS INT) ASC NULLS LAST, CAST({voter_epic_col} AS TEXT) ASC NULLS LAST"
        if (voter_sr_col and voter_epic_col)
        else (f"CAST({voter_sr_col} AS INT) ASC NULLS LAST" if voter_sr_col else (f"CAST({voter_epic_col} AS TEXT) ASC NULLS LAST" if voter_epic_col else voter_booth_no_col))
    )
    voters_rows = db.execute(
        text(
            f"""
            SELECT
                {voter_id_expr} AS voter_id,
                {voter_sr_col if voter_sr_col else 'NULL'} AS sl,
                {voter_epic_col if voter_epic_col else 'NULL'} AS epic_no,
                {voter_name_en_col if voter_name_en_col else 'NULL'} AS name_en,
                {voter_name_local_col if voter_name_local_col else 'NULL'} AS name_local,
                {voter_relation_en_col if voter_relation_en_col else 'NULL'} AS relation_name_en,
                {voter_relation_local_col if voter_relation_local_col else 'NULL'} AS relation_name_local,
                {voter_relation_type_col if voter_relation_type_col else 'NULL'} AS relation_type,
                {voter_house_col if voter_house_col else 'NULL'} AS house_no_en,
                {voter_gender_col if voter_gender_col else 'NULL'} AS gender,
                {voter_age_col if voter_age_col else 'NULL'} AS age
            FROM public.voters
            WHERE {where_clause}
            ORDER BY {order_expr}
            """
        ),
        params,
    ).all()

    voters = []
    male = 0
    female = 0
    for v in voters_rows:
        g = (v.gender or "").upper()
        if g.startswith("M"):
            male += 1
        if g.startswith("F"):
            female += 1
        voters.append(
            {
                "voterId": int(v.voter_id) if v.voter_id is not None else None,
                "sl": v.sl,
                "epicNo": v.epic_no,
                "firstMiddleNameEn": v.name_en,
                "lastNameEn": "",
                "firstMiddleNameLocal": v.name_local,
                "lastNameLocal": "",
                "relationFirstMiddleNameEn": v.relation_name_en,
                "relationFirstMiddleNameLocal": v.relation_name_local,
                "relationType": v.relation_type,
                "houseNoEn": str(v.house_no_en) if v.house_no_en is not None else None,
                "houseNoLocal": None,
                "gender": v.gender,
                "age": v.age,
                "dob": None,
                "mobile": None,
                "wardCode": str(booth_row.ward_code) if booth_row.ward_code is not None else None,
                "boothNo": str(booth_row.booth_no) if booth_row.booth_no is not None else None,
                "addressEn": None,
                "addressLocal": None,
                "status": None,
                "community": None,
                "caste": None,
                "residenceType": None,
                "civicIssue": None,
                "motherTongue": None,
                "team": None,
                "ownership": None,
                "education": None,
                "natureOfVoter": None,
                "latitude": None,
                "longitude": None,
            }
        )
    _merge_voter_payloads_with_enrichment(db, voters)

    return api_success(
        "Booth voters fetched",
        {
            "boothId": int(booth_row.booth_id),
            "boothNameEn": booth_row.booth_name_en,
            "boothNameLocal": booth_row.booth_name_local,
            "wardId": booth_row.ward_id,
            "wardCode": str(booth_row.ward_code) if booth_row.ward_code is not None else None,
            "wardNameEn": ward_name_en,
            "wardNameLocal": ward_name_local,
            "voterStats": {"total": len(voters), "male": male, "female": female},
            "voters": voters,
        },
    )


@app.get(f"{CONTEXT_PATH}/api/voters/snapshot")
def get_snapshot(
    assemblyCode: str,
    request: Request,
    includeVoters: bool = False,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("SUPER_ADMIN", "USER", "ADMIN")),
):
    try:
        scope = _resolve_access_scope_ids(db, current)
        if scope and not (scope.get("allowed_assembly_ids") or scope.get("allowed_ward_ids") or scope.get("allowed_booth_ids")):
            raise HTTPException(status_code=403, detail="No access scope defined for this user")
        public_snapshot = _build_public_snapshot(
            assemblyCode,
            db,
            include_voters=includeVoters,
            allowed_assembly_ids=scope.get("allowed_assembly_ids") if scope else None,
            allowed_ward_ids=scope.get("allowed_ward_ids") if scope else None,
            allowed_booth_ids=scope.get("allowed_booth_ids") if scope else None,
        )
        snapshot_id = _cache_snapshot(public_snapshot)
        snapshot_url = f"{_external_base_url(request)}{CONTEXT_PATH}/api/voters/snapshot/content/{snapshot_id}"
        payload = api_success("Snapshot fetched successfully", snapshot_url)
        payload["snapshotMode"] = "link"
        return JSONResponse(content=payload, headers={"X-Snapshot-Mode": "link"})
    except HTTPException as ex:
        return JSONResponse(status_code=ex.status_code, content=api_error("Snapshot failed", str(ex.detail)))
    except ValueError as ex:
        return JSONResponse(status_code=404, content=api_error("No snapshot found", str(ex)))
    except Exception as ex:
        print("[SNAPSHOT_ERROR]", ex)
        print(traceback.format_exc())
        return JSONResponse(status_code=500, content=api_error("Snapshot failed", str(ex)))


@app.get(f"{CONTEXT_PATH}/api/voters/snapshot/content/{{snapshot_id}}")
def get_snapshot_content(snapshot_id: str):
    _cleanup_snapshot_cache()
    cached = _snapshot_cache.get(snapshot_id)
    if not cached:
        raise HTTPException(status_code=404, detail="Snapshot link expired or not found")
    return cached["payload"]


@app.get(f"{CONTEXT_PATH}/api/association")
def list_association(boothId: int, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    rows = (
        db.query(Association)
        .filter(Association.booth_id == boothId, Association.tenant_id == current.tenantId)
        .all()
    )
    dtos = [{"associationId": r.association_id, "associationName": r.association_name} for r in rows]
    return api_success("Associations fetched", dtos)


@app.post(f"{CONTEXT_PATH}/api/association")
def create_association(payload: CreateAssociationRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN"))):
    booth = db.query(Booth).filter(Booth.booth_id == payload.boothId).first()
    if not booth:
        raise ValueError("Invalid booth ID")

    association = Association(
        association_name=payload.associationName,
        association_address=payload.associationAddress,
        association_head_name=payload.associationHeadName,
        phone=payload.phone,
        latitude=payload.latitude,
        longitude=payload.longitude,
        booth_id=booth.booth_id,
        tenant_id=current.tenantId,
    )
    db.add(association)
    db.commit()
    db.refresh(association)

    return api_success(
        "Association created",
        {
            "associationId": association.association_id,
            "associationName": association.association_name,
        },
    )


def _family_to_dto(db: Session, fam: Family) -> Dict[str, Any]:
    members = db.query(FamilyMember, Voter).join(Voter, FamilyMember.voter_id == Voter.voter_id).filter(FamilyMember.family_id == fam.familyId).all()
    m_dto = []
    for member, voter in members:
        full_name = f"{voter.first_middle_name_en or ''} {voter.last_name_en or ''}".strip()
        m_dto.append(
            {
                "memberId": member.member_id,
                "head": bool(member.is_head),
                "epicNo": voter.epic_no,
                "voterName": full_name,
            }
        )

    return {
        "familyId": fam.familyId,
        "tenantId": fam.tenant_id,
        "familyName": fam.family_name,
        "familyAddress": fam.family_address,
        "phone": fam.phone,
        "points": fam.points,
        "pointsProvided": fam.points_provided,
        "latitude": fam.latitude,
        "longitude": fam.longitude,
        "boothId": fam.booth_id,
        "associationId": fam.association_id,
        "headMemberId": fam.head_voter_id,
        "members": m_dto,
        "economicStatus": fam.economic_status,
        "familyNature": fam.family_nature,
    }


@app.post(f"{CONTEXT_PATH}/api/family")
def create_family(payload: CreateFamilyRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "USER"))):
    booth = db.query(Booth).filter(Booth.booth_id == payload.boothId).first()
    if not booth:
        raise ValueError("Invalid booth")

    association = None
    if payload.associationId is not None:
        association = db.query(Association).filter(Association.association_id == payload.associationId).first()
        if not association:
            raise ValueError("Invalid association")

    fam = Family(
        family_name=payload.familyName,
        family_address=payload.familyAddress,
        phone=payload.phone,
        points=payload.points,
        points_provided=payload.pointsProvided,
        latitude=payload.latitude,
        longitude=payload.longitude,
        economic_status=payload.economicStatus,
        family_nature=payload.familyNature,
        tenant_id=current.tenantId,
        booth_id=booth.booth_id,
        association_id=association.association_id if association else None,
        deleted=False,
    )
    db.add(fam)
    db.flush()

    head_member_id = None
    for epic in payload.memberEpicNos:
        voter = db.query(Voter).filter(Voter.epic_no == epic, Voter.tenant_id == current.tenantId).first()
        if not voter:
            raise ValueError(f"Voter not found: {epic}")

        is_head = payload.headEpicNo == epic
        member = FamilyMember(family_id=fam.familyId, voter_id=voter.voter_id, is_head=is_head)
        db.add(member)
        db.flush()
        if is_head:
            head_member_id = member.member_id

    fam.head_voter_id = head_member_id
    db.add(fam)
    db.commit()
    db.refresh(fam)

    return api_success("Family created", _family_to_dto(db, fam))


@app.put(f"{CONTEXT_PATH}/api/family/{{familyId}}")
def update_family(familyId: int, payload: UpdateFamilyRequest, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN", "USER"))):
    fam = db.query(Family).filter(Family.familyId == familyId).first()
    if not fam:
        raise ValueError(f"Family not found: {familyId}")

    fam.family_name = payload.familyName
    fam.family_address = payload.familyAddress
    fam.phone = payload.phone
    fam.points = payload.points
    fam.points_provided = payload.pointsProvided
    fam.latitude = payload.latitude
    fam.longitude = payload.longitude
    fam.economic_status = payload.economicStatus
    fam.family_nature = payload.familyNature

    if payload.associationId is not None:
        association = db.query(Association).filter(Association.association_id == payload.associationId).first()
        if not association:
            raise ValueError("Invalid association")
        fam.association_id = association.association_id
    else:
        fam.association_id = None

    existing = db.query(FamilyMember).join(Voter, FamilyMember.voter_id == Voter.voter_id).filter(FamilyMember.family_id == fam.familyId).all()
    keep = set(payload.memberEpicNos)

    for m in existing:
        voter = db.query(Voter).filter(Voter.voter_id == m.voter_id).first()
        if voter and voter.epic_no not in keep:
            db.delete(m)

    db.flush()

    all_members = db.query(FamilyMember).join(Voter, FamilyMember.voter_id == Voter.voter_id).filter(FamilyMember.family_id == fam.familyId).all()
    existing_epic = {
        db.query(Voter).filter(Voter.voter_id == m.voter_id).first().epic_no: m
        for m in all_members
    }

    for epic in payload.memberEpicNos:
        if epic not in existing_epic:
            voter = db.query(Voter).filter(Voter.epic_no == epic, Voter.tenant_id == current.tenantId).first()
            if not voter:
                raise ValueError(f"Voter not found: {epic}")
            member = FamilyMember(family_id=fam.familyId, voter_id=voter.voter_id, is_head=(epic == payload.headEpicNo))
            db.add(member)

    db.flush()

    members_after = db.query(FamilyMember).join(Voter, FamilyMember.voter_id == Voter.voter_id).filter(FamilyMember.family_id == fam.familyId).all()
    head_member_id = None
    for m in members_after:
        voter = db.query(Voter).filter(Voter.voter_id == m.voter_id).first()
        is_head = bool(voter and voter.epic_no == payload.headEpicNo)
        m.is_head = is_head
        if is_head:
            head_member_id = m.member_id

    fam.head_voter_id = head_member_id
    db.add(fam)
    db.commit()
    db.refresh(fam)

    return api_success("Family updated", _family_to_dto(db, fam))


@app.get(f"{CONTEXT_PATH}/api/family")
def list_families(
    boothId: int,
    page: int = 0,
    size: int = 10,
    search: Optional[str] = None,
    association: Optional[str] = None,
    db: Session = Depends(get_db),
    current: JwtUserDetails = Depends(require_roles("ADMIN", "USER")),
):
    q = db.query(Family).filter(Family.tenant_id == current.tenantId, Family.booth_id == boothId, Family.deleted.is_(False))

    association_filter = parse_optional_bool(association)
    if association_filter is not None:
        if association_filter:
            q = q.filter(Family.association_id.is_not(None))
        else:
            q = q.filter(Family.association_id.is_(None))

    if search and search.strip():
        s = f"%{search.lower().strip()}%"
        q = (
            q.outerjoin(FamilyMember, FamilyMember.family_id == Family.familyId)
            .outerjoin(Voter, FamilyMember.voter_id == Voter.voter_id)
            .filter(
                or_(
                    func.lower(Family.family_name).like(s),
                    func.lower(Family.family_address).like(s),
                    func.lower(Voter.first_middle_name_en).like(s),
                    func.lower(Voter.epic_no).like(s),
                )
            )
            .distinct()
        )

    total = q.count()
    families = q.offset(page * size).limit(size).all()
    return build_page([_family_to_dto(db, f) for f in families], page, size, total)


@app.get(f"{CONTEXT_PATH}/api/family/{{id}}")
def get_family(id: int, db: Session = Depends(get_db), _: JwtUserDetails = Depends(require_roles("ADMIN", "USER"))):
    fam = db.query(Family).filter(Family.familyId == id).first()
    if not fam:
        raise ValueError(f"Family not found: {id}")
    return _family_to_dto(db, fam)


@app.delete(f"{CONTEXT_PATH}/api/family/{{id}}")
def delete_family(id: int, db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN"))):
    fam = db.query(Family).filter(Family.familyId == id).first()
    if not fam:
        raise ValueError(f"Family not found: {id}")
    if fam.tenant_id != current.tenantId:
        raise ValueError("Unauthorized: Family does not belong to your tenant")

    fam.deleted = True
    db.commit()
    return api_success("Family deleted", {})


@app.get(f"{CONTEXT_PATH}/api/volunteers/dropdown")
def volunteer_dropdown(level: str, parentId: Optional[int] = None, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    level = level.upper()
    out = []

    if level == "ASSEMBLY":
        rows = (
            db.query(Assembly.assembly_id, Assembly.assembly_name_en)
            .order_by(Assembly.assembly_name_en)
            .all()
        )
        if current.tenantId is not None:
            rows = (
                db.query(Assembly.assembly_id, Assembly.assembly_name_en)
                .filter(Assembly.tenant_id == current.tenantId)
                .order_by(Assembly.assembly_name_en)
                .all()
            )
        if not rows:
            assembly_cols = _ensure_public_assembly_code(db)
            id_col = "id" if "id" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else ("assembly_no" if "assembly_no" in assembly_cols else None))
            name_col = "assembly_name_en" if "assembly_name_en" in assembly_cols else ("name_en" if "name_en" in assembly_cols else ("assembly_name_local" if "assembly_name_local" in assembly_cols else ("name_kannada" if "name_kannada" in assembly_cols else None)))
            code_expr = "assembly_code" if "assembly_code" in assembly_cols else "NULL"
            if id_col and name_col:
                rows = db.execute(
                    text(
                        f"""
                        SELECT {id_col} AS id, {name_col} AS name, {code_expr} AS code
                        FROM public.assembly
                        ORDER BY {name_col}
                        """
                    )
                ).all()
                for row in rows:
                    out.append({"id": int(row.id), "code": str(row.code or row.id), "name": row.name})
                return out
        for row in rows:
            out.append({"id": row[0], "code": str(row[0]), "name": row[1]})
    elif level == "WARD":
        if parentId is None:
            raise ValueError("assemblyId is required for WARD")
        rows = (
            db.query(Ward.ward_id, Ward.ward_name_en)
            .filter(Ward.assembly_id == parentId)
            .order_by(Ward.ward_name_en)
            .all()
        )
        if current.tenantId is not None:
            rows = (
                db.query(Ward.ward_id, Ward.ward_name_en)
                .filter(Ward.assembly_id == parentId, Ward.tenant_id == current.tenantId)
                .order_by(Ward.ward_name_en)
                .all()
            )
        if not rows:
            ward_cols = _get_table_columns(db, "public", "wards")
            id_col = "ward_id" if "ward_id" in ward_cols else ("ward_no" if "ward_no" in ward_cols else ("id" if "id" in ward_cols else None))
            name_col = "ward_name_en" if "ward_name_en" in ward_cols else ("name_en" if "name_en" in ward_cols else ("ward_name_local" if "ward_name_local" in ward_cols else ("name_kannada" if "name_kannada" in ward_cols else None)))
            assembly_ref = "assembly_id" if "assembly_id" in ward_cols else ("assembly_no" if "assembly_no" in ward_cols else None)
            if id_col and name_col:
                where = f"WHERE {assembly_ref} = :assembly_id" if assembly_ref else ""
                rows = db.execute(
                    text(
                        f"""
                        SELECT {id_col} AS id, {name_col} AS name
                        FROM public.wards
                        {where}
                        ORDER BY {name_col}
                        """
                    ),
                    {"assembly_id": parentId},
                ).all()
                for row in rows:
                    out.append({"id": int(row.id), "code": str(row.id), "name": row.name})
                return out
        for row in rows:
            out.append({"id": row[0], "code": str(row[0]), "name": row[1]})
    elif level == "BOOTH":
        if parentId is None:
            raise ValueError("wardId is required for BOOTH")
        rows = (
            db.query(Booth.booth_id, Booth.polling_station_adr_en)
            .filter(Booth.ward_id == parentId)
            .order_by(Booth.polling_station_adr_en)
            .all()
        )
        if current.tenantId is not None:
            rows = (
                db.query(Booth.booth_id, Booth.polling_station_adr_en)
                .filter(Booth.ward_id == parentId, Booth.tenant_id == current.tenantId)
                .order_by(Booth.polling_station_adr_en)
                .all()
            )
        if not rows:
            booth_cols = _get_table_columns(db, "public", "booths")
            id_col = "booth_id" if "booth_id" in booth_cols else ("booth_no" if "booth_no" in booth_cols else ("id" if "id" in booth_cols else None))
            name_col = "polling_station_adr_en" if "polling_station_adr_en" in booth_cols else ("booth_add_en" if "booth_add_en" in booth_cols else ("name_en" if "name_en" in booth_cols else None))
            ward_ref = "ward_id" if "ward_id" in booth_cols else ("ward_no" if "ward_no" in booth_cols else None)
            if id_col and name_col:
                where = f"WHERE {ward_ref} = :ward_id" if ward_ref else ""
                rows = db.execute(
                    text(
                        f"""
                        SELECT {id_col} AS id, {name_col} AS name
                        FROM public.booths
                        {where}
                        ORDER BY {name_col}
                        """
                    ),
                    {"ward_id": parentId},
                ).all()
                for row in rows:
                    out.append({"id": int(row.id), "code": str(row.id), "name": row.name})
                return out
        for row in rows:
            out.append({"id": row[0], "code": str(row[0]), "name": row[1]})
    else:
        raise ValueError("Invalid level")

    return out


@app.get(f"{CONTEXT_PATH}/api/volunteers/stats")
def volunteer_stats(level: Optional[str] = None, id: Optional[int] = None, db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    if level is not None and id is None:
        raise ValueError("assignmentId is required when assignmentType is provided")

    q = (
        db.query(
            func.count(User.id),
            func.coalesce(func.sum(case((User.blocked.is_(False), 1), else_=0)), 0),
            func.coalesce(func.sum(case((User.blocked.is_(True), 1), else_=0)), 0),
        )
        .join(Tenant, User.tenant_ref == Tenant.id)
        .filter(Tenant.tenant_id == current.tenantId, User.role == "USER")
    )

    if level is not None:
        q = q.filter(User.assignment_type == level)
    if id is not None:
        q = q.filter(User.assignment_id == id)

    total, active, inactive = q.one()
    total = int(total or 0)
    active = int(active or 0)
    inactive = int(inactive or 0)
    pending = 0

    def pct(v: int, t: int) -> int:
        return round((v * 100.0) / t) if t else 0

    return {
        "total": total,
        "active": {"count": active, "percentage": pct(active, total)},
        "pending": {"count": pending, "percentage": pct(pending, total)},
        "inactive": {"count": inactive, "percentage": pct(inactive, total)},
        "lastUpdated": datetime.now(timezone.utc).isoformat(),
    }


@app.get(f"{CONTEXT_PATH}/api/voters/stats/gender")
def voter_gender_stats(db: Session = Depends(get_db), current: JwtUserDetails = Depends(get_current_user)):
    total = db.query(func.count(Voter.voter_id)).filter(Voter.tenant_id == current.tenantId).scalar() or 0

    rows = db.query(Voter.gender, func.count(Voter.voter_id)).filter(Voter.tenant_id == current.tenantId).group_by(Voter.gender).all()
    male = female = other = 0

    for gender, count in rows:
        g = (gender or "").upper()
        if g == "M":
            male = int(count)
        elif g == "F":
            female = int(count)
        else:
            other += int(count)

    def pct(v: int, t: int) -> float:
        return round((v * 10000.0 / t), 0) / 100 if t else 0.0

    return api_success(
        "Voter gender statistics fetched",
        {
            "totalVoters": int(total),
            "maleCount": male,
            "malePercentage": pct(male, int(total)),
            "femaleCount": female,
            "femalePercentage": pct(female, int(total)),
            "otherCount": other,
            "otherPercentage": pct(other, int(total)),
        },
    )


def _build_voter_map(v: Voter) -> Dict[str, Any]:
    return {
        "voterId": v.voter_id,
        "srNo": v.sr_no,
        "epicNo": v.epic_no,
        "firstMiddleNameEn": v.first_middle_name_en,
        "lastNameEn": v.last_name_en,
        "firstMiddleNameLocal": v.first_middle_name_local,
        "lastNameLocal": v.last_name_local,
        "relationType": v.relation_type,
        "relationFirstMiddleNameEn": v.relation_first_middle_name_en,
        "relationLastNameEn": v.relation_last_name_en,
        "relationFirstMiddleNameLocal": v.relation_first_middle_name_local,
        "relationLastNameLocal": v.relation_last_name_local,
        "houseNoEn": v.house_no_en,
        "houseNoLocal": v.house_no_local,
        "gender": v.gender,
        "age": v.age,
        "dob": v.dob.isoformat() if v.dob else None,
        "mobile": v.mobile,
        "addressEn": v.address_en,
        "addressLocal": v.address_local,
        "status": v.status,
        "community": v.community,
        "caste": v.caste,
        "residenceType": v.residence_type,
        "civicIssue": v.civic_issue,
        "motherTongue": v.mother_tongue,
        "team": v.team,
        "ownership": v.ownership,
        "education": v.education,
        "natureOfVoter": v.nature_of_voter,
        "latitude": v.latitude,
        "longitude": v.longitude,
    }


def _get_table_columns(db: Session, schema: str, table: str) -> set[str]:
    rows = db.execute(
        text(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = :schema AND table_name = :table
            """
        ),
        {"schema": schema, "table": table},
    ).all()
    return {r[0] for r in rows}


def _get_public_voter_column_map(voter_cols: set[str]) -> Dict[str, str]:
    field_map: Dict[str, str] = {}
    if "mobile" in voter_cols:
        field_map["mobile"] = "mobile"
    if "gender" in voter_cols:
        field_map["gender"] = "gender"
    if "age" in voter_cols:
        field_map["age"] = "age"
    if "house" in voter_cols:
        field_map["houseNoEn"] = "house"
        field_map["houseNoLocal"] = "house"
    if "name_en" in voter_cols:
        field_map["firstMiddleNameEn"] = "name_en"
    if "name_kannada" in voter_cols:
        field_map["firstMiddleNameLocal"] = "name_kannada"
    if "rel_eng" in voter_cols:
        field_map["relationFirstMiddleNameEn"] = "rel_eng"
    if "rel_kannada" in voter_cols:
        field_map["relationFirstMiddleNameLocal"] = "rel_kannada"
    if "rel_type" in voter_cols:
        field_map["relationType"] = "rel_type"
    return field_map


VOTER_ENRICHMENT_FIELD_MAP: Dict[str, str] = {
    "firstMiddleNameEn": "first_middle_name_en",
    "lastNameEn": "last_name_en",
    "firstMiddleNameLocal": "first_middle_name_local",
    "lastNameLocal": "last_name_local",
    "relationType": "relation_type",
    "relationFirstMiddleNameEn": "relation_first_middle_name_en",
    "relationLastNameEn": "relation_last_name_en",
    "relationFirstMiddleNameLocal": "relation_first_middle_name_local",
    "relationLastNameLocal": "relation_last_name_local",
    "houseNoEn": "house_no_en",
    "houseNoLocal": "house_no_local",
    "gender": "gender",
    "age": "age",
    "dob": "dob",
    "mobile": "mobile",
    "addressEn": "address_en",
    "addressLocal": "address_local",
    "status": "status",
    "community": "community",
    "caste": "caste",
    "residenceType": "residence_type",
    "civicIssue": "civic_issue",
    "motherTongue": "mother_tongue",
    "team": "team",
    "ownership": "ownership",
    "education": "education",
    "natureOfVoter": "nature_of_voter",
    "voterPoints": "voter_points",
    "govtSchemeTracking": "govt_scheme_tracking",
    "engagementPotential": "engagement_potential",
    "ifShifted": "if_shifted",
    "notes": "notes",
    "presentAddress": "present_address",
    "newWard": "new_ward",
    "newBoothNo": "new_booth_no",
    "newSerialNo": "new_serial_no",
    "notAvailableReason": "not_available_reason",
    "latitude": "latitude",
    "longitude": "longitude",
}
VOTER_ENRICHMENT_JSON_FIELDS = {"govtSchemeTracking"}
VOTER_ENRICHMENT_INT_FIELDS = {"age"}
VOTER_ENRICHMENT_FLOAT_FIELDS = {"latitude", "longitude"}


def _parse_updated_fields(raw: Optional[str]) -> set[str]:
    if not raw:
        return set()
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            return {str(item) for item in parsed if str(item).strip()}
    except Exception:
        return set()
    return set()


def _serialize_enrichment_value(api_field: str, value: Any) -> Any:
    if api_field in VOTER_ENRICHMENT_JSON_FIELDS:
        return json.dumps(value or [])
    if api_field in VOTER_ENRICHMENT_INT_FIELDS:
        if value in (None, ""):
            return None
        return int(value)
    if api_field in VOTER_ENRICHMENT_FLOAT_FIELDS:
        if value in (None, ""):
            return None
        return float(value)
    if isinstance(value, str):
        return normalize_optional_text(value)
    return value


def _deserialize_enrichment_value(api_field: str, value: Any) -> Any:
    if api_field in VOTER_ENRICHMENT_JSON_FIELDS:
        if value in (None, ""):
            return []
        try:
            parsed = json.loads(value)
            return parsed if isinstance(parsed, list) else []
        except Exception:
            return []
    return value


def _build_voter_enrichment_payload(enrichment: VoterEnrichment) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "epicNo": enrichment.epic,
        "wardCode": enrichment.ward_code,
        "boothNo": enrichment.booth_no,
        "updatedFields": sorted(_parse_updated_fields(enrichment.updated_fields)),
    }
    for api_field, column_name in VOTER_ENRICHMENT_FIELD_MAP.items():
        payload[api_field] = _deserialize_enrichment_value(api_field, getattr(enrichment, column_name))
    return payload


def _get_voter_enrichments(db: Session, epics: Iterable[str]) -> Dict[str, Dict[str, Any]]:
    epic_list = sorted({str(epic).strip() for epic in epics if epic and str(epic).strip()})
    if not epic_list:
        return {}
    rows = db.query(VoterEnrichment).filter(VoterEnrichment.epic.in_(epic_list)).all()
    return {row.epic: _build_voter_enrichment_payload(row) for row in rows}


def _merge_voter_payload_with_enrichment(voter_payload: Dict[str, Any], enrichment: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not enrichment:
        return voter_payload
    updated_fields = set(enrichment.get("updatedFields") or [])
    for api_field in updated_fields:
        if api_field in enrichment:
            voter_payload[api_field] = enrichment.get(api_field)
    if enrichment.get("wardCode") and not voter_payload.get("wardCode"):
        voter_payload["wardCode"] = enrichment.get("wardCode")
    if enrichment.get("boothNo") and not voter_payload.get("boothNo"):
        voter_payload["boothNo"] = enrichment.get("boothNo")
    return voter_payload


def _merge_voter_payloads_with_enrichment(db: Session, voters: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    enrichments = _get_voter_enrichments(db, [v.get("epicNo") for v in voters])
    for voter in voters:
        _merge_voter_payload_with_enrichment(voter, enrichments.get(voter.get("epicNo")))
    return voters


def _build_public_voter_result(row: Any) -> Dict[str, Any]:
    return {
        "epicNo": row.epic,
        "wardCode": str(row.ward_code) if row.ward_code is not None else None,
        "boothNo": str(row.booth_no) if row.booth_no is not None else None,
        "srNo": row.sl,
        "houseNoEn": row.house,
        "houseNoLocal": row.house,
        "firstMiddleNameEn": row.name_en,
        "lastNameEn": "",
        "firstMiddleNameLocal": row.name_kannada,
        "lastNameLocal": "",
        "relationFirstMiddleNameEn": row.rel_eng,
        "relationLastNameEn": "",
        "relationFirstMiddleNameLocal": row.rel_kannada,
        "relationLastNameLocal": "",
        "relationType": row.rel_type,
        "gender": row.gender,
        "age": row.age,
        "mobile": row.mobile,
    }


def _ensure_public_assembly_code(db: Session) -> set[str]:
    assembly_cols = _get_table_columns(db, "public", "assembly")
    if not assembly_cols:
        raise ValueError("public.assembly table not found")
    if "assembly_code" in assembly_cols:
        return assembly_cols

    assembly_no_col = "assembly_no" if "assembly_no" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else None)
    if not assembly_no_col:
        raise ValueError("public.assembly missing assembly_code and assembly_no/assembly_id")

    try:
        db.execute(text("ALTER TABLE public.assembly ADD COLUMN IF NOT EXISTS assembly_code VARCHAR(12)"))
        db.execute(
            text(
                f"""
                UPDATE public.assembly
                SET assembly_code = LPAD(CAST({assembly_no_col} AS TEXT), 12, '0')
                WHERE assembly_code IS NULL OR TRIM(assembly_code) = ''
                """
            )
        )
        db.commit()
        return _get_table_columns(db, "public", "assembly")
    except Exception:
        db.rollback()
        return assembly_cols


def _build_public_snapshot(
    assembly_code: str,
    db: Session,
    include_voters: bool = True,
    allowed_assembly_ids: Optional[set[int]] = None,
    allowed_ward_ids: Optional[set[int]] = None,
    allowed_booth_ids: Optional[set[int]] = None,
) -> Dict[str, Any]:
    allowed_assembly_ids = {v for v in (allowed_assembly_ids or set()) if v is not None}
    allowed_ward_ids = {v for v in (allowed_ward_ids or set()) if v is not None}
    allowed_booth_ids = {v for v in (allowed_booth_ids or set()) if v is not None}
    assembly_cols = _ensure_public_assembly_code(db)
    booth_cols = _get_table_columns(db, "public", "booths")
    voter_cols = _get_table_columns(db, "public", "voters")
    ward_cols = _get_table_columns(db, "public", "wards")

    requested_assembly_code = normalize_assembly_code(assembly_code)
    assembly_pk_col = "id" if "id" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else None)
    assembly_no_col = "assembly_no" if "assembly_no" in assembly_cols else ("assembly_id" if "assembly_id" in assembly_cols else None)
    assembly_name_en_col = "assembly_name_en" if "assembly_name_en" in assembly_cols else ("name_en" if "name_en" in assembly_cols else None)
    assembly_name_local_col = "assembly_name_local" if "assembly_name_local" in assembly_cols else ("name_kannada" if "name_kannada" in assembly_cols else None)
    assembly_code_expr = "assembly_code" if "assembly_code" in assembly_cols else (
        f"LPAD(CAST({assembly_no_col} AS TEXT), 12, '0')" if assembly_no_col else "NULL"
    )

    assembly_row = db.execute(
        text(
            f"""
            SELECT
                {assembly_code_expr} AS assembly_code,
                {assembly_pk_col if assembly_pk_col else 'NULL'} AS assembly_pk,
                {assembly_no_col if assembly_no_col else 'NULL'} AS assembly_no,
                {assembly_name_en_col if assembly_name_en_col else 'NULL'} AS assembly_name_en,
                {assembly_name_local_col if assembly_name_local_col else 'NULL'} AS assembly_name_local
            FROM public.assembly
            WHERE {assembly_code_expr} = :assembly_code
               OR CAST({assembly_no_col if assembly_no_col else assembly_code_expr} AS TEXT) = :assembly_no_text
            LIMIT 1
            """
        ),
        {
            "assembly_code": requested_assembly_code,
            "assembly_no_text": str(int(requested_assembly_code)),
        },
    ).first()
    if not assembly_row:
        raise ValueError(f"Assembly not found in public.assembly for assemblyCode: {assembly_code}")
    if allowed_assembly_ids:
        assembly_pk = assembly_row.assembly_pk if assembly_row.assembly_pk is not None else assembly_row.assembly_no
        if assembly_pk not in allowed_assembly_ids and assembly_row.assembly_no not in allowed_assembly_ids:
            raise HTTPException(status_code=403, detail="Access denied for requested assembly")

    booth_id_col = "id" if "id" in booth_cols else ("booth_id" if "booth_id" in booth_cols else None)
    booth_ward_id_col = "ward_id" if "ward_id" in booth_cols else None
    booth_no_col = "booth_no" if "booth_no" in booth_cols else (booth_id_col or "id")
    booth_ward_code_col = "ward_code" if "ward_code" in booth_cols else None
    booth_name_en_col = "booth_add_en" if "booth_add_en" in booth_cols else ("polling_station_adr_en" if "polling_station_adr_en" in booth_cols else None)
    booth_name_local_col = "booth_add_local" if "booth_add_local" in booth_cols else ("polling_station_adr_local" if "polling_station_adr_local" in booth_cols else None)

    if not booth_id_col or not booth_ward_id_col:
        raise ValueError("public.booths missing required columns")

    booth_filters: List[str] = []
    booth_params: Dict[str, Any] = {}
    if allowed_booth_ids:
        clause, params = _build_in_clause(booth_id_col, sorted(allowed_booth_ids), "scope_booth")
        booth_filters.append(clause)
        booth_params.update(params)
    if allowed_ward_ids:
        clause, params = _build_in_clause(booth_ward_id_col, sorted(allowed_ward_ids), "scope_ward")
        booth_filters.append(clause)
        booth_params.update(params)
    booth_where = f"WHERE {' AND '.join(booth_filters)}" if booth_filters else ""
    booth_rows = db.execute(
        text(
            f"""
            SELECT
                {booth_id_col} AS booth_id,
                {booth_ward_id_col} AS ward_id,
                {booth_no_col} AS booth_no,
                {booth_ward_code_col if booth_ward_code_col else 'NULL'} AS ward_code,
                {booth_name_en_col if booth_name_en_col else 'NULL'} AS booth_name_en,
                {booth_name_local_col if booth_name_local_col else 'NULL'} AS booth_name_local
            FROM public.booths
            {booth_where}
            ORDER BY {booth_ward_id_col}, {booth_no_col}
            """
        ),
        booth_params,
    ).all()

    ward_id_col = "id" if "id" in ward_cols else ("ward_id" if "ward_id" in ward_cols else None)
    ward_code_col = "ward_code" if "ward_code" in ward_cols else ("ward_no" if "ward_no" in ward_cols else None)
    ward_name_en_col = "ward_name_en" if "ward_name_en" in ward_cols else ("name_en" if "name_en" in ward_cols else None)
    ward_name_local_col = "ward_name_local" if "ward_name_local" in ward_cols else ("name_kannada" if "name_kannada" in ward_cols else None)
    ward_assembly_id_col = "assembly_id" if "assembly_id" in ward_cols else None
    ward_assembly_no_col = "assembly_no" if "assembly_no" in ward_cols else None
    ward_assembly_code_col = "assembly_code" if "assembly_code" in ward_cols else None

    ward_map: Dict[Any, Dict[str, Any]] = {}
    if ward_id_col:
        ward_filters: List[str] = []
        ward_params: Dict[str, Any] = {}
        if ward_assembly_id_col and assembly_row.assembly_pk is not None:
            ward_filters.append(f"{ward_assembly_id_col} = :assembly_pk")
            ward_params["assembly_pk"] = assembly_row.assembly_pk
        elif ward_assembly_no_col and assembly_row.assembly_no is not None:
            ward_filters.append(f"{ward_assembly_no_col} = :assembly_no")
            ward_params["assembly_no"] = assembly_row.assembly_no
        elif ward_assembly_code_col:
            ward_filters.append(f"{ward_assembly_code_col} = :assembly_code")
            ward_params["assembly_code"] = assembly_row.assembly_code or requested_assembly_code
        if allowed_ward_ids:
            clause, params = _build_in_clause(ward_id_col, sorted(allowed_ward_ids), "scope_ward")
            ward_filters.append(clause)
            ward_params.update(params)

        ward_where = f"WHERE {' AND '.join(ward_filters)}" if ward_filters else ""

        ward_rows = db.execute(
            text(
                f"""
                SELECT
                    {ward_id_col} AS ward_id,
                    {ward_code_col if ward_code_col else 'NULL'} AS ward_code,
                    {ward_name_en_col if ward_name_en_col else 'NULL'} AS ward_name_en,
                    {ward_name_local_col if ward_name_local_col else 'NULL'} AS ward_name_local
                FROM public.wards
                {ward_where}
                """
            ),
            ward_params,
        ).all()
        for w in ward_rows:
            ward_map[w.ward_id] = {
                "wardId": w.ward_id,
                "wardCode": str(w.ward_code) if w.ward_code is not None else None,
                "wardNameEn": w.ward_name_en or (f"Ward {w.ward_id}"),
                "wardNameLocal": w.ward_name_local,
                "booths": [],
            }
    allowed_ward_ids = set(ward_map.keys())

    voter_id_col = "voter_id" if "voter_id" in voter_cols else ("id" if "id" in voter_cols else None)
    voter_sr_col = "sl" if "sl" in voter_cols else ("sr_no" if "sr_no" in voter_cols else None)
    voter_epic_col = "epic" if "epic" in voter_cols else ("epic_no" if "epic_no" in voter_cols else None)
    voter_name_local_col = "name_kannada" if "name_kannada" in voter_cols else ("first_middle_name_local" if "first_middle_name_local" in voter_cols else None)
    voter_house_col = "house" if "house" in voter_cols else ("house_no_en" if "house_no_en" in voter_cols else None)
    voter_gender_col = "gender" if "gender" in voter_cols else None
    voter_booth_no_col = "booth_no" if "booth_no" in voter_cols else ("booth_id" if "booth_id" in voter_cols else None)
    voter_ward_code_col = "ward_code" if "ward_code" in voter_cols else None

    if not voter_booth_no_col:
        raise ValueError("public.voters missing booth mapping column")

    voters_by_key: Dict[tuple, List[Dict[str, Any]]] = {}
    counts_by_key: Dict[tuple, Dict[str, int]] = {}
    relevant_booths = [
        b
        for b in booth_rows
        if (not allowed_ward_ids or b.ward_id in allowed_ward_ids)
        and (not allowed_booth_ids or b.booth_id in allowed_booth_ids)
    ]
    allowed_ward_codes = sorted(
        {
            str(w.get("wardCode"))
            for w in ward_map.values()
            if w.get("wardCode") is not None and str(w.get("wardCode")).strip() != ""
        }
    )
    allowed_booth_nos = sorted({str(b.booth_no) for b in relevant_booths if b.booth_no is not None})

    voter_where_parts: List[str] = []
    voter_where_params: Dict[str, Any] = {}
    if voter_ward_code_col and allowed_ward_codes:
        clause, params = _build_in_clause(f"CAST({voter_ward_code_col} AS TEXT)", allowed_ward_codes, "ward_code")
        voter_where_parts.append(clause)
        voter_where_params.update(params)
    if allowed_booth_nos:
        clause, params = _build_in_clause(f"CAST({voter_booth_no_col} AS TEXT)", allowed_booth_nos, "booth_no")
        voter_where_parts.append(clause)
        voter_where_params.update(params)

    if not voter_where_parts:
        return {
            "assembly": {
                "assemblyId": assembly_row.assembly_pk,
                "assemblyCode": assembly_row.assembly_code or requested_assembly_code,
                "assemblyNameEn": assembly_row.assembly_name_en,
                "assemblyNameLocal": assembly_row.assembly_name_local,
                "wards": sorted(list(ward_map.values()), key=lambda w: (w.get("wardCode") or str(w["wardId"]))),
            }
        }

    voter_where_sql = " WHERE " + " AND ".join(voter_where_parts)

    if include_voters:
        voter_id_expr = f"{voter_id_col}" if voter_id_col else "ROW_NUMBER() OVER ()"
        voter_rows = db.execute(
            text(
                f"""
                SELECT
                    {voter_id_expr} AS voter_id,
                    {voter_sr_col if voter_sr_col else 'NULL'} AS sr_no,
                    {voter_epic_col if voter_epic_col else 'NULL'} AS epic_no,
                    {voter_name_en_col if voter_name_en_col else 'NULL'} AS name_en,
                    {voter_name_local_col if voter_name_local_col else 'NULL'} AS name_local,
                    {voter_house_col if voter_house_col else 'NULL'} AS house_no_en,
                    {voter_gender_col if voter_gender_col else 'NULL'} AS gender,
                    {voter_booth_no_col} AS booth_no,
                    {voter_ward_code_col if voter_ward_code_col else 'NULL'} AS ward_code
                FROM public.voters
                {voter_where_sql}
                """
            ),
            voter_where_params,
        ).all()

        for v in voter_rows:
            key = (str(v.ward_code) if v.ward_code is not None else None, str(v.booth_no))
            voters_by_key.setdefault(key, []).append(
                {
                    "voterId": int(v.voter_id),
                    "srNo": v.sr_no,
                    "epicNo": v.epic_no,
                    "firstMiddleNameEn": v.name_en,
                    "lastNameEn": "",
                    "firstMiddleNameLocal": v.name_local,
                    "lastNameLocal": "",
                    "houseNoEn": str(v.house_no_en) if v.house_no_en is not None else None,
                    "houseNoLocal": None,
                    "gender": v.gender,
                    "age": None,
                    "dob": None,
                    "mobile": None,
                    "addressEn": None,
                    "addressLocal": None,
                    "status": None,
                    "community": None,
                    "caste": None,
                    "residenceType": None,
                    "civicIssue": None,
                    "motherTongue": None,
                    "team": None,
                    "ownership": None,
                    "education": None,
                    "natureOfVoter": None,
                    "latitude": None,
                    "longitude": None,
                }
            )
        for key, rows in voters_by_key.items():
            _merge_voter_payloads_with_enrichment(db, rows)
            male = sum(1 for r in rows if (r.get("gender") or "").upper().startswith("M"))
            female = sum(1 for r in rows if (r.get("gender") or "").upper().startswith("F"))
            counts_by_key[key] = {"total": len(rows), "male": male, "female": female}
    else:
        if voter_gender_col:
            counts_rows = db.execute(
                text(
                    f"""
                    SELECT
                    {voter_ward_code_col if voter_ward_code_col else 'NULL'} AS ward_code,
                    {voter_booth_no_col} AS booth_no,
                    COUNT(*)::int AS total_count,
                    SUM(CASE WHEN UPPER(COALESCE({voter_gender_col}, '')) LIKE 'M%' THEN 1 ELSE 0 END)::int AS male_count,
                    SUM(CASE WHEN UPPER(COALESCE({voter_gender_col}, '')) LIKE 'F%' THEN 1 ELSE 0 END)::int AS female_count
                    FROM public.voters
                    {voter_where_sql}
                    GROUP BY {voter_ward_code_col if voter_ward_code_col else 'NULL'}, {voter_booth_no_col}
                    """
                ),
                voter_where_params,
            ).all()
        else:
            counts_rows = db.execute(
                text(
                    f"""
                    SELECT
                    {voter_ward_code_col if voter_ward_code_col else 'NULL'} AS ward_code,
                    {voter_booth_no_col} AS booth_no,
                    COUNT(*)::int AS total_count,
                    0::int AS male_count,
                    0::int AS female_count
                    FROM public.voters
                    {voter_where_sql}
                    GROUP BY {voter_ward_code_col if voter_ward_code_col else 'NULL'}, {voter_booth_no_col}
                    """
                ),
                voter_where_params,
            ).all()
        for row in counts_rows:
            key = (str(row.ward_code) if row.ward_code is not None else None, str(row.booth_no))
            counts_by_key[key] = {
                "total": int(row.total_count or 0),
                "male": int(row.male_count or 0),
                "female": int(row.female_count or 0),
            }

    for b in booth_rows:
        ward_id = b.ward_id
        if allowed_ward_ids and ward_id not in allowed_ward_ids:
            continue
        if allowed_booth_ids and b.booth_id not in allowed_booth_ids:
            continue
        if ward_id not in ward_map:
            ward_map[ward_id] = {
                "wardId": ward_id,
                "wardCode": str(b.ward_code) if b.ward_code is not None else None,
                "wardNameEn": f"Ward {ward_id}",
                "wardNameLocal": None,
                "booths": [],
            }

        key = (str(b.ward_code) if b.ward_code is not None else None, str(b.booth_no))
        booth_entry = {
            "boothId": int(b.booth_id),
            "boothNo": int(b.booth_no) if b.booth_no is not None else None,
            "boothNameEn": b.booth_name_en or (f"Booth {b.booth_no}"),
            "boothNameLocal": b.booth_name_local,
            "voterStats": counts_by_key.get(key, {"total": 0, "male": 0, "female": 0}),
        }
        if include_voters:
            booth_entry["voters"] = voters_by_key.get(key, [])
        else:
            booth_entry["voters"] = []
        ward_map[ward_id]["booths"].append(booth_entry)

    return {
        "assembly": {
            "assemblyId": assembly_row.assembly_pk,
            "assemblyCode": assembly_row.assembly_code or requested_assembly_code,
            "assemblyNameEn": assembly_row.assembly_name_en,
            "assemblyNameLocal": assembly_row.assembly_name_local,
            "wards": sorted(
                [
                    {**ward, "booths": sorted(ward.get("booths", []), key=lambda b: (b.get("boothNo") or 0, b.get("boothId") or 0))}
                    for ward in ward_map.values()
                ],
                key=lambda w: (w.get("wardCode") or str(w["wardId"])),
            ),
            }
        }


def _upload_and_save_snapshot(db: Session, data: Dict[str, Any], key: str, tenant_id: str, assembly_code: str, ward_code: Optional[str], booth_id: Optional[int], level: str) -> None:
    s3_url = s3_upload_bytes(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"), "application/json", key)

    existing = (
        db.query(VoterSnapshot)
        .filter(
            VoterSnapshot.tenant_id == tenant_id,
            VoterSnapshot.assembly_code == assembly_code,
            VoterSnapshot.ward_code == ward_code,
            VoterSnapshot.booth_id == booth_id,
            VoterSnapshot.snapshot_level == level,
        )
        .first()
    )
    if existing:
        existing.s3_url = s3_url
        existing.updated_at = datetime.utcnow()
        existing.version = (existing.version or 1) + 1
        db.add(existing)
    else:
        db.add(
            VoterSnapshot(
                tenant_id=tenant_id,
                assembly_code=assembly_code,
                ward_code=ward_code,
                booth_id=booth_id,
                snapshot_level=level,
                s3_url=s3_url,
                version=1,
                created_at=datetime.utcnow(),
                updated_at=datetime.utcnow(),
            )
        )


def _build_assembly_json(
    db: Session,
    assembly: Assembly,
    wards: List[Ward],
    include_voters: bool,
    tenant_id: Optional[str],
) -> Dict[str, Any]:
    assembly_map: Dict[str, Any] = {
        "assemblyId": assembly.assembly_id,
        "assemblyNameEn": assembly.assembly_name_en,
        "assemblyNameLocal": assembly.assembly_name_local,
    }

    ward_list: List[Dict[str, Any]] = []
    for ward in wards:
        ward_map: Dict[str, Any] = {
            "wardId": ward.ward_id,
            "wardNameEn": ward.ward_name_en,
            "wardNameLocal": ward.ward_name_local,
        }

        booths_q = db.query(Booth).filter(Booth.ward_id == ward.ward_id)
        if tenant_id is not None:
            booths_q = booths_q.filter(Booth.tenant_id == tenant_id)
        booths = booths_q.all()
        booth_list: List[Dict[str, Any]] = []
        for booth in booths:
            booth_map: Dict[str, Any] = {
                "boothId": booth.booth_id,
                "boothNameEn": booth.polling_station_adr_en,
                "boothNameLocal": booth.polling_station_adr_local,
            }
            if include_voters:
                voters_q = db.query(Voter).filter(Voter.booth_id == booth.booth_id)
                if tenant_id is not None:
                    voters_q = voters_q.filter(Voter.tenant_id == tenant_id)
                voters = voters_q.all()
                booth_map["voters"] = [_build_voter_map(v) for v in voters]
            booth_list.append(booth_map)

        ward_map["booths"] = booth_list
        ward_list.append(ward_map)

    assembly_map["wards"] = ward_list
    return {"assembly": assembly_map}


def generate_snapshots(db: Session, assembly_id: int, tenant_id: str) -> None:
    assembly = db.query(Assembly).filter(Assembly.assembly_id == assembly_id).first()
    if not assembly:
        raise ValueError(f"Assembly not found: {assembly_id}")

    assembly_code = normalize_assembly_code(assembly.assembly_code)

    wards = db.query(Ward).filter(Ward.tenant_id == tenant_id, Ward.assembly_id == assembly.assembly_id).all()
    if wards:
        assembly_json = _build_assembly_json(db, assembly, wards, True, tenant_id)
        _upload_and_save_snapshot(db, assembly_json, f"snapshots/{tenant_id}/{assembly_code}/assembly.json", tenant_id, assembly_code, None, None, "ASSEMBLY")

    for ward in wards:
        ward_json = _build_assembly_json(db, assembly, [ward], True, tenant_id)
        _upload_and_save_snapshot(
            db,
            ward_json,
            f"snapshots/{tenant_id}/{assembly_code}/wards/ward_{ward.ward_code}.json",
            tenant_id,
            assembly_code,
            ward.ward_code,
            None,
            "WARD",
        )

    booths = (
        db.query(Booth)
        .join(Ward, Booth.ward_id == Ward.ward_id)
        .join(Assembly, Ward.assembly_id == Assembly.assembly_id)
        .filter(Assembly.assembly_code == assembly_code, Booth.tenant_id == tenant_id)
        .all()
    )

    for booth in booths:
        voters = db.query(Voter).filter(Voter.tenant_id == tenant_id, Voter.booth_id == booth.booth_id).all()
        if not voters:
            continue

        ward = db.query(Ward).filter(Ward.ward_id == booth.ward_id).first()
        if not ward:
            continue

        booth_json = {
            "assembly": {
                "assemblyId": assembly.assembly_id,
                "assemblyNameEn": assembly.assembly_name_en,
                "assemblyNameLocal": assembly.assembly_name_local,
                "wards": [
                    {
                        "wardId": ward.ward_id,
                        "wardNameEn": ward.ward_name_en,
                        "wardNameLocal": ward.ward_name_local,
                        "booths": [
                            {
                                "boothId": booth.booth_id,
                                "boothNameEn": booth.polling_station_adr_en,
                                "boothNameLocal": booth.polling_station_adr_local,
                                "voters": [_build_voter_map(v) for v in voters],
                            }
                        ],
                    }
                ],
            }
        }

        _upload_and_save_snapshot(
            db,
            booth_json,
            f"snapshots/{tenant_id}/{assembly_code}/booths/booth_{booth.booth_id}.json",
            tenant_id,
            assembly_code,
            ward.ward_code,
            booth.booth_id,
            "BOOTH",
        )


@app.post(f"{CONTEXT_PATH}/api/excel/upload")
def upload_excel(file: UploadFile = File(...), db: Session = Depends(get_db), current: JwtUserDetails = Depends(require_roles("ADMIN"))):
    try:
        wb = load_workbook(file.file, data_only=True)

        def sheet(name: str):
            if name not in wb.sheetnames:
                raise ValueError(f"Missing required sheet: {name}")
            return wb[name]

        def headers(ws) -> Dict[str, int]:
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            return {str(v).strip().upper(): i for i, v in enumerate(first_row) if v is not None}

        ws_assembly = sheet("ASSEMBLY")
        ws_ward = sheet("WARD")
        ws_booth = sheet("BOOTH")
        ws_data = sheet("DATA")

        h_assembly = headers(ws_assembly)
        h_ward = headers(ws_ward)
        h_booth = headers(ws_booth)
        h_data = headers(ws_data)

        for col in ["ASSEMBLY_NO", "ASSEMBLY_NAME_EN", "ASSEMBLY_NAME_LOCAL"]:
            if col not in h_assembly:
                raise ValueError(f"Required column missing in sheet 'ASSEMBLY': {col}")

        arow = next(ws_assembly.iter_rows(min_row=2, max_row=2, values_only=True))
        assembly_no = int(arow[h_assembly["ASSEMBLY_NO"]])
        assembly_code = normalize_assembly_code(assembly_no)
        assembly_name_en = arow[h_assembly["ASSEMBLY_NAME_EN"]]
        assembly_name_local = arow[h_assembly["ASSEMBLY_NAME_LOCAL"]]

        assembly = db.query(Assembly).filter(Assembly.assembly_code == assembly_code, Assembly.tenant_id == current.tenantId).first()
        if not assembly:
            assembly = Assembly(
                assembly_id=assembly_no,
                assembly_code=assembly_code,
                assembly_name_en=assembly_name_en or f"Assembly {assembly_no}",
                assembly_name_local=assembly_name_local,
                tenant_id=current.tenantId,
            )
            db.add(assembly)
            db.flush()

        # Wards
        for row in ws_ward.iter_rows(min_row=2, values_only=True):
            ward_code = row[h_ward.get("WARD_CODE", -1)]
            if ward_code is None or str(ward_code).strip() == "":
                continue
            exists = db.query(Ward).filter(Ward.ward_code == str(ward_code), Ward.tenant_id == current.tenantId).first()
            if exists:
                continue

            assembly_no_row = row[h_ward.get("ASSEMBLY_NO", -1)]
            if assembly_no_row is None:
                raise ValueError(f"Assembly number missing for ward: {ward_code}")

            assembly_ref = db.query(Assembly).filter(Assembly.assembly_id == int(assembly_no_row)).first()
            if not assembly_ref:
                assembly_ref = Assembly(
                    assembly_id=int(assembly_no_row),
                    assembly_code=normalize_assembly_code(int(assembly_no_row)),
                    assembly_name_en=f"Assembly {int(assembly_no_row)}",
                    tenant_id=current.tenantId,
                )
                db.add(assembly_ref)
                db.flush()

            db.add(
                Ward(
                    ward_code=str(ward_code),
                    ward_name_en=row[h_ward.get("WARD_NAME_EN", -1)] or f"Ward {ward_code}",
                    ward_name_local=row[h_ward.get("WARD_NAME_LOCAL", -1)],
                    tenant_id=current.tenantId,
                    assembly_id=assembly_ref.assembly_id,
                )
            )

        db.flush()

        # Booths
        for row in ws_booth.iter_rows(min_row=2, values_only=True):
            booth_no = row[h_booth.get("BOOTH_NO", -1)]
            ward_code = row[h_booth.get("WARD_CODE", -1)]
            if booth_no is None or ward_code is None:
                continue

            ward = db.query(Ward).filter(Ward.ward_code == str(ward_code), Ward.tenant_id == current.tenantId).first()
            if not ward:
                raise ValueError(f"Ward not found for booth {booth_no} (wardCode={ward_code})")

            booth = db.query(Booth).filter(Booth.booth_id == int(booth_no)).first()
            if not booth:
                db.add(
                    Booth(
                        booth_id=int(booth_no),
                        ward_id=ward.ward_id,
                        polling_station_adr_en=row[h_booth.get("BOOTH_ADD_EN", -1)],
                        polling_station_adr_local=row[h_booth.get("BOOTH_ADD_LOCAL", -1)],
                        tenant_id=current.tenantId,
                    )
                )

        db.flush()

        # Voters
        seen = set()
        incoming = []
        for row in ws_data.iter_rows(min_row=2, values_only=True):
            sr_no = row[h_data.get("SL", -1)] if "SL" in h_data else None
            booth_no = row[h_data.get("BOOTH_NO", -1)] if "BOOTH_NO" in h_data else None
            epic = row[h_data.get("EPIC", -1)] if "EPIC" in h_data else None
            if sr_no is None or booth_no is None or epic is None:
                continue
            epic = str(epic)
            if epic in seen:
                continue
            seen.add(epic)
            incoming.append((row, int(booth_no), epic, int(sr_no)))

        epic_set = {epic for _, _, epic, _ in incoming}
        existing_epics = set(
            r[0]
            for r in db.query(Voter.epic_no)
            .filter(Voter.tenant_id == current.tenantId, Voter.epic_no.in_(list(epic_set)))
            .all()
        )

        saved = 0
        for row, booth_no, epic, sr_no in incoming:
            if epic in existing_epics:
                continue

            booth = db.query(Booth).filter(Booth.booth_id == booth_no).first()
            if not booth:
                raise ValueError(f"Booth not found: {booth_no}")

            mobile = str(row[h_data["MOBILE"]]) if "MOBILE" in h_data and row[h_data["MOBILE"]] is not None else None
            if mobile is not None:
                mobile = re.sub(r"\D+", "", mobile)
                if len(mobile) != 10:
                    mobile = None

            voter = Voter(
                tenant_id=current.tenantId,
                booth_id=booth.booth_id,
                sr_no=sr_no,
                epic_no=epic,
                house_no_en=str(row[h_data["HOUSE"]]) if "HOUSE" in h_data and row[h_data["HOUSE"]] is not None else None,
                first_middle_name_en=str(row[h_data["NAME_EN"]]) if "NAME_EN" in h_data and row[h_data["NAME_EN"]] is not None else None,
                first_middle_name_local=str(row[h_data["NAME_KANNADA"]]) if "NAME_KANNADA" in h_data and row[h_data["NAME_KANNADA"]] is not None else None,
                gender=str(row[h_data["GENDER"]]) if "GENDER" in h_data and row[h_data["GENDER"]] is not None else None,
                age=int(row[h_data["AGE"]]) if "AGE" in h_data and row[h_data["AGE"]] is not None else None,
                relation_type=str(row[h_data["REL_TYPE"]]) if "REL_TYPE" in h_data and row[h_data["REL_TYPE"]] is not None else None,
                relation_first_middle_name_en=str(row[h_data["REL_ENG"]]) if "REL_ENG" in h_data and row[h_data["REL_ENG"]] is not None else None,
                relation_first_middle_name_local=str(row[h_data["REL_KANNADA"]]) if "REL_KANNADA" in h_data and row[h_data["REL_KANNADA"]] is not None else None,
                mobile=mobile,
            )
            db.add(voter)
            saved += 1

        db.flush()
        if saved > 0:
            generate_snapshots(db, assembly_no, current.tenantId)
        db.commit()

        return api_success("Successfully uploaded", saved)
    except Exception as ex:
        db.rollback()
        return JSONResponse(status_code=500, content=api_error("Excel upload failed", str(ex)))


@app.get("/health")
def health():
    return {"ok": True}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=int(os.getenv("PORT", "8081")), reload=True)
