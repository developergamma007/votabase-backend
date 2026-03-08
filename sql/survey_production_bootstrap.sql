-- Survey_Production bootstrap for Votebase backend compatibility
-- Safe-by-default: creates missing schemas/tables only, does not alter existing public survey tables.

BEGIN;

-- 1) Schemas expected by backend
CREATE SCHEMA IF NOT EXISTS metastore;
CREATE SCHEMA IF NOT EXISTS data;
CREATE SCHEMA IF NOT EXISTS snapshot;

-- 2) Metastore tables
CREATE TABLE IF NOT EXISTS metastore.tenant (
    id SERIAL PRIMARY KEY,
    tenant_id VARCHAR(20) UNIQUE,
    name VARCHAR(255) NOT NULL,
    description TEXT,
    contact_email VARCHAR(255) UNIQUE NOT NULL,
    contact_phone VARCHAR(50),
    active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);

CREATE TABLE IF NOT EXISTS metastore.users (
    id SERIAL PRIMARY KEY,
    tenant_id INTEGER REFERENCES metastore.tenant(id),
    role VARCHAR(30) NOT NULL,
    assignment_type VARCHAR(30),
    assignment_id INTEGER,
    first_name VARCHAR(100) UNIQUE NOT NULL,
    phone VARCHAR(10) UNIQUE NOT NULL,
    profile_pic_url TEXT,
    blocked BOOLEAN DEFAULT FALSE,
    deleted BOOLEAN DEFAULT FALSE
);

-- 3) Core data tables (new Survey names, old Votebase-compatible columns)
CREATE TABLE IF NOT EXISTS data.assembly (
    assembly_id INTEGER PRIMARY KEY,
    tenant_id VARCHAR(20) NOT NULL,
    assembly_name_en VARCHAR(255),
    assembly_name_local VARCHAR(255),
    assembly_code VARCHAR(12) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS data.wards (
    ward_id INTEGER PRIMARY KEY,
    assembly_id INTEGER REFERENCES data.assembly(assembly_id),
    tenant_id VARCHAR(20) NOT NULL,
    ward_name_en VARCHAR(255),
    ward_name_local VARCHAR(255),
    ward_code VARCHAR(20) UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS data.booths (
    booth_id INTEGER PRIMARY KEY,
    ward_id INTEGER REFERENCES data.wards(ward_id),
    tenant_id VARCHAR(20) NOT NULL,
    polling_station_adr_en VARCHAR(255),
    polling_station_adr_local VARCHAR(255)
);

CREATE TABLE IF NOT EXISTS data.voters (
    voter_id BIGINT PRIMARY KEY,
    tenant_id VARCHAR(20) NOT NULL,
    booth_id INTEGER REFERENCES data.booths(booth_id),
    sr_no INTEGER,
    epic_no VARCHAR(20) UNIQUE,
    first_middle_name_en VARCHAR(150),
    last_name_en VARCHAR(100),
    first_middle_name_local VARCHAR(150),
    last_name_local VARCHAR(100),
    relation_type VARCHAR(20),
    relation_first_middle_name_en VARCHAR(150),
    relation_last_name_en VARCHAR(100),
    relation_first_middle_name_local VARCHAR(150),
    relation_last_name_local VARCHAR(100),
    house_no_en VARCHAR(50),
    house_no_local VARCHAR(50),
    gender VARCHAR(10),
    age INTEGER,
    dob TIMESTAMP,
    mobile VARCHAR(15),
    address_en VARCHAR(255),
    address_local VARCHAR(255),
    status VARCHAR(20),
    community VARCHAR(100),
    caste VARCHAR(100),
    residence_type VARCHAR(100),
    civic_issue VARCHAR(255),
    mother_tongue VARCHAR(100),
    team VARCHAR(100),
    ownership VARCHAR(20),
    education VARCHAR(20),
    nature_of_voter VARCHAR(100),
    latitude DOUBLE PRECISION,
    longitude DOUBLE PRECISION
);

-- 4) Extra legacy tables from old Votebase (commonly missed)
CREATE TABLE IF NOT EXISTS data.association (
    association_id SERIAL PRIMARY KEY,
    association_name VARCHAR(100) NOT NULL,
    booth_id INTEGER REFERENCES data.booths(booth_id),
    association_address TEXT,
    association_head_name TEXT,
    phone VARCHAR(20),
    latitude DOUBLE PRECISION,
    longitude DOUBLE PRECISION,
    tenant_id VARCHAR(20) NOT NULL
);

CREATE TABLE IF NOT EXISTS data.family (
    family_id SERIAL PRIMARY KEY,
    tenant_id VARCHAR(20) NOT NULL,
    family_name VARCHAR(30) NOT NULL,
    family_address VARCHAR(555),
    head_voter_id INTEGER,
    phone VARCHAR(15),
    points INTEGER,
    points_provided INTEGER,
    latitude DOUBLE PRECISION,
    longitude DOUBLE PRECISION,
    booth_id INTEGER REFERENCES data.booths(booth_id),
    association_id INTEGER REFERENCES data.association(association_id),
    economic_status VARCHAR(50),
    family_nature VARCHAR(50),
    deleted BOOLEAN DEFAULT FALSE
);

CREATE TABLE IF NOT EXISTS data.family_members (
    member_id SERIAL PRIMARY KEY,
    family_id INTEGER REFERENCES data.family(family_id),
    voter_id BIGINT REFERENCES data.voters(voter_id),
    is_head BOOLEAN DEFAULT FALSE
);

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.table_constraints
        WHERE table_schema='data' AND table_name='family' AND constraint_name='fk_family_head_member'
    ) THEN
        ALTER TABLE data.family
            ADD CONSTRAINT fk_family_head_member
            FOREIGN KEY (head_voter_id) REFERENCES data.family_members(member_id);
    END IF;
END $$;

CREATE TABLE IF NOT EXISTS data.voter_changelog (
    id SERIAL PRIMARY KEY,
    tenant_id VARCHAR(20) NOT NULL,
    voter_id BIGINT REFERENCES data.voters(voter_id),
    updated_by INTEGER REFERENCES metastore.users(id),
    field_name VARCHAR(100) NOT NULL,
    old_value TEXT,
    new_value TEXT,
    updated_at TIMESTAMP,
    update_latitude DOUBLE PRECISION,
    update_longitude DOUBLE PRECISION
);

CREATE TABLE IF NOT EXISTS snapshot.voter_snapshot (
    id SERIAL PRIMARY KEY,
    tenant_id VARCHAR(20) NOT NULL,
    assembly_code VARCHAR(12) NOT NULL,
    ward_code VARCHAR(20),
    booth_id INTEGER,
    s3_url TEXT NOT NULL,
    snapshot_level VARCHAR(20) NOT NULL,
    version INTEGER DEFAULT 1,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);

-- 5) Helpful indexes
CREATE INDEX IF NOT EXISTS idx_users_tenant_role ON metastore.users (tenant_id, role);
CREATE INDEX IF NOT EXISTS idx_users_first_name ON metastore.users (first_name);
CREATE INDEX IF NOT EXISTS idx_users_phone ON metastore.users (phone);

CREATE INDEX IF NOT EXISTS idx_assembly_tenant_code ON data.assembly (tenant_id, assembly_code);
CREATE INDEX IF NOT EXISTS idx_wards_assembly ON data.wards (assembly_id);
CREATE INDEX IF NOT EXISTS idx_booths_ward ON data.booths (ward_id);
CREATE INDEX IF NOT EXISTS idx_voters_booth ON data.voters (booth_id);
CREATE INDEX IF NOT EXISTS idx_voters_tenant ON data.voters (tenant_id);
CREATE INDEX IF NOT EXISTS idx_voters_epic ON data.voters (epic_no);
CREATE INDEX IF NOT EXISTS idx_family_booth ON data.family (booth_id);
CREATE INDEX IF NOT EXISTS idx_association_booth ON data.association (booth_id);

COMMIT;
