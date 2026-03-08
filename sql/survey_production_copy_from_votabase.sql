-- Copy data from old Votabase_Production into Survey_Production (run on Survey_Production DB)
-- Requires privileges to create extension/server/user mapping.

BEGIN;

CREATE EXTENSION IF NOT EXISTS postgres_fdw;

-- 1) Foreign server to old Votabase_Production
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_foreign_server WHERE srvname = 'old_votabase_srv') THEN
        CREATE SERVER old_votabase_srv
        FOREIGN DATA WRAPPER postgres_fdw
        OPTIONS (host '65.0.75.0', port '5432', dbname 'postgres');
    END IF;
END $$;

-- Replace credentials if required.
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1
        FROM pg_user_mappings um
        JOIN pg_foreign_server s ON s.oid = um.srvid
        WHERE s.srvname = 'old_votabase_srv'
    ) THEN
        CREATE USER MAPPING FOR CURRENT_USER
        SERVER old_votabase_srv
        OPTIONS (user 'votabank_admin', password 'Tony$stark!1');
    END IF;
END $$;

CREATE SCHEMA IF NOT EXISTS old_data;
CREATE SCHEMA IF NOT EXISTS old_metastore;
CREATE SCHEMA IF NOT EXISTS old_snapshot;

-- 2) Import only required old tables
IMPORT FOREIGN SCHEMA data
LIMIT TO (assembly_details, ward_details, booth_details, voter_details, association, family, family_members, voter_changelog)
FROM SERVER old_votabase_srv INTO old_data;

IMPORT FOREIGN SCHEMA metastore
LIMIT TO (tenant, users)
FROM SERVER old_votabase_srv INTO old_metastore;

IMPORT FOREIGN SCHEMA snapshot
LIMIT TO (voter_snapshot)
FROM SERVER old_votabase_srv INTO old_snapshot;

-- 3) Upsert/copy metastore
INSERT INTO metastore.tenant (id, tenant_id, name, description, contact_email, contact_phone, active, created_at, updated_at)
SELECT id, tenant_id, name, description, contact_email, contact_phone, active, created_at, updated_at
FROM old_metastore.tenant
ON CONFLICT (id) DO NOTHING;

INSERT INTO metastore.users (id, tenant_id, role, assignment_type, assignment_id, first_name, phone, profile_pic_url, blocked, deleted)
SELECT id, tenant_id, role, assignment_type, assignment_id, first_name, phone, profile_pic_url, blocked, deleted
FROM old_metastore.users
ON CONFLICT (id) DO NOTHING;

-- Keep sequences aligned
SELECT setval('metastore.tenant_id_seq', COALESCE((SELECT MAX(id) FROM metastore.tenant), 1), true);
SELECT setval('metastore.users_id_seq', COALESCE((SELECT MAX(id) FROM metastore.users), 1), true);

-- 4) Upsert/copy core 4 tables (old names -> new names)
INSERT INTO data.assembly (assembly_id, tenant_id, assembly_name_en, assembly_name_local, assembly_code)
SELECT assembly_id, tenant_id, assembly_name_en, assembly_name_local, assembly_code
FROM old_data.assembly_details
ON CONFLICT (assembly_id) DO NOTHING;

INSERT INTO data.wards (ward_id, assembly_id, tenant_id, ward_name_en, ward_name_local, ward_code)
SELECT ward_id, assembly_id, tenant_id, ward_name_en, ward_name_local, ward_code
FROM old_data.ward_details
ON CONFLICT (ward_id) DO NOTHING;

INSERT INTO data.booths (booth_id, ward_id, tenant_id, polling_station_adr_en, polling_station_adr_local)
SELECT booth_id, ward_id, tenant_id, polling_station_adr_en, polling_station_adr_local
FROM old_data.booth_details
ON CONFLICT (booth_id) DO NOTHING;

INSERT INTO data.voters (
    voter_id, tenant_id, booth_id, sr_no, epic_no,
    first_middle_name_en, last_name_en, first_middle_name_local, last_name_local,
    relation_type, relation_first_middle_name_en, relation_last_name_en,
    relation_first_middle_name_local, relation_last_name_local,
    house_no_en, house_no_local, gender, age, dob, mobile,
    address_en, address_local, status, community, caste, residence_type,
    civic_issue, mother_tongue, team, ownership, education, nature_of_voter,
    latitude, longitude
)
SELECT
    voter_id, tenant_id, booth_id, sr_no, epic_no,
    first_middle_name_en, last_name_en, first_middle_name_local, last_name_local,
    relation_type, relation_first_middle_name_en, relation_last_name_en,
    relation_first_middle_name_local, relation_last_name_local,
    house_no_en, house_no_local, gender, age, dob, mobile,
    address_en, address_local, status, community, caste, residence_type,
    civic_issue, mother_tongue, team, ownership, education, nature_of_voter,
    latitude, longitude
FROM old_data.voter_details
ON CONFLICT (voter_id) DO NOTHING;

-- 5) Extra tables copy
INSERT INTO data.association (
    association_id, association_name, booth_id, association_address,
    association_head_name, phone, latitude, longitude, tenant_id
)
SELECT
    association_id, association_name, booth_id, association_address,
    association_head_name, phone, latitude, longitude, tenant_id
FROM old_data.association
ON CONFLICT (association_id) DO NOTHING;

INSERT INTO data.family (
    family_id, tenant_id, family_name, family_address, head_voter_id,
    phone, points, points_provided, latitude, longitude, booth_id,
    association_id, economic_status, family_nature, deleted
)
SELECT
    family_id, tenant_id, family_name, family_address, head_voter_id,
    phone, points, points_provided, latitude, longitude, booth_id,
    association_id, economic_status, family_nature, deleted
FROM old_data.family
ON CONFLICT (family_id) DO NOTHING;

INSERT INTO data.family_members (member_id, family_id, voter_id, is_head)
SELECT member_id, family_id, voter_id, is_head
FROM old_data.family_members
ON CONFLICT (member_id) DO NOTHING;

INSERT INTO data.voter_changelog (
    id, tenant_id, voter_id, updated_by, field_name, old_value,
    new_value, updated_at, update_latitude, update_longitude
)
SELECT
    id, tenant_id, voter_id, updated_by, field_name, old_value,
    new_value, updated_at, update_latitude, update_longitude
FROM old_data.voter_changelog
ON CONFLICT (id) DO NOTHING;

INSERT INTO snapshot.voter_snapshot (
    id, tenant_id, assembly_code, ward_code, booth_id,
    s3_url, snapshot_level, version, created_at, updated_at
)
SELECT
    id, tenant_id, assembly_code, ward_code, booth_id,
    s3_url, snapshot_level, version, created_at, updated_at
FROM old_snapshot.voter_snapshot
ON CONFLICT (id) DO NOTHING;

-- Keep sequences aligned
SELECT setval('data.association_association_id_seq', COALESCE((SELECT MAX(association_id) FROM data.association), 1), true);
SELECT setval('data.family_family_id_seq', COALESCE((SELECT MAX(family_id) FROM data.family), 1), true);
SELECT setval('data.family_members_member_id_seq', COALESCE((SELECT MAX(member_id) FROM data.family_members), 1), true);
SELECT setval('data.voter_changelog_id_seq', COALESCE((SELECT MAX(id) FROM data.voter_changelog), 1), true);
SELECT setval('snapshot.voter_snapshot_id_seq', COALESCE((SELECT MAX(id) FROM snapshot.voter_snapshot), 1), true);

COMMIT;
