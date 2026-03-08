-- Post-migration checks for Survey_Production

-- 1) Required tables exist
SELECT table_schema, table_name
FROM information_schema.tables
WHERE (table_schema, table_name) IN (
    ('metastore','tenant'), ('metastore','users'),
    ('data','assembly'), ('data','wards'), ('data','booths'), ('data','voters'),
    ('data','association'), ('data','family'), ('data','family_members'), ('data','voter_changelog'),
    ('snapshot','voter_snapshot')
)
ORDER BY table_schema, table_name;

-- 2) Row counts
SELECT 'data.assembly' AS table_name, COUNT(*) AS rows FROM data.assembly
UNION ALL SELECT 'data.wards', COUNT(*) FROM data.wards
UNION ALL SELECT 'data.booths', COUNT(*) FROM data.booths
UNION ALL SELECT 'data.voters', COUNT(*) FROM data.voters
UNION ALL SELECT 'data.association', COUNT(*) FROM data.association
UNION ALL SELECT 'data.family', COUNT(*) FROM data.family
UNION ALL SELECT 'data.family_members', COUNT(*) FROM data.family_members
UNION ALL SELECT 'metastore.tenant', COUNT(*) FROM metastore.tenant
UNION ALL SELECT 'metastore.users', COUNT(*) FROM metastore.users
UNION ALL SELECT 'snapshot.voter_snapshot', COUNT(*) FROM snapshot.voter_snapshot;

-- 3) Spot check tenant distribution in voters
SELECT tenant_id, COUNT(*) AS voters
FROM data.voters
GROUP BY tenant_id
ORDER BY voters DESC
LIMIT 20;

-- 4) Scope check: if API returns few voters, this explains it by assignment scope
SELECT u.first_name, u.role, u.assignment_type, u.assignment_id, t.tenant_id
FROM metastore.users u
LEFT JOIN metastore.tenant t ON t.id = u.tenant_id
WHERE COALESCE(u.blocked,false) = false AND COALESCE(u.deleted,false) = false
ORDER BY u.id
LIMIT 50;

-- 5) Assembly-wise voter volume (through booth/ward joins)
SELECT a.assembly_code, COUNT(v.voter_id) AS voters
FROM data.voters v
JOIN data.booths b ON b.booth_id = v.booth_id
JOIN data.wards w ON w.ward_id = b.ward_id
JOIN data.assembly a ON a.assembly_id = w.assembly_id
GROUP BY a.assembly_code
ORDER BY voters DESC
LIMIT 20;
