# Survey_Production Migration (Votabase Backend)

## Files
- `survey_production_bootstrap.sql`
  - Creates required backend schemas/tables in Survey_Production if missing.
  - Does not alter existing `public` survey tables.
- `survey_production_copy_from_votabase.sql`
  - Copies data from old `Votabase_Production` into new Survey_Production tables using `postgres_fdw`.
- `survey_production_post_migration_checks.sql`
  - Verifies table existence, row counts, tenant distribution, and assignment scope.

## Run order
1. Run `survey_production_bootstrap.sql` on Survey_Production.
2. Run `survey_production_copy_from_votabase.sql` on Survey_Production.
3. Run `survey_production_post_migration_checks.sql` on Survey_Production.

## Notes
- Update credentials/DB names in `survey_production_copy_from_votabase.sql` if needed.
- Backend code expects these schemas/tables:
  - `metastore.*`
  - `data.*`
  - `snapshot.*`
- `public.*` survey tables are not modified.
