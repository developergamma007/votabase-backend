from sqlalchemy import create_engine, text
import os

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("--- Srinagar in public.wards ---")
    # Search for Srinagar in public schema
    res = conn.execute(text("SELECT ward_id, ward_name_en, assembly_id FROM public.wards WHERE ward_name_en ILIKE '%Srinagar%'"))
    for row in res:
        print(row)
        
    print("\n--- All assemblies in public.assembly ---")
    res = conn.execute(text("SELECT id, assembly_name_en, assembly_code FROM public.assembly WHERE id = 170 OR assembly_name_en ILIKE '%Basavanagudi%'"))
    for row in res:
        print(row)
