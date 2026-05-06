
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("ASSEMBLY 170 in public.assembly:")
    res = conn.execute(text("SELECT assembly_code, assembly_no, assembly_name_en, tenant_id FROM public.assembly WHERE assembly_no = 170"))
    for row in res:
        print(f"  Code: {row[0]}, No: {row[1]}, Name: {row[2]}, Tenant: {row[3]}")
