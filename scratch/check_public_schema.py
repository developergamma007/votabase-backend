import os
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)

tables = ["assembly", "wards", "booths", "voters"]
with engine.connect() as conn:
    for table in tables:
        print(f"Checking table public.{table}...")
        try:
            res = conn.execute(text(f"SELECT * FROM public.{table} LIMIT 0"))
            cols = res.keys()
            print(f"  Columns: {list(cols)}")
            if "tenant_id" in cols:
                print(f"  [FOUND] tenant_id exists in public.{table}")
            else:
                print(f"  [MISSING] tenant_id NOT in public.{table}")
        except Exception as e:
            print(f"  [ERROR] {e}")
