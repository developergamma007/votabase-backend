import os
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)

tables = ["assembly", "wards", "booths", "voters"]
with engine.connect() as conn:
    trans = conn.begin()
    try:
        for table in tables:
            print(f"Migrating table public.{table}...")
            # Check if column exists first (safety)
            res = conn.execute(text(f"SELECT column_name FROM information_schema.columns WHERE table_schema='public' AND table_name='{table}' AND column_name='tenant_id'"))
            if not res.fetchone():
                print(f"  Adding tenant_id to public.{table}")
                conn.execute(text(f"ALTER TABLE public.{table} ADD COLUMN tenant_id VARCHAR(20)"))
            else:
                print(f"  tenant_id already exists in public.{table}")
        
        trans.commit()
        print("Migration COMPLETED.")
    except Exception as e:
        trans.rollback()
        print(f"Migration FAILED: {e}")
