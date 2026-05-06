import os
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    trans = conn.begin()
    try:
        print("Cleaning and Enforcing Primary Keys...")
        
        # Assembly
        try:
            print("  Handling Assembly...")
            conn.execute(text("ALTER TABLE public.assembly ADD PRIMARY KEY (assembly_no)"))
            print("  [SUCCESS] Assembly PK added.")
        except Exception as e:
            print(f"  [INFO] Assembly PK skipped: {e}")

        # Voters
        try:
            print("  Cleaning Voters (deleting null/empty epics)...")
            conn.execute(text("DELETE FROM public.voters WHERE epic IS NULL OR epic = ''"))
            print("  Trying to add PK to public.voters (epic)...")
            conn.execute(text("ALTER TABLE public.voters ADD PRIMARY KEY (epic)"))
            print("  [SUCCESS] Voters PK added.")
        except Exception as e:
            print(f"  [INFO] Voters PK skipped: {e}")

        trans.commit()
        print("Maintenance COMPLETE.")
    except Exception as ex:
        trans.rollback()
        print(f"FATAL: {ex}")
