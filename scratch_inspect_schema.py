
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("BOOTH COLUMNS:")
    res = conn.execute(text("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'booths'"))
    for row in res:
        print(f"  {row[0]}")
    
    print("\nWARD COLUMNS:")
    res = conn.execute(text("SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'wards'"))
    for row in res:
        print(f"  {row[0]}")
