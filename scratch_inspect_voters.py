
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("VOTERS SAMPLE (first 10):")
    # We don't know the column names for sure, so we select *
    # But we want to filter by something related to Assembly 170
    # Let's try to find voters where ward_code or booth_no is related
    res = conn.execute(text("SELECT ward_code, booth_no, COUNT(*) FROM public.voters GROUP BY ward_code, booth_no LIMIT 20"))
    for row in res:
        print(f"  Ward: {row[0]}, Booth: {row[1]}, Count: {row[2]}")
