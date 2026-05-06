
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("VOTERS FOR BOOTH 1 or similar in Ward 100 area:")
    res = conn.execute(text("SELECT ward_code, booth_no, epic FROM public.voters WHERE booth_no = '1' AND ward_code LIKE '%100%' LIMIT 5"))
    for row in res:
        print(f"  Ward: {row[0]}, Booth: {row[1]}, EPIC: {row[2]}")
