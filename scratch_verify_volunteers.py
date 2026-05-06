
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("VOLUNTEER USERS for Assignment 170:")
    res = conn.execute(text("SELECT id, first_name, phone, assignment_id, assignment_type, tenant_id FROM metastore.volunteer_users WHERE assignment_id = '170' OR assignment_id = '000000000170'"))
    for row in res:
        print(f"  ID: {row[0]}, Name: {row[1]}, Phone: {row[2]}, Assignment: {row[3]}, Type: {row[4]}, Tenant: {row[5]}")
