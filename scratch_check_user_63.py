
import sys
import os

# Add the app directory to the path
sys.path.append(os.path.abspath("."))

from sqlalchemy import create_engine, text

DATABASE_URL = "postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb"

engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    print("VOLUNTEER USER 63 DETAILS:")
    res = conn.execute(text("SELECT id, first_name, role, deleted, blocked, tenant_id, assignment_id FROM metastore.volunteer_users WHERE id = 63"))
    for row in res:
        print(f"  ID: {row[0]}, Name: {row[1]}, Role: {row[2]}, Deleted: {row[3]}, Blocked: {row[4]}, Tenant: {row[5]}, Assignment: {row[6]}")
