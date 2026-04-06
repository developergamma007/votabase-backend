import json
from sqlalchemy import create_engine, text

engine = create_engine("postgresql+psycopg2://surveyuser:surveyuser@13.233.40.235:5432/surveydb")
with engine.connect() as conn:
    print("Test89 deleted status:")
    vols = conn.execute(text("SELECT id, first_name, deleted FROM metastore.volunteer_users WHERE first_name = 'Test89'")).fetchall()
    for v in vols:
        print(dict(v._mapping))
