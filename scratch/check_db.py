
import os
from sqlalchemy import create_engine, text

def load_env():
    env_path = ".env"
    if os.path.exists(env_path):
        with open(env_path, "r") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                os.environ[k.strip()] = v.strip().strip("'\"")

load_env()
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)

with engine.connect() as conn:
    try:
        res = conn.execute(text("""
            SELECT column_default 
            FROM information_schema.columns 
            WHERE table_schema = 'data' AND table_name = 'voters' AND column_name = 'voter_id'
        """))
        print(f"data.voters.voter_id default: {res.scalar()}")
    except Exception as e:
        print(f"Error checking voters: {e}")
