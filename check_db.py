from app.database import SessionLocal
from sqlalchemy import text

db = SessionLocal()
epicNo = "UZJ6179527"
t = text("SELECT booth_no, ward_code FROM public.voters WHERE epic_no = :epic")
v = db.execute(t, {"epic": epicNo}).first()
print("voter:", v)

b = text("SELECT * FROM public.booths WHERE booth_no = '1'")
booths = db.execute(b).fetchall()
print("booths:", booths)

