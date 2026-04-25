from app.database import SessionLocal
from app.main import Voter, Booth, Ward, MessageTemplate
from sqlalchemy import text

db = SessionLocal()

epic = "UZJ6179527"
voter = db.query(Voter).filter(Voter.epic_no == epic).first()
print(f"VOTER: epic={epic}, booth_id={voter.booth_id if voter else 'NOT FOUND'}")

if voter and voter.booth_id:
    booth = db.query(Booth).filter(Booth.booth_id == voter.booth_id).first()
    print(f"BOOTH: id={booth.booth_id if booth else 'NOT FOUND'}, ward_id={booth.ward_id if booth else 'N/A'}")
    if booth and booth.ward_id:
        # Check MessageTemplate
        tpl = db.query(MessageTemplate).filter(MessageTemplate.ward_id == booth.ward_id).first()
        print(f"TEMPLATE for ward_id={booth.ward_id}: enabled={tpl.enabled if tpl else 'NOT FOUND'}")
        
        # Check Ward Info
        ward = db.query(Ward).filter(Ward.ward_id == booth.ward_id).first()
        print(f"WARD info: id={ward.ward_id if ward else 'N/A'}, name={ward.ward_name_en if ward else 'N/A'}")

