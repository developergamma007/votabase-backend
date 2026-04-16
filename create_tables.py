from app.main import engine, Base, Meeting, MeetingAttendance
print("Creating tables in database...")
Base.metadata.create_all(bind=engine)
print("Tables created successfully.")
