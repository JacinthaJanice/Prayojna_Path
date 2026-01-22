# create_admin.py
from app import app, db, User
from werkzeug.security import generate_password_hash

with app.app_context():
    # Check if admin user already exists
    existing_admin = User.query.filter_by(username='admin').first()
    if existing_admin:
        print("Admin user already exists:", existing_admin.username)
    else:
        # Create new admin user
        admin = User(username='admin', role='admin')
        admin.set_password('admin123')  # Set password to 'admin123'
        db.session.add(admin)
        db.session.commit()
        print("Admin user created with username: admin, password: admin123")