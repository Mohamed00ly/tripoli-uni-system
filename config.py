import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'tripoli-portal-secret-2024-change-in-production')
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        'DATABASE_URL',
        'postgresql://portal_user:portal_pass@localhost:5432/tripoli_portal'
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
