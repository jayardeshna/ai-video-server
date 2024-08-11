import os


class Config:
    UPLOAD_FOLDER = 'uploads/'
    SECRET_KEY = os.getenv('SECRET_KEY', 'your_secret_key')