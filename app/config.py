import os


class Config:
    UPLOAD_FOLDER = 'uploads/'
    PPT_FOLDER = 'uploads/ppt/'
    IMAGE_FOLDER = 'uploads/images/'
    AUDIO_FOLDER = 'uploads/audio'
    VIDEO_PATH = 'uploads/videos'
    SECRET_KEY = os.getenv('SECRET_KEY', 'your_secret_key')