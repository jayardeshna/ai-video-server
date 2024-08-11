from flask import Flask


def create_app():
    app = Flask(__name__)
    app.config.from_object('app.config.Config')

    # Register Blueprints
    from app.controller.file_controller import upload_bp
    app.register_blueprint(upload_bp)

    return app