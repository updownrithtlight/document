from app import app
import os
if __name__ == "__main__":
    # Use Gunicorn to serve the Flask app
    host = '0.0.0.0'
    port = int(os.environ.get('PORT', 5000))  # Allow Heroku to set the port
    app.run(host=host, port=port)
