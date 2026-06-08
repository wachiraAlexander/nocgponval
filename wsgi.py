"""WSGI entry point for the Flask application."""
import os
import sys

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(__file__))

from app import app, socketio

# For WSGI servers that don't support WebSockets natively
# The application object is what WSGI servers expect
application = app

if __name__ == "__main__":
    # This is for local testing with Waitress
    from waitress import serve
    serve(application, host='0.0.0.0', port=5000)
