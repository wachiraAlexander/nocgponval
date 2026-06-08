#!/bin/bash
# GPON Application - Waitress WSGI Server Startup Script for Linux/macOS

echo "Starting GPON Application with Waitress WSGI Server..."
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    echo "Please install Python 3.8+ from https://www.python.org/"
    exit 1
fi

# Activate virtual environment if it exists
if [ -f venv/bin/activate ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
fi

# Install/upgrade dependencies
echo "Installing dependencies..."
pip install -r requirements.txt

# Run the application with Waitress
echo ""
echo "========================================"
echo "GPON Application is starting..."
echo "Server: http://localhost:5000"
echo "Press Ctrl+C to stop the server"
echo "========================================"
echo ""

python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
