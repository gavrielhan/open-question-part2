#!/bin/bash
# Launcher script for Topic Generator Web App
# Double-click this file to start the application

# Get the directory where this script is located
cd "$(dirname "$0")"

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    source venv/bin/activate
elif [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# Check if .env file exists
if [ ! -f ".env" ]; then
    osascript -e 'display dialog "Error: .env file not found!\n\nPlease create a .env file with your API configuration." buttons {"OK"} default button "OK" with icon stop'
    exit 1
fi

# Start the web application
python3 topic_generator_app.py

# Keep terminal open if there's an error
if [ $? -ne 0 ]; then
    echo ""
    echo "Press any key to close..."
    read -n 1
fi

