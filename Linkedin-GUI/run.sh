#!/bin/bash

# Check if Python 3.12 is installed
if ! command -v python3.12 &>/dev/null; then
    echo "Python 3.12 is not installed. Installing Python 3.12 now..."
    if [ ! -f python-3.12.4-macos11.pkg ]; then
        echo "Installer not found! Make sure python-3.12.4-macos11.pkg is in the same directory as this script."
        exit 1
    fi
    sudo installer -pkg python-3.12.4-macos11.pkg -target /
fi

# Set up a virtual environment
python3.12 -m venv venv
source venv/bin/activate

# Upgrade pip
pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt

# Run the script
python3 linkedinGUI.py