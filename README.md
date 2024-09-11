# linkedin-bot

# LinkedIn Outreach Helper MVP

## Overview

The **LinkedIn Outreach Helper MVP** is a Python-based GUI application designed to assist with LinkedIn outreach efforts. The tool allows users to load LinkedIn profiles from a CSV file, categorize them as executive or non-executive, generate custom messages based on predefined templates, and manage profiles that have been contacted. The application also integrates with an external script to check the LinkedIn message box for recently contacted recipients and update the database accordingly.

## Features

- **Load CSV Files**: Import LinkedIn profiles from a CSV file.
- **Executive & Non-Executive Tabs**: Categorize profiles into executive and non-executive categories.
- **Custom Message Generation**: Automatically generate outreach messages based on the relationship strength (Strong, Moderate, Weak).
- **Mark as Sent**: Mark profiles as contacted and store the information in the database.
- **Skip Profiles**: Skip profiles that don't need to be contacted.
- **Check Recently Contacted**: Filter out recently contacted profiles based on a specified timeframe using an external LinkedIn script.
- **Settings**: Customize message templates and securely store LinkedIn credentials.

## Prerequisites

- **Python 3.x**
- **SQLite** (comes with Python standard library)
- **pip** (Python package manager)

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/linkedin-outreach-helper-mvp.git
cd linkedin-outreach-helper-mvp
```

### 2. Create a Virtual Environment (recommended)
```bash
python -m venv venv
source venv/bin/activate # On Windows use `venv\Scripts\activate`
```

### 3. Pip Install Dependencies
```bash
pip install -r requirements.txt
```


## Usage


### 1. Export connections from linkedin

Go to https://www.linkedin.com/mypreferences/d/categories/account
Profile > Settings > Data Privacy > Get a copy of your data > Connections

### 2. Run the Application
```bash
python linkedinGUI.py
```

### 3. Load Profiles
	•	Click on “Load CSV” to import your LinkedIn profiles from a CSV file.


## Troubleshooting
brew install python-tk