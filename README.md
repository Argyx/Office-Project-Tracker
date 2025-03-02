# OfficeInsights: Automated Real Estate Intelligence

An intelligent system that automatically scans the web for new office development projects, relocations, and commercial real estate transactions. Provides real-time email notifications about new office spaces, corporate moves, and real estate opportunities across Greece in both English and Greek.

## Key Features
- Automated daily web scanning for office-related news
- Multilingual support (English & Greek)
- Entity extraction for companies and locations
- Customizable email notifications
- Analytics and reporting dashboard
- No-code setup for non-technical users

Perfect for real estate professionals, commercial property investors, and business intelligence analysts tracking office market trends.


# User Guide

## Table of Contents
1. [Introduction](#introduction)
2. [What You'll Need](#what-youll-need)
3. [Installation Guide](#installation-guide)
    - [Windows Installation](#windows-installation)
    - [Mac Installation](#mac-installation)
    - [Linux Installation](#linux-installation)
4. [Getting API Keys](#getting-api-keys)
5. [Setting Up Email Notifications](#setting-up-email-notifications)
6. [Using the Application](#using-the-application)
    - [Desktop Application](#desktop-application)
    - [Background Service](#background-service)
7. [Viewing Your Results](#viewing-your-results)
8. [Troubleshooting](#troubleshooting)
9. [Frequently Asked Questions](#frequently-asked-questions)

---

## Introduction

The Office Project Tracker automatically searches the internet for information about new office projects and developments and sends you email notifications with what it finds. It works in Greek and English and can detect any company's office-related activities.

This system is designed to be simple to use - once set up, it can run automatically in the background, checking for new information daily and sending it straight to your email.

---

## What You'll Need

Before starting, make sure you have:

1. **A Google API Key** and **Custom Search Engine ID** (instructions provided below)
2. **An email account** to send notifications from (Gmail recommended)
3. **Internet connection**

Don't worry if you don't have these yet - the installation guide will help you get them.

---

## Installation Guide

### Windows Installation

**Option A: Simple Installation (Recommended)**

1. Download these files to your computer:
   - `office_tracker.py` (the main program)
   - `windows-installer.py` (the installer)
   - Save them in the same folder

2. Double-click on `windows-installer.py`

3. Follow the on-screen prompts:
   - The installer will automatically download and install Python if you don't have it
   - It will install all required packages
   - You'll be asked where to install the application
   - You can choose to have it start automatically when you turn on your computer

4. When installation is complete, you'll find a shortcut on your desktop labeled "Office Project Tracker"

**Option B: Manual Setup (For Advanced Users)**

If you already have Python installed:

1. Open Command Prompt
2. Run: `pip install requests beautifulsoup4 nltk python-dotenv langdetect`
3. Create a folder in your Documents called "OfficeTracker"
4. Copy the `office_tracker.py` and `office_tracker_gui.py` files to this folder

### Mac Installation

1. Download these files to your computer:
   - `office_tracker.py` (the main program)
   - `mac-service-setup.sh` (the setup script)
   - Save them in the same folder

2. If you don't have Python installed:
   - Go to [python.org/downloads](https://www.python.org/downloads/)
   - Download and install the latest Python version for Mac

3. Open Terminal (find it in Applications > Utilities)

4. Navigate to the folder where you saved the files:
   ```
   cd /path/to/your/folder
   ```

5. Run the setup script:
   ```
   bash mac-service-setup.sh
   ```

6. Follow the on-screen prompts to complete installation

### Linux Installation

1. Download these files to your computer:
   - `office_tracker.py` (the main program)
   - `mac-service-setup.sh` (works for Linux too)
   - Save them in the same folder

2. Make sure Python is installed (most Linux systems have it pre-installed):
   ```
   python3 --version
   ```
   If not installed, use your distribution's package manager, e.g.:
   ```
   sudo apt install python3 python3-pip    # For Ubuntu/Debian
   ```

3. Open Terminal

4. Navigate to the folder where you saved the files:
   ```
   cd /path/to/your/folder
   ```

5. Run the setup script:
   ```
   bash mac-service-setup.sh
   ```

6. Follow the on-screen prompts to complete installation

---

## Getting API Keys

You'll need a Google API Key and Custom Search Engine ID to use this application. Here's how to get them:

### Google API Key

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)

2. Sign in with your Google account

3. Create a new project:
   - Click on the project dropdown at the top of the page
   - Click "New Project"
   - Enter a name (e.g., "Office Tracker")
   - Click "Create"

4. Enable the Custom Search API:
   - In the left sidebar, click "APIs & Services" > "Library"
   - Search for "Custom Search API"
   - Click on it and then click "Enable"

5. Create an API Key:
   - In the left sidebar, click "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "API key"
   - Copy the API key that appears (you'll need it later)

### Custom Search Engine ID

1. Go to [programmablesearchengine.google.com](https://programmablesearchengine.google.com/)

2. Click "Add" to create a new search engine

3. In the "Sites to search" field, enter `*` to search the entire web

4. Give your search engine a name (e.g., "Office Projects Search")

5. Click "Create"

6. On the next page, click "Customize" for your new search engine

7. Make sure "Search the entire web" is enabled

8. Find your Search Engine ID (it will look something like `012345678901234567890:abcdefghijk`)

9. Copy the Search Engine ID (you'll need it later)

---

## Setting Up Email Notifications

The application will send you email notifications about new office projects. Here's how to set it up:

### Gmail Setup (Recommended)

1. You'll need a Gmail account to send notifications

2. For security, Google requires you to use an "App Password" instead of your regular password:
   - Go to [myaccount.google.com/security](https://myaccount.google.com/security)
   - Make sure 2-Step Verification is enabled
   - Once enabled, go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Select "Other" from the dropdown, type "Office Tracker" and click "Generate"
   - Copy the 16-character password (you'll use this in the application)

3. You'll enter these settings in the application:
   - Email Username: Your full Gmail address (e.g., `yourname@gmail.com`)
   - Email Password: The App Password you generated
   - SMTP Server: `smtp.gmail.com`
   - SMTP Port: `587`

### Other Email Providers

If you use another email provider, you'll need:
- Your email address
- Your password (or app password if required)
- The SMTP server address
- The SMTP port

Contact your email provider for these details if unsure.

---

## Using the Application

There are two ways to use the Office Project Tracker:

### Desktop Application

If you installed using the Windows installer:

1. Double-click the "Office Project Tracker" shortcut on your desktop

2. When you first run the app, go to the "Settings" tab:
   - Enter your Google API Key and Custom Search Engine ID
   - Enter your email settings
   - Choose your preferred language (English or Greek)
   - Click "Save Settings"

3. Return to the "Dashboard" tab:
   - Click "Run Now" to perform a search immediately
   - Click "Start Scheduler" to have the application check for new projects automatically
   - Set how often you want it to check (6, 12, or 24 hours)

4. The "Recent Results" tab shows you the latest projects found

### Background Service

If you installed using the service setup script:

1. The service is configured to run automatically twice daily (9 AM and 3 PM)

2. You need to configure your settings first:
   - Edit the `.env` file in your OfficeTracker folder:
     - On Windows: `C:\Users\YourUsername\Documents\OfficeTracker\.env`
     - On Mac/Linux: `~/Documents/OfficeTracker/.env`
   
   - Enter your API keys, email settings, and preferences:
     ```
     GOOGLE_API_KEY=your_api_key_here
     GOOGLE_CSE_ID=your_search_engine_id_here
     EMAIL_USERNAME=your_email@gmail.com
     EMAIL_PASSWORD=your_app_password_here
     RECEIVER_EMAIL=where_to_send@example.com
     SMTP_SERVER=smtp.gmail.com
     SMTP_PORT=587
     EMAIL_LANGUAGE=en
     ```

3. To run the service manually:
   - Windows: Double-click `Start Office Tracker.bat` in your OfficeTracker folder
   - Mac/Linux: Run `~/Documents/OfficeTracker/run_now.sh`

---

## Viewing Your Results

The Office Project Tracker will send you email notifications when it finds new office projects. These emails will include:

- The company name
- The location of the project
- A description of the project
- A link to the source of the information

If you're using the desktop application, you can also view recent findings in the "Recent Results" tab.

---

## Troubleshooting

### The application won't start

- Make sure you have Python installed (the Windows installer should have handled this)
- Try running the installer again

### No emails are being received

- Check your spam/junk folder
- Verify your email settings are correct
- Make sure your app password is entered correctly
- Try the "Test Email" button in settings

### No results are being found

- Verify your Google API Key and Custom Search Engine ID are entered correctly
- Make sure your internet connection is working
- Try running the search manually by clicking "Run Now"

### "API key not valid" error

- Your Google API key may have expired or been disabled
- Go to the Google Cloud Console and check your API key status
- You may need to create a new API key

### Windows Defender or antivirus warning

- This can happen because the application connects to the internet
- You can safely add the application to your antivirus exceptions
- The code is open-source and doesn't contain any harmful content

---

## Frequently Asked Questions

**Q: How much does this cost to use?**  
A: Google provides a free quota of 100 searches per day with the Custom Search API. This should be more than enough for daily use.

**Q: Will this work on my computer?**  
A: Yes, the application works on Windows, Mac, and Linux systems.

**Q: Do I need to keep my computer on all the time?**  
A: No. If you're using the background service, it will run whenever your computer is on. It will check for new projects the next time it runs according to schedule.

**Q: How do I stop the automatic checks?**  
A: In the desktop application, click "Stop Scheduler". For the background service:
- Windows: Open Task Scheduler, find "OfficeProjectTracker" and disable it
- Mac: Run `launchctl unload ~/Library/LaunchAgents/com.officetracker.plist`
- Linux: Run `systemctl --user disable officetracker.timer`

**Q: Can I change what kind of projects it looks for?**  
A: The system is pre-configured to look for office projects. Advanced users can modify the search queries in the `office_tracker.py` file.

**Q: How often should I run the tracker?**  
A: We recommend daily checks. Running too frequently might exceed your Google API free quota.

**Q: Is my Google API key secure?**  
A: Yes, your API key is stored only on your computer and is used only to perform searches.

**Q: Can I use this for other types of projects?**  
A: Yes, advanced users can modify the code to search for different types of projects.

---