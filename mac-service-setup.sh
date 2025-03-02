#!/bin/bash

# Office Project Tracker Service Setup
# This script sets up the Office Project Tracker to run as a background service

# Colors for pretty output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}"
echo "================================================="
echo "  Office Project Tracker - Background Service Setup"
echo "================================================="
echo -e "${NC}"

# Detect OS
if [[ "$OSTYPE" == "darwin"* ]]; then
    OS="mac"
    echo "Mac OS detected"
    SERVICE_DIR="$HOME/Library/LaunchAgents"
    SERVICE_FILE="com.officetracker.plist"
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="linux"
    echo "Linux detected"
    SERVICE_DIR="$HOME/.config/systemd/user"
    SERVICE_FILE="officetracker.service"
else
    echo -e "${RED}Unsupported operating system. This script works on Mac and Linux only.${NC}"
    exit 1
fi

# Determine installation directory
INSTALL_DIR="$HOME/Documents/OfficeTracker"
mkdir -p "$INSTALL_DIR"

echo -e "\n${YELLOW}Step 1: Preparing installation directory${NC}"
echo "Installing to: $INSTALL_DIR"

# Check for Python
echo -e "\n${YELLOW}Step 2: Checking for Python 3${NC}"
if command -v python3 &>/dev/null; then
    echo -e "${GREEN}Python 3 found!${NC}"
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    # Check if it's Python 3
    PY_VERSION=$(python --version 2>&1)
    if [[ $PY_VERSION == *"Python 3"* ]]; then
        echo -e "${GREEN}Python 3 found!${NC}"
        PYTHON_CMD="python"
    else
        echo -e "${RED}Python 3 is required but not found. Please install Python 3 and try again.${NC}"
        exit 1
    fi
else
    echo -e "${RED}Python 3 is required but not found. Please install Python 3 and try again.${NC}"
    exit 1
fi

# Check and install required packages
echo -e "\n${YELLOW}Step 3: Installing required Python packages${NC}"
$PYTHON_CMD -m pip install --user requests beautifulsoup4 nltk python-dotenv langdetect

# Install NLTK data
echo -e "\n${YELLOW}Step 4: Downloading NLTK data${NC}"
$PYTHON_CMD -c "import nltk; nltk.download('punkt', quiet=True); nltk.download('stopwords', quiet=True)"

# Create the script to copy provided office_tracker.py
echo -e "\n${YELLOW}Step 5: Creating application files${NC}"

# Writing the entry point script
cat > "$INSTALL_DIR/run_tracker.py" << 'EOF'
#!/usr/bin/env python3
"""
Office Project Tracker - Automated Runner Script
This script runs the office tracker in background mode.
"""
import os
import sys
import logging
from datetime import datetime
import time

# Set up logging
log_dir = os.path.expanduser("~/Documents/OfficeTracker/logs")
os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, f"tracker_{datetime.now().strftime('%Y%m%d')}.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Add script directory to path
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

try:
    import office_tracker
    logging.info("Starting Office Project Tracker...")
    office_tracker.main()
    logging.info("Office Project Tracker completed successfully")
except Exception as e:
    logging.error(f"Error running Office Project Tracker: {e}", exc_info=True)

print("Office Project Tracker run completed. Check logs at:", log_file)
EOF

chmod +x "$INSTALL_DIR/run_tracker.py"

# Create .env file template
cat > "$INSTALL_DIR/.env.template" << 'EOF'
# Google API Settings
GOOGLE_API_KEY=your_google_api_key
GOOGLE_CSE_ID=your_custom_search_engine_id

# Email Settings
EMAIL_USERNAME=your_email@gmail.com
EMAIL_PASSWORD=your_email_password
RECEIVER_EMAIL=panos.bompolas@inmind.com.gr
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Application Settings
EMAIL_LANGUAGE=en
MAX_SEARCH_QUERIES=30
SEND_ANALYTICS=false
EOF

echo -e "${GREEN}Created template .env file. You'll need to edit this with your settings.${NC}"

# Copy the main office tracker file if it exists in current directory
if [ -f "office_tracker.py" ]; then
    cp "office_tracker.py" "$INSTALL_DIR/"
    echo -e "${GREEN}Copied office_tracker.py to $INSTALL_DIR/${NC}"
else
    echo -e "${RED}Warning: office_tracker.py not found in current directory.${NC}"
    echo "Please manually copy the office_tracker.py file to $INSTALL_DIR/ before running the service."
fi

# Set up the service based on OS
echo -e "\n${YELLOW}Step 6: Setting up background service${NC}"

if [[ "$OS" == "mac" ]]; then
    # Create macOS LaunchAgent
    mkdir -p "$SERVICE_DIR"
    
    cat > "$SERVICE_DIR/$SERVICE_FILE" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.officetracker</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PYTHON_CMD</string>
        <string>$INSTALL_DIR/run_tracker.py</string>
    </array>
    <key>StartCalendarInterval</key>
    <array>
        <dict>
            <key>Hour</key>
            <integer>9</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
        <dict>
            <key>Hour</key>
            <integer>15</integer>
            <key>Minute</key>
            <integer>0</integer>
        </dict>
    </array>
    <key>StandardOutPath</key>
    <string>$INSTALL_DIR/logs/stdout.log</string>
    <key>StandardErrorPath</key>
    <string>$INSTALL_DIR/logs/stderr.log</string>
    <key>WorkingDirectory</key>
    <string>$INSTALL_DIR</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin</string>
    </dict>
</dict>
</plist>
EOF

    echo -e "${GREEN}Created LaunchAgent at $SERVICE_DIR/$SERVICE_FILE${NC}"
    echo "This will run the tracker twice daily at 9:00 AM and 3:00 PM."
    
    # Load the service
    echo -e "\n${YELLOW}Loading and starting the service...${NC}"
    launchctl unload "$SERVICE_DIR/$SERVICE_FILE" 2>/dev/null
    launchctl load "$SERVICE_DIR/$SERVICE_FILE"
    
    echo -e "${GREEN}Service installed and started!${NC}"
    echo -e "To manually start: ${BLUE}launchctl start com.officetracker${NC}"
    echo -e "To manually stop: ${BLUE}launchctl stop com.officetracker${NC}"
    echo -e "To remove: ${BLUE}launchctl unload $SERVICE_DIR/$SERVICE_FILE${NC}"

elif [[ "$OS" == "linux" ]]; then
    # Create Linux systemd user service
    mkdir -p "$SERVICE_DIR"
    
    cat > "$SERVICE_DIR/$SERVICE_FILE" << EOF
[Unit]
Description=Office Project Tracker
After=network.target

[Service]
Type=oneshot
ExecStart=$PYTHON_CMD $INSTALL_DIR/run_tracker.py
WorkingDirectory=$INSTALL_DIR
StandardOutput=append:$INSTALL_DIR/logs/stdout.log
StandardError=append:$INSTALL_DIR/logs/stderr.log
Environment="PATH=/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"

[Install]
WantedBy=default.target
EOF

    # Create timer for scheduled runs
    cat > "$SERVICE_DIR/officetracker.timer" << EOF
[Unit]
Description=Run Office Project Tracker twice daily

[Timer]
OnCalendar=*-*-* 9,15:00:00
Persistent=true

[Install]
WantedBy=timers.target
EOF

    echo -e "${GREEN}Created systemd service at $SERVICE_DIR/$SERVICE_FILE${NC}"
    echo "This will run the tracker twice daily at 9:00 AM and 3:00 PM."
    
    # Load the service
    echo -e "\n${YELLOW}Loading and starting the service...${NC}"
    systemctl --user daemon-reload
    systemctl --user enable officetracker.timer
    systemctl --user start officetracker.timer
    
    echo -e "${GREEN}Service installed and started!${NC}"
    echo -e "To manually run once: ${BLUE}systemctl --user start officetracker.service${NC}"
    echo -e "To check timer status: ${BLUE}systemctl --user status officetracker.timer${NC}"
    echo -e "To disable: ${BLUE}systemctl --user disable officetracker.timer${NC}"
fi

# Create a manual run script
cat > "$INSTALL_DIR/run_now.sh" << EOF
#!/bin/bash
cd "\$(dirname "\$0")"
$PYTHON_CMD run_tracker.py
EOF

chmod +x "$INSTALL_DIR/run_now.sh"

echo -e "\n${GREEN}===== Installation Complete =====${NC}"
echo -e "The Office Project Tracker has been set up to run automatically."
echo -e "\nIMPORTANT: You must configure your settings by editing:"
echo -e "${BLUE}$INSTALL_DIR/.env.template${NC}"
echo -e "Rename it to ${BLUE}.env${NC} after editing."
echo -e "\nTo run the tracker manually, use:"
echo -e "${BLUE}$INSTALL_DIR/run_now.sh${NC}"
echo -e "\nLogs can be found in: ${BLUE}$INSTALL_DIR/logs/${NC}"

# Ask if user wants to edit the .env file now
echo -e "\n${YELLOW}Would you like to create and edit the .env file now? (y/n)${NC}"
read -r response
if [[ "$response" =~ ^([yY][eE][sS]|[yY])$ ]]; then
    cp "$INSTALL_DIR/.env.template" "$INSTALL_DIR/.env"
    
    # Try to determine the best editor
    if command -v nano &>/dev/null; then
        EDITOR="nano"
    elif command -v vim &>/dev/null; then
        EDITOR="vim"
    elif command -v vi &>/dev/null; then
        EDITOR="vi"
    elif [[ "$OS" == "mac" ]] && command -v open &>/dev/null; then
        EDITOR="open -t"
    else
        EDITOR=""
    fi
    
    if [[ -n "$EDITOR" ]]; then
        $EDITOR "$INSTALL_DIR/.env"
    else
        echo -e "${YELLOW}No text editor found. Please edit the file manually:${NC}"
        echo -e "${BLUE}$INSTALL_DIR/.env${NC}"
    fi
else
    echo -e "${YELLOW}Remember to create and edit .env file before running the service.${NC}"
fi

echo -e "\n${GREEN}Setup complete!${NC}"