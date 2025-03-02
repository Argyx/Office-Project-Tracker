#!/usr/bin/env python3
"""
Office Project Tracker - Uninstallation Script
This script helps uninstall the Office Project Tracker by removing database files,
logs, and other related files.
"""

import os
import sys
import shutil
from pathlib import Path
import sqlite3

def main():
    print("=" * 60)
    print("OFFICE PROJECT TRACKER - UNINSTALLATION".center(60))
    print("=" * 60)
    print("\nThis script will remove all files created by the Office Project Tracker.")
    
    # Get current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Files to be removed from the current directory
    files_to_remove = [
        "office_projects.db",        # Database file
        "office_tracker.log",        # Log file
        ".env",                      # Environment variables
        "office_tracker_gui.py",     # GUI script if present
        "run_tracker.py",            # Runner script if present
    ]
    
    # Additional log files with date patterns
    log_pattern = "office_tracker_*.log"
    
    # Ask for confirmation
    print("\nThe following items will be removed:")
    print("1. Database file (office_projects.db)")
    print("2. Log files (office_tracker.log and date-based logs)")
    print("3. Environment configuration file (.env)")
    
    confirmation = input("\n⚠️  Do you want to proceed with uninstallation? (yes/no): ")
    
    if confirmation.lower() not in ["yes", "y"]:
        print("\nUninstallation cancelled.")
        return
    
    # Remove specific files
    files_removed = 0
    for file_name in files_to_remove:
        file_path = os.path.join(current_dir, file_name)
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"✓ Removed: {file_name}")
                files_removed += 1
            except Exception as e:
                print(f"✗ Failed to remove {file_name}: {e}")
    
    # Remove log files matching pattern
    log_files = list(Path(current_dir).glob(log_pattern))
    for log_file in log_files:
        try:
            os.remove(log_file)
            print(f"✓ Removed: {log_file.name}")
            files_removed += 1
        except Exception as e:
            print(f"✗ Failed to remove {log_file.name}: {e}")
    
    # Check if data directory exists in user's Documents folder
    data_dir = os.path.join(str(Path.home()), "Documents", "OfficeTracker")
    if os.path.exists(data_dir):
        print(f"\nFound data directory at: {data_dir}")
        remove_data = input("Do you want to remove this directory and all its contents? (yes/no): ")
        
        if remove_data.lower() in ["yes", "y"]:
            try:
                shutil.rmtree(data_dir)
                print(f"✓ Removed data directory: {data_dir}")
            except Exception as e:
                print(f"✗ Failed to remove data directory: {e}")
    
    # Check for application code
    remove_code = input("\nDo you want to remove the Office Tracker code files? (yes/no): ")
    if remove_code.lower() in ["yes", "y"]:
        code_files = [
            "office_tracker.py",
            "uninstall.py"  # This script itself
        ]
        
        for file_name in code_files:
            file_path = os.path.join(current_dir, file_name)
            if os.path.exists(file_path) and file_name != sys.argv[0]:
                try:
                    os.remove(file_path)
                    print(f"✓ Removed: {file_name}")
                    files_removed += 1
                except Exception as e:
                    print(f"✗ Failed to remove {file_name}: {e}")
        
        # Remove this script last
        if os.path.basename(sys.argv[0]) in code_files:
            print(f"\nNote: This script ({os.path.basename(sys.argv[0])}) will be removed last.")
            with open("remove_script.bat" if os.name == "nt" else "remove_script.sh", "w") as f:
                if os.name == "nt":  # Windows
                    f.write(f"@echo off\ntimeout /t 1 /nobreak >nul\ndel {os.path.basename(sys.argv[0])}\ndel remove_script.bat\n")
                else:  # macOS/Linux
                    f.write(f"#!/bin/bash\nsleep 1\nrm {os.path.basename(sys.argv[0])}\nrm remove_script.sh\n")
            
            if os.name != "nt":  # Make the script executable on Unix
                os.chmod("remove_script.sh", 0o755)
            
            print(f"A script has been created to remove this uninstaller.")
            print(f"Run 'remove_script.{'bat' if os.name == 'nt' else 'sh'}' after this script completes to finish the uninstallation.")
    
    # Remove scheduler entries if applicable
    if os.name == "nt":  # Windows
        print("\nChecking for Windows scheduled tasks...")
        import subprocess
        try:
            # Check if task exists
            check_process = subprocess.run(["schtasks", "/query", "/tn", "OfficeProjectTracker"], 
                                          stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            if check_process.returncode == 0:
                print("Found scheduled task: OfficeProjectTracker")
                remove_task = input("Do you want to remove the scheduled task? (yes/no): ")
                
                if remove_task.lower() in ["yes", "y"]:
                    subprocess.run(["schtasks", "/delete", "/tn", "OfficeProjectTracker", "/f"])
                    print("✓ Removed scheduled task: OfficeProjectTracker")
        except:
            print("No scheduled tasks found or unable to check.")
    
    elif sys.platform == "darwin":  # macOS
        plist_path = os.path.expanduser("~/Library/LaunchAgents/com.officetracker.plist")
        if os.path.exists(plist_path):
            print(f"\nFound macOS LaunchAgent: {plist_path}")
            remove_agent = input("Do you want to remove the LaunchAgent? (yes/no): ")
            
            if remove_agent.lower() in ["yes", "y"]:
                try:
                    # Unload the service first
                    os.system(f"launchctl unload {plist_path}")
                    os.remove(plist_path)
                    print(f"✓ Removed LaunchAgent: {plist_path}")
                except Exception as e:
                    print(f"✗ Failed to remove LaunchAgent: {e}")
    
    else:  # Linux
        systemd_path = os.path.expanduser("~/.config/systemd/user/officetracker.service")
        timer_path = os.path.expanduser("~/.config/systemd/user/officetracker.timer")
        
        if os.path.exists(systemd_path) or os.path.exists(timer_path):
            print("\nFound systemd service/timer files.")
            remove_service = input("Do you want to remove the systemd service files? (yes/no): ")
            
            if remove_service.lower() in ["yes", "y"]:
                try:
                    os.system("systemctl --user stop officetracker.timer")
                    os.system("systemctl --user disable officetracker.timer")
                    
                    if os.path.exists(systemd_path):
                        os.remove(systemd_path)
                        print(f"✓ Removed service file: {systemd_path}")
                    
                    if os.path.exists(timer_path):
                        os.remove(timer_path)
                        print(f"✓ Removed timer file: {timer_path}")
                    
                    os.system("systemctl --user daemon-reload")
                except Exception as e:
                    print(f"✗ Failed to remove systemd files: {e}")
    
    # Uninstall dependencies
    uninstall_deps = input("\nDo you want to uninstall the Python dependencies used by the tracker? (yes/no): ")
    if uninstall_deps.lower() in ["yes", "y"]:
        print("\nUninstalling dependencies...")
        try:
            import pip
            dependencies = [
                "requests", 
                "beautifulsoup4", 
                "nltk", 
                "python-dotenv", 
                "langdetect",
                "tqdm"  # For progress bars
            ]
            
            for package in dependencies:
                print(f"Uninstalling {package}...")
                try:
                    pip.main(["uninstall", "-y", package])
                    print(f"✓ Uninstalled {package}")
                except:
                    print(f"✗ Failed to uninstall {package}")
        except:
            print("❌ Failed to uninstall dependencies. You may need to run: pip uninstall requests beautifulsoup4 nltk python-dotenv langdetect tqdm")
    
    # Complete
    print("\n" + "=" * 60)
    print("UNINSTALLATION COMPLETE".center(60))
    print("=" * 60)
    print(f"\nRemoved {files_removed} files related to Office Project Tracker.")
    
    if os.path.basename(sys.argv[0]) in ["uninstall.py", "uninstaller.py"]:
        print("\nDon't forget to run the removal script to delete this uninstaller.")
    
    print("\nThank you for using Office Project Tracker!")

if __name__ == "__main__":
    main()
    
    # If running from command line, wait for key press before exiting
    if sys.stdin.isatty():  # Only if running in terminal, not as service
        print("\nPress Enter to exit...")
        input()