import os
import sys
import subprocess
import shutil
import tempfile
import winreg
import ctypes
import urllib.request
import zipfile
from pathlib import Path

# ASCII art banner
BANNER = """
 ╔═════════════════════════════════════════════════════╗
 ║                                                     ║
 ║          OFFICE PROJECT TRACKER INSTALLER           ║
 ║                                                     ║
 ╚═════════════════════════════════════════════════════╝
"""

# Configuration
APP_NAME = "Office Project Tracker"
APP_VERSION = "1.0.0"
APP_DATA_DIR = os.path.join(os.path.expanduser("~"), "Documents", "OfficeTracker")
STARTUP_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")

# URLs for dependencies
PYTHON_URL = "https://www.python.org/ftp/python/3.10.8/python-3.10.8-embed-amd64.zip"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

# Required packages
REQUIRED_PACKAGES = [
    "requests",
    "beautifulsoup4",
    "nltk",
    "python-dotenv",
    "langdetect",
    "tkinter",
    "pywin32"
]

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def print_step(message):
    """Print a step message."""
    print("\n" + "=" * 70)
    print(f"  {message}")
    print("=" * 70)

def create_desktop_shortcut(target_path, name, description=""):
    """Create a desktop shortcut for the application."""
    try:
        import winshell
        from win32com.client import Dispatch
        
        desktop = winshell.desktop()
        shortcut_path = os.path.join(desktop, f"{name}.lnk")
        
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.Description = description
        shortcut.IconLocation = target_path
        shortcut.save()
        
        print(f"Created desktop shortcut: {shortcut_path}")
        return True
    except Exception as e:
        print(f"Warning: Could not create desktop shortcut: {e}")
        return False

def create_startup_entry(target_path, args=""):
    """Create a startup entry for the application."""
    try:
        import winshell
        from win32com.client import Dispatch
        
        if not os.path.exists(STARTUP_DIR):
            os.makedirs(STARTUP_DIR)
        
        shortcut_path = os.path.join(STARTUP_DIR, f"{APP_NAME}.lnk")
        
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.Arguments = args
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.save()
        
        print(f"Created startup shortcut: {shortcut_path}")
        return True
    except Exception as e:
        print(f"Warning: Could not create startup entry: {e}")
        return False

def install_python():
    """Download and install Python if needed."""
    print_step("Checking Python installation")
    
    python_installed = shutil.which("python") is not None
    
    if python_installed:
        print("Python is already installed.")
        return True
    
    print("Python not found. Downloading and installing...")
    
    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()
    python_zip = os.path.join(temp_dir, "python.zip")
    
    try:
        # Download Python
        print("Downloading Python...")
        urllib.request.urlretrieve(PYTHON_URL, python_zip)
        
        # Extract Python
        print("Extracting Python...")
        python_dir = os.path.join(APP_DATA_DIR, "python")
        os.makedirs(python_dir, exist_ok=True)
        
        with zipfile.ZipFile(python_zip, 'r') as zip_ref:
            zip_ref.extractall(python_dir)
        
        # Add to PATH
        print("Adding Python to PATH...")
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_ALL_ACCESS)
            path_value, _ = winreg.QueryValueEx(key, "PATH")
            
            if python_dir not in path_value:
                new_path = f"{path_value};{python_dir}"
                winreg.SetValueEx(key, "PATH", 0, winreg.REG_EXPAND_SZ, new_path)
                winreg.CloseKey(key)
        except Exception as e:
            print(f"Warning: Could not update PATH: {e}")
        
        # Download and install pip
        print("Installing pip...")
        get_pip = os.path.join(temp_dir, "get-pip.py")
        urllib.request.urlretrieve(GET_PIP_URL, get_pip)
        
        subprocess.run([os.path.join(python_dir, "python.exe"), get_pip], check=True)
        
        print("Python installation complete!")
        return True
    
    except Exception as e:
        print(f"Error installing Python: {e}")
        return False
    finally:
        # Clean up temporary directory
        shutil.rmtree(temp_dir, ignore_errors=True)

def install_dependencies():
    """Install required Python packages."""
    print_step("Installing required packages")
    
    python_exe = shutil.which("python")
    if not python_exe:
        python_exe = os.path.join(APP_DATA_DIR, "python", "python.exe")
    
    pip_exe = shutil.which("pip")
    if not pip_exe:
        pip_exe = os.path.join(APP_DATA_DIR, "python", "Scripts", "pip.exe")
    
    if not os.path.exists(pip_exe):
        print(f"Error: pip not found at {pip_exe}")
        return False
    
    try:
        for package in REQUIRED_PACKAGES:
            print(f"Installing {package}...")
            subprocess.run([pip_exe, "install", package], check=True)
        
        # Download NLTK data
        print("Downloading NLTK data...")
        subprocess.run([
            python_exe, "-c", 
            "import nltk; nltk.download('punkt'); nltk.download('stopwords')"
        ], check=True)
        
        print("All packages installed successfully!")
        return True
    except Exception as e:
        print(f"Error installing packages: {e}")
        return False

def copy_application_files():
    """Copy application files to the application directory."""
    print_step("Installing application files")
    
    os.makedirs(APP_DATA_DIR, exist_ok=True)
    
    # Copy the main application files
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Copy office_tracker.py
        source_file = os.path.join(script_dir, "office_tracker.py")
        target_file = os.path.join(APP_DATA_DIR, "office_tracker.py")
        shutil.copy2(source_file, target_file)
        print(f"Copied {source_file} to {target_file}")
        
        # Copy GUI application
        source_file = os.path.join(script_dir, "office_tracker_gui.py")
        target_file = os.path.join(APP_DATA_DIR, "office_tracker_gui.py")
        shutil.copy2(source_file, target_file)
        print(f"Copied {source_file} to {target_file}")
        
        # Create a launcher script
        launcher_path = os.path.join(APP_DATA_DIR, "run_tracker.py")
        with open(launcher_path, "w") as f:
            f.write("""import sys
import os
import runpy

# Add the current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Run the GUI application
runpy.run_path('office_tracker_gui.py', run_name='__main__')
""")
        print(f"Created launcher script at {launcher_path}")
        
        # Create a .env file template
        env_path = os.path.join(APP_DATA_DIR, ".env.template")
        with open(env_path, "w") as f:
            f.write("""# Google API Settings
GOOGLE_API_KEY=your_google_api_key
GOOGLE_CSE_ID=your_custom_search_engine_id

# Email Settings
EMAIL_USERNAME=your_email@gmail.com
EMAIL_PASSWORD=your_email_password
RECEIVER_EMAIL=recipient@example.com
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Application Settings
EMAIL_LANGUAGE=en
MAX_SEARCH_QUERIES=30
SEND_ANALYTICS=false
""")
        print(f"Created .env template at {env_path}")
        
        # Create a bat file for easy launching
        bat_path = os.path.join(APP_DATA_DIR, "Start Office Tracker.bat")
        with open(bat_path, "w") as f:
            python_exe = shutil.which("python")
            if not python_exe:
                python_exe = os.path.join(APP_DATA_DIR, "python", "python.exe")
            
            f.write(f"""@echo off
echo Starting Office Project Tracker...
"{python_exe}" "{launcher_path}"
""")
        print(f"Created BAT launcher at {bat_path}")
        
        print("Application files installed successfully!")
        return True, bat_path
    except Exception as e:
        print(f"Error copying application files: {e}")
        return False, None

def create_task_scheduler_job():
    """Create a Windows Task Scheduler job to run the tracker in the background."""
    print_step("Setting up background Task Scheduler job")
    
    try:
        python_exe = shutil.which("python")
        if not python_exe:
            python_exe = os.path.join(APP_DATA_DIR, "python", "python.exe")
        
        launcher_path = os.path.join(APP_DATA_DIR, "office_tracker.py")
        
        # Create XML file for the task
        xml_path = os.path.join(APP_DATA_DIR, "office_tracker_task.xml")
        
        # Get current username
        username = os.environ.get("USERNAME")
        
        with open(xml_path, "w") as f:
            f.write(f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Runs Office Project Tracker on a schedule to find and report on new office projects</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2023-01-01T09:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>"{python_exe}"</Command>
      <Arguments>"{launcher_path}"</Arguments>
      <WorkingDirectory>"{APP_DATA_DIR}"</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
""")
        
        # Import the task using schtasks
        task_name = "OfficeProjectTracker"
        subprocess.run([
            "schtasks", "/create", "/tn", task_name, 
            "/xml", xml_path, "/f"
        ], check=True)
        
        print(f"Created scheduled task '{task_name}' to run daily at 9:00 AM")
        return True
    except Exception as e:
        print(f"Warning: Could not create scheduled task: {e}")
        return False

def main():
    """Main installer function."""
    print(BANNER)
    print(f"\nWelcome to the {APP_NAME} Installer (v{APP_VERSION})\n")
    
    # Check if running as administrator
    if not is_admin():
        print("Warning: This installer is not running with administrator privileges.")
        print("Some features may not work correctly.")
        response = input("Continue anyway? (y/n): ")
        if response.lower() != 'y':
            print("Installation cancelled.")
            return
    
    # Create application directory
    print_step("Creating application directory")
    os.makedirs(APP_DATA_DIR, exist_ok=True)
    print(f"Application will be installed to: {APP_DATA_DIR}")
    
    # Install Python
    if not install_python():
        print("Error: Failed to install Python. Installation cannot continue.")
        return
    
    # Install dependencies
    if not install_dependencies():
        print("Error: Failed to install dependencies. Installation cannot continue.")
        return
    
    # Copy application files
    success, bat_path = copy_application_files()
    if not success:
        print("Error: Failed to copy application files. Installation cannot continue.")
        return
    
    # Create shortcuts
    print_step("Creating shortcuts")
    
    # Desktop shortcut
    create_desktop_shortcut(bat_path, APP_NAME, "Track and monitor office projects")
    
    # Startup shortcut (optional)
    response = input("Do you want the application to start automatically when you log in? (y/n): ")
    if response.lower() == 'y':
        create_startup_entry(bat_path)
    
    # Create scheduled task (optional)
    response = input("Do you want to set up a background task to check for projects daily? (y/n): ")
    if response.lower() == 'y':
        create_task_scheduler_job()
    
    print_step("Installation Complete!")
    print(f"\n{APP_NAME} has been successfully installed!")
    print("\nYou can now:")
    print(f"  1. Launch the application from your desktop shortcut")
    print(f"  2. Launch from: {bat_path}")
    print("\nOn first launch, you'll need to configure your API keys and email settings.")
    
    # Open application
    response = input("\nDo you want to start the application now? (y/n): ")
    if response.lower() == 'y':
        print(f"Starting {APP_NAME}...")
        subprocess.Popen([bat_path], shell=True)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInstallation cancelled by user.")
    except Exception as e:
        print(f"\n\nAn unexpected error occurred: {e}")
    
    input("\nPress Enter to exit...")