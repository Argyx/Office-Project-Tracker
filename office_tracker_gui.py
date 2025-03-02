import sys
import os
import time
import threading
import logging
from datetime import datetime, timedelta
from pathlib import Path
import json
import subprocess
import shutil
import traceback

# Import the tracker code (we'll keep it in a separate file)
import office_tracker

# GUI libraries
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

class OfficeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office Project Tracker")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # Set icon (if available)
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # Variables
        self.is_running = False
        self.last_run = None
        self.scheduler_thread = None
        self.stop_event = threading.Event()
        self.projects_found = 0
        self.run_interval = tk.IntVar(value=24)  # Hours
        
        # Create a settings file path in user's documents
        self.app_data_dir = Path.home() / "Documents" / "OfficeTracker"
        self.app_data_dir.mkdir(exist_ok=True, parents=True)
        self.settings_file = self.app_data_dir / "settings.json"
        self.log_file = self.app_data_dir / "tracker.log"
        
        # Configure logging
        logging.basicConfig(
            filename=self.log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        # Create handlers for the log
        self.console_handler = logging.StreamHandler(sys.stdout)
        self.console_handler.setLevel(logging.INFO)
        logging.getLogger().addHandler(self.console_handler)
        
        # Load settings
        self.settings = self.load_settings()
        
        # Create GUI
        self.create_widgets()
        
        # Update UI with settings
        self.update_settings_ui()
        
        # Check for first run
        if not self.settings.get("api_key") or not self.settings.get("cse_id"):
            self.show_first_run_wizard()
        
        # Start automatic updating of the UI
        self.update_ui()
        
    def load_settings(self):
        default_settings = {
            "api_key": "",
            "cse_id": "",
            "email_username": "",
            "email_password": "",
            "receiver_email": "",
            "smtp_server": "smtp.gmail.com",
            "smtp_port": "587",
            "run_interval": 24,
            "run_on_startup": False,
            "auto_start": False,
            "language": "en"
        }
        
        if self.settings_file.exists():
            try:
                with open(self.settings_file, "r") as f:
                    settings = json.load(f)
                    # Update with any missing default settings
                    for key, value in default_settings.items():
                        if key not in settings:
                            settings[key] = value
                    return settings
            except:
                logging.error("Failed to load settings, using defaults")
                return default_settings
        else:
            return default_settings
    
    def save_settings(self):
        with open(self.settings_file, "w") as f:
            json.dump(self.settings, f, indent=4)
        
        # Update environment variables
        os.environ["GOOGLE_API_KEY"] = self.settings.get("api_key", "")
        os.environ["GOOGLE_CSE_ID"] = self.settings.get("cse_id", "")
        os.environ["EMAIL_USERNAME"] = self.settings.get("email_username", "")
        os.environ["EMAIL_PASSWORD"] = self.settings.get("email_password", "")
        os.environ["RECEIVER_EMAIL"] = self.settings.get("receiver_email", "")
        os.environ["SMTP_SERVER"] = self.settings.get("smtp_server", "smtp.gmail.com")
        os.environ["SMTP_PORT"] = self.settings.get("smtp_port", "587")
        os.environ["EMAIL_LANGUAGE"] = self.settings.get("language", "en")
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Dashboard tab
        dashboard_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(dashboard_frame, text="Dashboard")
        
        # Status frame
        status_frame = ttk.LabelFrame(dashboard_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.X, pady=5)
        
        # Status indicators
        self.status_label = ttk.Label(status_frame, text="Status: Not Running")
        self.status_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.last_run_label = ttk.Label(status_frame, text="Last Run: Never")
        self.last_run_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        self.next_run_label = ttk.Label(status_frame, text="Next Run: Not Scheduled")
        self.next_run_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        
        self.projects_label = ttk.Label(status_frame, text="Projects Found: 0")
        self.projects_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        
        # Control frame
        control_frame = ttk.LabelFrame(dashboard_frame, text="Controls", padding=10)
        control_frame.pack(fill=tk.X, pady=10)
        
        # Run now button
        self.run_button = ttk.Button(control_frame, text="Run Now", command=self.run_once)
        self.run_button.grid(row=0, column=0, padx=5, pady=5)
        
        # Start/Stop scheduler button
        self.scheduler_button = ttk.Button(control_frame, text="Start Scheduler", command=self.toggle_scheduler)
        self.scheduler_button.grid(row=0, column=1, padx=5, pady=5)
        
        # Interval selection
        ttk.Label(control_frame, text="Run every:").grid(row=0, column=2, padx=5, pady=5)
        
        interval_frame = ttk.Frame(control_frame)
        interval_frame.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Radiobutton(interval_frame, text="6h", variable=self.run_interval, value=6).pack(side=tk.LEFT)
        ttk.Radiobutton(interval_frame, text="12h", variable=self.run_interval, value=12).pack(side=tk.LEFT)
        ttk.Radiobutton(interval_frame, text="24h", variable=self.run_interval, value=24).pack(side=tk.LEFT)
        
        # Set run interval from settings
        self.run_interval.set(self.settings.get("run_interval", 24))
        
        # Log viewer
        log_frame = ttk.LabelFrame(dashboard_frame, text="Activity Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # Results tab
        results_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(results_frame, text="Recent Results")
        
        # Results tree view
        results_tree_frame = ttk.Frame(results_frame)
        results_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.results_tree = ttk.Treeview(results_tree_frame, columns=("company", "location", "date", "relevance"))
        self.results_tree.heading("#0", text="Project")
        self.results_tree.heading("company", text="Company")
        self.results_tree.heading("location", text="Location")
        self.results_tree.heading("date", text="Date Added")
        self.results_tree.heading("relevance", text="Relevance")
        
        self.results_tree.column("#0", width=300)
        self.results_tree.column("company", width=150)
        self.results_tree.column("location", width=100)
        self.results_tree.column("date", width=100)
        self.results_tree.column("relevance", width=80)
        
        scrollbar = ttk.Scrollbar(results_tree_frame, orient="vertical", command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Double-click to open URL
        self.results_tree.bind("<Double-1>", self.open_selected_result)
        
        # Settings tab
        settings_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_frame, text="Settings")
        
        # API settings
        api_frame = ttk.LabelFrame(settings_frame, text="Google API Settings", padding=10)
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="Google API Key:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.api_key_entry = ttk.Entry(api_frame, width=50, show="*")
        self.api_key_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(api_frame, text="Custom Search Engine ID:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.cse_id_entry = ttk.Entry(api_frame, width=50)
        self.cse_id_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Email settings
        email_frame = ttk.LabelFrame(settings_frame, text="Email Settings", padding=10)
        email_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(email_frame, text="Email Username:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_username_entry = ttk.Entry(email_frame, width=50)
        self.email_username_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(email_frame, text="Email Password:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.email_password_entry = ttk.Entry(email_frame, width=50, show="*")
        self.email_password_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(email_frame, text="Receiver Email:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.receiver_email_entry = ttk.Entry(email_frame, width=50)
        self.receiver_email_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(email_frame, text="SMTP Server:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.smtp_server_entry = ttk.Entry(email_frame, width=50)
        self.smtp_server_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(email_frame, text="SMTP Port:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.smtp_port_entry = ttk.Entry(email_frame, width=10)
        self.smtp_port_entry.grid(row=4, column=1, sticky="w", padx=5, pady=5)
        
        # Language settings
        language_frame = ttk.LabelFrame(settings_frame, text="Language Settings", padding=10)
        language_frame.pack(fill=tk.X, pady=10)
        
        self.language_var = tk.StringVar(value=self.settings.get("language", "en"))
        ttk.Radiobutton(language_frame, text="English", variable=self.language_var, value="en").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(language_frame, text="Greek", variable=self.language_var, value="el").grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # Other settings
        other_frame = ttk.LabelFrame(settings_frame, text="Other Settings", padding=10)
        other_frame.pack(fill=tk.X, pady=10)
        
        self.startup_var = tk.BooleanVar(value=self.settings.get("run_on_startup", False))
        ttk.Checkbutton(other_frame, text="Run on system startup", variable=self.startup_var).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        self.autostart_var = tk.BooleanVar(value=self.settings.get("auto_start", False))
        ttk.Checkbutton(other_frame, text="Start scheduler automatically when app opens", variable=self.autostart_var).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(settings_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings_from_ui).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Test Email", command=self.test_email).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Open Database", command=self.open_database_location).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Create Startup Shortcut", command=self.create_startup_shortcut).pack(side=tk.LEFT, padx=5)
        
        # Help tab
        help_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(help_frame, text="Help")
        
        help_text = """
        Office Project Tracker
        
        This application automatically searches for office projects and sends you notifications by email.
        
        Quick Start:
        1. Go to the Settings tab and enter your Google API Key and Custom Search Engine ID
        2. Enter your email settings to receive notifications
        3. Click "Save Settings"
        4. Go back to the Dashboard tab and click "Run Now" to test, or "Start Scheduler" for continuous operation
        
        Getting API Keys:
        1. Create a Google API Key at: https://console.cloud.google.com/
        2. Enable the Custom Search API in the Google Cloud Console
        3. Create a Custom Search Engine at: https://programmablesearchengine.google.com/
        4. Ensure you enable "Search the entire web" in your CSE settings
        
        For Gmail:
        You need to enable "Less secure app access" or create an App Password.
        
        Need help? Contact support at: support@example.com
        """
        
        help_text_widget = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD)
        help_text_widget.pack(fill=tk.BOTH, expand=True)
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.config(state=tk.DISABLED)
        
        # Status bar at the bottom
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # If autostart is enabled, start the scheduler
        if self.settings.get("auto_start", False):
            self.root.after(2000, self.toggle_scheduler)
    
    def update_settings_ui(self):
        """Update settings UI with values from settings"""
        self.api_key_entry.delete(0, tk.END)
        self.api_key_entry.insert(0, self.settings.get("api_key", ""))
        
        self.cse_id_entry.delete(0, tk.END)
        self.cse_id_entry.insert(0, self.settings.get("cse_id", ""))
        
        self.email_username_entry.delete(0, tk.END)
        self.email_username_entry.insert(0, self.settings.get("email_username", ""))
        
        self.email_password_entry.delete(0, tk.END)
        self.email_password_entry.insert(0, self.settings.get("email_password", ""))
        
        self.receiver_email_entry.delete(0, tk.END)
        self.receiver_email_entry.insert(0, self.settings.get("receiver_email", ""))
        
        self.smtp_server_entry.delete(0, tk.END)
        self.smtp_server_entry.insert(0, self.settings.get("smtp_server", "smtp.gmail.com"))
        
        self.smtp_port_entry.delete(0, tk.END)
        self.smtp_port_entry.insert(0, self.settings.get("smtp_port", "587"))
        
        self.language_var.set(self.settings.get("language", "en"))
        self.startup_var.set(self.settings.get("run_on_startup", False))
        self.autostart_var.set(self.settings.get("auto_start", False))
        self.run_interval.set(int(self.settings.get("run_interval", 24)))
    
    def save_settings_from_ui(self):
        """Save settings from UI to settings file"""
        self.settings["api_key"] = self.api_key_entry.get()
        self.settings["cse_id"] = self.cse_id_entry.get()
        self.settings["email_username"] = self.email_username_entry.get()
        self.settings["email_password"] = self.email_password_entry.get()
        self.settings["receiver_email"] = self.receiver_email_entry.get()
        self.settings["smtp_server"] = self.smtp_server_entry.get()
        self.settings["smtp_port"] = self.smtp_port_entry.get()
        self.settings["language"] = self.language_var.get()
        self.settings["run_on_startup"] = self.startup_var.get()
        self.settings["auto_start"] = self.autostart_var.get()
        self.settings["run_interval"] = self.run_interval.get()
        
        self.save_settings()
        
        if self.startup_var.get():
            self.create_startup_shortcut()
        
        messagebox.showinfo("Settings Saved", "Your settings have been saved successfully.")
        self.log("Settings updated")
    
    def show_first_run_wizard(self):
        """Show a wizard for first run setup"""
        messagebox.showinfo("Welcome", "Welcome to Office Project Tracker! Let's set up your application.")
        self.notebook.select(2)  # Switch to settings tab
    
    def toggle_scheduler(self):
        """Start or stop the scheduler"""
        if self.is_running:
            # Stop the scheduler
            self.stop_event.set()
            self.is_running = False
            self.scheduler_button.config(text="Start Scheduler")
            self.status_label.config(text="Status: Not Running")
            self.next_run_label.config(text="Next Run: Not Scheduled")
            self.log("Scheduler stopped")
        else:
            # Check if settings are configured
            if not self.settings.get("api_key") or not self.settings.get("cse_id"):
                messagebox.showerror("Configuration Error", "Please configure your API keys in the Settings tab first.")
                self.notebook.select(2)  # Switch to settings tab
                return
            
            # Start the scheduler
            self.stop_event.clear()
            self.is_running = True
            self.scheduler_button.config(text="Stop Scheduler")
            self.status_label.config(text="Status: Running")
            
            # Start the scheduler thread
            self.scheduler_thread = threading.Thread(target=self.scheduler_loop)
            self.scheduler_thread.daemon = True
            self.scheduler_thread.start()
            
            self.log("Scheduler started")
    
    def scheduler_loop(self):
        """Run the scheduler loop in a separate thread"""
        next_run = datetime.now() + timedelta(seconds=10)  # First run after 10 seconds
        
        while not self.stop_event.is_set():
            now = datetime.now()
            
            # Update the next run label in the UI thread
            self.root.after(0, lambda: self.next_run_label.config(text=f"Next Run: {next_run.strftime('%Y-%m-%d %H:%M:%S')}"))
            
            if now >= next_run:
                # Time to run
                self.root.after(0, self.run_search_job)
                
                # Calculate next run time
                hours = self.run_interval.get()
                next_run = datetime.now() + timedelta(hours=hours)
            
            # Sleep for a bit to avoid hogging the CPU
            time.sleep(1)
    
    def run_once(self):
        """Run search once immediately"""
        # Check if settings are configured
        if not self.settings.get("api_key") or not self.settings.get("cse_id"):
            messagebox.showerror("Configuration Error", "Please configure your API keys in the Settings tab first.")
            self.notebook.select(2)  # Switch to settings tab
            return
        
        # Disable the run button to prevent multiple runs
        self.run_button.config(state=tk.DISABLED)
        
        # Run in a separate thread
        threading.Thread(target=self.run_search_job).start()
    
    def run_search_job(self):
        """Run the actual search job"""
        try:
            self.log("Starting office project search...")
            
            # Set status in UI thread
            self.root.after(0, lambda: self.status_bar.config(text="Searching for office projects..."))
            
            # Make sure settings are loaded into environment variables
            self.save_settings()
            
            # Run the search
            result = office_tracker.main()
            
            # Update the last run time
            self.last_run = datetime.now()
            self.root.after(0, lambda: self.last_run_label.config(text=f"Last Run: {self.last_run.strftime('%Y-%m-%d %H:%M:%S')}"))
            
            # Get project count
            self.root.after(1000, self.update_project_count)
            self.root.after(1000, self.update_results_tree)
            
            self.log("Search completed successfully")
            
            # Show notification
            self.root.after(0, lambda: self.status_bar.config(text="Search completed successfully"))
        except Exception as e:
            error_msg = f"Error running search: {str(e)}"
            self.log(error_msg)
            traceback.print_exc()
            self.root.after(0, lambda: self.status_bar.config(text=error_msg))
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred while running the search:\n\n{str(e)}"))
        finally:
            # Re-enable the run button in the UI thread
            self.root.after(0, lambda: self.run_button.config(state=tk.NORMAL))
    
    def update_project_count(self):
        """Update the project count from the database"""
        try:
            conn = office_tracker.sqlite3.connect(os.path.join(self.app_data_dir, "office_projects.db"))
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM projects")
            count = cursor.fetchone()[0]
            conn.close()
            
            self.projects_found = count
            self.projects_label.config(text=f"Projects Found: {count}")
        except Exception as e:
            self.log(f"Error updating project count: {str(e)}")
    
    def update_results_tree(self):
        """Update the results tree view with latest projects"""
        try:
            # Clear existing items
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # Connect to the database
            conn = office_tracker.sqlite3.connect(os.path.join(self.app_data_dir, "office_projects.db"))
            conn.row_factory = office_tracker.sqlite3.Row
            cursor = conn.cursor()
            
            # Get the latest 100 projects
            cursor.execute("""
                SELECT * FROM projects 
                ORDER BY date_added DESC 
                LIMIT 100
            """)
            
            projects = cursor.fetchall()
            conn.close()
            
            # Add to tree view
            for project in projects:
                self.results_tree.insert(
                    "", "end", 
                    text=project["source_title"], 
                    values=(
                        project["company_name"], 
                        project["location"], 
                        project["date_added"].split(" ")[0] if " " in project["date_added"] else project["date_added"],
                        project["relevance_score"]
                    ),
                    tags=(project["source_url"],)
                )
        except Exception as e:
            self.log(f"Error updating results tree: {str(e)}")
    
    def open_selected_result(self, event):
        """Open the selected result in the default web browser"""
        item = self.results_tree.selection()[0]
        source_url = self.results_tree.item(item, "tags")[0]
        
        if source_url:
            import webbrowser
            webbrowser.open(source_url)
    
    def log(self, message):
        """Add a log message to the log viewer"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        # Add to the log viewer in the UI thread
        self.root.after(0, lambda: self.update_log_text(log_message))
        
        # Log to the console and file
        logging.info(message)
    
    def update_log_text(self, message):
        """Update the log text in the UI thread"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def test_email(self):
        """Send a test email to verify email settings"""
        # Save current settings
        self.save_settings_from_ui()
        
        try:
            self.log("Sending test email...")
            
            # Create a simple test email
            import smtplib
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            
            sender_email = self.settings.get("email_username", "")
            sender_password = self.settings.get("email_password", "")
            receiver_email = self.settings.get("receiver_email", "")
            smtp_server = self.settings.get("smtp_server", "smtp.gmail.com")
            smtp_port = int(self.settings.get("smtp_port", "587"))
            
            if not all([sender_email, sender_password, receiver_email]):
                raise ValueError("Please fill in all email settings")
            
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = "Office Project Tracker - Test Email"
            
            body = """
            This is a test email from the Office Project Tracker application.
            
            If you received this email, your email settings are configured correctly.
            
            ---
            Sent on: {}
            """.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
            
            self.log("Test email sent successfully")
            messagebox.showinfo("Success", "Test email sent successfully!")
        except Exception as e:
            error_msg = f"Error sending test email: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def create_startup_shortcut(self):
        """Create a shortcut to run the application at startup"""
        try:
            import platform
            system = platform.system()
            
            if system == "Windows":
                # Windows startup folder
                import winshell
                from win32com.client import Dispatch
                
                startup_folder = winshell.startup()
                shortcut_path = os.path.join(startup_folder, "Office Project Tracker.lnk")
                
                # Get the path to the current executable
                target = sys.executable
                if target.endswith("python.exe"):
                    # Running from script, create a shortcut to the script
                    wdir = os.path.dirname(os.path.abspath(__file__))
                    script = os.path.join(wdir, "run_tracker.py")
                    
                    # Create the shortcut
                    shell = Dispatch('WScript.Shell')
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.Targetpath = target
                    shortcut.Arguments = f'"{script}"'
                    shortcut.WorkingDirectory = wdir
                    shortcut.save()
                else:
                    # Running from executable, create a shortcut to the exe
                    wdir = os.path.dirname(os.path.abspath(target))
                    
                    # Create the shortcut
                    shell = Dispatch('WScript.Shell')
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.Targetpath = target
                    shortcut.WorkingDirectory = wdir
                    shortcut.save()
                
                self.log(f"Created startup shortcut at {shortcut_path}")
                messagebox.showinfo("Success", "Startup shortcut created successfully!")
            
            elif system == "Linux":
                # Linux startup file
                home = os.path.expanduser("~")
                autostart_dir = os.path.join(home, ".config", "autostart")
                
                if not os.path.exists(autostart_dir):
                    os.makedirs(autostart_dir)
                
                desktop_file = os.path.join(autostart_dir, "office-tracker.desktop")
                
                # Get the path to the current executable
                target = sys.executable
                if target.endswith("python"):
                    # Running from script
                    wdir = os.path.dirname(os.path.abspath(__file__))
                    script = os.path.join(wdir, "run_tracker.py")
                    exec_line = f"{target} {script}"
                else:
                    # Running from executable
                    exec_line = target
                
                # Create the desktop file
                with open(desktop_file, "w") as f:
                    f.write(f"""[Desktop Entry]
Type=Application
Exec={exec_line}
Hidden=false
NoDisplay=false
X-GNOME-Autostart-enabled=true
Name=Office Project Tracker
Comment=Automatically track office projects
""")
                
                # Make it executable
                os.chmod(desktop_file, 0o755)
                
                self.log(f"Created startup file at {desktop_file}")
                messagebox.showinfo("Success", "Startup file created successfully!")
            
            elif system == "Darwin":  # macOS
                # macOS LaunchAgents
                home = os.path.expanduser("~")
                launch_agents_dir = os.path.join(home, "Library", "LaunchAgents")
                
                if not os.path.exists(launch_agents_dir):
                    os.makedirs(launch_agents_dir)
                
                plist_file = os.path.join(launch_agents_dir, "com.officetracker.startup.plist")
                
                # Get the path to the current executable
                target = sys.executable
                if "python" in target:
                    # Running from script
                    wdir = os.path.dirname(os.path.abspath(__file__))
                    script = os.path.join(wdir, "run_tracker.py")
                    program_arg = script
                else:
                    # Running from executable
                    program_arg = target
                
                # Create the plist file
                with open(plist_file, "w") as f:
                    f.write(f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.officetracker.startup</string>
    <key>ProgramArguments</key>
    <array>
        <string>{target}</string>
        <string>{program_arg}</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>
""")
                
                self.log(f"Created LaunchAgent at {plist_file}")
                messagebox.showinfo("Success", "Startup item created successfully!")
            
            else:
                messagebox.showwarning("Unsupported OS", f"Creating startup shortcuts on {system} is not supported yet.")
        
        except Exception as e:
            error_msg = f"Error creating startup shortcut: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def open_database_location(self):
        """Open the database file location in file explorer"""
        db_path = os.path.join(self.app_data_dir, "office_projects.db")
        
        if not os.path.exists(db_path):
            messagebox.showinfo("Database Not Found", "The database file hasn't been created yet. Run the tracker first.")
            return
        
        # Open the folder containing the database
        import platform
        system = platform.system()
        
        try:
            if system == "Windows":
                os.startfile(os.path.dirname(db_path))
            elif system == "Darwin":  # macOS
                subprocess.run(["open", os.path.dirname(db_path)])
            else:  # Linux
                subprocess.run(["xdg-open", os.path.dirname(db_path)])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open database location: {str(e)}")
    
    def update_ui(self):
        """Update UI elements periodically"""
        # Check for log file changes
        self.update_log_from_file()
        
        # Schedule next update
        self.root.after(1000, self.update_ui)
    
    def update_log_from_file(self):
        """Update log text from log file"""
        if os.path.exists(self.log_file):
            try:
                with open(self.log_file, "r") as f:
                    # Get the last 20 lines
                    lines = f.readlines()[-20:]
                    
                    # Update the log text
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.delete(1.0, tk.END)
                    self.log_text.insert(tk.END, "".join(lines))
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
            except:
                pass

def run_gui():
    # Copy the office_tracker module to the app data directory
    app_data_dir = Path.home() / "Documents" / "OfficeTracker"
    app_data_dir.mkdir(exist_ok=True, parents=True)
    
    # Copy the office_tracker.py file
    src_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "office_tracker.py")
    dst_file = os.path.join(app_data_dir, "office_tracker.py")
    
    if not os.path.exists(dst_file) or os.path.getmtime(src_file) > os.path.getmtime(dst_file):
        shutil.copy2(src_file, dst_file)
    
    # Add app_data_dir to path so we can import office_tracker
    sys.path.insert(0, str(app_data_dir))
    
    # Create the root window
    root = tk.Tk()
    app = OfficeTrackerApp(root)
    
    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    run_gui()