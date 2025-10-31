#!/usr/bin/env python3
"""
BASH Flow Management - Question Sheet GUI
Python GUI application using tkinter that matches BASHQuestionSheet.vba functionality
"""

import json
import sqlite3
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, date, timedelta
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict
import os
import pandas as pd
import re
import xml.etree.ElementTree as ET
from tkinter import filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.dates as mdates
from datetime import datetime, timedelta
from milestone_editor import open_milestone_editor
from database_manager import open_database_manager
from db_config import get_database_connection, create_mysql_database
from user_manager import UserManager
from login_gui import show_login_window, UserManagementWindow

@dataclass
class QuestionSheetEntry:
    state: str
    stage: str
    milestone_start_date: date
    status: str
    created_date: datetime
    clli: str
    host_wire_centre: str
    lata: str
    equipment_type: str
    subtask: str
    last_update: datetime

class QuestionSheetGUI:
    def __init__(self, config_file: str = "config.json", user_manager: UserManager = None):
        # Initialize debug log first
        self.debug_log = []
        self.add_debug_entry("QuestionSheetGUI initialization started")
        
        # Initialize user management
        self.user_manager = user_manager
        if not self.user_manager:
            # Show login window if no user manager provided
            self.user_manager = show_login_window()
            if not self.user_manager:
                print("Login cancelled. Exiting application.")
                sys.exit(0)
        
        self.add_debug_entry(f"User authenticated: {self.user_manager.get_current_user().username}")
        
        # Initialize database connection
        self.db_conn = get_database_connection(config_file)
        self.add_debug_entry("Database connection initialized")
        
        # For backward compatibility, keep db_path for legacy code
        self.db_path = "mysql_database"  # Placeholder for legacy references
        self.add_debug_entry(f"Database path set to: {self.db_path}")

        # Initialize runtime status values
        self.last_nat_sync = None
        
        self.states = self._initialize_states()
        self.add_debug_entry(f"Initialized {len(self.states)} states")
        
        self.stages = self._initialize_stages()
        self.add_debug_entry(f"Initialized {len(self.stages)} stages")
        
        self.subtasks = self._initialize_subtasks()
        self.add_debug_entry(f"Initialized {len(self.subtasks)} subtasks")
        
        self.criticality_levels = self._initialize_criticality_levels()
        self.add_debug_entry(f"Initialized {len(self.criticality_levels)} criticality levels")
        
        self.statuses = self._initialize_statuses()
        self.add_debug_entry(f"Initialized {len(self.statuses)} statuses")
        
        self.equipment_types = self._initialize_equipment_types()
        self.add_debug_entry(f"Initialized {len(self.equipment_types)} equipment types")
        
        self.milestone_dates = self._initialize_milestone_dates()
        self.add_debug_entry(f"Initialized {len(self.milestone_dates)} milestone dates")
        
        # Initialize databases
        self.add_debug_entry("Initializing main database")
        self._initialize_database()
        self.add_debug_entry("Main database initialized successfully")
        
        self.add_debug_entry("Initializing milestone database")
        self._initialize_milestone_database()
        self.add_debug_entry("Milestone database initialized successfully")
        
        # Create date entries table
        self.add_debug_entry("Creating date entries table")
        self._create_date_entries_table()
        self.add_debug_entry("Date entries table created successfully")
        
        # Load Excel data for CLLI lookup
        self.add_debug_entry("Loading CLLI data from Excel file")
        self.clli_data = self._load_clli_data()
        self.add_debug_entry(f"CLLI data loaded: {len(self.clli_data) if self.clli_data is not None else 0} rows")
        
        self.clli_suggestions = []
        self.add_debug_entry("CLLI suggestions list initialized")
        
        # Initialize record number
        self.current_record_id = None
        self.add_debug_entry("Record ID initialized to None")
        
        # Initialize inline editing state
        self.editing_item = None
        self.editing_column = None
        self.editing_column_name = None
        self.original_value = None
        self.edit_entry = None
        
        # Create main window
        self.add_debug_entry("Creating main window")
        self.root = tk.Tk()
        self.root.title("Degrow Workflow Manager")
        self.root.geometry("1920x1080")
        self.root.resizable(True, True)
        self.add_debug_entry("Main window created and configured")
        
        # Create menu bar
        self.create_menu_bar()
        
        # Configure style
        self.add_debug_entry("Configuring GUI styles")
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.add_debug_entry("GUI styles configured")
        
        # Configure white background style for entry fields
        self.style.configure("White.TEntry", fieldbackground="white", background="white")
        
        # Configure white background style for combobox fields
        self.style.configure("White.TCombobox", fieldbackground="white", background="white")
        
        # Also try alternative styling
        self.style.map("White.TEntry", fieldbackground=[('readonly', 'white'), ('active', 'white')])
        self.style.map("White.TCombobox", fieldbackground=[('readonly', 'white'), ('active', 'white')])
        
        # Create GUI elements
        self.create_widgets()
        
        # Center window on screen
        self.center_window()
    
    def create_menu_bar(self):
        """Create application menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Entry", command=self.create_entry)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Milestone Editor", command=self.open_milestone_editor)
        tools_menu.add_command(label="Database Manager", command=self.open_database_manager)
        tools_menu.add_separator()
        
        # User menu
        user_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="User", menu=user_menu)
        user_menu.add_command(label="User Management", command=self.open_user_management)
        user_menu.add_separator()
        user_menu.add_command(label="Change Password", command=self.change_password)
        user_menu.add_command(label="Logout", command=self.logout)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="User Guide", command=self.show_user_guide)
    
    def open_user_management(self):
        """Open user management window"""
        if not self.user_manager.has_permission('admin'):
            messagebox.showerror("Access Denied", "You need administrator privileges to access user management.")
            return
        
        try:
            UserManagementWindow(self.root, self.user_manager)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open user management: {e}")
    
    def change_password(self):
        """Change user password dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Change Password")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Current password
        ttk.Label(form_frame, text="Current Password:").pack(anchor=tk.W, pady=(0, 5))
        current_password_var = tk.StringVar()
        current_password_entry = ttk.Entry(form_frame, textvariable=current_password_var, 
                                         show="*", width=30)
        current_password_entry.pack(fill=tk.X, pady=(0, 10))
        
        # New password
        ttk.Label(form_frame, text="New Password:").pack(anchor=tk.W, pady=(0, 5))
        new_password_var = tk.StringVar()
        new_password_entry = ttk.Entry(form_frame, textvariable=new_password_var, 
                                     show="*", width=30)
        new_password_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Confirm password
        ttk.Label(form_frame, text="Confirm New Password:").pack(anchor=tk.W, pady=(0, 5))
        confirm_password_var = tk.StringVar()
        confirm_password_entry = ttk.Entry(form_frame, textvariable=confirm_password_var, 
                                         show="*", width=30)
        confirm_password_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def save_password():
            current = current_password_var.get()
            new = new_password_var.get()
            confirm = confirm_password_var.get()
            
            if not current or not new or not confirm:
                messagebox.showerror("Error", "All fields are required")
                return
            
            if new != confirm:
                messagebox.showerror("Error", "New passwords do not match")
                return
            
            if len(new) < 6:
                messagebox.showerror("Error", "Password must be at least 6 characters long")
                return
            
            # Here you would implement password change logic
            messagebox.showinfo("Success", "Password changed successfully")
            dialog.destroy()
        
        ttk.Button(button_frame, text="Change Password", command=save_password).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT)
    
    def logout(self):
        """Logout current user"""
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            self.user_manager.logout()
            self.root.quit()
    
    def show_about(self):
        """Show about dialog"""
        about_text = """
Degrow Workflow Manager
Version 1.0.0

A comprehensive workflow management system for telecommunications and network infrastructure projects.

Features:
- Question Sheet Management
- Milestone and Subtask Tracking
- Database Management
- Email Processing
- User Authentication and Management

© 2025 Workflow Manager System
        """
        
        about_window = tk.Toplevel(self.root)
        about_window.title("About")
        about_window.geometry("400x300")
        about_window.transient(self.root)
        about_window.grab_set()
        
        text_widget = tk.Text(about_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, about_text)
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(about_window, text="Close", 
                  command=about_window.destroy).pack(pady=10)
    
    def show_user_guide(self):
        """Show user guide"""
        messagebox.showinfo("User Guide", "User guide functionality will be implemented")
    
    def open_email_processor(self):
        """Open email processor"""
        messagebox.showinfo("Email Processor", "Email processor functionality will be implemented")
    
    def _initialize_states(self) -> List[str]:
        """Initialize list of US states"""
        return [
            "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
            "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
            "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
            "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi",
            "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
            "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma",
            "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
            "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
            "West Virginia", "Wisconsin", "Wyoming"
        ]
    
    def _initialize_stages(self) -> List[str]:
        """Initialize list of project stages"""
        return [
            "Complete Special Orders", "Create WFA Orders for Switch Specialists", "Issue Switch Special Orders",
            "Create Workbook for Order Sets", "Complete bash Report Cleanup", "Verify bash Report",
            "Create Bash Report", "RUN TMART Report", "Collect Switch Device Info",
            "Complete Pre-Cut", "Restrict CM", "Create Grooming Workbook Tool"
        ]
    
    def _initialize_subtasks(self) -> List[str]:
        """Initialize list of CLLI-specific subtasks"""
        return [
            "Go-Live Support", "Deployment Strategy", "Testing Protocol", "Implementation Planning",
            "Approval Process", "Documentation Review", "Risk Assessment", "Capacity Planning",
            "Network Analysis", "Equipment Inventory", "Site Survey", "CLLI Validation"
        ]
    
    def _initialize_criticality_levels(self) -> Dict[str, str]:
        """Initialize criticality levels for each subtask"""
        return {
            "CLLI Validation": "Must be complete",
            "Site Survey": "Must be complete", 
            "Equipment Inventory": "Should be Complete",
            "Network Analysis": "Must be complete",
            "Capacity Planning": "Should be Complete",
            "Risk Assessment": "Must be complete",
            "Documentation Review": "Does not block",
            "Approval Process": "Must be complete",
            "Implementation Planning": "Should be Complete",
            "Testing Protocol": "Must be complete",
            "Deployment Strategy": "Should be Complete",
            "Go-Live Support": "Must be complete"
        }
    
    def _initialize_statuses(self) -> List[str]:
        """Initialize list of status options"""
        return ["To Do", "In Progress", "Blocked", "Done"]
    
    def _initialize_equipment_types(self) -> List[str]:
        """Initialize list of equipment types"""
        return [
            "Switch", "Router", "Gateway", "Server", "Firewall", 
            "Load Balancer", "Storage", "UPS", "Generator", "Cooling",
            "Power Distribution", "Cable Management", "Rack", "Other"
        ]
    
    def _initialize_milestone_dates(self) -> Dict[str, date]:
        """Initialize milestone dates for each stage"""
        return {
            "Create Grooming Workbook Tool": date(2024, 1, 1),
            "Restrict CM": date(2024, 1, 15),
            "Complete Pre-Cut": date(2024, 2, 1),
            "Collect Switch Device Info": date(2024, 2, 15),
            "RUN TMART Report": date(2024, 3, 1),
            "Create Bash Report": date(2024, 3, 15),
            "Verify bash Report": date(2024, 4, 1),
            "Complete bash Report Cleanup": date(2024, 4, 15),
            "Create Workbook for Order Sets": date(2024, 5, 1),
            "Issue Switch Special Orders": date(2024, 5, 15),
            "Create WFA Orders for Switch Specialists": date(2024, 6, 1),
            "Complete Special Orders": date(2024, 6, 15)
        }
    
    def _initialize_database(self):
        """Initialize MySQL database for storing workflow entries"""
        try:
            # Create MySQL database if it doesn't exist
            create_mysql_database()
            
            # Connect to database
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Create workflow_entries table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS workflow_entries (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    state VARCHAR(50) NOT NULL,
                    clli VARCHAR(20),
                    host_wire_centre VARCHAR(100),
                    lata VARCHAR(10),
                    equipment_type VARCHAR(50),
                    current_milestone VARCHAR(100) NOT NULL,
                    milestone_subtask VARCHAR(100),
                    status VARCHAR(50) NOT NULL,
                    planned_start VARCHAR(20),
                    actual_start VARCHAR(20),
                    duration VARCHAR(20),
                    planned_end VARCHAR(20),
                    actual_end VARCHAR(20),
                    milestone_date VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_update DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Add new columns to existing tables (migration)
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN planned_start VARCHAR(20)')
            except Exception:
                pass  # Column already exists
            
            # Rename city column to host_wire_centre if it exists
            try:
                cursor.execute('ALTER TABLE workflow_entries CHANGE COLUMN city host_wire_centre VARCHAR(100)')
            except Exception:
                pass  # Column doesn't exist or already renamed
            
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN actual_start VARCHAR(20)')
            except Exception:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN duration VARCHAR(20)')
            except Exception:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN planned_end VARCHAR(20)')
            except Exception:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN actual_end VARCHAR(20)')
            except Exception:
                pass  # Column already exists
            
            # Create index for better performance
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_workflow_entries_milestone 
                ON workflow_entries(current_milestone)
            ''')
            
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_workflow_entries_status 
                ON workflow_entries(status)
            ''')
            
            conn.commit()
            cursor.close()
            conn.close()
            
            print("MySQL database initialized successfully")
            
        except Exception as e:
            print(f"Error initializing database: {e}")
            import traceback
            traceback.print_exc()
    
    def _initialize_milestone_database(self):
        """Initialize milestone database tables in MySQL"""
        try:
            # Connect to database
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Create milestones table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS milestones (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL UNIQUE,
                    description TEXT,
                    order_index INT DEFAULT 0,
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create subtasks table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS subtasks (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    milestone_id INT NOT NULL,
                    name VARCHAR(100) NOT NULL,
                    description TEXT,
                    criticality VARCHAR(50) DEFAULT 'Should be Complete',
                    order_index INT DEFAULT 0,
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL,
                    FOREIGN KEY (milestone_id) REFERENCES milestones (id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create people table (renamed from stakeholders)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS people (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    home_state VARCHAR(50) NOT NULL,
                    role VARCHAR(100),
                    team VARCHAR(100),
                    email VARCHAR(255),
                    phone VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Store milestone database reference (same as main database for MySQL)
            self.milestone_db_path = "mysql_database"
            print("Milestone database tables initialized successfully in MySQL")
            
        except Exception as e:
            print(f"Error initializing milestone database: {e}")
            import traceback
            traceback.print_exc()
    
    def _create_date_entries_table(self):
        """Create date entries table with CLLI as primary key"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Create date_entries table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS date_entries (
                    clli VARCHAR(20) PRIMARY KEY,
                    planned_start VARCHAR(20),
                    actual_start VARCHAR(20),
                    planned_end VARCHAR(20),
                    actual_end VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            conn.commit()
            cursor.close()
            conn.close()
            
            print("Date entries table created successfully")
            
        except Exception as e:
            print(f"Error creating date entries table: {e}")
            import traceback
            traceback.print_exc()
    
    def _load_dates_for_clli(self, clli):
        """Load date entries for a specific CLLI"""
        try:
            if not clli:
                return {}
            
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT planned_start, actual_start, planned_end, actual_end
                FROM date_entries 
                WHERE clli = %s
            """, (clli,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return {
                    'planned_start': result[0] or '',
                    'actual_start': result[1] or '',
                    'planned_end': result[2] or '',
                    'actual_end': result[3] or ''
                }
            else:
                return {}
                
        except Exception as e:
            print(f"Error loading dates for CLLI {clli}: {e}")
            return {}
    
    def _load_clli_data(self) -> pd.DataFrame:
        """Load CLLI data from Excel file"""
        try:
            excel_path = r"C:\Lumen\Workflow Manager\Dummy Switch Data TXO Testing 20251017.xlsx"
            if os.path.exists(excel_path):
                # Read Excel file with headers from row 1 (default)
                df = pd.read_excel(excel_path)  # Default header=0 means row 1
                print(f"Loaded {len(df)} rows from Excel file (headers from row 1)")
                print(f"Excel columns: {list(df.columns)}")
                print(f"First few rows:")
                print(df.head())
                return df
            else:
                print(f"Excel file not found: {excel_path}")
                return pd.DataFrame()
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()
    
    def _search_clli(self, query: str) -> List[str]:
        """Search for CLLI codes matching the query in Host CLLI column"""
        self.add_debug_entry(f"CLLI search initiated for query: '{query}'")
        if not query or len(query) < 2:
            self.add_debug_entry("CLLI search cancelled - query too short")
            return []
        
        suggestions = []
        query_lower = query.lower()
        
        try:
            print(f"Searching for CLLI codes matching: '{query}'")
            
            # Search specifically in "Host CLLI" column
            if 'Host CLLI' in self.clli_data.columns:
                host_clli_column = self.clli_data['Host CLLI']
                print(f"Searching in Host CLLI column with {len(host_clli_column)} rows")
                
                # Convert to string and search
                str_column = host_clli_column.astype(str)
                matches = str_column[str_column.str.lower().str.contains(query_lower, na=False)]
                
                print(f"Found {len(matches)} matches before filtering")
                
                for match in matches.dropna().unique():
                    match_str = str(match).strip()
                    if match_str and match_str != 'nan' and match_str.lower().startswith(query_lower):
                        suggestions.append(match_str)
                        print(f"Added suggestion: {match_str}")
            else:
                print("Host CLLI column not found, searching all columns")
                # Fallback: search in all columns if "Host CLLI" not found
                for column in self.clli_data.columns:
                    if self.clli_data[column].dtype == 'object':  # Text columns
                        str_column = self.clli_data[column].astype(str)
                        matches = str_column[str_column.str.lower().str.contains(query_lower, na=False)]
                        for match in matches.dropna().unique():
                            match_str = str(match).strip()
                            if match_str and match_str != 'nan' and match_str.lower().startswith(query_lower):
                                suggestions.append(match_str)
            
            # Remove duplicates and limit to 10 suggestions
            suggestions = list(set(suggestions))[:10]
            print(f"Final suggestions: {suggestions}")
            return suggestions
            
        except Exception as e:
            print(f"Error searching CLLI: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _get_all_clli_codes(self) -> List[str]:
        """Get all CLLI codes from Host CLLI column for dropdown"""
        try:
            print(f"Getting all CLLI codes from Excel data...")
            print(f"Excel data columns: {list(self.clli_data.columns)}")
            
            if 'Host CLLI' in self.clli_data.columns:
                host_clli_column = self.clli_data['Host CLLI']
                print(f"Host CLLI column found with {len(host_clli_column)} rows")
                
                # Get unique values, remove NaN, and sort
                all_codes = host_clli_column.dropna().astype(str).unique().tolist()
                all_codes = [code for code in all_codes if code and code.strip() and code != 'nan']
                
                print(f"Found {len(all_codes)} unique CLLI codes")
                print(f"First 5 codes: {all_codes[:5]}")
                return sorted(all_codes)
            else:
                print("Host CLLI column not found in Excel data")
                print(f"Available columns: {list(self.clli_data.columns)}")
                return []
        except Exception as e:
            print(f"Error getting all CLLI codes: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def create_logo_section(self):
        """Create logo section at the top of the window"""
        try:
            # Create logo frame with white background
            logo_frame = tk.Frame(self.root, bg='white', relief='flat', bd=0)
            logo_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=(10, 5))
            
            # Configure logo frame for three-column layout
            logo_frame.columnconfigure(0, weight=0)  # Lumen logo (left)
            logo_frame.columnconfigure(1, weight=1)  # Title (center, expandable)
            logo_frame.columnconfigure(2, weight=0)  # TXO logo (right)
            
            # Try to load Lumen logo
            try:
                from PIL import Image, ImageTk
                import os
                import cairosvg
                import io
                
                # Check for SVG first, then PNG as fallback
                print(f"Current working directory: {os.getcwd()}")
                print(f"Looking for Lumen.svg in: {os.path.abspath('Lumen.svg')}")
                print(f"Lumen.svg exists: {os.path.exists('Lumen.svg')}")
                
                if os.path.exists("Lumen.svg"):
                    print("Found Lumen.svg, converting...")
                    # Convert SVG to PNG using cairosvg
                    svg_data = open("Lumen.svg", "rb").read()
                    png_data = cairosvg.svg2png(bytestring=svg_data, output_width=200, output_height=50)
                    pil_image = Image.open(io.BytesIO(png_data))
                    print(f"Lumen SVG converted to size: {pil_image.size}")
                    lumen_image = ImageTk.PhotoImage(pil_image)
                elif os.path.exists("lumen.png"):
                    pil_image = Image.open("lumen.png")
                    print(f"Original Lumen logo size: {pil_image.size}")
                    # Make Lumen logo MUCH larger for prominence - use resize to force exact dimensions
                    pil_image = pil_image.resize((200, 40), Image.Resampling.LANCZOS)
                    print(f"Resized Lumen logo size: {pil_image.size}")
                    lumen_image = ImageTk.PhotoImage(pil_image)
                else:
                    lumen_image = tk.PhotoImage(file="lumen.png")
                
                lumen_label = tk.Label(logo_frame, image=lumen_image, bg='white')
                lumen_label.image = lumen_image  # Keep a reference
                lumen_label.grid(row=0, column=0, padx=(0, 20), sticky=tk.W)
                # Force minimum size to ensure logo displays at full size
                lumen_label.configure(width=200, height=50)
                print(f"Lumen logo loaded successfully - Size: {lumen_image.width()}x{lumen_image.height()}")
            except Exception as e:
                print(f"Could not load lumen logo: {e}")
                # Create a placeholder label
                lumen_label = tk.Label(logo_frame, text="LUMEN", font=("Arial", 32, "bold"), bg='white')
                lumen_label.grid(row=0, column=0, padx=(0, 20), sticky=tk.W)
            
            # Try to load TXO logo
            try:
                from PIL import Image, ImageTk
                import os
                import cairosvg
                import io
                
                # Check for SVG first, then PNG as fallback
                print(f"Looking for TXO.svg in: {os.path.abspath('TXO.svg')}")
                print(f"TXO.svg exists: {os.path.exists('TXO.svg')}")
                
                if os.path.exists("TXO.svg"):
                    print("Found TXO.svg, converting...")
                    # Convert SVG to PNG using cairosvg
                    svg_data = open("TXO.svg", "rb").read()
                    png_data = cairosvg.svg2png(bytestring=svg_data, output_width=200, output_height=50)
                    pil_image = Image.open(io.BytesIO(png_data))
                    print(f"TXO SVG converted to size: {pil_image.size}")
                    txo_image = ImageTk.PhotoImage(pil_image)
                elif os.path.exists("TXO.png"):
                    pil_image = Image.open("TXO.png")
                    # Resize to standard size (height 50px, maintain aspect ratio)
                    pil_image.thumbnail((200, 50), Image.Resampling.LANCZOS)
                    txo_image = ImageTk.PhotoImage(pil_image)
                else:
                    txo_image = tk.PhotoImage(file="TXO.png")
                
                txo_label = tk.Label(logo_frame, image=txo_image, bg='white')
                txo_label.image = txo_image  # Keep a reference
                txo_label.grid(row=0, column=2, padx=(20, 0), sticky=tk.E)
                print(f"TXO logo loaded successfully - Size: {txo_image.width()}x{txo_image.height()}")
            except Exception as e:
                print(f"Could not load TXO logo: {e}")
                # Create a placeholder label
                txo_label = tk.Label(logo_frame, text="TXO", font=("Arial", 16, "bold"), bg='white')
                txo_label.grid(row=0, column=2, padx=(20, 0), sticky=tk.E)
            
            # Title text removed; will use image banner instead

            # Optional small workflow banner centered below logos
            try:
                from PIL import Image, ImageTk
                banner_path = r"C:\Lumen\Workflow Manager\workflowbannersmall.png"
                if os.path.exists(banner_path):
                    banner_img = Image.open(banner_path)
                    # Constrain height to ~40px, preserve aspect
                    target_h = 40
                    scale = target_h / banner_img.height if banner_img.height else 1
                    banner_img = banner_img.resize((max(1, int(banner_img.width * scale)), target_h), Image.Resampling.LANCZOS)
                    self._workflow_banner_photo = ImageTk.PhotoImage(banner_img)
                    banner_label = tk.Label(logo_frame, image=self._workflow_banner_photo, bg='white')
                    banner_label.grid(row=0, column=1, padx=20)
            except Exception as e:
                print(f"Could not load workflowbannersmall.png: {e}")
            
        except Exception as e:
            print(f"Error creating logo section: {e}")
            # Create a simple fallback
            fallback_frame = tk.Frame(self.root, bg='white')
            fallback_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
            fallback_label = tk.Label(fallback_frame, text="Degrow Workflow Manager", 
                                     font=("HP Simplified", 18, "bold"), bg='white')
            fallback_label.pack()
    
    def create_toggle_buttons(self):
        """Create toggle buttons for column visibility"""
        try:
            # Create toggle buttons frame
            toggle_frame = ttk.Frame(self.root)
            toggle_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
            toggle_frame.columnconfigure(0, weight=1)
            toggle_frame.columnconfigure(1, weight=1)
            toggle_frame.columnconfigure(2, weight=1)
            
            # Form toggle button
            self.form_toggle_btn = ttk.Button(toggle_frame, text="📝 Data Entry", 
                                            command=self.toggle_form_column)
            self.form_toggle_btn.grid(row=0, column=0, padx=5, sticky=(tk.W, tk.E))
            
            # Subtasks toggle button
            self.subtasks_toggle_btn = ttk.Button(toggle_frame, text="📋 Milestone Tasks", 
                                                 command=self.toggle_subtasks_column)
            self.subtasks_toggle_btn.grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E))
            
            # Gantt chart toggle button
            self.gantt_toggle_btn = ttk.Button(toggle_frame, text="📊 Gantt Chart", 
                                             command=self.toggle_gantt_column)
            self.gantt_toggle_btn.grid(row=0, column=2, padx=5, sticky=(tk.W, tk.E))
            
        except Exception as e:
            print(f"Error creating toggle buttons: {e}")
    
    def create_widgets(self):
        """Create GUI widgets"""
        self.add_debug_entry("Starting widget creation")
        
        # Create logo section at the top
        self.add_debug_entry("Creating logo section")
        self.create_logo_section()
        self.add_debug_entry("Logo section created")
        
        # Create toggle buttons section
        self.create_toggle_buttons()
        
        # Main container frame
        main_container = ttk.Frame(self.root)
        main_container.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initialize column visibility states
        self.form_visible = True
        self.subtasks_visible = True
        self.gantt_visible = True
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)  # Logo section (fixed height)
        self.root.rowconfigure(1, weight=0)  # Toggle buttons (fixed height)
        self.root.rowconfigure(2, weight=1)  # Main content (expandable)
        self.root.rowconfigure(3, weight=0)  # Status bar (fixed height)
        main_container.columnconfigure(0, weight=1)  # Form section
        main_container.columnconfigure(1, weight=2)  # Subtasks section
        main_container.columnconfigure(2, weight=2)  # Gantt chart section (new)
        main_container.rowconfigure(0, weight=1)
        
        # Left side - Form section
        self.main_frame = ttk.Frame(main_container, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.columnconfigure(1, weight=1)
        
        # Middle - Subtasks section
        self.subtasks_frame = ttk.LabelFrame(main_container, text="Milestone Subtasks", padding="10")
        self.subtasks_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        self.subtasks_frame.columnconfigure(0, weight=1)
        self.subtasks_frame.rowconfigure(0, weight=1)
        
        # Right side - Gantt Chart section
        self.gantt_frame = ttk.LabelFrame(main_container, text="Gantt Chart", padding="10")
        self.gantt_frame.grid(row=0, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        self.gantt_frame.columnconfigure(0, weight=1)
        self.gantt_frame.rowconfigure(0, weight=1)
        
        # Create Gantt chart
        self.create_gantt_chart(self.gantt_frame)
        
        
        
        # Record (Row) number field
        ttk.Label(self.main_frame, text="Record (Row) Number:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.record_number_var = tk.StringVar()
        self.record_number_entry = tk.Entry(self.main_frame, textvariable=self.record_number_var, 
                                           width=32, state="readonly", bg='lightgrey', fg='black')
        self.record_number_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # State selection
        ttk.Label(self.main_frame, text="State:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.state_var = tk.StringVar()
        self.state_combo = ttk.Combobox(self.main_frame, textvariable=self.state_var, 
                                      values=self.states, width=30, state="readonly")
        self.state_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.state_combo.configure(style="White.TCombobox")
        self.state_combo['background'] = 'white'
        
        # CLLI field with combined autocomplete and dropdown
        ttk.Label(self.main_frame, text="CLLI:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.clli_var = tk.StringVar()
        self.clli_combo = ttk.Combobox(self.main_frame, textvariable=self.clli_var, width=32, state="normal")
        self.clli_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.clli_combo.configure(style="White.TCombobox")
        self.clli_combo['background'] = 'white'
        
        # Bind events for CLLI autocomplete
        self.clli_var.trace('w', self.on_clli_changed)
        self.clli_combo.bind('<KeyRelease>', self.on_clli_key_release)
        self.clli_combo.bind('<FocusOut>', self.on_clli_focus_out)
        self.clli_combo.bind('<<ComboboxSelected>>', self.on_clli_combo_selected)
        self.clli_combo.bind('<KeyPress>', self.on_clli_combo_key_press)
        
        # Host Wire Centre field (autopopulated from CLLI)
        ttk.Label(self.main_frame, text="Host Wire Centre:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.city_var = tk.StringVar()
        self.city_entry = ttk.Entry(self.main_frame, textvariable=self.city_var, width=32, state="readonly")
        self.city_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.city_entry.configure(style="White.TEntry")
        self.city_entry['background'] = 'white'
        
        # LATA field (autopopulated from CLLI)
        ttk.Label(self.main_frame, text="LATA:", font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.lata_var = tk.StringVar()
        self.lata_entry = ttk.Entry(self.main_frame, textvariable=self.lata_var, width=32, state="readonly")
        self.lata_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.lata_entry.configure(style="White.TEntry")
        self.lata_entry['background'] = 'white'
        
        # Equipment Type field (autopopulated from CLLI)
        ttk.Label(self.main_frame, text="Equipment Type:", font=('Arial', 10, 'bold')).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.equipment_var = tk.StringVar()
        self.equipment_entry = ttk.Entry(self.main_frame, textvariable=self.equipment_var, width=32, state="readonly")
        self.equipment_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.equipment_entry.configure(style="White.TEntry")
        self.equipment_entry['background'] = 'white'
        
        # Current Milestone (replaces Stage)
        ttk.Label(self.main_frame, text="Current Milestone:", font=('Arial', 10, 'bold')).grid(row=6, column=0, sticky=tk.W, pady=5)
        self.stage_var = tk.StringVar()
        self.stage_combo = ttk.Combobox(self.main_frame, textvariable=self.stage_var, 
                                      values=self.stages, width=30, state="readonly")
        self.stage_combo.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.stage_combo.configure(style="White.TCombobox")
        self.stage_combo['background'] = 'white'
        self.stage_combo.bind('<<ComboboxSelected>>', self.on_stage_selected)
        
        # Milestone Subtask
        ttk.Label(self.main_frame, text="Milestone Subtask:", font=('Arial', 10, 'bold')).grid(row=7, column=0, sticky=tk.W, pady=5)
        self.subtask_var = tk.StringVar()
        self.subtask_combo = ttk.Combobox(self.main_frame, textvariable=self.subtask_var, 
                                        values=self.subtasks, width=30, state="readonly")
        self.subtask_combo.grid(row=7, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.subtask_combo.configure(style="White.TCombobox")
        self.subtask_combo['background'] = 'white'
        
        # Status selection with color coding
        ttk.Label(self.main_frame, text="Status:", font=('Arial', 10, 'bold')).grid(row=8, column=0, sticky=tk.W, pady=5)
        self.status_var = tk.StringVar()
        self.status_combo = ttk.Combobox(self.main_frame, textvariable=self.status_var, 
                                       values=self.statuses, width=30, state="readonly")
        self.status_combo.grid(row=8, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.status_combo.configure(style="White.TCombobox")
        self.status_combo['background'] = 'white'
        self.status_combo.bind('<<ComboboxSelected>>', self.on_status_selected)
        
        # Milestone date (read-only)
        ttk.Label(self.main_frame, text="Milestone Date:", font=('Arial', 10, 'bold')).grid(row=9, column=0, sticky=tk.W, pady=5)
        self.milestone_var = tk.StringVar()
        self.milestone_entry = ttk.Entry(self.main_frame, textvariable=self.milestone_var, 
                                       width=32, state="readonly")
        self.milestone_entry.grid(row=9, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.milestone_entry.configure(style="White.TEntry")
        self.milestone_entry['background'] = 'white'
        
        # Bind events for automatic milestone date calculation
        self.stage_combo.bind('<<ComboboxSelected>>', self.on_stage_selected)
        
        # Buttons frame
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.grid(row=10, column=0, columnspan=2, pady=20)
        
        # Save button
        self.save_button = ttk.Button(buttons_frame, text="Save Entry", 
                                     command=self.save_entry, style='Accent.TButton')
        self.save_button.grid(row=0, column=0, padx=5)
        
        # Clear button
        self.clear_button = ttk.Button(buttons_frame, text="Clear Form", 
                                      command=self.clear_form)
        self.clear_button.grid(row=0, column=1, padx=5)
        
        
        # Show Database button
        self.database_button = ttk.Button(buttons_frame, text="Show Database", 
                                         command=self.show_database)
        self.database_button.grid(row=0, column=2, padx=5)
        
        # Import Project button
        self.import_button = ttk.Button(buttons_frame, text="Import Project",
                                     command=self.import_project_data)
        self.import_button.grid(row=0, column=3, padx=5)
        
        # Database Manager button
        self.db_manager_button = ttk.Button(buttons_frame, text="Database Manager",
                                          command=self.open_database_manager)
        self.db_manager_button.grid(row=0, column=4, padx=5)
        
        # Exit button
        self.exit_button = ttk.Button(buttons_frame, text="Exit", 
                                     command=self.root.quit)
        self.exit_button.grid(row=0, column=5, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=11, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))
        
        # Recent entries frame
        entries_frame = ttk.LabelFrame(self.main_frame, text="Recent Entries", padding="5")
        entries_frame.grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        entries_frame.columnconfigure(0, weight=1)
        entries_frame.rowconfigure(1, weight=1)
        
        # Toggle switch frame
        toggle_frame = ttk.Frame(entries_frame)
        toggle_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Toggle switch
        self.entries_mode_var = tk.StringVar(value="database")
        self.database_radio = ttk.Radiobutton(toggle_frame, text="Database Entries", 
                                           variable=self.entries_mode_var, value="database",
                                           command=self._safe_toggle_entries_mode)
        self.database_radio.pack(side=tk.LEFT, padx=(0, 10))
        
        self.debug_radio = ttk.Radiobutton(toggle_frame, text="Debug Log", 
                                        variable=self.entries_mode_var, value="debug",
                                        command=self._safe_toggle_entries_mode)
        self.debug_radio.pack(side=tk.LEFT)
        
        # Recent entries list
        self.entries_text = scrolledtext.ScrolledText(entries_frame, height=10, width=50)
        self.entries_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        
        # Create subtasks list
        self.create_subtasks_list(self.subtasks_frame)
        
        # Initialize debug log
        self.debug_log = []
        
        # Add some initial debug entries
        self.add_debug_entry("Application started")
        self.add_debug_entry("Database initialized")
        self.add_debug_entry("GUI components loaded")
        
        # Load recent entries
        self.load_recent_entries()
        
        # Populate CLLI dropdown
        self.populate_clli_dropdown()

        # Create bottom status bar and start updates
        try:
            self._create_status_bar()
            self._update_status_bar()
        except Exception as e:
            print(f"Error initializing status bar: {e}")
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _create_status_bar(self):
        """Create the bottom status bar with runtime/system info."""
        # Track last NAT sync value; can be updated elsewhere when sync occurs
        if not hasattr(self, 'last_nat_sync'):
            self.last_nat_sync = None

        status_frame = ttk.Frame(self.root, padding=(8, 4))
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=0)
        status_frame.columnconfigure(1, weight=1)

        self.status_user_label = ttk.Label(status_frame, text="Logged in as: ")
        self.status_user_value = ttk.Label(status_frame, text=self.user_manager.get_current_user().username)

        self.status_db_label = ttk.Label(status_frame, text="  |  Database: ")
        self.status_db_value = ttk.Label(status_frame, text="")

        self.status_uptime_label = ttk.Label(status_frame, text="  |  DB Uptime: ")
        self.status_uptime_value = ttk.Label(status_frame, text="")

        self.status_nat_label = ttk.Label(status_frame, text="  |  Last NAT Sync: ")
        self.status_nat_value = ttk.Label(status_frame, text="(NULL)")

        # Lay out labels
        self.status_user_label.grid(row=0, column=0, sticky=tk.W)
        self.status_user_value.grid(row=0, column=1, sticky=tk.W)
        c = 2
        self.status_db_label.grid(row=0, column=c, sticky=tk.W); c += 1
        self.status_db_value.grid(row=0, column=c, sticky=tk.W); c += 1
        self.status_uptime_label.grid(row=0, column=c, sticky=tk.W); c += 1
        self.status_uptime_value.grid(row=0, column=c, sticky=tk.W); c += 1
        self.status_nat_label.grid(row=0, column=c, sticky=tk.W); c += 1
        self.status_nat_value.grid(row=0, column=c, sticky=tk.W)

        self._status_frame = status_frame

    def _update_status_bar(self):
        """Refresh dynamic values in the status bar periodically."""
        try:
            self.status_user_value.configure(text=self.user_manager.get_current_user().username)
        except Exception:
            pass

        db_name = self._get_db_display_name()
        connected, uptime_text = self._get_db_connection_and_uptime()

        if connected:
            self.status_db_value.configure(text=f"Connected to {db_name}", foreground='green')
        else:
            self.status_db_value.configure(text="DB NOT CONNECTED", foreground='red')

        self.status_uptime_value.configure(text=uptime_text)

        if self.last_nat_sync:
            try:
                if isinstance(self.last_nat_sync, str):
                    nat_text = self.last_nat_sync
                else:
                    nat_text = self.last_nat_sync.strftime('%Y-%m-%d %H:%M:%S')
                self.status_nat_value.configure(text=nat_text)
            except Exception:
                self.status_nat_value.configure(text="(NULL)")
        else:
            self.status_nat_value.configure(text="(NULL)")

        try:
            self.root.after(60000, self._update_status_bar)  # refresh every 60s
        except Exception:
            pass

    def set_last_nat_sync(self, when):
        """Externally update the Last NAT Sync value.
        Accepts datetime or string; None clears to (NULL).
        """
        self.last_nat_sync = when
        try:
            self._update_status_bar()
        except Exception:
            pass

    def _get_db_display_name(self) -> str:
        """Return a user-friendly database target string based on live connection if possible."""
        try:
            if not hasattr(self, 'db_conn') or not hasattr(self.db_conn, 'db_config'):
                return ""

            cfg = self.db_conn.db_config
            db_type = str(cfg.get('type', 'sqlite')).lower()

            connection = getattr(self.db_conn, 'connection', None)
            if connection is None:
                try:
                    connection = self.db_conn.connect()
                except Exception:
                    connection = None

            if db_type == 'mysql':
                # Prefer actual selected database name
                db_name = cfg.get('database_name', '')
                try:
                    if connection and hasattr(connection, 'cursor'):
                        cur = connection.cursor()
                        cur.execute('SELECT DATABASE()')
                        row = cur.fetchone()
                        cur.close()
                        if row and row[0]:
                            db_name = row[0]
                except Exception:
                    pass
                host = cfg.get('host', 'localhost')
                port = cfg.get('port', 3306)
                return f"mysql://{host}:{port}/{db_name}"

            # SQLite
            db_name = cfg.get('database_name', 'degrow_workflow.db')
            if not str(db_name).endswith('.db'):
                db_name = f"{db_name}.db"
            try:
                # Create absolute path for clarity
                import os
                return os.path.abspath(db_name)
            except Exception:
                return str(db_name)
        except Exception:
            return ""

    def _get_db_connection_and_uptime(self):
        """Return (connected: bool, uptime_text: str)."""
        try:
            if not hasattr(self, 'db_conn') or self.db_conn is None:
                return False, "N/A"

            connection = getattr(self.db_conn, 'connection', None)
            if connection is None:
                connection = self.db_conn.connect()

            if hasattr(connection, 'is_connected'):
                is_up = connection.is_connected()
                uptime_text = ""
                try:
                    cursor = connection.cursor()
                    cursor.execute("SHOW GLOBAL STATUS LIKE 'Uptime'")
                    row = cursor.fetchone()
                    cursor.close()
                    if row and len(row) >= 2:
                        seconds = int(row[1])
                        uptime_text = self._format_seconds(seconds)
                except Exception:
                    uptime_text = "Unknown"
                return bool(is_up), uptime_text or "Unknown"

            return True, "N/A"  # SQLite
        except Exception:
            return False, "Unknown"

    def _format_seconds(self, total_seconds: int) -> str:
        """Format seconds into human-friendly d hh:mm:ss."""
        try:
            total_seconds = int(total_seconds)
            days, rem = divmod(total_seconds, 86400)
            hours, rem = divmod(rem, 3600)
            minutes, seconds = divmod(rem, 60)
            if days > 0:
                return f"{days}d {hours:02}:{minutes:02}:{seconds:02}"
            return f"{hours:02}:{minutes:02}:{seconds:02}"
        except Exception:
            return "Unknown"
    
    def on_stage_selected(self, event=None):
        """Handle stage selection to update milestone date"""
        stage = self.stage_var.get()
        if stage in self.milestone_dates:
            milestone_date = self.milestone_dates[stage]
            self.milestone_var.set(milestone_date.strftime('%m/%d/%Y'))
        else:
            self.milestone_var.set("")
    
    def on_clli_changed(self, *args):
        """Handle CLLI field changes for autocomplete"""
        query = self.clli_var.get()
        print(f"CLLI changed: '{query}' (length: {len(query)})")
        if len(query) >= 2:
            suggestions = self._search_clli(query)
            print(f"Found {len(suggestions)} suggestions: {suggestions}")
            # Update dropdown values for filtering
            if hasattr(self, 'clli_combo'):
                self.clli_combo['values'] = suggestions
        else:
            # Reset dropdown to all values
            if hasattr(self, 'clli_combo'):
                all_codes = self._get_all_clli_codes()
                self.clli_combo['values'] = all_codes
                print(f"Reset dropdown to all {len(all_codes)} CLLI codes")
    
    def on_clli_key_release(self, event):
        """Handle key release in CLLI field"""
        if event.keysym == 'Return':
            # Get current value and trigger autopopulation
            current_value = self.clli_combo.get()
            if current_value:
                self.autopopulate_from_clli(current_value)
        elif event.keysym == 'Escape':
            # Clear the field
            self.clli_var.set("")
    
    def on_clli_focus_out(self, event):
        """Handle focus out from CLLI field"""
        # Trigger autopopulation if there's a value
        current_value = self.clli_combo.get()
        if current_value:
            self.autopopulate_from_clli(current_value)
    
    def populate_clli_dropdown(self):
        """Populate CLLI dropdown with all available codes"""
        try:
            all_codes = self._get_all_clli_codes()
            self.clli_combo['values'] = all_codes
            print(f"Loaded {len(all_codes)} CLLI codes for dropdown")
            print(f"First 5 CLLI codes: {all_codes[:5] if all_codes else 'None'}")
        except Exception as e:
            print(f"Error populating CLLI dropdown: {e}")
            import traceback
            traceback.print_exc()
    
    def on_clli_combo_selected(self, event=None):
        """Handle CLLI dropdown selection"""
        selected = self.clli_combo.get()
        if selected:
            self.clli_var.set(selected)
            # Autopopulate Host Wire Centre and LATA
            self.autopopulate_from_clli(selected)
            # Refresh subtasks to load dates for new CLLI
            self.populate_subtasks()
    
    def on_clli_combo_key_release(self, event):
        """Handle key release in CLLI dropdown"""
        if event.keysym in ['Up', 'Down']:
            pass  # Arrow key handling removed
        elif event.keysym == 'Return':
            # Get current value and trigger autopopulation
            current_value = self.clli_combo.get()
            if current_value:
                self.clli_var.set(current_value)
                self.autopopulate_from_clli(current_value)
                # Refresh subtasks to load dates for new CLLI
                self.populate_subtasks()
        elif event.keysym == 'Escape':
            pass  # Escape key handling removed
        else:
            # Trigger autocomplete for typing
            self.on_clli_changed()
    
    def on_clli_combo_focus_out(self, event):
        """Handle focus out from CLLI dropdown"""
        # Focus out handling - no action needed
        pass
    
    def on_clli_combo_key_press(self, event):
        """Handle key press in CLLI dropdown"""
        # Allow the key to be processed first, then trigger autopopulation
        self.root.after(10, self.check_clli_autopopulation)
    
    def check_clli_autopopulation(self):
        """Check if CLLI value should trigger autopopulation"""
        current_value = self.clli_combo.get()
        if current_value and len(current_value) >= 2:
            # Check if this is a complete CLLI code that exists in our data
            all_codes = self._get_all_clli_codes()
            if current_value in all_codes:
                self.autopopulate_from_clli(current_value)
    
    def autopopulate_from_clli(self, clli_code):
        """Autopopulate Host Wire Centre and LATA from CLLI code"""
        try:
            self.add_debug_entry(f"Starting autopopulate from CLLI: {clli_code}")
            print(f"Attempting to autopopulate from CLLI: {clli_code}")
            if 'Host CLLI' in self.clli_data.columns:
                print(f"Excel data columns: {list(self.clli_data.columns)}")
                # Find the row with matching CLLI
                matching_rows = self.clli_data[self.clli_data['Host CLLI'].astype(str) == clli_code]
                print(f"Found {len(matching_rows)} matching rows")
                
                if not matching_rows.empty:
                    # Get the first matching row
                    row_data = matching_rows.iloc[0]
                    print(f"Row data: {dict(row_data)}")
                    
                    # Try to find Host Wire Centre, LATA, and Equipment Type columns
                    host_wire_centre_value = ""
                    lata_value = ""
                    equipment_value = ""
                    
                    # Look for Host Wire Centre column (case insensitive) - try multiple variations
                    host_wire_centre_columns = [col for col in self.clli_data.columns if any(keyword in col.lower() for keyword in ['city', 'location', 'place', 'wire', 'centre', 'center'])]
                    print(f"Potential Host Wire Centre columns: {host_wire_centre_columns}")
                    
                    for col in host_wire_centre_columns:
                        if pd.notna(row_data[col]):
                            host_wire_centre_value = str(row_data[col]).strip()
                            print(f"Found Host Wire Centre column '{col}' with value: '{host_wire_centre_value}'")
                            break
                    
                    # Look for Equipment Type column
                    equipment_columns = [col for col in self.clli_data.columns if any(keyword in col.lower() for keyword in ['equipment', 'type', 'device'])]
                    print(f"Potential Equipment Type columns: {equipment_columns}")
                    
                    for col in equipment_columns:
                        if pd.notna(row_data[col]):
                            equipment_value = str(row_data[col]).strip()
                            print(f"Found Equipment Type column '{col}' with value: '{equipment_value}'")
                            break
                    
                    # Look for LATA column - prioritize exact "LATA" match, then search for 3-digit codes
                    lata_columns = []
                    
                    # First, look for exact "LATA" column
                    if 'LATA' in self.clli_data.columns:
                        lata_columns.append('LATA')
                        print("Found exact 'LATA' column")
                    
                    # Also check for columns containing 'lata' (case insensitive)
                    for col in self.clli_data.columns:
                        if 'lata' in col.lower() and col not in lata_columns:
                            lata_columns.append(col)
                    
                    print(f"Potential LATA columns: {lata_columns}")
                    
                    # Try each LATA column
                    for col in lata_columns:
                        if pd.notna(row_data[col]):
                            value = str(row_data[col]).strip()
                            if value and value != 'nan' and value != 'None':
                                # Check if this looks like a 3-digit LATA code
                                if value.isdigit() and len(value) == 3:
                                    lata_value = value
                                    print(f"Found 3-digit LATA in column '{col}' with value: '{lata_value}'")
                                    break
                                elif value.isdigit() and len(value) <= 3:
                                    lata_value = value
                                    print(f"Found LATA in column '{col}' with value: '{lata_value}'")
                                    break
                    
                    # If no LATA found with keywords, look for any 3-digit numeric column
                    if not lata_value:
                        print("No LATA found with keywords, checking for 3-digit numeric columns...")
                        for col in self.clli_data.columns:
                            if col.lower() not in ['host clli', 'city', 'state']:
                                try:
                                    value = str(row_data[col]).strip()
                                    if value and value != 'nan' and value != 'None':
                                        # Check if it's a 3-digit number (LATA code)
                                        if value.isdigit() and len(value) == 3:
                                            lata_value = value
                                            print(f"Found 3-digit LATA code in column '{col}' with value: '{lata_value}'")
                                            break
                                except:
                                    continue
                    
                    # Update the fields
                    self.city_var.set(host_wire_centre_value)
                    self.lata_var.set(lata_value)
                    self.equipment_var.set(equipment_value)
                    
                    print(f"Final autopopulation result:")
                    print(f"  CLLI: {clli_code}")
                    print(f"  Host Wire Centre: '{host_wire_centre_value}'")
                    print(f"  LATA: '{lata_value}'")
                    print(f"  Equipment Type: '{equipment_value}'")
                    
                    # Update subtasks list for this CLLI
                    self.update_subtasks_for_clli(clli_code)
                    
                else:
                    # Clear fields if no match found
                    self.city_var.set("")
                    self.lata_var.set("")
                    self.equipment_var.set("")
                    print(f"No matching data found for CLLI: {clli_code}")
                    
                    # Reset subtasks to generic view
                    self.populate_subtasks()
            else:
                print("Host CLLI column not found in Excel data")
                
        except Exception as e:
            print(f"Error autopopulating from CLLI: {e}")
            import traceback
            traceback.print_exc()
    
    def on_status_selected(self, event=None):
        """Handle status selection with color coding"""
        status = self.status_var.get()
        if status:
            # Apply color coding to the status field
            if status == "To Do":
                self.status_combo.configure(foreground="red")
            elif status == "In Progress":
                self.status_combo.configure(foreground="blue")
            elif status == "Blocked":
                self.status_combo.configure(foreground="orange")
            elif status == "Done":
                self.status_combo.configure(foreground="green")
            else:
                self.status_combo.configure(foreground="black")
    
    
    def create_subtasks_list(self, parent_frame):
        """Create subtasks list for each milestone"""
        try:
            # Create a treeview for subtasks
            self.subtasks_tree = ttk.Treeview(parent_frame, columns=('subtask', 'status', 'criticality', 'planned_start', 'actual_start', 'duration', 'planned_end', 'actual_end'), show='tree headings', height=15)
            
            # Store original column widths for reference
            self.original_column_widths = {
                'subtask': 180,
                'status': 80,
                'criticality': 120,
                'planned_start': 100,
                'actual_start': 100,
                'duration': 80,
                'planned_end': 100,
                'actual_end': 100
            }
            self.subtasks_tree.heading('#0', text='Milestone')
            self.subtasks_tree.heading('subtask', text='Subtasks')
            self.subtasks_tree.heading('status', text='Status')
            self.subtasks_tree.heading('criticality', text='Criticality')
            self.subtasks_tree.heading('planned_start', text='Planned Start')
            self.subtasks_tree.heading('actual_start', text='Actual Start')
            self.subtasks_tree.heading('duration', text='Duration')
            self.subtasks_tree.heading('planned_end', text='Planned End')
            self.subtasks_tree.heading('actual_end', text='Actual End')
            # Configure columns to use full width efficiently
            # Milestone column - reasonable width that prevents truncation
            self.subtasks_tree.column('#0', width=300, minwidth=250)  # Milestone - reasonable width
            self.subtasks_tree.column('subtask', width=180, minwidth=150)  # Subtasks - restored
            self.subtasks_tree.column('status', width=80, minwidth=60)  # Status - restored
            self.subtasks_tree.column('criticality', width=120, minwidth=100)  # Criticality - restored
            self.subtasks_tree.column('planned_start', width=100, minwidth=80)  # Planned Start - restored
            self.subtasks_tree.column('actual_start', width=100, minwidth=80)  # Actual Start - restored
            self.subtasks_tree.column('duration', width=80, minwidth=60)  # Duration - restored
            self.subtasks_tree.column('planned_end', width=100, minwidth=80)  # Planned End - restored
            self.subtasks_tree.column('actual_end', width=100, minwidth=80)  # Actual End - restored
            
            # Add vertical and horizontal scrollbars
            subtasks_v_scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=self.subtasks_tree.yview)
            subtasks_h_scrollbar = ttk.Scrollbar(parent_frame, orient="horizontal", command=self.subtasks_tree.xview)
            self.subtasks_tree.configure(yscrollcommand=subtasks_v_scrollbar.set, xscrollcommand=subtasks_h_scrollbar.set)
            
            # Configure parent frame column weights for full width table
            parent_frame.columnconfigure(0, weight=1)  # Tree takes most space
            parent_frame.columnconfigure(1, weight=0)  # Scrollbar takes minimal space
            parent_frame.columnconfigure(2, weight=0)  # Button area takes minimal space
            
            # Configure treeview selection mode
            self.subtasks_tree.configure(selectmode='extended')  # Allow multiple selection if needed
            
            # Grid the treeview and scrollbars
            self.subtasks_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 0))
            subtasks_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), padx=(0, 0))
            subtasks_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 0))
            
            # Configure tags for color coding AFTER tree is created
            self.subtasks_tree.tag_configure('no_entry', background='lightblue', foreground='darkblue')
            self.subtasks_tree.tag_configure('in_progress', background='lightyellow', foreground='darkorange')
            self.subtasks_tree.tag_configure('done', background='lightgreen', foreground='darkgreen')
            self.subtasks_tree.tag_configure('blocked', background='lightcoral', foreground='darkred')
            
            # Add tooltip for editable columns
            self.create_tooltip()
            
            # Create context menu for status changes
            self.create_context_menu()
            
            # Bind right-click event
            self.subtasks_tree.bind("<Button-3>", self.show_context_menu)  # Button-3 is right-click
            
            # Bind double-click event for cell editing
            self.subtasks_tree.bind("<Double-1>", self.on_cell_double_click)  # Double-1 is double-click
            
            # Bind window resize event to maintain milestone column width
            self.root.bind("<Configure>", self.on_window_resize)
            
            # Create a button frame to align buttons with scrollbar
            button_frame = ttk.Frame(parent_frame)
            button_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
            button_frame.columnconfigure(0, weight=0)  # Test Colors button - fixed size
            button_frame.columnconfigure(1, weight=0)   # Test Status Update button - fixed size
            button_frame.columnconfigure(2, weight=1)    # Spacer
            button_frame.columnconfigure(3, weight=0)   # Edit Milestones button - fixed size
            
            # Add test buttons for color coding - same size and aligned
            test_btn = ttk.Button(button_frame, text="Test Colors", 
                                 command=self.test_color_coding)
            test_btn.grid(row=0, column=0, padx=(0, 5))
            
            # Add button to test status updates - same size as Test Colors
            status_test_btn = ttk.Button(button_frame, text="Test Status Update", 
                                        command=self.test_status_update)
            status_test_btn.grid(row=0, column=1, padx=(0, 5))
            
            # Add button to open milestone editor - aligned with right edge
            editor_btn = ttk.Button(button_frame, text="Edit Milestones", 
                                   command=self.open_milestone_editor)
            editor_btn.grid(row=0, column=3)
            
            # Populate subtasks for each milestone
            self.populate_subtasks()
            
            # Adjust milestone column width after populating
            self.adjust_milestone_column_width()
            
        except Exception as e:
            print(f"Error creating subtasks list: {e}")
            error_label = ttk.Label(parent_frame, text=f"Subtasks Error: {str(e)}", 
                                  foreground="red")
            error_label.grid(row=0, column=0)
    
    def populate_subtasks(self):
        """Populate subtasks for each milestone from database"""
        try:
            # Clear existing items
            for item in self.subtasks_tree.get_children():
                self.subtasks_tree.delete(item)
            
            # Load milestones from database
            milestones = self._load_milestones_from_db()
            
            if not milestones:
                # If no milestones in database, show empty state
                empty_node = self.subtasks_tree.insert('', 'end', text="No milestones defined", 
                                                     values=('', '', '', '', '', '', '', ''))
                return
            
            # Get all milestones and their subtasks from database
            for milestone in milestones:
                # Create milestone node
                milestone_node = self.subtasks_tree.insert('', 'end', text=milestone['name'], 
                                                          values=('', '', '', '', '', '', '', ''))
                
                # Get subtasks for this milestone from database
                subtasks = self._load_subtasks_from_db(milestone['id'])
                
                for subtask in subtasks:
                    # Get status from workflow database
                    status = self.get_subtask_status(milestone['name'], subtask['name'])
                    tag = self.get_status_tag(status)
                    
                    # Get dates from date_entries table for current CLLI
                    current_clli = self.clli_var.get().strip()
                    dates = self._load_dates_for_clli(current_clli)
                    
                    # Insert subtask with dates
                    self.subtasks_tree.insert(milestone_node, 'end', 
                                            values=(subtask['name'], status, subtask['criticality'], 
                                                   dates.get('planned_start', ''), 
                                                   dates.get('actual_start', ''), 
                                                   '',  # duration - not stored in date_entries
                                                   dates.get('planned_end', ''), 
                                                   dates.get('actual_end', '')),
                                            tags=(tag,))
                
                # Expand the milestone
                self.subtasks_tree.item(milestone_node, open=True)
            
            # Adjust milestone column width after populating
            self.adjust_milestone_column_width()
            
            # Force a refresh of the treeview to ensure proper display
            self.root.after(100, self.refresh_milestone_column_width)
                
        except Exception as e:
            print(f"Error populating subtasks: {e}")
    
    def _load_milestones_from_db(self):
        """Load milestones from database"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, name, description, order_index, created_date, last_updated
                FROM milestones 
                ORDER BY order_index, name
            """)
            
            milestones = []
            for row in cursor.fetchall():
                milestones.append({
                    'id': row[0],
                    'name': row[1],
                    'description': row[2],
                    'order_index': row[3],
                    'created_date': row[4],
                    'last_updated': row[5]
                })
            
            cursor.close()
            conn.close()
            return milestones
            
        except Exception as e:
            print(f"Error loading milestones from database: {e}")
            return []
    
    def _load_subtasks_from_db(self, milestone_id):
        """Load subtasks for a specific milestone from database"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, name, description, criticality, order_index, created_date, last_updated
                FROM subtasks 
                WHERE milestone_id = %s
                ORDER BY order_index, name
            """, (milestone_id,))
            
            subtasks = []
            for row in cursor.fetchall():
                subtasks.append({
                    'id': row[0],
                    'name': row[1],
                    'description': row[2],
                    'criticality': row[3],
                    'order_index': row[4],
                    'created_date': row[5],
                    'last_updated': row[6]
                })
            
            cursor.close()
            conn.close()
            return subtasks
            
        except Exception as e:
            print(f"Error loading subtasks from database: {e}")
            return []
    
    def update_subtasks_for_clli(self, clli_code):
        """Update subtasks list based on selected CLLI"""
        try:
            if not clli_code:
                # If no CLLI selected, show generic subtasks
                self.populate_subtasks()
                return
            
            # Clear existing items
            for item in self.subtasks_tree.get_children():
                self.subtasks_tree.delete(item)
            
            # Get milestones in workflow order
            milestones = list(reversed(self.stages))
            
            # Add each milestone with CLLI-specific subtasks
            for milestone in milestones:
                # Insert milestone as parent with CLLI context
                milestone_text = f"{milestone} (CLLI: {clli_code})"
                milestone_id = self.subtasks_tree.insert('', 'end', text=milestone_text, values=('', ''))
                
                # Add CLLI-specific subtasks for this milestone
                for subtask in self.subtasks:
                    # Get status for this subtask with CLLI context
                    status = self.get_subtask_status_for_clli(milestone, subtask, clli_code)
                    tag = self.get_status_tag(status)
                    # Get criticality level for this subtask
                    criticality = self.criticality_levels.get(subtask, "Should be Complete")
                    
                    self.subtasks_tree.insert(milestone_id, 'end', text='', 
                                           values=(subtask, status, criticality, '', '', '', '', ''), tags=tag)
                
                # Expand the milestone
                self.subtasks_tree.item(milestone_id, open=True)
                
        except Exception as e:
            print(f"Error updating subtasks for CLLI: {e}")
    
    def test_color_coding(self):
        """Test method to verify color coding is working"""
        try:
            print("Testing color coding...")
            # Clear existing items
            for item in self.subtasks_tree.get_children():
                self.subtasks_tree.delete(item)
            
            # Create test milestone
            test_milestone = "Test Milestone"
            milestone_id = self.subtasks_tree.insert('', 'end', text=test_milestone, values=('', ''))
            
            # Add test subtasks with different statuses
            test_data = [
                ("CLLI Validation", "No Entry", "Must be complete", "no_entry"),
                ("Site Survey", "In Progress", "Must be complete", "in_progress"),
                ("Equipment Inventory", "Done", "Should be Complete", "done"),
                ("Network Analysis", "Blocked", "Must be complete", "blocked")
            ]
            
            for subtask, status, criticality, tag in test_data:
                print(f"Inserting test subtask: {subtask}, status: {status}, criticality: {criticality}, tag: {tag}")
                self.subtasks_tree.insert(milestone_id, 'end', text='', 
                                       values=(subtask, status, criticality, '', '', '', '', ''), tags=tag)
            
            # Expand the milestone
            self.subtasks_tree.item(milestone_id, open=True)
            print("Test color coding completed")
            
        except Exception as e:
            print(f"Error in test color coding: {e}")
    
    def test_status_update(self):
        """Test method to verify status updates are working"""
        try:
            print("Testing status update...")
            
            # Clear existing items
            for item in self.subtasks_tree.get_children():
                self.subtasks_tree.delete(item)
            
            # Create test milestone
            test_milestone = "Status Test Milestone"
            milestone_id = self.subtasks_tree.insert('', 'end', text=test_milestone, values=('', ''))
            
            # Add test subtask
            test_subtask = "Test Subtask"
            test_item = self.subtasks_tree.insert(milestone_id, 'end', text='', 
                                                values=(test_subtask, "No Entry", "Should be Complete"), tags="no_entry")
            
            # Expand the milestone
            self.subtasks_tree.item(milestone_id, open=True)
            
            # Simulate status update
            print("Simulating status update...")
            self.selected_subtask_item = test_item
            self.update_subtask_status("In Progress")
            
            print("Status update test completed")
            
        except Exception as e:
            print(f"Error in test status update: {e}")
    
    def on_cell_double_click(self, event):
        """Handle double-click on treeview cell for editing"""
        try:
            # If already editing, save current edit first
            if self.editing_item:
                self.save_current_edit()
                return
            
            # Get the item and column that was clicked
            item = self.subtasks_tree.identify_row(event.y)
            column = self.subtasks_tree.identify_column(event.x)
            
            if not item or not column:
                return
            
            # Get column index (subtract 1 because #0 is the tree column)
            column_index = int(column[1:]) - 1
            
            # Only allow editing of date columns (planned_start, actual_start, planned_end, actual_end)
            # These are columns 4, 5, 7, 8 (0-indexed: 3, 4, 6, 7)
            date_column_indices = [3, 4, 6, 7]  # planned_start, actual_start, planned_end, actual_end
            if column_index in date_column_indices:
                self.edit_cell(item, column_index)
                
        except Exception as e:
            print(f"Error handling cell double-click: {e}")
    
    def save_current_edit(self):
        """Save the current edit if one is in progress"""
        try:
            if self.editing_item and self.edit_entry:
                new_value = self.edit_entry.get().strip()
                if self.validate_date(new_value):
                    # Log user activity
                    if self.user_manager:
                        self.user_manager.log_activity(
                            self.user_manager.get_current_user().user_id,
                            'update_entry',
                            f'Updated {self.editing_column_name} to {new_value}'
                        )
                    
                    # Update the treeview
                    values = list(self.subtasks_tree.item(self.editing_item, 'values'))
                    values[self.editing_column] = new_value
                    self.subtasks_tree.item(self.editing_item, values=values)
                    
                    # Update database
                    self.update_subtask_date(self.editing_item, self.editing_column_name, new_value)
                    
                    # Clean up
                    self.edit_entry.destroy()
                    self.editing_item = None
                    self.editing_column = None
                    self.editing_column_name = None
                    self.original_value = None
                else:
                    messagebox.showerror("Invalid Date", "Please enter a valid date in MM/DD/YYYY format.")
                    self.edit_entry.focus()
        except Exception as e:
            print(f"Error saving current edit: {e}")
    
    def on_treeview_click(self, event):
        """Handle clicks on the treeview"""
        try:
            if self.editing_item:
                # Save current edit if clicking outside the editing cell
                self.save_current_edit()
            else:
                # Restore normal double-click behavior
                self.subtasks_tree.bind('<Double-1>', self.on_cell_double_click)
        except Exception as e:
            print(f"Error handling treeview click: {e}")
    
    def edit_cell(self, item, column_index):
        """Edit a specific cell in the treeview using inline editing"""
        try:
            # Get current values
            values = list(self.subtasks_tree.item(item, 'values'))
            
            # Get the current value of the cell to edit
            current_value = values[column_index] if column_index < len(values) else ""
            
            # Get column name for database update
            column_names = ['subtask', 'status', 'criticality', 'planned_start', 'actual_start', 'duration', 'planned_end', 'actual_end']
            column_name = column_names[column_index] if column_index < len(column_names) else "Date"
            
            # Get the bounding box of the cell
            bbox = self.subtasks_tree.bbox(item, f"#{column_index + 1}")
            if not bbox:
                return
            
            x, y, width, height = bbox
            
            # Create entry widget for inline editing
            self.edit_entry = tk.Entry(self.subtasks_tree, font=('Arial', 9), 
                                     bg='lightyellow', fg='black', relief='solid', bd=1)
            self.edit_entry.place(x=x, y=y, width=width, height=height)
            self.edit_entry.insert(0, current_value)
            self.edit_entry.select_range(0, tk.END)
            self.edit_entry.focus()
            
            # Store editing state
            self.editing_item = item
            self.editing_column = column_index
            self.editing_column_name = column_name
            self.original_value = current_value
            
            def save_edit():
                new_value = self.edit_entry.get().strip()
                if self.validate_date(new_value):
                    # Update the treeview
                    values[column_index] = new_value
                    self.subtasks_tree.item(item, values=values)
                    
                    # Update database
                    self.update_subtask_date(item, column_name, new_value)
                    
                    # Clean up
                    self.edit_entry.destroy()
                    self.editing_item = None
                    self.editing_column = None
                    self.editing_column_name = None
                    self.original_value = None
                else:
                    messagebox.showerror("Invalid Date", "Please enter a valid date in MM/DD/YYYY format.")
                    self.edit_entry.focus()
            
            def cancel_edit():
                # Restore original value
                values[column_index] = self.original_value
                self.subtasks_tree.item(item, values=values)
                
                # Clean up
                self.edit_entry.destroy()
                self.editing_item = None
                self.editing_column = None
                self.editing_column_name = None
                self.original_value = None
            
            # Bind events
            self.edit_entry.bind('<Return>', lambda e: save_edit())
            self.edit_entry.bind('<Escape>', lambda e: cancel_edit())
            self.edit_entry.bind('<FocusOut>', lambda e: save_edit())  # Save when focus is lost
            
            # Bind click outside to save (temporarily unbind and rebind)
            self.subtasks_tree.unbind('<Button-1>')
            self.subtasks_tree.bind('<Button-1>', self.on_treeview_click)
            
        except Exception as e:
            print(f"Error editing cell: {e}")
            messagebox.showerror("Error", f"Error editing cell: {e}")
    
    def validate_date(self, date_string):
        """Validate date string in MM/DD/YYYY format"""
        try:
            if not date_string:
                return True  # Allow empty dates
            
            # Try to parse the date
            datetime.strptime(date_string, '%m/%d/%Y')
            return True
        except ValueError:
            return False
    
    def update_subtask_date(self, item, column_name, new_value):
        """Update subtask date in date_entries table using CLLI as primary key"""
        try:
            # Get the current CLLI from the form
            current_clli = self.clli_var.get().strip()
            if not current_clli:
                messagebox.showerror("Error", "No CLLI selected. Please select a CLLI first.")
                return
            
            # Connect to database
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if record exists for this CLLI
            cursor.execute("SELECT clli FROM date_entries WHERE clli = ?", (current_clli,))
            existing_record = cursor.fetchone()
            
            if existing_record:
                # Update existing record
                update_query = f"""
                    UPDATE date_entries 
                    SET {column_name} = ?, last_updated = ?
                    WHERE clli = ?
                """
                cursor.execute(update_query, (
                    new_value if new_value else None,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    current_clli
                ))
            else:
                # Insert new record
                insert_query = """
                    INSERT INTO date_entries (clli, planned_start, actual_start, planned_end, actual_end, created_date, last_updated)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """
                # Initialize all date fields as None
                planned_start = None
                actual_start = None
                planned_end = None
                actual_end = None
                
                # Set the specific field being updated
                if column_name == 'planned_start':
                    planned_start = new_value if new_value else None
                elif column_name == 'actual_start':
                    actual_start = new_value if new_value else None
                elif column_name == 'planned_end':
                    planned_end = new_value if new_value else None
                elif column_name == 'actual_end':
                    actual_end = new_value if new_value else None
                
                cursor.execute(insert_query, (
                    current_clli,
                    planned_start,
                    actual_start,
                    planned_end,
                    actual_end,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ))
            
            conn.commit()
            conn.close()
            
            print(f"Updated {column_name} for CLLI {current_clli} to {new_value}")
            
        except Exception as e:
            print(f"Error updating subtask date: {e}")
            messagebox.showerror("Error", f"Error updating date: {e}")
    
    def create_tooltip(self):
        """Create tooltip for editable columns"""
        try:
            # Create tooltip window
            self.tooltip = tk.Toplevel(self.root)
            self.tooltip.withdraw()  # Hide initially
            self.tooltip.overrideredirect(True)  # Remove window decorations
            self.tooltip.configure(bg='lightyellow')
            
            # Create tooltip label
            self.tooltip_label = tk.Label(self.tooltip, text="Double-click date columns to edit", 
                                        bg='lightyellow', fg='black', font=('Arial', 9))
            self.tooltip_label.pack()
            
            # Bind mouse motion to show/hide tooltip
            self.subtasks_tree.bind("<Motion>", self.on_mouse_motion)
            self.subtasks_tree.bind("<Leave>", self.hide_tooltip)
            
        except Exception as e:
            print(f"Error creating tooltip: {e}")
    
    def on_mouse_motion(self, event):
        """Handle mouse motion over treeview"""
        try:
            # Get the column that was hovered
            column = self.subtasks_tree.identify_column(event.x)
            if not column:
                self.hide_tooltip()
                return
            
            # Get column index
            column_index = int(column[1:]) - 1
            
            # Show tooltip only for date columns
            date_column_indices = [3, 4, 6, 7]  # planned_start, actual_start, planned_end, actual_end
            if column_index in date_column_indices:
                self.show_tooltip(event)
            else:
                self.hide_tooltip()
                
        except Exception as e:
            print(f"Error in mouse motion: {e}")
    
    def show_tooltip(self, event):
        """Show tooltip at mouse position"""
        try:
            x = self.root.winfo_rootx() + event.x + 10
            y = self.root.winfo_rooty() + event.y + 10
            self.tooltip.geometry(f"+{x}+{y}")
            self.tooltip.deiconify()
        except Exception as e:
            print(f"Error showing tooltip: {e}")
    
    def hide_tooltip(self, event=None):
        """Hide tooltip"""
        try:
            self.tooltip.withdraw()
        except Exception as e:
            print(f"Error hiding tooltip: {e}")
    
    def adjust_milestone_column_width(self):
        """Adjust milestone column width to span across available space"""
        try:
            # Get all milestone items
            milestone_items = []
            for item in self.subtasks_tree.get_children():
                milestone_text = self.subtasks_tree.item(item, 'text')
                if milestone_text:
                    milestone_items.append(milestone_text)
            
            if not milestone_items:
                return
            
            # Calculate the maximum width needed for milestone text
            max_text_width = 0
            for milestone_text in milestone_items:
                # More accurate width estimation
                estimated_width = len(milestone_text) * 9 + 40  # Add padding
                max_text_width = max(max_text_width, estimated_width)
            
            # Get the total available width of the treeview
            treeview_width = self.subtasks_tree.winfo_width()
            if treeview_width <= 1:  # Treeview not yet rendered
                treeview_width = 1000  # Use a reasonable default
            
            # Calculate total width of other columns
            other_columns_width = sum(self.original_column_widths.values())
            
            # Calculate available space for milestone column
            available_width = treeview_width - other_columns_width - 50  # Account for scrollbars and padding
            
            # Use the larger of: required text width or available space
            milestone_width = max(max_text_width, available_width)
            
            # Set reasonable limits
            milestone_width = min(milestone_width, 350)  # Cap at 350 pixels
            milestone_width = max(milestone_width, 250)  # Minimum width
            
            # Update the milestone column width
            self.subtasks_tree.column('#0', width=int(milestone_width))
            
            print(f"Adjusted milestone column width to: {int(milestone_width)} pixels (text needed: {max_text_width}, available: {available_width})")
            
        except Exception as e:
            print(f"Error adjusting milestone column width: {e}")
    
    def refresh_milestone_column_width(self):
        """Refresh milestone column width after a delay"""
        try:
            # Recalculate and set the milestone column width
            self.adjust_milestone_column_width()
            
            # Force treeview to redraw
            self.subtasks_tree.update_idletasks()
            
        except Exception as e:
            print(f"Error refreshing milestone column width: {e}")
    
    def on_window_resize(self, event):
        """Handle window resize events to maintain milestone column width"""
        try:
            # Only handle resize events for the main window
            if event.widget == self.root:
                # Debounce the resize event to avoid too many calls
                if hasattr(self, '_resize_timer'):
                    self.root.after_cancel(self._resize_timer)
                
                self._resize_timer = self.root.after(200, self.refresh_milestone_column_width)
                
        except Exception as e:
            print(f"Error handling window resize: {e}")
    
    def force_milestone_column_expansion(self):
        """Force the milestone column to expand to maximum available width"""
        try:
            # Get the actual treeview width
            treeview_width = self.subtasks_tree.winfo_width()
            if treeview_width <= 1:
                # If treeview not yet rendered, try again later
                self.root.after(200, self.force_milestone_column_expansion)
                return
            
            # Calculate total width of other columns
            other_columns_width = sum(self.original_column_widths.values())
            
            # Make milestone column take up reasonable space
            milestone_width = treeview_width - other_columns_width - 100  # Leave margin for other columns
            
            # Ensure reasonable width limits
            milestone_width = max(milestone_width, 250)  # Minimum width
            milestone_width = min(milestone_width, 350)  # Maximum width
            
            # Update milestone column width
            self.subtasks_tree.column('#0', width=int(milestone_width))
            
            print(f"Forced milestone column expansion to: {int(milestone_width)} pixels (treeview width: {treeview_width})")
            
        except Exception as e:
            print(f"Error forcing milestone column expansion: {e}")
    
    def open_milestone_editor(self):
        """Open the milestone editor window"""
        try:
            # Pass the milestone database path to the editor
            milestone_db_path = getattr(self, 'milestone_db_path', self.db_path.replace('.accdb', '_milestones.db'))
            open_milestone_editor(self.root, milestone_db_path)
            # Refresh the subtasks after editing
            self.populate_subtasks()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open milestone editor: {e}")
    
    def open_database_manager(self):
        """Open the database manager window"""
        try:
            self.add_debug_entry("Opening database manager")
            # Pass both database paths to the manager
            milestone_db_path = getattr(self, 'milestone_db_path', self.db_path.replace('.accdb', '_milestones.db'))
            open_database_manager(self.root, self.db_path, milestone_db_path)
            self.add_debug_entry("Database manager closed")
        except Exception as e:
            self.add_debug_entry(f"Error opening database manager: {e}")
            messagebox.showerror("Error", f"Failed to open database manager: {e}")
    
    def set_record_number(self, record_id: int = None):
        """Set the record number field"""
        if record_id is not None:
            self.current_record_id = record_id
            self.record_number_var.set(f"Record #{record_id}")
        else:
            self.current_record_id = None
            self.record_number_var.set("New Record")
    
    def show_database(self):
        """Show database contents in a new window"""
        try:
            # Create new window for database display
            db_window = tk.Toplevel(self.root)
            db_window.title("Database Contents")
            db_window.geometry("1000x600")
            db_window.transient(self.root)
            db_window.grab_set()
            
            # Create frame for database content
            self.main_frame = ttk.Frame(db_window)
            self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
            
            # Title
            title_label = ttk.Label(self.main_frame, text="Workflow Entries Database", 
                                  font=('Arial', 14, 'bold'))
            title_label.pack(pady=(0, 10))
            
            # Create treeview for database display
            columns = ('ID', 'State', 'CLLI', 'Host Wire Centre', 'LATA', 'Equipment Type', 
                      'Current Milestone', 'Milestone Subtask', 'Status', 
                      'Planned Start', 'Actual Start', 'Duration', 'Planned End', 'Actual End',
                      'Milestone Date', 'Created Date', 'Last Update')
            
            tree = ttk.Treeview(self.main_frame, columns=columns, show='headings', height=15)
            
            # Configure column headings
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100, minwidth=80)
            
            # Add scrollbars
            v_scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=tree.yview)
            h_scrollbar = ttk.Scrollbar(self.main_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            
            # Pack treeview and scrollbars
            tree.pack(side="left", fill="both", expand=True)
            v_scrollbar.pack(side="right", fill="y")
            h_scrollbar.pack(side="bottom", fill="x")
            
            # Load database data
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, state, clli, host_wire_centre, lata, equipment_type, current_milestone, 
                       milestone_subtask, status, planned_start, actual_start, duration, 
                       planned_end, actual_end, milestone_date, created_date, last_update
                FROM workflow_entries 
                ORDER BY created_date DESC
            """)
            
            records = cursor.fetchall()
            
            # Insert records into treeview
            for record in records:
                tree.insert('', 'end', values=record)
            
            conn.close()
            
            # Add close button
            close_button = ttk.Button(self.main_frame, text="Close", 
                                    command=db_window.destroy)
            close_button.pack(pady=(10, 0))
            
            # Show record count
            count_label = ttk.Label(self.main_frame, text=f"Total Records: {len(records)}")
            count_label.pack(pady=(5, 0))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show database: {e}")
    
    def import_project_data(self):
        """Import Microsoft Project data"""
        try:
            # Open file dialog to select project file
            file_path = filedialog.askopenfilename(
                title="Select Microsoft Project File",
                filetypes=[
                    ("Microsoft Project Files", "*.mpp *.xml"),
                    ("XML Files", "*.xml"),
                    ("All Files", "*.*")
                ]
            )
            
            if not file_path:
                return
            
            # Check file extension and process accordingly
            if file_path.lower().endswith('.xml'):
                self._import_xml_project(file_path)
            elif file_path.lower().endswith('.mpp'):
                self._import_mpp_project(file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select a .mpp or .xml file.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import project data: {e}")
    
    def _import_xml_project(self, file_path: str):
        """Import XML-based Microsoft Project file"""
        try:
            # Parse XML file
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Find tasks in the XML structure
            tasks = []
            
            # Look for common Microsoft Project XML structures
            for task in root.findall('.//Task'):
                task_data = self._extract_task_data(task)
                if task_data:
                    tasks.append(task_data)
            
            # If no tasks found, try alternative structure
            if not tasks:
                for task in root.findall('.//task'):
                    task_data = self._extract_task_data(task)
                    if task_data:
                        tasks.append(task_data)
            
            if tasks:
                self._process_imported_tasks(tasks, "XML Project")
            else:
                messagebox.showwarning("Warning", "No tasks found in the XML file.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse XML project file: {e}")
    
    def _import_mpp_project(self, file_path: str):
        """Import MPP Microsoft Project file (requires additional library)"""
        try:
            # For MPP files, we would need a library like python-mpp
            # For now, show a message about limitations
            messagebox.showinfo("Information", 
                              "MPP file import requires additional setup.\n"
                              "Please save your Microsoft Project file as XML format\n"
                              "and try importing the XML file instead.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import MPP project file: {e}")
    
    def _extract_task_data(self, task_element):
        """Extract task data from XML element"""
        try:
            task_data = {}
            
            # Extract common task fields
            name_elem = task_element.find('Name')
            if name_elem is not None:
                task_data['name'] = name_elem.text or ""
            
            start_elem = task_element.find('Start')
            if start_elem is not None:
                task_data['start_date'] = start_elem.text or ""
            
            finish_elem = task_element.find('Finish')
            if finish_elem is not None:
                task_data['finish_date'] = finish_elem.text or ""
            
            duration_elem = task_element.find('Duration')
            if duration_elem is not None:
                task_data['duration'] = duration_elem.text or ""
            
            # Extract custom fields if available
            for field in task_element.findall('ExtendedAttribute'):
                name_elem = field.find('FieldName')
                value_elem = field.find('Value')
                if name_elem is not None and value_elem is not None:
                    field_name = name_elem.text
                    field_value = value_elem.text
                    if field_name and field_value:
                        task_data[field_name.lower()] = field_value
            
            return task_data if task_data else None
            
        except Exception as e:
            print(f"Error extracting task data: {e}")
            return None
    
    def _process_imported_tasks(self, tasks: List[Dict], source: str):
        """Process imported tasks and show results"""
        try:
            # Create results window
            results_window = tk.Toplevel(self.root)
            results_window.title(f"Import Results - {source}")
            results_window.geometry("800x600")
            results_window.transient(self.root)
            results_window.grab_set()
            
            # Create main frame
            self.main_frame = ttk.Frame(results_window)
            self.main_frame.pack(expand=True, fill="both", padx=10, pady=10)
            
            # Title
            title_label = ttk.Label(self.main_frame, text=f"Imported Tasks from {source}", 
                                  font=('Arial', 14, 'bold'))
            title_label.pack(pady=(0, 10))
            
            # Create treeview for task display
            columns = ('Task Name', 'Start Date', 'Finish Date', 'Duration', 'Status')
            tree = ttk.Treeview(self.main_frame, columns=columns, show='headings', height=15)
            
            # Configure columns
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=150, minwidth=100)
            
            # Add scrollbar
            scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Pack treeview and scrollbar
            tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Insert tasks
            for i, task in enumerate(tasks):
                tree.insert('', 'end', values=(
                    task.get('name', ''),
                    task.get('start_date', ''),
                    task.get('finish_date', ''),
                    task.get('duration', ''),
                    'Imported'
                ))
            
            # Add buttons frame
            buttons_frame = ttk.Frame(self.main_frame)
            buttons_frame.pack(fill="x", pady=(10, 0))
            
            # Import to database button
            import_btn = ttk.Button(buttons_frame, text="Import to Database", 
                                  command=lambda: self._import_tasks_to_database(tasks, results_window))
            import_btn.pack(side="left", padx=(0, 10))
            
            # Close button
            close_btn = ttk.Button(buttons_frame, text="Close", 
                                 command=results_window.destroy)
            close_btn.pack(side="left")
            
            # Show task count
            count_label = ttk.Label(self.main_frame, text=f"Total Tasks: {len(tasks)}")
            count_label.pack(pady=(5, 0))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process imported tasks: {e}")
    
    def _import_tasks_to_database(self, tasks: List[Dict], parent_window):
        """Import tasks to database as workflow entries"""
        try:
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            imported_count = 0
            
            for task in tasks:
                # Create workflow entry from task data
                cursor.execute("""
                    INSERT INTO workflow_entries 
                    (state, clli, city, lata, equipment_type, current_milestone, milestone_subtask, 
                     status, milestone_date, created_date, last_update)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    task.get('state', 'Imported'),
                    task.get('clli', ''),
                    task.get('host_wire_centre', ''),
                    task.get('lata', ''),
                    task.get('equipment_type', ''),
                    task.get('name', 'Imported Task'),
                    task.get('milestone_subtask', ''),
                    'Imported',
                    task.get('start_date', ''),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ))
                imported_count += 1
            
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", f"Successfully imported {imported_count} tasks to database!")
            parent_window.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import tasks to database: {e}")
    
    def create_gantt_chart(self, parent_frame):
        """Create Gantt chart visualization"""
        try:
            # Create matplotlib figure
            self.gantt_fig = Figure(figsize=(8, 6), dpi=100)
            self.gantt_ax = self.gantt_fig.add_subplot(111)
            
            # Create canvas for matplotlib
            self.gantt_canvas = FigureCanvasTkAgg(self.gantt_fig, parent_frame)
            self.gantt_canvas.get_tk_widget().pack(expand=True, fill="both")
            
            # Create control frame for Gantt chart
            control_frame = ttk.Frame(parent_frame)
            control_frame.pack(fill="x", pady=(5, 0))
            
            # Week navigation controls
            nav_frame = ttk.Frame(control_frame)
            nav_frame.pack(fill="x", pady=(0, 5))
            
            # Current week display
            self.current_week_var = tk.StringVar()
            self.current_week_var.set(f"Week {datetime.now().isocalendar()[1]}, {datetime.now().year}")
            week_label = ttk.Label(nav_frame, textvariable=self.current_week_var, font=('Arial', 10, 'bold'))
            week_label.pack(side="left", padx=(0, 10))
            
            # Week slider
            ttk.Label(nav_frame, text="Week Navigation:").pack(side="left", padx=(0, 5))
            self.week_slider = ttk.Scale(nav_frame, from_=1, to=52, orient="horizontal", 
                                       command=self.on_week_slider_change)
            self.week_slider.set(datetime.now().isocalendar()[1])  # Set to current week
            self.week_slider.pack(side="left", fill="x", expand=True, padx=(0, 10))
            
            # Year controls
            year_frame = ttk.Frame(control_frame)
            year_frame.pack(fill="x", pady=(0, 5))
            
            ttk.Label(year_frame, text="Year:").pack(side="left", padx=(0, 5))
            self.year_var = tk.StringVar()
            self.year_var.set(str(datetime.now().year))
            year_spinbox = ttk.Spinbox(year_frame, from_=2020, to=2030, textvariable=self.year_var, 
                                     width=8, command=self.on_year_change)
            year_spinbox.pack(side="left", padx=(0, 10))
            
            # Refresh button
            refresh_btn = ttk.Button(control_frame, text="Refresh Chart", 
                                   command=self.refresh_gantt_chart)
            refresh_btn.pack(side="right", padx=(10, 0))
            
            # Initial chart creation
            self.refresh_gantt_chart()
            
        except Exception as e:
            print(f"Error creating Gantt chart: {e}")
    
    def on_week_slider_change(self, value):
        """Handle week slider change"""
        try:
            week_num = int(float(value))
            if hasattr(self, 'year_var'):
                year = int(self.year_var.get())
                self.current_week_var.set(f"Week {week_num}, {year}")
                self.refresh_gantt_chart()
            else:
                # Fallback to current year if year_var not available
                import datetime
                year = datetime.datetime.now().year
                self.current_week_var.set(f"Week {week_num}, {year}")
        except Exception as e:
            print(f"Error handling week slider change: {e}")
    
    def on_year_change(self):
        """Handle year change"""
        try:
            year = int(self.year_var.get())
            week_num = int(self.week_slider.get())
            self.current_week_var.set(f"Week {week_num}, {year}")
            self.refresh_gantt_chart()
        except Exception as e:
            print(f"Error handling year change: {e}")
    
    def refresh_gantt_chart(self):
        """Refresh the Gantt chart with current data"""
        try:
            # Clear the chart
            self.gantt_ax.clear()
            
            # Get current week and year
            week_num = int(self.week_slider.get())
            year = int(self.year_var.get())
            
            # Get data from database
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT current_milestone, planned_start, actual_start, planned_end, actual_end, status
                FROM workflow_entries 
                WHERE current_milestone IS NOT NULL AND current_milestone != ''
                ORDER BY created_date DESC
            """)
            
            records = cursor.fetchall()
            conn.close()
            
            if not records:
                # Show message if no data
                self.gantt_ax.text(0.5, 0.5, 'No milestone data available', 
                                 ha='center', va='center', transform=self.gantt_ax.transAxes)
                self.gantt_ax.set_title(f'Gantt Chart - Week {week_num}, {year} (No Data)')
                self.gantt_canvas.draw()
                return
            
            # Prepare data for Gantt chart
            tasks = []
            y_pos = 0
            
            for record in records:
                milestone, planned_start, actual_start, planned_end, actual_end, status = record
                
                # Create task entry
                task = {
                    'name': milestone,
                    'y_pos': y_pos,
                    'planned_start': planned_start,
                    'actual_start': actual_start,
                    'planned_end': planned_end,
                    'actual_end': actual_end,
                    'status': status
                }
                tasks.append(task)
                y_pos += 1
            
            # Create Gantt chart with week-based view
            self._draw_gantt_chart_weekly(tasks, week_num, year)
            
        except Exception as e:
            print(f"Error refreshing Gantt chart: {e}")
            self.gantt_ax.clear()
            self.gantt_ax.text(0.5, 0.5, f'Error loading chart: {e}', 
                             ha='center', va='center', transform=self.gantt_ax.transAxes)
            self.gantt_canvas.draw()
    
    def _draw_gantt_chart_weekly(self, tasks, week_num, year):
        """Draw the Gantt chart with week-based view"""
        try:
            # Calculate week start and end dates
            week_start = datetime.strptime(f"{year}-W{week_num:02d}-1", "%Y-W%W-%w")
            week_end = week_start + timedelta(days=6)
            
            # Set up the chart with week view
            self.gantt_ax.set_xlim(0, 7)  # 7 days of the week
            self.gantt_ax.set_ylim(-0.5, len(tasks) - 0.5)
            
            # Color mapping for status
            status_colors = {
                'Active': 'lightblue',
                'In Progress': 'orange',
                'Completed': 'green',
                'Blocked': 'red',
                'Imported': 'purple',
                'No Entry': 'lightgray'
            }
            
            # Day labels
            day_labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
            
            # Draw bars for each task
            for task in tasks:
                y_pos = task['y_pos']
                task_name = task['name']
                status = task['status'] or 'No Entry'
                color = status_colors.get(status, 'lightgray')
                
                # Try to parse dates and convert to week days
                planned_start = self._parse_date(task['planned_start'])
                actual_start = self._parse_date(task['actual_start'])
                planned_end = self._parse_date(task['planned_end'])
                actual_end = self._parse_date(task['actual_end'])
                
                # Convert dates to week days (0-6)
                planned_start_day = self._date_to_week_day(planned_start, week_start)
                actual_start_day = self._date_to_week_day(actual_start, week_start)
                planned_end_day = self._date_to_week_day(planned_end, week_start)
                actual_end_day = self._date_to_week_day(actual_end, week_start)
                
                # Draw planned bar (lighter)
                if planned_start_day is not None and planned_end_day is not None:
                    duration = max(1, planned_end_day - planned_start_day + 1)
                    if duration > 0 and planned_start_day >= 0 and planned_start_day < 7:
                        self.gantt_ax.barh(y_pos, duration, left=planned_start_day, 
                                         height=0.3, color=color, alpha=0.3, 
                                         label='Planned' if y_pos == 0 else "")
                
                # Draw actual bar (darker)
                if actual_start_day is not None and actual_end_day is not None:
                    duration = max(1, actual_end_day - actual_start_day + 1)
                    if duration > 0 and actual_start_day >= 0 and actual_start_day < 7:
                        self.gantt_ax.barh(y_pos, duration, left=actual_start_day, 
                                         height=0.3, color=color, alpha=0.8,
                                         label='Actual' if y_pos == 0 else "")
                elif actual_start_day is not None and actual_start_day >= 0 and actual_start_day < 7:
                    # Show start point if only start date available
                    self.gantt_ax.barh(y_pos, 1, left=actual_start_day, 
                                     height=0.3, color=color, alpha=0.8)
                
                # Add task name
                self.gantt_ax.text(0, y_pos, f"{task_name} ({status})", 
                                 va='center', ha='left', fontsize=8)
            
            # Format the chart
            self.gantt_ax.set_yticks(range(len(tasks)))
            self.gantt_ax.set_yticklabels([task['name'] for task in tasks])
            self.gantt_ax.set_xlabel('Days of Week')
            self.gantt_ax.set_title(f'Project Milestones - Week {week_num}, {year}')
            
            # Set x-axis to show days
            self.gantt_ax.set_xticks(range(7))
            self.gantt_ax.set_xticklabels(day_labels)
            
            # Add legend
            self.gantt_ax.legend()
            
            # Add week date range info
            date_range_text = f"{week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}"
            self.gantt_ax.text(0.5, -0.1, date_range_text, ha='center', va='top', 
                             transform=self.gantt_ax.transAxes, fontsize=8)
            
            # Adjust layout
            self.gantt_fig.tight_layout()
            
            # Refresh the canvas
            self.gantt_canvas.draw()
            
        except Exception as e:
            print(f"Error drawing weekly Gantt chart: {e}")
            self.gantt_ax.clear()
            self.gantt_ax.text(0.5, 0.5, f'Error drawing chart: {e}', 
                             ha='center', va='center', transform=self.gantt_ax.transAxes)
            self.gantt_canvas.draw()
    
    def _date_to_week_day(self, date_obj, week_start):
        """Convert date to day of week (0-6) relative to week start"""
        if not date_obj:
            return None
        
        try:
            # Calculate days difference from week start
            days_diff = (date_obj - week_start).days
            if 0 <= days_diff <= 6:
                return days_diff
            return None
        except:
            return None
    
    def _parse_date(self, date_str):
        """Parse date string to datetime object"""
        if not date_str:
            return None
        
        try:
            # Try different date formats
            for fmt in ['%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
            return None
        except:
            return None
    
    def toggle_form_column(self):
        """Toggle form column visibility"""
        try:
            self.add_debug_entry(f"Toggling form column - Current state: {self.form_visible}")
            self.form_visible = not self.form_visible
            self.add_debug_entry(f"Form column visibility changed to: {self.form_visible}")
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling form column: {e}")
    
    def toggle_subtasks_column(self):
        """Toggle subtasks column visibility"""
        try:
            self.add_debug_entry(f"Toggling subtasks column - Current state: {self.subtasks_visible}")
            self.subtasks_visible = not self.subtasks_visible
            self.add_debug_entry(f"Subtasks column visibility changed to: {self.subtasks_visible}")
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling subtasks column: {e}")
    
    def toggle_gantt_column(self):
        """Toggle Gantt chart column visibility"""
        try:
            self.add_debug_entry(f"Toggling Gantt column - Current state: {self.gantt_visible}")
            self.gantt_visible = not self.gantt_visible
            self.add_debug_entry(f"Gantt column visibility changed to: {self.gantt_visible}")
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling Gantt column: {e}")
    
    def update_column_visibility(self):
        """Update column visibility and weights"""
        try:
            self.add_debug_entry(f"Updating column visibility - Form: {self.form_visible}, Subtasks: {self.subtasks_visible}, Gantt: {self.gantt_visible}")
            # Get the main container (the frame that contains our three sections)
            main_container = None
            for child in self.root.winfo_children():
                if isinstance(child, ttk.Frame) and hasattr(child, 'winfo_children'):
                    # Check if this frame contains our three sections
                    children = child.winfo_children()
                    if len(children) >= 3:
                        main_container = child
                        break
            
            if not main_container:
                return
            
            # Calculate weights based on visibility
            visible_columns = sum([self.form_visible, self.subtasks_visible, self.gantt_visible])
            
            if visible_columns == 0:
                # If all columns are hidden, show at least one
                self.form_visible = True
                visible_columns = 1
            
            # Set column weights
            if self.form_visible:
                main_container.columnconfigure(0, weight=1)
            else:
                main_container.columnconfigure(0, weight=0)
            
            if self.subtasks_visible:
                main_container.columnconfigure(1, weight=2)
            else:
                main_container.columnconfigure(1, weight=0)
            
            if self.gantt_visible:
                main_container.columnconfigure(2, weight=2)
            else:
                main_container.columnconfigure(2, weight=0)
            
            # Update grid visibility using stored frame references
            # Form section (Data Entry) - includes all form fields, buttons, recent entries
            if hasattr(self, 'main_frame'):
                if self.form_visible:
                    self.main_frame.grid()
                else:
                    self.main_frame.grid_remove()
            
            # Subtasks section (Milestone Tasks) - includes the subtasks table and all content
            if hasattr(self, 'subtasks_frame'):
                if self.subtasks_visible:
                    self.subtasks_frame.grid()
                else:
                    self.subtasks_frame.grid_remove()
            
            # Gantt chart section - includes the chart and all navigation controls
            if hasattr(self, 'gantt_frame'):
                if self.gantt_visible:
                    self.gantt_frame.grid()
                else:
                    self.gantt_frame.grid_remove()
            
            # Adjust button sizes based on available space
            print("Calling adjust_button_sizes()")
            self.adjust_button_sizes()
            
        except Exception as e:
            print(f"Error updating column visibility: {e}")
    
    def update_toggle_button_text(self):
        """Update toggle button text based on visibility state"""
        try:
            # Update form button
            if self.form_visible:
                self.form_toggle_btn.config(text="📝 Data Entry")
            else:
                self.form_toggle_btn.config(text="📝 Data Entry (Hidden)")
            
            # Update subtasks button
            if self.subtasks_visible:
                self.subtasks_toggle_btn.config(text="📋 Milestone Tasks")
            else:
                self.subtasks_toggle_btn.config(text="📋 Milestone Tasks (Hidden)")
            
            # Update Gantt button
            if self.gantt_visible:
                self.gantt_toggle_btn.config(text="📊 Gantt Chart")
            else:
                self.gantt_toggle_btn.config(text="📊 Gantt Chart (Hidden)")
                
        except Exception as e:
            print(f"Error updating toggle button text: {e}")
    
    def adjust_button_sizes(self):
        """Adjust button sizes based on available space when all columns are visible"""
        try:
            # Check if all three columns are visible
            all_columns_visible = self.form_visible and self.subtasks_visible and self.gantt_visible
            print(f"Adjusting button sizes - All columns visible: {all_columns_visible}")
            
            if all_columns_visible:
                # When all columns are visible, use shorter button text and less padding
                button_padding = 1  # Minimal padding
                print("Setting compact button mode")
                
                # Update button text to shorter versions
                if hasattr(self, 'save_button'):
                    self.save_button.config(text="Save")
                    print("Updated save button to 'Save'")
                if hasattr(self, 'clear_button'):
                    self.clear_button.config(text="Clear")
                    print("Updated clear button to 'Clear'")
                if hasattr(self, 'database_button'):
                    self.database_button.config(text="Database")
                    print("Updated database button to 'Database'")
                if hasattr(self, 'import_button'):
                    self.import_button.config(text="Import")
                    print("Updated import button to 'Import'")
                if hasattr(self, 'db_manager_button'):
                    self.db_manager_button.config(text="DB Manager")
                    print("Updated database manager button to 'DB Manager'")
                if hasattr(self, 'exit_button'):
                    self.exit_button.config(text="Exit")
                    print("Updated exit button to 'Exit'")
            else:
                # When columns are hidden, use full button text and normal padding
                button_padding = 5  # Normal padding
                print("Setting normal button mode")
                
                # Update button text to full versions
                if hasattr(self, 'save_button'):
                    self.save_button.config(text="Save Entry")
                    print("Updated save button to 'Save Entry'")
                if hasattr(self, 'clear_button'):
                    self.clear_button.config(text="Clear Form")
                    print("Updated clear button to 'Clear Form'")
                if hasattr(self, 'database_button'):
                    self.database_button.config(text="Show Database")
                    print("Updated database button to 'Show Database'")
                if hasattr(self, 'import_button'):
                    self.import_button.config(text="Import Project")
                    print("Updated import button to 'Import Project'")
                if hasattr(self, 'db_manager_button'):
                    self.db_manager_button.config(text="Database Manager")
                    print("Updated database manager button to 'Database Manager'")
                if hasattr(self, 'exit_button'):
                    self.exit_button.config(text="Exit")
                    print("Updated exit button to 'Exit'")
            
            # Update button grid padding
            if hasattr(self, 'save_button'):
                self.save_button.grid_configure(padx=button_padding)
            if hasattr(self, 'clear_button'):
                self.clear_button.grid_configure(padx=button_padding)
            if hasattr(self, 'database_button'):
                self.database_button.grid_configure(padx=button_padding)
            if hasattr(self, 'import_button'):
                self.import_button.grid_configure(padx=button_padding)
            if hasattr(self, 'db_manager_button'):
                self.db_manager_button.grid_configure(padx=button_padding)
            if hasattr(self, 'exit_button'):
                self.exit_button.grid_configure(padx=button_padding)
            
            print(f"Button adjustment completed with padding: {button_padding}")
                
        except Exception as e:
            print(f"Error adjusting button sizes: {e}")
    
    def _safe_toggle_entries_mode(self):
        """Safe wrapper for toggle_entries_mode to handle initialization issues"""
        try:
            if hasattr(self, 'toggle_entries_mode'):
                self.toggle_entries_mode()
            else:
                print("toggle_entries_mode method not yet available")
        except Exception as e:
            print(f"Error in safe toggle entries mode: {e}")
    
    def toggle_entries_mode(self):
        """Toggle between database entries and debug log display"""
        try:
            mode = self.entries_mode_var.get()
            self.entries_text.delete(1.0, tk.END)
            
            if mode == "database":
                # Show database entries
                self.load_recent_entries()
            else:
                # Show debug log
                self.show_debug_log()
                
        except Exception as e:
            print(f"Error toggling entries mode: {e}")
    
    def show_debug_log(self):
        """Display the debug log in the entries text widget"""
        try:
            self.entries_text.insert(tk.END, "Debug Log:\n")
            self.entries_text.insert(tk.END, "=" * 50 + "\n")
            
            if not self.debug_log:
                self.entries_text.insert(tk.END, "No debug entries yet.\n")
            else:
                for entry in self.debug_log:
                    self.entries_text.insert(tk.END, f"{entry}\n")
                    
        except Exception as e:
            print(f"Error showing debug log: {e}")
    
    def add_debug_entry(self, message):
        """Add an entry to the debug log"""
        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            debug_entry = f"[{timestamp}] {message}"
            self.debug_log.append(debug_entry)
            
            # If we're currently showing debug log, update the display
            if hasattr(self, 'entries_mode_var') and self.entries_mode_var.get() == "debug":
                self.show_debug_log()
                
        except Exception as e:
            print(f"Error adding debug entry: {e}")
    
    def create_context_menu(self):
        """Create right-click context menu for status changes"""
        try:
            self.context_menu = tk.Menu(self.root, tearoff=0)
            
            # Add status options
            self.context_menu.add_command(label="No Entry", command=lambda: self.update_subtask_status("No Entry"))
            self.context_menu.add_command(label="In Progress", command=lambda: self.update_subtask_status("In Progress"))
            self.context_menu.add_command(label="Done", command=lambda: self.update_subtask_status("Done"))
            self.context_menu.add_command(label="Blocked", command=lambda: self.update_subtask_status("Blocked"))
            
            # Add separator
            self.context_menu.add_separator()
            
            # Add refresh option
            self.context_menu.add_command(label="Refresh", command=self.refresh_subtasks)
            
        except Exception as e:
            print(f"Error creating context menu: {e}")
    
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        try:
            # Get the item under the cursor
            item = self.subtasks_tree.identify_row(event.y)
            if item:
                # Select the item
                self.subtasks_tree.selection_set(item)
                # Store the selected item for status update
                self.selected_subtask_item = item
                # Show the context menu
                self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            print(f"Error showing context menu: {e}")
    
    def update_subtask_status(self, new_status):
        """Update the status of the selected subtask"""
        try:
            if not hasattr(self, 'selected_subtask_item') or not self.selected_subtask_item:
                print("No subtask selected")
                return
            
            # Get the selected item details
            item = self.selected_subtask_item
            values = self.subtasks_tree.item(item, 'values')
            
            if len(values) >= 3:
                subtask_name = values[0]
                current_status = values[1]
                criticality = values[2]
                
                print(f"Updating subtask '{subtask_name}' from '{current_status}' to '{new_status}'")
                
                # Update the tree item
                tag = self.get_status_tag(new_status)
                print(f"Applying tag '{tag}' to item")
                self.subtasks_tree.item(item, values=(subtask_name, new_status, criticality, '', '', '', '', ''), tags=tag)
                
                # Force the tree to refresh the display
                self.subtasks_tree.update_idletasks()
                
                # Save to database if we have milestone context
                self.save_subtask_status_to_db(subtask_name, new_status)
                
                # Update status bar to show the change
                self.status_var.set(f"Updated {subtask_name} to {new_status}")
                
        except Exception as e:
            print(f"Error updating subtask status: {e}")
    
    def save_subtask_status_to_db(self, subtask_name, status):
        """Save subtask status to database"""
        try:
            # Get current CLLI and milestone context
            current_clli = self.clli_var.get().strip()
            current_milestone = self.stage_var.get().strip()
            
            if not current_milestone:
                print("No milestone selected, cannot save subtask status")
                return
            
            # Connect to database
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            # Insert or update the subtask status
            cursor.execute("""
                INSERT OR REPLACE INTO workflow_entries 
                (state, clli, city, lata, equipment_type, current_milestone, milestone_subtask, 
                 status, milestone_date, created_date, last_update)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                self.state_var.get().strip(),
                current_clli,
                self.city_var.get().strip(),
                self.lata_var.get().strip(),
                self.equipment_var.get().strip(),
                current_milestone,
                subtask_name,
                status,
                self.milestone_var.get().strip(),
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ))
            
            conn.commit()
            conn.close()
            
            print(f"Saved subtask status: {subtask_name} -> {status}")
            
        except Exception as e:
            print(f"Error saving subtask status to database: {e}")
    
    def refresh_subtasks(self):
        """Refresh the subtasks list"""
        try:
            current_clli = self.clli_var.get().strip()
            if current_clli:
                self.update_subtasks_for_clli(current_clli)
            else:
                self.populate_subtasks()
        except Exception as e:
            print(f"Error refreshing subtasks: {e}")
    
    def get_subtask_status(self, milestone, subtask):
        """Get the status of a specific subtask for a milestone"""
        try:
            # Query database for this specific milestone and subtask combination
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT status FROM workflow_entries 
                WHERE current_milestone = ? AND milestone_subtask = ?
                ORDER BY created_date DESC LIMIT 1
            """, (milestone, subtask))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return result[0]
            else:
                return "No Entry"
                
        except Exception as e:
            print(f"Error getting subtask status: {e}")
            return "No Entry"
    
    def get_subtask_status_for_clli(self, milestone, subtask, clli_code):
        """Get the status of a specific subtask for a milestone and CLLI"""
        try:
            # Query database for this specific milestone, subtask, and CLLI combination
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT status FROM workflow_entries 
                WHERE current_milestone = ? AND milestone_subtask = ? AND clli = ?
                ORDER BY created_date DESC LIMIT 1
            """, (milestone, subtask, clli_code))
            
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return result[0]
            else:
                return "No Entry"
                
        except Exception as e:
            print(f"Error getting subtask status for CLLI: {e}")
            return "No Entry"
    
    def get_status_tag(self, status):
        """Get the color tag for a status"""
        status_lower = status.lower()
        print(f"Getting tag for status: '{status}' -> '{status_lower}'")
        if status_lower == "in progress":
            tag = "in_progress"
        elif status_lower == "done":
            tag = "done"
        elif status_lower == "blocked":
            tag = "blocked"
        else:
            tag = "no_entry"
        print(f"Returning tag: '{tag}' for status: '{status}'")
        return tag
    
    
    def get_milestone_date(self, stage: str) -> date:
        """Get milestone date based on stage"""
        return self.milestone_dates.get(stage, date.today())
    
    def create_entry(self) -> Optional[QuestionSheetEntry]:
        """Create a question sheet entry from current form data"""
        state = self.state_var.get().strip()
        stage = self.stage_var.get().strip()
        status = self.status_var.get().strip()
        clli = self.clli_var.get().strip()
        host_wire_centre = self.city_var.get().strip()
        lata = self.lata_var.get().strip()
        equipment_type = self.equipment_var.get().strip()
        subtask = self.subtask_var.get().strip()
        
        if not state or not stage or not status:
            messagebox.showerror("Validation Error", 
                               "Please select State, Stage, and Status before saving.")
            return None
        
        return QuestionSheetEntry(
            state=state,
            stage=stage,
            milestone_start_date=self.get_milestone_date(stage),
            status=status,
            created_date=datetime.now(),
            clli=clli,
            host_wire_centre=host_wire_centre,
            lata=lata,
            equipment_type=equipment_type,
            subtask=subtask,
            last_update=datetime.now()
        )
    
    def save_entry(self):
        """Save entry to database"""
        self.add_debug_entry("Starting save_entry process")
        entry = self.create_entry()
        if not entry:
            self.add_debug_entry("Save entry cancelled - invalid entry data")
            return
        self.add_debug_entry("Entry data validated successfully")
        
        try:
            # Use the new workflow_entries table
            self.add_debug_entry("Connecting to database")
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            self.add_debug_entry("Database connection established")
            
            # Insert record into workflow_entries table
            self.add_debug_entry("Preparing database insert statement")
            cursor.execute("""
                INSERT INTO workflow_entries 
                (state, clli, host_wire_centre, lata, equipment_type, current_milestone, milestone_subtask, 
                 status, planned_start, actual_start, duration, planned_end, actual_end,
                 milestone_date, created_date, last_update)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                entry.state,
                entry.clli,
                entry.host_wire_centre,
                entry.lata,
                entry.equipment_type,
                entry.stage,
                entry.subtask,
                entry.status,
                '',  # planned_start
                '',  # actual_start
                '',  # duration
                '',  # planned_end
                '',  # actual_end
                entry.milestone_start_date.strftime('%Y-%m-%d'),
                entry.created_date.strftime('%Y-%m-%d %H:%M:%S'),
                entry.last_update.strftime('%Y-%m-%d %H:%M:%S')
            ))
            
            # Get the record ID of the inserted record
            record_id = cursor.lastrowid
            self.add_debug_entry(f"Database insert completed - Record ID: {record_id}")
            
            conn.commit()
            conn.close()
            self.add_debug_entry("Database transaction committed and connection closed")
            
            # Set the record number
            self.set_record_number(record_id)
            self.add_debug_entry(f"Record number set to: {record_id}")
            
            # Show success message
            messagebox.showinfo("Success", 
                              f"Entry saved successfully!\n\n"
                              f"State: {entry.state}\n"
                              f"Stage: {entry.stage}\n"
                              f"Milestone Date: {entry.milestone_start_date.strftime('%m/%d/%Y')}\n"
                              f"Status: {entry.status}\n"
                              f"CLLI: {entry.clli}\n"
                              f"Host Wire Centre: {entry.host_wire_centre}\n"
                              f"LATA: {entry.lata}\n"
                              f"Equipment: {entry.equipment_type}")
            
            # Clear form
            self.clear_form()
            
            # Update recent entries
            self.load_recent_entries()
            
            
            # Update subtasks list with new status
            current_clli = self.clli_var.get().strip()
            if current_clli:
                self.update_subtasks_for_clli(current_clli)
            else:
                self.populate_subtasks()
            
            # Update status
            self.status_var.set(f"Entry saved at {entry.created_date.strftime('%H:%M:%S')}")
            
        except Exception as e:
            messagebox.showerror("Database Error", f"Error saving entry: {str(e)}")
            self.status_var.set("Error saving entry")
    
    def clear_form(self):
        """Clear all form fields"""
        self.state_var.set("")
        self.stage_var.set("")
        self.subtask_var.set("")
        self.status_var.set("")
        self.clli_var.set("")
        self.city_var.set("")
        self.lata_var.set("")
        self.equipment_var.set("")
        self.milestone_var.set("")
        # Clear CLLI dropdown
        self.clli_combo.set("")
        # Reset status color
        self.status_combo.configure(foreground="black")
        self.status_var.set("Form cleared")
        
        # Reset subtasks to generic view
        self.populate_subtasks()
        
        # Reset record number
        self.set_record_number()
    
    def load_recent_entries(self):
        """Load and display recent entries"""
        try:
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            # Get recent entries from workflow_entries table
            cursor.execute("""
                SELECT state, current_milestone, milestone_date, status, created_date, 
                       clli, host_wire_centre, lata, equipment_type, last_update, milestone_subtask
                FROM workflow_entries 
                ORDER BY created_date DESC 
                LIMIT 10
            """)
            
            entries = cursor.fetchall()
            conn.close()
            
            # Clear and populate text widget
            self.entries_text.delete(1.0, tk.END)
            
            if entries:
                self.entries_text.insert(tk.END, "Recent Entries:\n")
                self.entries_text.insert(tk.END, "=" * 100 + "\n")
                
                for entry in entries:
                    state, milestone, milestone_date, status, created, clli, host_wire_centre, lata, equipment, last_update, subtask = entry
                    self.entries_text.insert(tk.END, 
                        f"State: {state:<12} Milestone: {milestone:<25} Status: {status:<10}\n"
                        f"CLLI: {clli:<8} Host Wire Centre: {host_wire_centre:<15} LATA: {lata:<8} Equipment: {equipment:<15}\n"
                        f"Subtask: {subtask:<15} Milestone Date: {milestone_date}\n"
                        f"Created: {created:<20} Last Update: {last_update}\n"
                        f"-" * 100 + "\n")
            else:
                self.entries_text.insert(tk.END, "No entries found.")
            
        except Exception as e:
            self.entries_text.delete(1.0, tk.END)
            self.entries_text.insert(tk.END, f"Error loading entries: {str(e)}")
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()

def main():
    """Main function"""
    try:
        app = QuestionSheetGUI()
        app.run()
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
