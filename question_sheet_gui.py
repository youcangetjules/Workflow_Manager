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
    def __init__(self, db_path: str = r"C:\BASHFlowSandbox\TestDatabase.accdb"):
        self.db_path = db_path
        self.states = self._initialize_states()
        self.stages = self._initialize_stages()
        self.subtasks = self._initialize_subtasks()
        self.criticality_levels = self._initialize_criticality_levels()
        self.statuses = self._initialize_statuses()
        self.equipment_types = self._initialize_equipment_types()
        self.milestone_dates = self._initialize_milestone_dates()
        
        # Initialize databases
        self._initialize_database()
        self._initialize_milestone_database()
        
        # Load Excel data for CLLI lookup
        self.clli_data = self._load_clli_data()
        self.clli_suggestions = []
        
        # Initialize record number
        self.current_record_id = None
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("Degrow Workflow Manager")
        self.root.geometry("1920x1080")
        self.root.resizable(True, True)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
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
        """Initialize SQLite database for storing workflow entries"""
        try:
            # Convert .accdb path to .db path for SQLite
            sqlite_path = self.db_path.replace('.accdb', '.db')
            
            # Create database directory if it doesn't exist
            db_dir = os.path.dirname(sqlite_path)
            if db_dir and not os.path.exists(db_dir):
                os.makedirs(db_dir)
            
            # Connect to SQLite database
            conn = sqlite3.connect(sqlite_path)
            cursor = conn.cursor()
            
            # Create workflow_entries table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS workflow_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    state TEXT NOT NULL,
                    clli TEXT,
                    host_wire_centre TEXT,
                    lata TEXT,
                    equipment_type TEXT,
                    current_milestone TEXT NOT NULL,
                    milestone_subtask TEXT,
                    status TEXT NOT NULL,
                    planned_start TEXT,
                    actual_start TEXT,
                    duration TEXT,
                    planned_end TEXT,
                    actual_end TEXT,
                    milestone_date TEXT,
                    created_date TEXT NOT NULL,
                    last_update TEXT NOT NULL
                )
            ''')
            
            # Add new columns to existing tables (migration)
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN planned_start TEXT')
            except sqlite3.OperationalError:
                pass  # Column already exists
            
            # Rename city column to host_wire_centre if it exists
            try:
                cursor.execute('ALTER TABLE workflow_entries RENAME COLUMN city TO host_wire_centre')
            except sqlite3.OperationalError:
                pass  # Column doesn't exist or already renamed
            
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN actual_start TEXT')
            except sqlite3.OperationalError:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN duration TEXT')
            except sqlite3.OperationalError:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN planned_end TEXT')
            except sqlite3.OperationalError:
                pass  # Column already exists
                
            try:
                cursor.execute('ALTER TABLE workflow_entries ADD COLUMN actual_end TEXT')
            except sqlite3.OperationalError:
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
            conn.close()
            
            print(f"Database initialized successfully: {sqlite_path}")
            
        except Exception as e:
            print(f"Error initializing database: {e}")
            import traceback
            traceback.print_exc()
    
    def _initialize_milestone_database(self):
        """Initialize milestone database"""
        try:
            # Create milestone database path
            milestone_db_path = self.db_path.replace('.accdb', '_milestones.db')
            
            conn = sqlite3.connect(milestone_db_path)
            cursor = conn.cursor()
            
            # Create milestones table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS milestones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    description TEXT,
                    order_index INTEGER DEFAULT 0,
                    created_date TEXT NOT NULL,
                    last_updated TEXT NOT NULL
                )
            ''')
            
            # Create subtasks table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS subtasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    milestone_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    description TEXT,
                    criticality TEXT DEFAULT 'Should be Complete',
                    order_index INTEGER DEFAULT 0,
                    created_date TEXT NOT NULL,
                    last_updated TEXT NOT NULL,
                    FOREIGN KEY (milestone_id) REFERENCES milestones (id) ON DELETE CASCADE
                )
            ''')
            
            conn.commit()
            conn.close()
            
            # Store milestone database path
            self.milestone_db_path = milestone_db_path
            print(f"Milestone database initialized successfully: {milestone_db_path}")
            
        except Exception as e:
            print(f"Error initializing milestone database: {e}")
            import traceback
            traceback.print_exc()
    
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
        if not query or len(query) < 2:
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
                if os.path.exists("lumen.png"):
                    pil_image = Image.open("lumen.png")
                    # Make Lumen logo MUCH larger for prominence
                    pil_image.thumbnail((600, 120), Image.Resampling.LANCZOS)
                    lumen_image = ImageTk.PhotoImage(pil_image)
                else:
                    lumen_image = tk.PhotoImage(file="lumen.png")
                lumen_label = tk.Label(logo_frame, image=lumen_image, bg='white')
                lumen_label.image = lumen_image  # Keep a reference
                lumen_label.grid(row=0, column=0, padx=(0, 20), sticky=tk.W)
                print(f"Lumen logo loaded successfully - Size: {lumen_image.width()}x{lumen_image.height()}")
            except Exception as e:
                print(f"Could not load lumen.png: {e}")
                # Create a placeholder label
                lumen_label = tk.Label(logo_frame, text="LUMEN", font=("Arial", 32, "bold"), bg='white')
                lumen_label.grid(row=0, column=0, padx=(0, 20), sticky=tk.W)
            
            # Try to load TXO logo
            try:
                from PIL import Image, ImageTk
                import os
                if os.path.exists("TXO.png"):
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
                print(f"Could not load TXO.png: {e}")
                # Create a placeholder label
                txo_label = tk.Label(logo_frame, text="TXO", font=("Arial", 16, "bold"), bg='white')
                txo_label.grid(row=0, column=2, padx=(20, 0), sticky=tk.E)
            
            # Add title in the center
            title_label = tk.Label(logo_frame, text="Degrow Workflow Manager", 
                                   font=("Arial", 18, "bold"), bg='white')
            title_label.grid(row=0, column=1, padx=20)
            
        except Exception as e:
            print(f"Error creating logo section: {e}")
            # Create a simple fallback
            fallback_frame = tk.Frame(self.root, bg='white')
            fallback_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=5)
            fallback_label = tk.Label(fallback_frame, text="Degrow Workflow Manager", 
                                     font=("Arial", 18, "bold"), bg='white')
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
        # Create logo section at the top
        self.create_logo_section()
        
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
        
        # Milestone Editor button
        self.editor_button = ttk.Button(buttons_frame, text="Milestone Editor", 
                                       command=self.open_milestone_editor)
        self.editor_button.grid(row=0, column=2, padx=5)
        
        # Show Database button
        self.database_button = ttk.Button(buttons_frame, text="Show Database", 
                                         command=self.show_database)
        self.database_button.grid(row=0, column=3, padx=5)
        
        # Import Project button
        self.import_button = ttk.Button(buttons_frame, text="Import Project", 
                                      command=self.import_project_data)
        self.import_button.grid(row=0, column=4, padx=5)
        
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
        entries_frame.rowconfigure(0, weight=1)
        
        # Recent entries list
        self.entries_text = scrolledtext.ScrolledText(entries_frame, height=10, width=50)
        self.entries_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        
        # Create subtasks list
        self.create_subtasks_list(self.subtasks_frame)
        
        # Load recent entries
        self.load_recent_entries()
        
        # Populate CLLI dropdown
        self.populate_clli_dropdown()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
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
    
    def on_clli_combo_key_release(self, event):
        """Handle key release in CLLI dropdown"""
        if event.keysym in ['Up', 'Down']:
            pass  # Arrow key handling removed
        elif event.keysym == 'Return':
            # Get current value and trigger autopopulation
            current_value = self.clli_combo.get()
            if current_value:
                self.autopopulate_from_clli(current_value)
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
            self.subtasks_tree.column('#0', width=180, minwidth=120)  # Milestone
            self.subtasks_tree.column('subtask', width=200, minwidth=150)  # Subtasks
            self.subtasks_tree.column('status', width=80, minwidth=60)  # Status
            self.subtasks_tree.column('criticality', width=120, minwidth=100)  # Criticality
            self.subtasks_tree.column('planned_start', width=100, minwidth=80)  # Planned Start
            self.subtasks_tree.column('actual_start', width=100, minwidth=80)  # Actual Start
            self.subtasks_tree.column('duration', width=80, minwidth=60)  # Duration
            self.subtasks_tree.column('planned_end', width=100, minwidth=80)  # Planned End
            self.subtasks_tree.column('actual_end', width=100, minwidth=80)  # Actual End
            
            # Add vertical and horizontal scrollbars
            subtasks_v_scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=self.subtasks_tree.yview)
            subtasks_h_scrollbar = ttk.Scrollbar(parent_frame, orient="horizontal", command=self.subtasks_tree.xview)
            self.subtasks_tree.configure(yscrollcommand=subtasks_v_scrollbar.set, xscrollcommand=subtasks_h_scrollbar.set)
            
            # Configure parent frame column weights for full width table
            parent_frame.columnconfigure(0, weight=1)  # Tree takes most space
            parent_frame.columnconfigure(1, weight=0)  # Scrollbar takes minimal space
            parent_frame.columnconfigure(2, weight=0)  # Button area takes minimal space
            
            # Grid the treeview and scrollbars
            self.subtasks_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 0))
            subtasks_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), padx=(0, 0))
            subtasks_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 0))
            
            # Configure tags for color coding AFTER tree is created
            self.subtasks_tree.tag_configure('no_entry', background='lightblue', foreground='darkblue')
            self.subtasks_tree.tag_configure('in_progress', background='lightyellow', foreground='darkorange')
            self.subtasks_tree.tag_configure('done', background='lightgreen', foreground='darkgreen')
            self.subtasks_tree.tag_configure('blocked', background='lightcoral', foreground='darkred')
            
            # Create context menu for status changes
            self.create_context_menu()
            
            # Bind right-click event
            self.subtasks_tree.bind("<Button-3>", self.show_context_menu)  # Button-3 is right-click
            
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
                    
                    # Insert subtask
                    self.subtasks_tree.insert(milestone_node, 'end', 
                                            values=(subtask['name'], status, subtask['criticality'], '', '', '', '', ''),
                                            tags=(tag,))
                
                # Expand the milestone
                self.subtasks_tree.item(milestone_node, open=True)
                
        except Exception as e:
            print(f"Error populating subtasks: {e}")
    
    def _load_milestones_from_db(self):
        """Load milestones from database"""
        try:
            if not hasattr(self, 'milestone_db_path'):
                return []
            
            conn = sqlite3.connect(self.milestone_db_path)
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
            
            conn.close()
            return milestones
            
        except Exception as e:
            print(f"Error loading milestones from database: {e}")
            return []
    
    def _load_subtasks_from_db(self, milestone_id):
        """Load subtasks for a specific milestone from database"""
        try:
            if not hasattr(self, 'milestone_db_path'):
                return []
            
            conn = sqlite3.connect(self.milestone_db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, name, description, criticality, order_index, created_date, last_updated
                FROM subtasks 
                WHERE milestone_id = ?
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
            year = int(self.year_var.get())
            self.current_week_var.set(f"Week {week_num}, {year}")
            self.refresh_gantt_chart()
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
            self.form_visible = not self.form_visible
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling form column: {e}")
    
    def toggle_subtasks_column(self):
        """Toggle subtasks column visibility"""
        try:
            self.subtasks_visible = not self.subtasks_visible
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling subtasks column: {e}")
    
    def toggle_gantt_column(self):
        """Toggle Gantt chart column visibility"""
        try:
            self.gantt_visible = not self.gantt_visible
            self.update_column_visibility()
            self.update_toggle_button_text()
        except Exception as e:
            print(f"Error toggling Gantt column: {e}")
    
    def update_column_visibility(self):
        """Update column visibility and weights"""
        try:
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
        entry = self.create_entry()
        if not entry:
            return
        
        try:
            # Use the new workflow_entries table
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            # Insert record into workflow_entries table
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
            
            conn.commit()
            conn.close()
            
            # Set the record number
            self.set_record_number(record_id)
            
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
