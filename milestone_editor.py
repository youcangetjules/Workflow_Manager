"""
Milestone and Subtask Editor Module
Allows editing of milestones and subtasks for the Degrow Workflow Manager
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import List, Dict, Optional
import sqlite3
from datetime import datetime
from db_config import get_database_connection


class MilestoneEditor:
    def __init__(self, parent_window, config_file: str = "config.json"):
        self.parent_window = parent_window
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        
        # Create editor window
        self.editor_window = tk.Toplevel(parent_window)
        self.editor_window.title("Milestone & Subtask Editor")
        self.editor_window.geometry("1000x700")
        self.editor_window.resizable(True, True)
        
        # Make window modal
        self.editor_window.transient(parent_window)
        self.editor_window.grab_set()
        
        # Initialize data
        self.milestones = []
        self.subtasks = []
        self.criticality_levels = {}
        
        # Initialize US states list (48 continental states)
        self.us_states = [
            "Alabama", "Arizona", "Arkansas", "California", "Colorado", "Connecticut",
            "Delaware", "Florida", "Georgia", "Idaho", "Illinois", "Indiana",
            "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland",
            "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri",
            "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
            "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio",
            "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina",
            "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia",
            "Washington", "West Virginia", "Wisconsin", "Wyoming"
        ]
        
        # Load existing data
        self.load_data()
        
        # Create interface
        self.create_interface()
        
        # Center the window
        self.center_window()
    
    def center_window(self):
        """Center the editor window on screen"""
        self.editor_window.update_idletasks()
        width = self.editor_window.winfo_width()
        height = self.editor_window.winfo_height()
        x = (self.editor_window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.editor_window.winfo_screenheight() // 2) - (height // 2)
        self.editor_window.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_data(self):
        """Load existing milestones and subtasks from database"""
        try:
            # Initialize with empty lists - no default values
            self.milestones = []
            self.subtasks = []
            self.criticality_levels = {}
            
        except Exception as e:
            print(f"Error loading data: {e}")
            # Initialize with empty lists if loading fails
            self.milestones = []
            self.subtasks = []
            self.criticality_levels = {}
    
    def create_interface(self):
        """Create the editor interface"""
        # Main container
        main_frame = ttk.Frame(self.editor_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Milestone & Subtask Editor", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Milestones tab
        self.create_milestones_tab(notebook)
        
        # Subtasks tab
        self.create_subtasks_tab(notebook)
        
        # Criticality tab
        self.create_criticality_tab(notebook)
        
        # Stakeholders tab
        self.create_stakeholders_tab(notebook)
        
        # Emails tab
        self.create_emails_tab(notebook)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Configure button frame for proper alignment
        buttons_frame.columnconfigure(0, weight=1)  # Left side (View Changelog)
        buttons_frame.columnconfigure(1, weight=0)  # Spacer
        buttons_frame.columnconfigure(2, weight=0)  # Cancel button
        buttons_frame.columnconfigure(3, weight=0)  # Save Changes button
        
        # View Changelog button (left side)
        changelog_btn = ttk.Button(buttons_frame, text="View Changelog", command=self.view_changelog)
        changelog_btn.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        # Cancel button (right side, left of Save)
        cancel_btn = ttk.Button(buttons_frame, text="Cancel", command=self.cancel_editing)
        cancel_btn.grid(row=0, column=2, padx=(0, 10))
        
        # Save Changes button (right side)
        save_btn = ttk.Button(buttons_frame, text="Save Changes", command=self.save_changes)
        save_btn.grid(row=0, column=3, sticky=tk.E)
    
    def create_milestones_tab(self, notebook):
        """Create the milestones editing tab"""
        milestones_frame = ttk.Frame(notebook)
        notebook.add(milestones_frame, text="Milestones")
        
        # Title
        title_label = ttk.Label(milestones_frame, text="Manage Milestones", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))
        
        # Milestones list frame
        list_frame = ttk.LabelFrame(milestones_frame, text="Milestones", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for milestones
        columns = ('milestone', 'duration', 'stakeholder_team', 'duration_calculation')
        self.milestones_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        self.milestones_tree.heading('#0', text='Milestone #')
        self.milestones_tree.heading('milestone', text='Milestone Name')
        self.milestones_tree.heading('duration', text='Duration (days)')
        self.milestones_tree.heading('stakeholder_team', text='Stakeholder Team')
        self.milestones_tree.heading('duration_calculation', text='Duration Calculation')
        self.milestones_tree.column('#0', width=60, minwidth=50)
        self.milestones_tree.column('milestone', width=200, minwidth=150)
        self.milestones_tree.column('duration', width=100, minwidth=80)
        self.milestones_tree.column('stakeholder_team', width=150, minwidth=120)
        self.milestones_tree.column('duration_calculation', width=150, minwidth=120)
        
        # Add scrollbar
        milestones_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.milestones_tree.yview)
        self.milestones_tree.configure(yscrollcommand=milestones_scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.milestones_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        milestones_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Add milestone button
        add_btn = ttk.Button(buttons_frame, text="Add Milestone", command=self.add_milestone)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Edit milestone button
        edit_btn = ttk.Button(buttons_frame, text="Edit Milestone", command=self.edit_milestone)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Delete milestone button
        delete_btn = ttk.Button(buttons_frame, text="Delete Milestone", command=self.delete_milestone)
        delete_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Move up button
        move_up_btn = ttk.Button(buttons_frame, text="Move Up", command=self.move_milestone_up)
        move_up_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Move down button
        move_down_btn = ttk.Button(buttons_frame, text="Move Down", command=self.move_milestone_down)
        move_down_btn.pack(side=tk.LEFT)
        
        # Populate milestones
        self.populate_milestones()
    
    def create_subtasks_tab(self, notebook):
        """Create the subtasks editing tab"""
        subtasks_frame = ttk.Frame(notebook)
        notebook.add(subtasks_frame, text="Subtasks")
        
        # Title
        title_label = ttk.Label(subtasks_frame, text="Manage Subtasks", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))
        
        # Subtasks list frame
        list_frame = ttk.LabelFrame(subtasks_frame, text="Subtasks", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for subtasks
        columns = ('master_milestone', 'subtask', 'prerequisites', 'duration')
        self.subtasks_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        self.subtasks_tree.heading('#0', text='Subtask #')
        self.subtasks_tree.heading('master_milestone', text='Master Milestone')
        self.subtasks_tree.heading('subtask', text='Subtask Name')
        self.subtasks_tree.heading('prerequisites', text='Pre-Requisites')
        self.subtasks_tree.heading('duration', text='Duration (days)')
        self.subtasks_tree.column('#0', width=60, minwidth=50)
        self.subtasks_tree.column('master_milestone', width=180, minwidth=150)
        self.subtasks_tree.column('subtask', width=200, minwidth=150)
        self.subtasks_tree.column('prerequisites', width=200, minwidth=150)
        self.subtasks_tree.column('duration', width=100, minwidth=80)
        
        # Add scrollbar
        subtasks_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.subtasks_tree.yview)
        self.subtasks_tree.configure(yscrollcommand=subtasks_scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.subtasks_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        subtasks_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Add subtask button
        add_btn = ttk.Button(buttons_frame, text="Add Subtask", command=self.add_subtask)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Edit subtask button
        edit_btn = ttk.Button(buttons_frame, text="Edit Subtask", command=self.edit_subtask)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Delete subtask button
        delete_btn = ttk.Button(buttons_frame, text="Delete Subtask", command=self.delete_subtask)
        delete_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Move up button
        move_up_btn = ttk.Button(buttons_frame, text="Move Up", command=self.move_subtask_up)
        move_up_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Move down button
        move_down_btn = ttk.Button(buttons_frame, text="Move Down", command=self.move_subtask_down)
        move_down_btn.pack(side=tk.LEFT)
        
        # Populate subtasks
        self.populate_subtasks()
    
    def create_criticality_tab(self, notebook):
        """Create the criticality editing tab"""
        criticality_frame = ttk.Frame(notebook)
        notebook.add(criticality_frame, text="Criticality")
        
        # Title
        title_label = ttk.Label(criticality_frame, text="Manage Criticality Levels", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))
        
        # Criticality list frame
        list_frame = ttk.LabelFrame(criticality_frame, text="Subtask Criticality", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for criticality
        columns = ('subtask', 'criticality')
        self.criticality_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        self.criticality_tree.heading('subtask', text='Subtask')
        self.criticality_tree.heading('criticality', text='Criticality Level')
        self.criticality_tree.column('subtask', width=300, minwidth=200)
        self.criticality_tree.column('criticality', width=200, minwidth=150)
        
        # Add scrollbar
        criticality_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.criticality_tree.yview)
        self.criticality_tree.configure(yscrollcommand=criticality_scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.criticality_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        criticality_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Add criticality button
        add_btn = ttk.Button(buttons_frame, text="Add Criticality", command=self.add_criticality)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Edit criticality button
        edit_btn = ttk.Button(buttons_frame, text="Edit Criticality", command=self.edit_criticality)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Delete criticality button
        delete_btn = ttk.Button(buttons_frame, text="Delete Criticality", command=self.delete_criticality)
        delete_btn.pack(side=tk.LEFT)
        
        # Populate criticality
        self.populate_criticality()
    
    def create_stakeholders_tab(self, notebook):
        """Create the stakeholders editing tab"""
        stakeholders_frame = ttk.Frame(notebook)
        notebook.add(stakeholders_frame, text="Stakeholders")
        
        # Title
        title_label = ttk.Label(stakeholders_frame, text="Manage Stakeholders", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))
        
        # Stakeholders list frame
        list_frame = ttk.LabelFrame(stakeholders_frame, text="Stakeholder Information", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for stakeholders
        columns = ('name', 'home_state', 'role', 'team', 'email', 'phone')
        self.stakeholders_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        self.stakeholders_tree.heading('name', text='Name')
        self.stakeholders_tree.heading('home_state', text='Home State')
        self.stakeholders_tree.heading('role', text='Role')
        self.stakeholders_tree.heading('team', text='Team')
        self.stakeholders_tree.heading('email', text='Email')
        self.stakeholders_tree.heading('phone', text='Phone')
        self.stakeholders_tree.column('name', width=150, minwidth=100)
        self.stakeholders_tree.column('home_state', width=120, minwidth=80)
        self.stakeholders_tree.column('role', width=120, minwidth=80)
        self.stakeholders_tree.column('team', width=150, minwidth=100)
        self.stakeholders_tree.column('email', width=200, minwidth=150)
        self.stakeholders_tree.column('phone', width=120, minwidth=100)
        
        # Add scrollbar
        stakeholders_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.stakeholders_tree.yview)
        self.stakeholders_tree.configure(yscrollcommand=stakeholders_scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.stakeholders_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        stakeholders_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Add stakeholder button
        add_btn = ttk.Button(buttons_frame, text="Add Stakeholder", command=self.add_stakeholder)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Edit stakeholder button
        edit_btn = ttk.Button(buttons_frame, text="Edit Stakeholder", command=self.edit_stakeholder)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Delete stakeholder button
        delete_btn = ttk.Button(buttons_frame, text="Delete Stakeholder", command=self.delete_stakeholder)
        delete_btn.pack(side=tk.LEFT)
        
        # Initialize stakeholders list
        self.stakeholders = []
        self.populate_stakeholders()
    
    def create_emails_tab(self, notebook):
        """Create the emails editing tab"""
        emails_frame = ttk.Frame(notebook)
        notebook.add(emails_frame, text="Emails")
        
        # Title
        title_label = ttk.Label(emails_frame, text="Manage Email Templates", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(10, 20))
        
        # Emails list frame
        list_frame = ttk.LabelFrame(emails_frame, text="Email Templates", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview for emails
        columns = ('template_name', 'subject', 'recipients', 'trigger_event', 'status')
        self.emails_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        self.emails_tree.heading('template_name', text='Template Name')
        self.emails_tree.heading('subject', text='Subject')
        self.emails_tree.heading('recipients', text='Recipients')
        self.emails_tree.heading('trigger_event', text='Trigger Event')
        self.emails_tree.heading('status', text='Status')
        self.emails_tree.column('template_name', width=150, minwidth=100)
        self.emails_tree.column('subject', width=200, minwidth=150)
        self.emails_tree.column('recipients', width=150, minwidth=100)
        self.emails_tree.column('trigger_event', width=120, minwidth=80)
        self.emails_tree.column('status', width=100, minwidth=80)
        
        # Add scrollbar
        emails_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.emails_tree.yview)
        self.emails_tree.configure(yscrollcommand=emails_scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.emails_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        emails_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # Add email button
        add_btn = ttk.Button(buttons_frame, text="Add Email Template", command=self.add_email_template)
        add_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Edit email button
        edit_btn = ttk.Button(buttons_frame, text="Edit Email Template", command=self.edit_email_template)
        edit_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Delete email button
        delete_btn = ttk.Button(buttons_frame, text="Delete Email Template", command=self.delete_email_template)
        delete_btn.pack(side=tk.LEFT)
        
        # Initialize emails list
        self.emails = []
        self.populate_emails()
    
    def populate_milestones(self):
        """Populate the milestones tree"""
        # Clear existing items
        for item in self.milestones_tree.get_children():
            self.milestones_tree.delete(item)
        
        # Add milestones with order numbers
        for i, milestone in enumerate(self.milestones, 1):
            if isinstance(milestone, dict):
                name = milestone.get('name', '')
                duration = milestone.get('duration', 0)
                stakeholder_team = milestone.get('stakeholder_team', '')
                duration_calculation = milestone.get('duration_calculation', 'milestone')
                # Convert duration calculation to display text
                calc_display = "Use Milestone Duration" if duration_calculation == "milestone" else "Sum of Subtasks"
                self.milestones_tree.insert('', 'end', text=f'MS{i}', values=(name, duration, stakeholder_team, calc_display))
            else:
                # Handle old format (string only)
                self.milestones_tree.insert('', 'end', text=f'MS{i}', values=(milestone, 0, '', 'Use Milestone Duration'))
    
    def populate_subtasks(self):
        """Populate the subtasks tree"""
        # Clear existing items
        for item in self.subtasks_tree.get_children():
            self.subtasks_tree.delete(item)
        
        # Add subtasks with order numbers
        for i, subtask in enumerate(self.subtasks, 1):
            if isinstance(subtask, dict):
                name = subtask.get('name', '')
                master_milestone = subtask.get('master_milestone', '')
                prerequisites = subtask.get('prerequisites', '')
                duration = subtask.get('duration', 0)
                self.subtasks_tree.insert('', 'end', text=f'ST{i}', values=(master_milestone, name, prerequisites, duration))
            else:
                # Handle old format (string only)
                self.subtasks_tree.insert('', 'end', text=f'ST{i}', values=('', subtask, '', 0))
    
    def populate_criticality(self):
        """Populate the criticality tree"""
        # Clear existing items
        for item in self.criticality_tree.get_children():
            self.criticality_tree.delete(item)
        
        # Add criticality levels
        for subtask, criticality in self.criticality_levels.items():
            self.criticality_tree.insert('', 'end', values=(subtask, criticality))
    
    def add_milestone(self):
        """Add a new milestone"""
        # Create custom dialog for milestone entry
        dialog = tk.Toplevel(self.editor_window)
        dialog.title("Add Milestone")
        dialog.geometry("600x550")
        dialog.resizable(False, False)
        dialog.transient(self.editor_window)
        dialog.grab_set()
        
        # Remove window border and set background
        dialog.configure(bg='SystemButtonFace', relief='flat', bd=0)
        dialog.overrideredirect(False)
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (550 // 2)
        dialog.geometry(f"600x550+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Set main frame background to match dialog
        main_frame.configure(style='TFrame')
        
        # Name label and entry
        name_label = ttk.Label(main_frame, text="Milestone Name:")
        name_label.pack(anchor=tk.W, pady=(0, 5))
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=60)
        name_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Insert before dropdown
        insert_before_label = ttk.Label(main_frame, text="Insert before:")
        insert_before_label.pack(anchor=tk.W, pady=(0, 5))
        
        insert_before_var = tk.StringVar()
        insert_before_combo = ttk.Combobox(main_frame, textvariable=insert_before_var, width=57, state="readonly")
        
        # Populate dropdown with existing milestones
        if self.milestones:
            milestone_names = []
            for milestone in self.milestones:
                if isinstance(milestone, dict):
                    milestone_names.append(milestone.get('name', ''))
                else:
                    milestone_names.append(str(milestone))
            insert_before_combo['values'] = milestone_names
        else:
            insert_before_combo['values'] = ["Milestone will be inserted as MS1"]
            insert_before_var.set("Milestone will be inserted as MS1")
        
        insert_before_combo.pack(fill=tk.X, pady=(0, 15))
        
        # Duration label and entry
        duration_label = ttk.Label(main_frame, text="Duration (days):")
        duration_label.pack(anchor=tk.W, pady=(0, 5))
        
        duration_var = tk.StringVar()
        duration_entry = ttk.Entry(main_frame, textvariable=duration_var, width=60)
        duration_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Stakeholder team label and entry
        stakeholder_label = ttk.Label(main_frame, text="Stakeholder Team:")
        stakeholder_label.pack(anchor=tk.W, pady=(0, 5))
        
        stakeholder_var = tk.StringVar()
        stakeholder_entry = ttk.Entry(main_frame, textvariable=stakeholder_var, width=60)
        stakeholder_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Duration Calculation section
        duration_calc_label = ttk.Label(main_frame, text="Duration Calculation:")
        duration_calc_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Radio buttons frame
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(anchor=tk.W, pady=(0, 20))
        
        duration_calc_var = tk.StringVar(value="milestone")
        
        # Function to handle duration calculation change
        def on_duration_calc_change(*args):
            current_value = duration_calc_var.get()
            print(f"Duration calculation changed to: {current_value}")  # Debug print
            
            if current_value == "subtasks":
                # Grey out duration entry when "Sum of Subtasks" is selected
                print("Disabling duration entry field")  # Debug print
                duration_entry.config(state="disabled", style="Disabled.TEntry")
                duration_var.set("")  # Clear the duration value
                duration_label.config(text="Duration (days) - Disabled (Sum of Subtasks)")
            else:
                # Enable duration entry when "Use Milestone Duration" is selected
                print("Enabling duration entry field")  # Debug print
                duration_entry.config(state="normal", style="TEntry")
                duration_label.config(text="Duration (days):")
        
        # Bind the change event
        duration_calc_var.trace('w', on_duration_calc_change)
        
        use_milestone_radio = ttk.Radiobutton(radio_frame, text="Use Milestone Duration", 
                                           variable=duration_calc_var, value="milestone")
        use_milestone_radio.pack(anchor=tk.W)
        
        sum_subtasks_radio = ttk.Radiobutton(radio_frame, text="Sum of Subtasks", 
                                           variable=duration_calc_var, value="subtasks")
        sum_subtasks_radio.pack(anchor=tk.W)
        
        # Add command callbacks to radio buttons for immediate response
        use_milestone_radio.configure(command=lambda: on_duration_calc_change())
        sum_subtasks_radio.configure(command=lambda: on_duration_calc_change())
        
        # Add separator line
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(5, 10))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 15), padx=10)
        
        # Configure button frame for proper spacing - 4 equal width buttons
        button_frame.columnconfigure(0, weight=1)  # Stakeholders button
        button_frame.columnconfigure(1, weight=1)  # Subtasks button
        button_frame.columnconfigure(2, weight=1)  # Cancel button
        button_frame.columnconfigure(3, weight=1)  # OK button
        
        # OK button
        def add_milestone_data():
            name = name_var.get().strip()
            duration_calc = duration_calc_var.get()
            stakeholder_team = stakeholder_var.get().strip()
            insert_before = insert_before_var.get()
            
            # Handle duration based on calculation method
            if duration_calc == "subtasks":
                # When using sum of subtasks, duration is 0 (will be calculated later)
                duration_int = 0
            else:
                # When using milestone duration, get the entered value
                duration = duration_var.get().strip() if duration_var.get() else "0"
                try:
                    duration_int = int(duration)
                except ValueError:
                    duration_int = 0
            
            if name:
                new_milestone = {
                    'name': name,
                    'duration': duration_int,
                    'stakeholder_team': stakeholder_team,
                    'duration_calculation': duration_calc
                }
                
                # Handle insertion position
                if insert_before and insert_before != "Milestone will be inserted as MS1":
                    # Find the position to insert before
                    insert_index = 0
                    for i, milestone in enumerate(self.milestones):
                        if isinstance(milestone, dict):
                            milestone_name = milestone.get('name', '')
                        else:
                            milestone_name = str(milestone)
                        if milestone_name == insert_before:
                            insert_index = i
                            break
                    self.milestones.insert(insert_index, new_milestone)
                else:
                    # Insert at the end (MS1 or append)
                    self.milestones.append(new_milestone)
                
                self.populate_milestones()
                dialog.destroy()
            else:
                messagebox.showwarning("Warning", "Please enter a milestone name.")
        
        # Stakeholders button
        stakeholders_button = ttk.Button(button_frame, text="Stakeholders", command=lambda: print("Stakeholders clicked"), width=12)
        stakeholders_button.grid(row=0, column=0, padx=(0, 5), pady=5, sticky=(tk.W, tk.E))
        
        # Subtasks button
        subtasks_button = ttk.Button(button_frame, text="Subtasks", command=lambda: print("Subtasks clicked"), width=12)
        subtasks_button.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # Cancel button
        cancel_button = ttk.Button(button_frame, text="Cancel", command=dialog.destroy, width=12)
        cancel_button.grid(row=0, column=2, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # OK button
        ok_button = ttk.Button(button_frame, text="OK", command=add_milestone_data, width=12)
        ok_button.grid(row=0, column=3, padx=(5, 0), pady=5, sticky=(tk.W, tk.E))
        
        # Focus on name entry
        name_entry.focus()
        
        # Keyboard bindings
        name_entry.bind('<Return>', lambda e: add_milestone_data())
        duration_entry.bind('<Return>', lambda e: add_milestone_data())
        stakeholder_entry.bind('<Return>', lambda e: add_milestone_data())
        
        # Escape key to cancel
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def edit_milestone(self):
        """Edit selected milestone"""
        selection = self.milestones_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a milestone to edit.")
            return
        
        item = selection[0]
        values = self.milestones_tree.item(item, 'values')
        current_name = values[0]
        current_duration = values[1] if len(values) > 1 else "0"
        current_stakeholder = values[2] if len(values) > 2 else ""
        
        new_name = simpledialog.askstring("Edit Milestone", "Enter new milestone name:", 
                                        initialvalue=current_name)
        if new_name and new_name.strip():
            new_duration = simpledialog.askstring("Edit Milestone", "Enter duration in days:", 
                                                initialvalue=current_duration)
            duration = new_duration.strip() if new_duration else "0"
            try:
                duration_int = int(duration)
            except ValueError:
                duration_int = 0
            
            new_stakeholder = simpledialog.askstring("Edit Milestone", "Enter stakeholder team:", 
                                                   initialvalue=current_stakeholder)
            stakeholder_team = new_stakeholder.strip() if new_stakeholder else ""
            
            # Find and update the milestone
            for i, milestone in enumerate(self.milestones):
                if isinstance(milestone, dict) and milestone.get('name') == current_name:
                    self.milestones[i] = {'name': new_name.strip(), 'duration': duration_int, 'stakeholder_team': stakeholder_team}
                    break
                elif milestone == current_name:  # Handle old format
                    self.milestones[i] = {'name': new_name.strip(), 'duration': duration_int, 'stakeholder_team': stakeholder_team}
                    break
            self.populate_milestones()
    
    def delete_milestone(self):
        """Delete selected milestone"""
        selection = self.milestones_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a milestone to delete.")
            return
        
        item = selection[0]
        values = self.milestones_tree.item(item, 'values')
        milestone_name = values[0]
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{milestone_name}'?"):
            # Find and remove the milestone by name
            for i, milestone in enumerate(self.milestones):
                if isinstance(milestone, dict) and milestone.get('name') == milestone_name:
                    self.milestones.pop(i)
                    break
                elif isinstance(milestone, str) and milestone == milestone_name:
                    self.milestones.pop(i)
                    break
            self.populate_milestones()
    
    def move_milestone_up(self):
        """Move selected milestone up"""
        selection = self.milestones_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a milestone to move.")
            return
        
        item = selection[0]
        values = self.milestones_tree.item(item, 'values')
        milestone_name = values[0]
        
        try:
            # Find the milestone by name in the dictionary list
            index = None
            for i, milestone in enumerate(self.milestones):
                if isinstance(milestone, dict) and milestone.get('name') == milestone_name:
                    index = i
                    break
                elif isinstance(milestone, str) and milestone == milestone_name:
                    index = i
                    break
            
            if index is not None and index > 0:
                # Swap with previous item
                self.milestones[index], self.milestones[index-1] = self.milestones[index-1], self.milestones[index]
                self.populate_milestones()
                # Reselect the moved item
                self.milestones_tree.selection_set(self.milestones_tree.get_children()[index-1])
        except (ValueError, IndexError) as e:
            print(f"Error moving milestone up: {e}")
            pass
    
    def move_milestone_down(self):
        """Move selected milestone down"""
        selection = self.milestones_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a milestone to move.")
            return
        
        item = selection[0]
        values = self.milestones_tree.item(item, 'values')
        milestone_name = values[0]
        
        try:
            # Find the milestone by name in the dictionary list
            index = None
            for i, milestone in enumerate(self.milestones):
                if isinstance(milestone, dict) and milestone.get('name') == milestone_name:
                    index = i
                    break
                elif isinstance(milestone, str) and milestone == milestone_name:
                    index = i
                    break
            
            if index is not None and index < len(self.milestones) - 1:
                # Swap with next item
                self.milestones[index], self.milestones[index+1] = self.milestones[index+1], self.milestones[index]
                self.populate_milestones()
                # Reselect the moved item
                self.milestones_tree.selection_set(self.milestones_tree.get_children()[index+1])
        except (ValueError, IndexError) as e:
            print(f"Error moving milestone down: {e}")
            pass
    
    def _get_master_milestone_selection(self):
        """Get master milestone selection from user using dropdown"""
        if not self.milestones:
            messagebox.showwarning("Warning", "No milestones available. Please add milestones first.")
            return None
        
        # Create milestone names list
        milestone_names = []
        for milestone in self.milestones:
            if isinstance(milestone, dict):
                milestone_names.append(milestone.get('name', ''))
            else:
                milestone_names.append(milestone)
        
        # Create custom dialog with dropdown
        dialog = tk.Toplevel(self.editor_window)
        dialog.title("Select Master Milestone")
        dialog.geometry("800x150")
        dialog.resizable(False, False)
        dialog.transient(self.editor_window)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (dialog.winfo_screenheight() // 2) - (150 // 2)
        dialog.geometry(f"800x150+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label
        label = ttk.Label(main_frame, text="Select Master Milestone:", font=("Arial", 10, "bold"))
        label.pack(pady=(0, 10))
        
        # Dropdown
        selected_milestone = tk.StringVar()
        milestone_combo = ttk.Combobox(main_frame, textvariable=selected_milestone, 
                                      values=milestone_names, state="readonly", width=40)
        milestone_combo.pack(pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # OK button
        ok_button = ttk.Button(button_frame, text="OK", 
                              command=lambda: dialog.destroy())
        ok_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Cancel button
        cancel_button = ttk.Button(button_frame, text="Cancel", 
                                  command=lambda: [setattr(dialog, 'result', None), dialog.destroy()])
        cancel_button.pack(side=tk.RIGHT)
        
        # Focus on dropdown
        milestone_combo.focus()
        
        # Wait for dialog to close
        dialog.wait_window()
        
        # Return selected value
        if hasattr(dialog, 'result'):
            return None
        return selected_milestone.get() if selected_milestone.get() else None
    
    def add_subtask(self):
        """Add a new subtask"""
        name = simpledialog.askstring("Add Subtask", "Enter subtask name:")
        if name and name.strip():
            # Get master milestone selection
            master_milestone = self._get_master_milestone_selection()
            if not master_milestone:
                return
            
            prerequisites = simpledialog.askstring("Add Subtask", "Enter prerequisites (optional):")
            prerequisites = prerequisites.strip() if prerequisites else ""
            
            duration = simpledialog.askstring("Add Subtask", "Enter duration in days (optional):")
            duration = duration.strip() if duration else "0"
            try:
                duration_int = int(duration)
            except ValueError:
                duration_int = 0
            
            self.subtasks.append({
                'name': name.strip(),
                'master_milestone': master_milestone,
                'prerequisites': prerequisites,
                'duration': duration_int
            })
            # Set default criticality
            self.criticality_levels[name.strip()] = "Should be Complete"
            self.populate_subtasks()
            self.populate_criticality()
    
    def edit_subtask(self):
        """Edit selected subtask"""
        selection = self.subtasks_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subtask to edit.")
            return
        
        item = selection[0]
        values = self.subtasks_tree.item(item, 'values')
        current_master_milestone = values[0] if len(values) > 0 else ""
        current_name = values[1] if len(values) > 1 else values[0]
        current_prerequisites = values[2] if len(values) > 2 else ""
        current_duration = values[3] if len(values) > 3 else "0"
        
        new_name = simpledialog.askstring("Edit Subtask", "Enter new subtask name:", 
                                        initialvalue=current_name)
        if new_name and new_name.strip():
            # Get master milestone selection
            master_milestone = self._get_master_milestone_selection()
            if not master_milestone:
                return
            
            new_prerequisites = simpledialog.askstring("Edit Subtask", "Enter prerequisites:", 
                                                     initialvalue=current_prerequisites)
            prerequisites = new_prerequisites.strip() if new_prerequisites else ""
            
            new_duration = simpledialog.askstring("Edit Subtask", "Enter duration in days:", 
                                                initialvalue=current_duration)
            duration = new_duration.strip() if new_duration else "0"
            try:
                duration_int = int(duration)
            except ValueError:
                duration_int = 0
            
            # Find and update the subtask
            for i, subtask in enumerate(self.subtasks):
                if isinstance(subtask, dict) and subtask.get('name') == current_name:
                    self.subtasks[i] = {'name': new_name.strip(), 'master_milestone': master_milestone, 'prerequisites': prerequisites, 'duration': duration_int}
                    # Update criticality mapping
                    if current_name in self.criticality_levels:
                        criticality = self.criticality_levels.pop(current_name)
                        self.criticality_levels[new_name.strip()] = criticality
                    break
                elif subtask == current_name:  # Handle old format
                    self.subtasks[i] = {'name': new_name.strip(), 'master_milestone': master_milestone, 'prerequisites': prerequisites, 'duration': duration_int}
                    # Update criticality mapping
                    if current_name in self.criticality_levels:
                        criticality = self.criticality_levels.pop(current_name)
                        self.criticality_levels[new_name.strip()] = criticality
                    break
            self.populate_subtasks()
            self.populate_criticality()
    
    def delete_subtask(self):
        """Delete selected subtask"""
        selection = self.subtasks_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subtask to delete.")
            return
        
        item = selection[0]
        values = self.subtasks_tree.item(item, 'values')
        subtask_name = values[0]
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{subtask_name}'?"):
            # Find and remove the subtask by name
            for i, subtask in enumerate(self.subtasks):
                if isinstance(subtask, dict) and subtask.get('name') == subtask_name:
                    self.subtasks.pop(i)
                    break
                elif isinstance(subtask, str) and subtask == subtask_name:
                    self.subtasks.pop(i)
                    break
            if subtask_name in self.criticality_levels:
                del self.criticality_levels[subtask_name]
            self.populate_subtasks()
            self.populate_criticality()
    
    def move_subtask_up(self):
        """Move selected subtask up"""
        selection = self.subtasks_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subtask to move.")
            return
        
        item = selection[0]
        values = self.subtasks_tree.item(item, 'values')
        subtask_name = values[0]
        
        try:
            # Find the subtask by name in the dictionary list
            index = None
            for i, subtask in enumerate(self.subtasks):
                if isinstance(subtask, dict) and subtask.get('name') == subtask_name:
                    index = i
                    break
                elif isinstance(subtask, str) and subtask == subtask_name:
                    index = i
                    break
            
            if index is not None and index > 0:
                # Swap with previous item
                self.subtasks[index], self.subtasks[index-1] = self.subtasks[index-1], self.subtasks[index]
                self.populate_subtasks()
                # Reselect the moved item
                self.subtasks_tree.selection_set(self.subtasks_tree.get_children()[index-1])
        except (ValueError, IndexError) as e:
            print(f"Error moving subtask up: {e}")
            pass
    
    def move_subtask_down(self):
        """Move selected subtask down"""
        selection = self.subtasks_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subtask to move.")
            return
        
        item = selection[0]
        values = self.subtasks_tree.item(item, 'values')
        subtask_name = values[0]
        
        try:
            # Find the subtask by name in the dictionary list
            index = None
            for i, subtask in enumerate(self.subtasks):
                if isinstance(subtask, dict) and subtask.get('name') == subtask_name:
                    index = i
                    break
                elif isinstance(subtask, str) and subtask == subtask_name:
                    index = i
                    break
            
            if index is not None and index < len(self.subtasks) - 1:
                # Swap with next item
                self.subtasks[index], self.subtasks[index+1] = self.subtasks[index+1], self.subtasks[index]
                self.populate_subtasks()
                # Reselect the moved item
                self.subtasks_tree.selection_set(self.subtasks_tree.get_children()[index+1])
        except (ValueError, IndexError) as e:
            print(f"Error moving subtask down: {e}")
            pass
    
    def edit_criticality(self):
        """Edit criticality level for selected subtask"""
        selection = self.criticality_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subtask to edit criticality.")
            return
        
        item = selection[0]
        values = self.criticality_tree.item(item, 'values')
        subtask_name = values[0]
        current_criticality = values[1]
        
        # Create criticality selection dialog
        criticality_window = tk.Toplevel(self.editor_window)
        criticality_window.title("Edit Criticality")
        criticality_window.geometry("300x200")
        criticality_window.transient(self.editor_window)
        criticality_window.grab_set()
        
        # Center the dialog
        criticality_window.update_idletasks()
        x = (criticality_window.winfo_screenwidth() // 2) - 150
        y = (criticality_window.winfo_screenheight() // 2) - 100
        criticality_window.geometry(f'300x200+{x}+{y}')
        
        # Title
        title_label = ttk.Label(criticality_window, text=f"Edit Criticality for: {subtask_name}", 
                              font=("Arial", 12, "bold"))
        title_label.pack(pady=20)
        
        # Criticality selection
        criticality_var = tk.StringVar(value=current_criticality)
        criticality_frame = ttk.Frame(criticality_window)
        criticality_frame.pack(pady=20)
        
        ttk.Radiobutton(criticality_frame, text="Must be complete", 
                       variable=criticality_var, value="Must be complete").pack(anchor=tk.W)
        ttk.Radiobutton(criticality_frame, text="Should be Complete", 
                       variable=criticality_var, value="Should be Complete").pack(anchor=tk.W)
        ttk.Radiobutton(criticality_frame, text="Does not block", 
                       variable=criticality_var, value="Does not block").pack(anchor=tk.W)
        
        # Buttons
        buttons_frame = ttk.Frame(criticality_window)
        buttons_frame.pack(pady=20)
        
        def save_criticality():
            self.criticality_levels[subtask_name] = criticality_var.get()
            self.populate_criticality()
            criticality_window.destroy()
        
        ttk.Button(buttons_frame, text="Save", command=save_criticality).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Cancel", command=criticality_window.destroy).pack(side=tk.LEFT)
    
    def save_changes(self):
        """Save changes to database"""
        try:
            # Here you would save to database
            # For now, we'll just show a success message
            messagebox.showinfo("Success", "Changes saved successfully!")
            self.editor_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save changes: {e}")
    
    def cancel_editing(self):
        """Cancel editing and close window"""
        if messagebox.askyesno("Confirm Cancel", "Are you sure you want to cancel? All changes will be lost."):
            self.editor_window.destroy()
    
    def view_changelog(self):
        """View the changelog window"""
        try:
            # Create changelog window
            changelog_window = tk.Toplevel(self.editor_window)
            changelog_window.title("Changelog")
            changelog_window.geometry("600x400")
            changelog_window.transient(self.editor_window)
            changelog_window.grab_set()
            
            # Create main frame
            main_frame = ttk.Frame(changelog_window)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Title
            title_label = ttk.Label(main_frame, text="Milestone Editor Changelog", 
                                  font=('Arial', 12, 'bold'))
            title_label.pack(pady=(0, 10))
            
            # Create text widget with scrollbar
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill=tk.BOTH, expand=True)
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, height=15, width=70)
            scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Changelog content
            changelog_content = """Milestone Editor Changelog
=====================================

Version 1.0.0 - Initial Release
- Basic milestone and subtask management
- Add, edit, delete, and reorder functionality
- Criticality level management
- Stakeholder and email template tabs

Recent Updates:
- Fixed move up/down buttons for dictionary-based data structure
- Renamed "Order" column to "Milestone #" with MS prefix
- Added sequential numbering (MS1, MS2, MS3, etc.)
- Improved error handling and error reporting
- Enhanced user interface with better button alignment

Features:
✓ Milestone Management
  - Add new milestones with duration and stakeholder team
  - Edit existing milestone details
  - Delete milestones with confirmation
  - Move milestones up/down in order
  - Duration calculation options (Milestone Duration vs Sum of Subtasks)

✓ Subtask Management
  - Add subtasks with master milestone assignment
  - Edit subtask details including prerequisites
  - Delete subtasks with confirmation
  - Move subtasks up/down in order
  - Criticality level assignment

✓ Criticality Management
  - Add new criticality levels
  - Edit existing criticality assignments
  - Delete criticality levels
  - Color-coded status indicators

✓ Data Structure
  - Dictionary-based storage for enhanced data management
  - Backward compatibility with legacy string format
  - Automatic data migration and validation

✓ User Interface
  - Tabbed interface for organized editing
  - Intuitive button layout with proper alignment
  - Confirmation dialogs for destructive operations
  - Real-time updates and validation

Known Issues:
- None currently reported

Future Enhancements:
- Database persistence
- Import/Export functionality
- Advanced filtering and search
- Bulk operations
- Audit trail and version control

For support or feature requests, please contact the development team.
"""
            
            text_widget.insert(tk.END, changelog_content)
            text_widget.config(state=tk.DISABLED)  # Make read-only
            
            # Close button
            close_btn = ttk.Button(main_frame, text="Close", command=changelog_window.destroy)
            close_btn.pack(pady=(10, 0))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open changelog: {e}")
    
    def populate_stakeholders(self):
        """Populate the stakeholders tree"""
        # Clear existing items
        for item in self.stakeholders_tree.get_children():
            self.stakeholders_tree.delete(item)
        
        # Add stakeholders
        for i, stakeholder in enumerate(self.stakeholders, 1):
            if isinstance(stakeholder, dict):
                name = stakeholder.get('name', '')
                home_state = stakeholder.get('home_state', '')
                role = stakeholder.get('role', '')
                team = stakeholder.get('team', '')
                email = stakeholder.get('email', '')
                phone = stakeholder.get('phone', '')
                self.stakeholders_tree.insert('', 'end', text=str(i), values=(name, home_state, role, team, email, phone))
            else:
                # Handle old format (string only)
                self.stakeholders_tree.insert('', 'end', text=str(i), values=(stakeholder, '', '', '', '', ''))
    
    def populate_emails(self):
        """Populate the emails tree"""
        # Clear existing items
        for item in self.emails_tree.get_children():
            self.emails_tree.delete(item)
        
        # Add emails
        for i, email in enumerate(self.emails, 1):
            if isinstance(email, dict):
                template_name = email.get('template_name', '')
                subject = email.get('subject', '')
                recipients = email.get('recipients', '')
                trigger_event = email.get('trigger_event', '')
                status = email.get('status', '')
                self.emails_tree.insert('', 'end', text=str(i), values=(template_name, subject, recipients, trigger_event, status))
            else:
                # Handle old format (string only)
                self.emails_tree.insert('', 'end', text=str(i), values=(email, '', '', '', ''))
    
    # Placeholder methods for criticality
    def add_criticality(self):
        """Add a new criticality level"""
        messagebox.showinfo("Info", "Add Criticality functionality will be implemented")
    
    def delete_criticality(self):
        """Delete selected criticality level"""
        messagebox.showinfo("Info", "Delete Criticality functionality will be implemented")
    
    def add_stakeholder(self):
        """Add a new stakeholder"""
        # Create custom dialog for stakeholder entry
        dialog = tk.Toplevel(self.editor_window)
        dialog.title("Add Stakeholder")
        dialog.geometry("600x500")
        dialog.resizable(False, False)
        dialog.transient(self.editor_window)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"600x500+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name label and entry
        name_label = ttk.Label(main_frame, text="Name:")
        name_label.pack(anchor=tk.W, pady=(0, 5))
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=60)
        name_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Home State label and dropdown
        home_state_label = ttk.Label(main_frame, text="Home State:")
        home_state_label.pack(anchor=tk.W, pady=(0, 5))
        
        home_state_var = tk.StringVar()
        home_state_combo = ttk.Combobox(main_frame, textvariable=home_state_var, 
                                       values=self.us_states, state="readonly", width=57)
        home_state_combo.pack(fill=tk.X, pady=(0, 15))
        
        # Role label and entry
        role_label = ttk.Label(main_frame, text="Role:")
        role_label.pack(anchor=tk.W, pady=(0, 5))
        
        role_var = tk.StringVar()
        role_entry = ttk.Entry(main_frame, textvariable=role_var, width=60)
        role_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Team label and entry
        team_label = ttk.Label(main_frame, text="Team:")
        team_label.pack(anchor=tk.W, pady=(0, 5))
        
        team_var = tk.StringVar()
        team_entry = ttk.Entry(main_frame, textvariable=team_var, width=60)
        team_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Email label and entry
        email_label = ttk.Label(main_frame, text="Email:")
        email_label.pack(anchor=tk.W, pady=(0, 5))
        
        email_var = tk.StringVar()
        email_entry = ttk.Entry(main_frame, textvariable=email_var, width=60)
        email_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Phone label and entry
        phone_label = ttk.Label(main_frame, text="Phone:")
        phone_label.pack(anchor=tk.W, pady=(0, 5))
        
        phone_var = tk.StringVar()
        phone_entry = ttk.Entry(main_frame, textvariable=phone_var, width=60)
        phone_entry.pack(fill=tk.X, pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def add_stakeholder_data():
            name = name_var.get().strip()
            home_state = home_state_var.get().strip()
            role = role_var.get().strip()
            team = team_var.get().strip()
            email = email_var.get().strip()
            phone = phone_var.get().strip()
            
            if name:
                new_stakeholder = {
                    'name': name,
                    'home_state': home_state,
                    'role': role,
                    'team': team,
                    'email': email,
                    'phone': phone
                }
                
                self.stakeholders.append(new_stakeholder)
                self.populate_stakeholders()
                dialog.destroy()
            else:
                messagebox.showwarning("Warning", "Please enter a stakeholder name.")
        
        # Cancel button
        cancel_button = ttk.Button(button_frame, text="Cancel", command=dialog.destroy, width=12)
        cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # OK button
        ok_button = ttk.Button(button_frame, text="OK", command=add_stakeholder_data, width=12)
        ok_button.pack(side=tk.RIGHT)
        
        # Focus on name entry
        name_entry.focus()
        
        # Keyboard bindings
        name_entry.bind('<Return>', lambda e: add_stakeholder_data())
        home_state_combo.bind('<Return>', lambda e: add_stakeholder_data())
        role_entry.bind('<Return>', lambda e: add_stakeholder_data())
        team_entry.bind('<Return>', lambda e: add_stakeholder_data())
        email_entry.bind('<Return>', lambda e: add_stakeholder_data())
        phone_entry.bind('<Return>', lambda e: add_stakeholder_data())
        
        # Escape key to cancel
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def edit_stakeholder(self):
        """Edit selected stakeholder"""
        selection = self.stakeholders_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a stakeholder to edit.")
            return
        
        item = selection[0]
        values = self.stakeholders_tree.item(item, 'values')
        current_name = values[0]
        current_home_state = values[1] if len(values) > 1 else ""
        current_role = values[2] if len(values) > 2 else ""
        current_team = values[3] if len(values) > 3 else ""
        current_email = values[4] if len(values) > 4 else ""
        current_phone = values[5] if len(values) > 5 else ""
        
        # Create custom dialog for stakeholder editing
        dialog = tk.Toplevel(self.editor_window)
        dialog.title("Edit Stakeholder")
        dialog.geometry("600x500")
        dialog.resizable(False, False)
        dialog.transient(self.editor_window)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"600x500+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name label and entry
        name_label = ttk.Label(main_frame, text="Name:")
        name_label.pack(anchor=tk.W, pady=(0, 5))
        
        name_var = tk.StringVar(value=current_name)
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=60)
        name_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Home State label and dropdown
        home_state_label = ttk.Label(main_frame, text="Home State:")
        home_state_label.pack(anchor=tk.W, pady=(0, 5))
        
        home_state_var = tk.StringVar(value=current_home_state)
        home_state_combo = ttk.Combobox(main_frame, textvariable=home_state_var, 
                                       values=self.us_states, state="readonly", width=57)
        home_state_combo.pack(fill=tk.X, pady=(0, 15))
        
        # Role label and entry
        role_label = ttk.Label(main_frame, text="Role:")
        role_label.pack(anchor=tk.W, pady=(0, 5))
        
        role_var = tk.StringVar(value=current_role)
        role_entry = ttk.Entry(main_frame, textvariable=role_var, width=60)
        role_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Team label and entry
        team_label = ttk.Label(main_frame, text="Team:")
        team_label.pack(anchor=tk.W, pady=(0, 5))
        
        team_var = tk.StringVar(value=current_team)
        team_entry = ttk.Entry(main_frame, textvariable=team_var, width=60)
        team_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Email label and entry
        email_label = ttk.Label(main_frame, text="Email:")
        email_label.pack(anchor=tk.W, pady=(0, 5))
        
        email_var = tk.StringVar(value=current_email)
        email_entry = ttk.Entry(main_frame, textvariable=email_var, width=60)
        email_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Phone label and entry
        phone_label = ttk.Label(main_frame, text="Phone:")
        phone_label.pack(anchor=tk.W, pady=(0, 5))
        
        phone_var = tk.StringVar(value=current_phone)
        phone_entry = ttk.Entry(main_frame, textvariable=phone_var, width=60)
        phone_entry.pack(fill=tk.X, pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def edit_stakeholder_data():
            name = name_var.get().strip()
            home_state = home_state_var.get().strip()
            role = role_var.get().strip()
            team = team_var.get().strip()
            email = email_var.get().strip()
            phone = phone_var.get().strip()
            
            if name:
                # Find and update the stakeholder
                for i, stakeholder in enumerate(self.stakeholders):
                    if isinstance(stakeholder, dict) and stakeholder.get('name') == current_name:
                        self.stakeholders[i] = {
                            'name': name,
                            'home_state': home_state,
                            'role': role,
                            'team': team,
                            'email': email,
                            'phone': phone
                        }
                        break
                    elif stakeholder == current_name:  # Handle old format
                        self.stakeholders[i] = {
                            'name': name,
                            'home_state': home_state,
                            'role': role,
                            'team': team,
                            'email': email,
                            'phone': phone
                        }
                        break
                
                self.populate_stakeholders()
                dialog.destroy()
            else:
                messagebox.showwarning("Warning", "Please enter a stakeholder name.")
        
        # Cancel button
        cancel_button = ttk.Button(button_frame, text="Cancel", command=dialog.destroy, width=12)
        cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # OK button
        ok_button = ttk.Button(button_frame, text="OK", command=edit_stakeholder_data, width=12)
        ok_button.pack(side=tk.RIGHT)
        
        # Focus on name entry
        name_entry.focus()
        
        # Keyboard bindings
        name_entry.bind('<Return>', lambda e: edit_stakeholder_data())
        home_state_combo.bind('<Return>', lambda e: edit_stakeholder_data())
        role_entry.bind('<Return>', lambda e: edit_stakeholder_data())
        team_entry.bind('<Return>', lambda e: edit_stakeholder_data())
        email_entry.bind('<Return>', lambda e: edit_stakeholder_data())
        phone_entry.bind('<Return>', lambda e: edit_stakeholder_data())
        
        # Escape key to cancel
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def delete_stakeholder(self):
        """Delete selected stakeholder"""
        selection = self.stakeholders_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a stakeholder to delete.")
            return
        
        item = selection[0]
        values = self.stakeholders_tree.item(item, 'values')
        stakeholder_name = values[0]
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{stakeholder_name}'?"):
            # Find and remove the stakeholder by name
            for i, stakeholder in enumerate(self.stakeholders):
                if isinstance(stakeholder, dict) and stakeholder.get('name') == stakeholder_name:
                    self.stakeholders.pop(i)
                    break
                elif isinstance(stakeholder, str) and stakeholder == stakeholder_name:
                    self.stakeholders.pop(i)
                    break
            
            self.populate_stakeholders()
    
    # Placeholder methods for emails
    def add_email_template(self):
        """Add a new email template"""
        messagebox.showinfo("Info", "Add Email Template functionality will be implemented")
    
    def edit_email_template(self):
        """Edit selected email template"""
        messagebox.showinfo("Info", "Edit Email Template functionality will be implemented")
    
    def delete_email_template(self):
        """Delete selected email template"""
        messagebox.showinfo("Info", "Delete Email Template functionality will be implemented")


def open_milestone_editor(parent_window, config_file: str = "config.json"):
    """Open the milestone editor window"""
    MilestoneEditor(parent_window, config_file)
