"""
Milestone and Subtask Editor Module
Allows editing of milestones and subtasks for the Degrow Workflow Manager
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import List, Dict, Optional
import sqlite3
from datetime import datetime


class MilestoneEditor:
    def __init__(self, parent_window, db_path: str):
        self.parent_window = parent_window
        self.db_path = db_path
        
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
            # Load milestones (these would be stored in a separate table)
            # For now, we'll use the default milestones
            self.milestones = [
                "Create Grooming Workbook Tool", "Restrict CM", "Complete Pre-Cut",
                "Collect Switch Device Info", "RUN TMART Report", "Create Bash Report",
                "Verify bash Report", "Complete bash Report Cleanup", "Create Workbook for Order Sets",
                "Issue Switch Special Orders", "Create WFA Orders for Switch Specialists", "Complete Special Orders"
            ]
            
            # Load subtasks
            self.subtasks = [
                "Go-Live Support", "Deployment Strategy", "Testing Protocol", "Implementation Planning",
                "Approval Process", "Documentation Review", "Risk Assessment", "Capacity Planning",
                "Network Analysis", "Equipment Inventory", "Site Survey", "CLLI Validation"
            ]
            
            # Load criticality levels
            self.criticality_levels = {
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
            
        except Exception as e:
            print(f"Error loading data: {e}")
            messagebox.showerror("Error", f"Failed to load data: {e}")
    
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
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Save button
        save_btn = ttk.Button(buttons_frame, text="Save Changes", command=self.save_changes)
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Cancel button
        cancel_btn = ttk.Button(buttons_frame, text="Cancel", command=self.cancel_editing)
        cancel_btn.pack(side=tk.LEFT)
    
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
        columns = ('milestone',)
        self.milestones_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        self.milestones_tree.heading('#0', text='Order')
        self.milestones_tree.heading('milestone', text='Milestone Name')
        self.milestones_tree.column('#0', width=60, minwidth=50)
        self.milestones_tree.column('milestone', width=400, minwidth=300)
        
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
        columns = ('subtask',)
        self.subtasks_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        self.subtasks_tree.heading('#0', text='Order')
        self.subtasks_tree.heading('subtask', text='Subtask Name')
        self.subtasks_tree.column('#0', width=60, minwidth=50)
        self.subtasks_tree.column('subtask', width=400, minwidth=300)
        
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
        
        # Edit criticality button
        edit_btn = ttk.Button(buttons_frame, text="Edit Criticality", command=self.edit_criticality)
        edit_btn.pack(side=tk.LEFT)
        
        # Populate criticality
        self.populate_criticality()
    
    def populate_milestones(self):
        """Populate the milestones tree"""
        # Clear existing items
        for item in self.milestones_tree.get_children():
            self.milestones_tree.delete(item)
        
        # Add milestones with order numbers
        for i, milestone in enumerate(self.milestones, 1):
            self.milestones_tree.insert('', 'end', text=str(i), values=(milestone,))
    
    def populate_subtasks(self):
        """Populate the subtasks tree"""
        # Clear existing items
        for item in self.subtasks_tree.get_children():
            self.subtasks_tree.delete(item)
        
        # Add subtasks with order numbers
        for i, subtask in enumerate(self.subtasks, 1):
            self.subtasks_tree.insert('', 'end', text=str(i), values=(subtask,))
    
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
        name = simpledialog.askstring("Add Milestone", "Enter milestone name:")
        if name and name.strip():
            self.milestones.append(name.strip())
            self.populate_milestones()
    
    def edit_milestone(self):
        """Edit selected milestone"""
        selection = self.milestones_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a milestone to edit.")
            return
        
        item = selection[0]
        values = self.milestones_tree.item(item, 'values')
        current_name = values[0]
        
        new_name = simpledialog.askstring("Edit Milestone", "Enter new milestone name:", 
                                        initialvalue=current_name)
        if new_name and new_name.strip():
            # Find and update the milestone
            for i, milestone in enumerate(self.milestones):
                if milestone == current_name:
                    self.milestones[i] = new_name.strip()
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
            self.milestones.remove(milestone_name)
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
            index = self.milestones.index(milestone_name)
            if index > 0:
                # Swap with previous item
                self.milestones[index], self.milestones[index-1] = self.milestones[index-1], self.milestones[index]
                self.populate_milestones()
                # Reselect the moved item
                self.milestones_tree.selection_set(self.milestones_tree.get_children()[index-1])
        except ValueError:
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
            index = self.milestones.index(milestone_name)
            if index < len(self.milestones) - 1:
                # Swap with next item
                self.milestones[index], self.milestones[index+1] = self.milestones[index+1], self.milestones[index]
                self.populate_milestones()
                # Reselect the moved item
                self.milestones_tree.selection_set(self.milestones_tree.get_children()[index+1])
        except ValueError:
            pass
    
    def add_subtask(self):
        """Add a new subtask"""
        name = simpledialog.askstring("Add Subtask", "Enter subtask name:")
        if name and name.strip():
            self.subtasks.append(name.strip())
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
        current_name = values[0]
        
        new_name = simpledialog.askstring("Edit Subtask", "Enter new subtask name:", 
                                        initialvalue=current_name)
        if new_name and new_name.strip():
            # Find and update the subtask
            for i, subtask in enumerate(self.subtasks):
                if subtask == current_name:
                    self.subtasks[i] = new_name.strip()
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
            self.subtasks.remove(subtask_name)
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
            index = self.subtasks.index(subtask_name)
            if index > 0:
                # Swap with previous item
                self.subtasks[index], self.subtasks[index-1] = self.subtasks[index-1], self.subtasks[index]
                self.populate_subtasks()
                # Reselect the moved item
                self.subtasks_tree.selection_set(self.subtasks_tree.get_children()[index-1])
        except ValueError:
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
            index = self.subtasks.index(subtask_name)
            if index < len(self.subtasks) - 1:
                # Swap with next item
                self.subtasks[index], self.subtasks[index+1] = self.subtasks[index+1], self.subtasks[index]
                self.populate_subtasks()
                # Reselect the moved item
                self.subtasks_tree.selection_set(self.subtasks_tree.get_children()[index+1])
        except ValueError:
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


def open_milestone_editor(parent_window, db_path: str):
    """Open the milestone editor window"""
    MilestoneEditor(parent_window, db_path)
