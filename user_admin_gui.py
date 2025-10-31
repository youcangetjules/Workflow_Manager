"""
User Administration GUI
=======================

Graphical User Interface for comprehensive user administration.
This interface provides full user management capabilities including user creation,
modification, suspension, password management, and privilege control.

WARNING: This GUI is for administrative purposes only and should NOT be 
accessible from the main application interface.

Author: Workflow Manager System
Version: 2.0.0
Date: 2025-10-31
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime, timedelta
from typing import Optional, List
import logging
from user_admin import UserAdministration, AdminUser, UserRole, UserStatus, PrivilegeLevel

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class UserAdminGUI:
    """Main User Administration GUI Application"""
    
    def __init__(self):
        """Initialize the User Administration GUI"""
        self.root = tk.Tk()
        self.root.title("User Administration System - ADMINISTRATIVE ACCESS ONLY")
        self.root.geometry("1400x900")
        
        # Set theme colors
        self.bg_color = "#f0f0f0"
        self.primary_color = "#2c3e50"
        self.secondary_color = "#3498db"
        self.success_color = "#27ae60"
        self.warning_color = "#f39c12"
        self.danger_color = "#e74c3c"
        
        self.root.configure(bg=self.bg_color)
        
        # Initialize user administration
        self.admin = UserAdministration()
        self.current_admin_user = None
        
        # Authentication required
        if not self.authenticate_admin():
            self.root.destroy()
            return
        
        # Initialize GUI components
        self.init_gui()
        
        # Load users
        self.refresh_user_list()
    
    def authenticate_admin(self) -> bool:
        """Authenticate the administrator before allowing access"""
        auth_window = tk.Toplevel(self.root)
        auth_window.title("Administrative Authentication")
        auth_window.geometry("400x250")
        auth_window.transient(self.root)
        auth_window.grab_set()
        
        # Center the window
        auth_window.update_idletasks()
        x = (auth_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (auth_window.winfo_screenheight() // 2) - (250 // 2)
        auth_window.geometry(f"400x250+{x}+{y}")
        
        authenticated = [False]
        
        # Warning label
        warning_frame = tk.Frame(auth_window, bg="#e74c3c")
        warning_frame.pack(fill=tk.X)
        
        warning_label = tk.Label(
            warning_frame,
            text="⚠ ADMINISTRATIVE ACCESS REQUIRED ⚠",
            font=("Arial", 12, "bold"),
            bg="#e74c3c",
            fg="white",
            pady=10
        )
        warning_label.pack()
        
        # Form frame
        form_frame = tk.Frame(auth_window, bg="white", padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(
            form_frame,
            text="Username:",
            font=("Arial", 10),
            bg="white"
        ).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        username_entry = tk.Entry(form_frame, font=("Arial", 10), width=30)
        username_entry.grid(row=0, column=1, pady=5)
        username_entry.focus()
        
        tk.Label(
            form_frame,
            text="Password:",
            font=("Arial", 10),
            bg="white"
        ).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        password_entry = tk.Entry(form_frame, font=("Arial", 10), width=30, show="*")
        password_entry.grid(row=1, column=1, pady=5)
        
        def try_authenticate():
            username = username_entry.get()
            password = password_entry.get()
            
            if not username or not password:
                messagebox.showerror("Error", "Please enter username and password")
                return
            
            # For initial setup, accept superadmin credentials
            # In production, you would verify against the database
            user = self.admin.get_user_by_username(username)
            
            if user and user.role in ['super_admin', 'admin']:
                # In a real implementation, you would verify the password
                # For now, we'll accept the user if they exist and are admin
                self.current_admin_user = user
                authenticated[0] = True
                auth_window.destroy()
            else:
                messagebox.showerror("Authentication Failed", 
                                   "Invalid credentials or insufficient privileges.\n"
                                   "Only administrators can access this system.")
        
        def cancel_auth():
            auth_window.destroy()
        
        # Buttons
        button_frame = tk.Frame(form_frame, bg="white")
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        login_btn = tk.Button(
            button_frame,
            text="Login",
            command=try_authenticate,
            bg=self.secondary_color,
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=5,
            cursor="hand2"
        )
        login_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            command=cancel_auth,
            bg=self.danger_color,
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5,
            cursor="hand2"
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key
        password_entry.bind('<Return>', lambda e: try_authenticate())
        
        self.root.wait_window(auth_window)
        
        return authenticated[0]
    
    def init_gui(self):
        """Initialize GUI components"""
        # Header
        self.create_header()
        
        # Main container with notebook (tabs)
        main_container = tk.Frame(self.root, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.create_user_list_tab()
        self.create_user_details_tab()
        self.create_statistics_tab()
        self.create_audit_log_tab()
        
        # Status bar
        self.create_status_bar()
    
    def create_header(self):
        """Create header with title and admin info"""
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🔐 User Administration System",
            font=("Arial", 18, "bold"),
            bg=self.primary_color,
            fg="white"
        )
        title_label.pack(side=tk.LEFT, padx=20)
        
        if self.current_admin_user:
            admin_info = tk.Label(
                header_frame,
                text=f"Admin: {self.current_admin_user.username} | Role: {self.current_admin_user.role}",
                font=("Arial", 10),
                bg=self.primary_color,
                fg="white"
            )
            admin_info.pack(side=tk.RIGHT, padx=20)
    
    def create_user_list_tab(self):
        """Create user list/management tab"""
        tab = tk.Frame(self.notebook, bg="white")
        self.notebook.add(tab, text="User Management")
        
        # Toolbar
        toolbar = tk.Frame(tab, bg="white", pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        # Search and filter
        search_frame = tk.Frame(toolbar, bg="white")
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Label(search_frame, text="Search:", bg="white").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_users())
        
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(search_frame, text="Status:", bg="white").pack(side=tk.LEFT, padx=5)
        self.status_filter_var = tk.StringVar(value="All")
        status_combo = ttk.Combobox(
            search_frame,
            textvariable=self.status_filter_var,
            values=["All", "active", "suspended", "locked", "pending_activation", "deactivated"],
            width=15,
            state="readonly"
        )
        status_combo.pack(side=tk.LEFT, padx=5)
        status_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_users())
        
        tk.Label(search_frame, text="Role:", bg="white").pack(side=tk.LEFT, padx=5)
        self.role_filter_var = tk.StringVar(value="All")
        role_combo = ttk.Combobox(
            search_frame,
            textvariable=self.role_filter_var,
            values=["All", "super_admin", "admin", "manager", "supervisor", "user", "guest"],
            width=15,
            state="readonly"
        )
        role_combo.pack(side=tk.LEFT, padx=5)
        role_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_users())
        
        # Action buttons
        button_frame = tk.Frame(toolbar, bg="white")
        button_frame.pack(side=tk.RIGHT)
        
        self.create_button(button_frame, "➕ New User", self.create_new_user, 
                          self.success_color).pack(side=tk.LEFT, padx=2)
        self.create_button(button_frame, "🔄 Refresh", self.refresh_user_list, 
                          self.secondary_color).pack(side=tk.LEFT, padx=2)
        
        # User list
        list_frame = tk.Frame(tab, bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create Treeview
        columns = ("ID", "Username", "Email", "Role", "Status", "Company", "Position", "Created Date")
        self.user_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=20)
        
        # Define column headings and widths
        column_widths = {"ID": 50, "Username": 120, "Email": 200, "Role": 100, 
                        "Status": 120, "Company": 150, "Position": 150, "Created Date": 150}
        
        for col in columns:
            self.user_tree.heading(col, text=col, command=lambda c=col: self.sort_tree(c))
            self.user_tree.column(col, width=column_widths.get(col, 100))
        
        # Scrollbars
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.user_tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.user_tree.xview)
        self.user_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.user_tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Context menu for user tree
        self.user_tree.bind('<Button-3>', self.show_user_context_menu)
        self.user_tree.bind('<Double-1>', self.view_user_details)
        
        # Action buttons for selected user
        action_frame = tk.Frame(tab, bg="white", pady=10)
        action_frame.pack(fill=tk.X, padx=10)
        
        self.create_button(action_frame, "👁 View Details", self.view_user_details,
                          self.secondary_color).pack(side=tk.LEFT, padx=2)
        self.create_button(action_frame, "✏ Edit User", self.edit_user,
                          self.warning_color).pack(side=tk.LEFT, padx=2)
        self.create_button(action_frame, "🔒 Change Password", self.change_user_password,
                          self.warning_color).pack(side=tk.LEFT, padx=2)
        self.create_button(action_frame, "⏸ Suspend User", self.suspend_user,
                          self.danger_color).pack(side=tk.LEFT, padx=2)
        self.create_button(action_frame, "▶ Activate User", self.activate_user,
                          self.success_color).pack(side=tk.LEFT, padx=2)
        self.create_button(action_frame, "🗑 Delete User", self.delete_user,
                          self.danger_color).pack(side=tk.LEFT, padx=2)
    
    def create_user_details_tab(self):
        """Create user details tab"""
        tab = tk.Frame(self.notebook, bg="white")
        self.notebook.add(tab, text="User Details")
        
        # Details will be shown here when a user is selected
        self.details_frame = tk.Frame(tab, bg="white")
        self.details_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        label = tk.Label(
            self.details_frame,
            text="Select a user from the User Management tab to view details",
            font=("Arial", 12),
            bg="white",
            fg="#666"
        )
        label.pack(expand=True)
    
    def create_statistics_tab(self):
        """Create statistics tab"""
        tab = tk.Frame(self.notebook, bg="white")
        self.notebook.add(tab, text="Statistics")
        
        self.stats_frame = tk.Frame(tab, bg="white")
        self.stats_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Refresh button
        refresh_btn = self.create_button(
            self.stats_frame,
            "🔄 Refresh Statistics",
            self.refresh_statistics,
            self.secondary_color
        )
        refresh_btn.pack(pady=10)
        
        # Statistics will be displayed here
        self.refresh_statistics()
    
    def create_audit_log_tab(self):
        """Create audit log tab"""
        tab = tk.Frame(self.notebook, bg="white")
        self.notebook.add(tab, text="Audit Log")
        
        # Toolbar
        toolbar = tk.Frame(tab, bg="white", pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        tk.Label(toolbar, text="Filter by User ID:", bg="white").pack(side=tk.LEFT, padx=5)
        self.audit_user_id_var = tk.StringVar()
        audit_user_entry = tk.Entry(toolbar, textvariable=self.audit_user_id_var, width=10)
        audit_user_entry.pack(side=tk.LEFT, padx=5)
        
        self.create_button(toolbar, "🔍 Filter", self.refresh_audit_log,
                          self.secondary_color).pack(side=tk.LEFT, padx=5)
        self.create_button(toolbar, "🔄 Show All", self.refresh_audit_log,
                          self.secondary_color).pack(side=tk.LEFT, padx=5)
        
        # Audit log list
        log_frame = tk.Frame(tab, bg="white")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("ID", "User", "Action", "Description", "Performed By", "IP Address", "Timestamp")
        self.audit_tree = ttk.Treeview(log_frame, columns=columns, show="headings", height=25)
        
        column_widths = {"ID": 50, "User": 120, "Action": 150, "Description": 250,
                        "Performed By": 120, "IP Address": 120, "Timestamp": 150}
        
        for col in columns:
            self.audit_tree.heading(col, text=col)
            self.audit_tree.column(col, width=column_widths.get(col, 100))
        
        vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.audit_tree.yview)
        hsb = ttk.Scrollbar(log_frame, orient="horizontal", command=self.audit_tree.xview)
        self.audit_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.audit_tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)
        
        # Load audit log
        self.refresh_audit_log()
    
    def create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = tk.Label(
            self.root,
            text="Ready",
            bg=self.primary_color,
            fg="white",
            anchor=tk.W,
            font=("Arial", 9),
            padx=10
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_button(self, parent, text, command, bg_color):
        """Create a styled button"""
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg_color,
            fg="white",
            font=("Arial", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2",
            relief=tk.FLAT
        )
        return button
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_bar.config(text=message)
        self.root.update_idletasks()
    
    def refresh_user_list(self):
        """Refresh the user list"""
        try:
            self.update_status("Loading users...")
            
            # Clear existing items
            for item in self.user_tree.get_children():
                self.user_tree.delete(item)
            
            # Get all users
            users = self.admin.get_all_users()
            
            for user in users:
                self.user_tree.insert("", tk.END, values=(
                    user.user_id,
                    user.username,
                    user.email,
                    user.role,
                    user.status,
                    user.company or "",
                    user.position or "",
                    user.created_date.strftime("%Y-%m-%d %H:%M") if user.created_date else ""
                ))
            
            self.update_status(f"Loaded {len(users)} users")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load users: {e}")
            logger.error(f"Error refreshing user list: {e}")
    
    def filter_users(self):
        """Filter users based on search and filter criteria"""
        search_term = self.search_var.get().lower()
        status_filter = self.status_filter_var.get()
        role_filter = self.role_filter_var.get()
        
        # Clear tree
        for item in self.user_tree.get_children():
            self.user_tree.delete(item)
        
        # Get filtered users
        try:
            users = self.admin.get_all_users()
            
            for user in users:
                # Apply filters
                if status_filter != "All" and user.status != status_filter:
                    continue
                
                if role_filter != "All" and user.role != role_filter:
                    continue
                
                if search_term:
                    if (search_term not in user.username.lower() and
                        search_term not in user.email.lower() and
                        (not user.company or search_term not in user.company.lower())):
                        continue
                
                self.user_tree.insert("", tk.END, values=(
                    user.user_id,
                    user.username,
                    user.email,
                    user.role,
                    user.status,
                    user.company or "",
                    user.position or "",
                    user.created_date.strftime("%Y-%m-%d %H:%M") if user.created_date else ""
                ))
            
        except Exception as e:
            logger.error(f"Error filtering users: {e}")
    
    def sort_tree(self, col):
        """Sort tree by column"""
        # This is a simple implementation
        items = [(self.user_tree.set(item, col), item) for item in self.user_tree.get_children('')]
        items.sort()
        
        for index, (val, item) in enumerate(items):
            self.user_tree.move(item, '', index)
    
    def get_selected_user_id(self) -> Optional[int]:
        """Get the ID of the selected user"""
        selection = self.user_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a user first")
            return None
        
        item = self.user_tree.item(selection[0])
        return int(item['values'][0])
    
    def create_new_user(self):
        """Create a new user"""
        dialog = UserCreateDialog(self.root, self.admin, self.current_admin_user.user_id)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.refresh_user_list()
            messagebox.showinfo("Success", f"User '{dialog.result}' created successfully")
    
    def view_user_details(self, event=None):
        """View detailed user information"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        # Clear details frame
        for widget in self.details_frame.winfo_children():
            widget.destroy()
        
        # Create details view
        canvas = tk.Canvas(self.details_frame, bg="white")
        scrollbar = ttk.Scrollbar(self.details_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # User details
        details_data = [
            ("User ID", user.user_id),
            ("Username", user.username),
            ("Email", user.email),
            ("Role", user.role),
            ("Privilege Level", user.privilege_level),
            ("Status", user.status),
            ("Company", user.company or "N/A"),
            ("Position", user.position or "N/A"),
            ("Department", user.department or "N/A"),
            ("Employee ID", user.employee_id or "N/A"),
            ("Phone Number", user.phone_number or "N/A"),
            ("Manager ID", user.manager_id or "N/A"),
            ("Created Date", user.created_date.strftime("%Y-%m-%d %H:%M:%S") if user.created_date else "N/A"),
            ("Last Login", user.last_login.strftime("%Y-%m-%d %H:%M:%S") if user.last_login else "Never"),
            ("Suspension Date", user.suspension_date.strftime("%Y-%m-%d %H:%M:%S") if user.suspension_date else "N/A"),
            ("Suspension Reason", user.suspension_reason or "N/A"),
        ]
        
        for i, (label, value) in enumerate(details_data):
            tk.Label(
                scrollable_frame,
                text=f"{label}:",
                font=("Arial", 10, "bold"),
                bg="white",
                anchor=tk.W
            ).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            
            tk.Label(
                scrollable_frame,
                text=str(value),
                font=("Arial", 10),
                bg="white",
                anchor=tk.W
            ).grid(row=i, column=1, sticky=tk.W, padx=10, pady=5)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Switch to details tab
        self.notebook.select(1)
    
    def edit_user(self):
        """Edit user information"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        dialog = UserEditDialog(self.root, self.admin, user, self.current_admin_user.user_id)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.refresh_user_list()
            messagebox.showinfo("Success", "User updated successfully")
    
    def change_user_password(self):
        """Change user password"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        dialog = PasswordChangeDialog(self.root, user.username)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            if self.admin.change_password(user_id, dialog.result, self.current_admin_user.user_id):
                messagebox.showinfo("Success", "Password changed successfully")
            else:
                messagebox.showerror("Error", "Failed to change password")
    
    def suspend_user(self):
        """Suspend user account"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        if user.status == 'suspended':
            messagebox.showinfo("Info", "User is already suspended")
            return
        
        reason = simpledialog.askstring(
            "Suspension Reason",
            f"Enter reason for suspending {user.username}:",
            parent=self.root
        )
        
        if reason:
            if self.admin.suspend_user(user_id, reason, self.current_admin_user.user_id):
                self.refresh_user_list()
                messagebox.showinfo("Success", f"User {user.username} has been suspended")
            else:
                messagebox.showerror("Error", "Failed to suspend user")
    
    def activate_user(self):
        """Activate user account"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        if user.status == 'active':
            messagebox.showinfo("Info", "User is already active")
            return
        
        if messagebox.askyesno("Confirm Activation", f"Activate user {user.username}?"):
            if self.admin.activate_user(user_id, self.current_admin_user.user_id):
                self.refresh_user_list()
                messagebox.showinfo("Success", f"User {user.username} has been activated")
            else:
                messagebox.showerror("Error", "Failed to activate user")
    
    def delete_user(self):
        """Delete user account"""
        user_id = self.get_selected_user_id()
        if not user_id:
            return
        
        user = self.admin.get_user_by_id(user_id)
        if not user:
            messagebox.showerror("Error", "User not found")
            return
        
        if user.user_id == self.current_admin_user.user_id:
            messagebox.showerror("Error", "You cannot delete your own account")
            return
        
        if not messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to DELETE user {user.username}?\n\n"
            "This action cannot be undone!\n"
            "Consider suspending or deactivating the user instead.",
            icon='warning'
        ):
            return
        
        # Require password confirmation for deletion
        password = simpledialog.askstring(
            "Confirm Deletion",
            "Enter your admin password to confirm deletion:",
            show='*',
            parent=self.root
        )
        
        if password:
            if self.admin.delete_user(user_id, self.current_admin_user.user_id):
                self.refresh_user_list()
                messagebox.showinfo("Success", f"User {user.username} has been deleted")
            else:
                messagebox.showerror("Error", "Failed to delete user")
    
    def show_user_context_menu(self, event):
        """Show context menu for user"""
        # Select the item under cursor
        item = self.user_tree.identify_row(event.y)
        if item:
            self.user_tree.selection_set(item)
            
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="View Details", command=self.view_user_details)
            menu.add_command(label="Edit User", command=self.edit_user)
            menu.add_separator()
            menu.add_command(label="Change Password", command=self.change_user_password)
            menu.add_separator()
            menu.add_command(label="Suspend User", command=self.suspend_user)
            menu.add_command(label="Activate User", command=self.activate_user)
            menu.add_separator()
            menu.add_command(label="Delete User", command=self.delete_user)
            
            menu.post(event.x_root, event.y_root)
    
    def refresh_statistics(self):
        """Refresh statistics display"""
        try:
            # Clear existing widgets
            for widget in self.stats_frame.winfo_children():
                if widget.winfo_class() != 'Button':
                    widget.destroy()
            
            stats = self.admin.get_user_statistics()
            
            # Create statistics cards
            cards_frame = tk.Frame(self.stats_frame, bg="white")
            cards_frame.pack(fill=tk.BOTH, expand=True, pady=20)
            
            # Total users card
            self.create_stat_card(
                cards_frame,
                "Total Users",
                str(stats.get('total_users', 0)),
                self.primary_color,
                0, 0
            )
            
            # Active sessions
            self.create_stat_card(
                cards_frame,
                "Active Sessions",
                str(stats.get('active_sessions', 0)),
                self.success_color,
                0, 1
            )
            
            # Recently created
            self.create_stat_card(
                cards_frame,
                "New Users (30 days)",
                str(stats.get('recently_created', 0)),
                self.secondary_color,
                0, 2
            )
            
            # Users by status
            status_frame = tk.LabelFrame(
                self.stats_frame,
                text="Users by Status",
                font=("Arial", 12, "bold"),
                bg="white",
                padx=20,
                pady=20
            )
            status_frame.pack(fill=tk.X, padx=20, pady=10)
            
            by_status = stats.get('by_status', {})
            for i, (status, count) in enumerate(by_status.items()):
                tk.Label(
                    status_frame,
                    text=f"{status}:",
                    font=("Arial", 10, "bold"),
                    bg="white"
                ).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
                
                tk.Label(
                    status_frame,
                    text=str(count),
                    font=("Arial", 10),
                    bg="white"
                ).grid(row=i, column=1, sticky=tk.W, padx=10, pady=5)
            
            # Users by role
            role_frame = tk.LabelFrame(
                self.stats_frame,
                text="Users by Role",
                font=("Arial", 12, "bold"),
                bg="white",
                padx=20,
                pady=20
            )
            role_frame.pack(fill=tk.X, padx=20, pady=10)
            
            by_role = stats.get('by_role', {})
            for i, (role, count) in enumerate(by_role.items()):
                tk.Label(
                    role_frame,
                    text=f"{role}:",
                    font=("Arial", 10, "bold"),
                    bg="white"
                ).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
                
                tk.Label(
                    role_frame,
                    text=str(count),
                    font=("Arial", 10),
                    bg="white"
                ).grid(row=i, column=1, sticky=tk.W, padx=10, pady=5)
            
            self.update_status("Statistics updated")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load statistics: {e}")
            logger.error(f"Error refreshing statistics: {e}")
    
    def create_stat_card(self, parent, title, value, color, row, col):
        """Create a statistics card"""
        card = tk.Frame(parent, bg=color, padx=20, pady=20)
        card.grid(row=row, column=col, padx=10, pady=10, sticky=tk.NSEW)
        
        tk.Label(
            card,
            text=title,
            font=("Arial", 10),
            bg=color,
            fg="white"
        ).pack()
        
        tk.Label(
            card,
            text=value,
            font=("Arial", 24, "bold"),
            bg=color,
            fg="white"
        ).pack()
        
        parent.grid_columnconfigure(col, weight=1)
    
    def refresh_audit_log(self):
        """Refresh audit log"""
        try:
            # Clear existing items
            for item in self.audit_tree.get_children():
                self.audit_tree.delete(item)
            
            # Get user ID filter if specified
            user_id = None
            user_id_str = self.audit_user_id_var.get().strip()
            if user_id_str:
                try:
                    user_id = int(user_id_str)
                except ValueError:
                    pass
            
            # Get audit log
            logs = self.admin.get_audit_log(user_id=user_id, limit=500)
            
            for log in logs:
                self.audit_tree.insert("", tk.END, values=(
                    log['id'],
                    log['username'],
                    log['action_type'],
                    log['description'],
                    log['performed_by_name'] or "System",
                    log['ip_address'] or "N/A",
                    log['timestamp']
                ))
            
            self.update_status(f"Loaded {len(logs)} audit log entries")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load audit log: {e}")
            logger.error(f"Error refreshing audit log: {e}")
    
    def run(self):
        """Run the GUI application"""
        self.root.mainloop()


class UserCreateDialog:
    """Dialog for creating a new user"""
    
    def __init__(self, parent, admin: UserAdministration, created_by: int):
        self.admin = admin
        self.created_by = created_by
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Create New User")
        self.dialog.geometry("500x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Create form
        form = tk.Frame(self.dialog, bg="white", padx=20, pady=20)
        form.pack(fill=tk.BOTH, expand=True)
        
        fields = [
            ("Username*", "username"),
            ("Email*", "email"),
            ("Password*", "password"),
            ("Confirm Password*", "confirm_password"),
            ("Company", "company"),
            ("Position", "position"),
            ("Department", "department"),
            ("Employee ID", "employee_id"),
            ("Phone Number", "phone_number"),
        ]
        
        self.entries = {}
        row = 0
        
        for label_text, field_name in fields:
            tk.Label(
                form,
                text=label_text,
                font=("Arial", 10),
                bg="white"
            ).grid(row=row, column=0, sticky=tk.W, pady=5)
            
            if field_name in ["password", "confirm_password"]:
                entry = tk.Entry(form, font=("Arial", 10), width=30, show="*")
            else:
                entry = tk.Entry(form, font=("Arial", 10), width=30)
            
            entry.grid(row=row, column=1, pady=5, sticky=tk.W)
            self.entries[field_name] = entry
            row += 1
        
        # Role dropdown
        tk.Label(form, text="Role*", font=("Arial", 10), bg="white").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.role_var = tk.StringVar(value="user")
        role_combo = ttk.Combobox(
            form,
            textvariable=self.role_var,
            values=["super_admin", "admin", "manager", "supervisor", "user", "guest"],
            width=28,
            state="readonly"
        )
        role_combo.grid(row=row, column=1, pady=5, sticky=tk.W)
        row += 1
        
        # Privilege level dropdown
        tk.Label(form, text="Privilege Level*", font=("Arial", 10), bg="white").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.privilege_var = tk.StringVar(value="limited")
        privilege_combo = ttk.Combobox(
            form,
            textvariable=self.privilege_var,
            values=["full_access", "read_write", "read_only", "limited", "none"],
            width=28,
            state="readonly"
        )
        privilege_combo.grid(row=row, column=1, pady=5, sticky=tk.W)
        row += 1
        
        # Auto-activate checkbox
        self.auto_activate_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            form,
            text="Activate user immediately",
            variable=self.auto_activate_var,
            bg="white",
            font=("Arial", 10)
        ).grid(row=row, column=0, columnspan=2, pady=10, sticky=tk.W)
        row += 1
        
        # Buttons
        button_frame = tk.Frame(form, bg="white")
        button_frame.grid(row=row, column=0, columnspan=2, pady=20)
        
        tk.Button(
            button_frame,
            text="Create User",
            command=self.create_user,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Cancel",
            command=self.dialog.destroy,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
    
    def create_user(self):
        """Create the user"""
        try:
            username = self.entries['username'].get().strip()
            email = self.entries['email'].get().strip()
            password = self.entries['password'].get()
            confirm_password = self.entries['confirm_password'].get()
            
            if not username or not email or not password:
                messagebox.showerror("Error", "Please fill in all required fields")
                return
            
            if password != confirm_password:
                messagebox.showerror("Error", "Passwords do not match")
                return
            
            user_id = self.admin.create_user(
                username=username,
                email=email,
                password=password,
                role=self.role_var.get(),
                privilege_level=self.privilege_var.get(),
                company=self.entries['company'].get().strip() or None,
                position=self.entries['position'].get().strip() or None,
                department=self.entries['department'].get().strip() or None,
                employee_id=self.entries['employee_id'].get().strip() or None,
                phone_number=self.entries['phone_number'].get().strip() or None,
                created_by=self.created_by,
                auto_activate=self.auto_activate_var.get()
            )
            
            if user_id:
                self.result = username
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to create user. Check logs for details.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create user: {e}")


class UserEditDialog:
    """Dialog for editing user information"""
    
    def __init__(self, parent, admin: UserAdministration, user: AdminUser, modified_by: int):
        self.admin = admin
        self.user = user
        self.modified_by = modified_by
        self.result = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Edit User: {user.username}")
        self.dialog.geometry("500x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Create form
        form = tk.Frame(self.dialog, bg="white", padx=20, pady=20)
        form.pack(fill=tk.BOTH, expand=True)
        
        fields = [
            ("Username", "username", user.username),
            ("Email", "email", user.email),
            ("Company", "company", user.company or ""),
            ("Position", "position", user.position or ""),
            ("Department", "department", user.department or ""),
            ("Employee ID", "employee_id", user.employee_id or ""),
            ("Phone Number", "phone_number", user.phone_number or ""),
        ]
        
        self.entries = {}
        row = 0
        
        for label_text, field_name, value in fields:
            tk.Label(
                form,
                text=label_text,
                font=("Arial", 10),
                bg="white"
            ).grid(row=row, column=0, sticky=tk.W, pady=5)
            
            entry = tk.Entry(form, font=("Arial", 10), width=30)
            entry.insert(0, value)
            entry.grid(row=row, column=1, pady=5, sticky=tk.W)
            self.entries[field_name] = entry
            row += 1
        
        # Role dropdown
        tk.Label(form, text="Role", font=("Arial", 10), bg="white").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.role_var = tk.StringVar(value=user.role)
        role_combo = ttk.Combobox(
            form,
            textvariable=self.role_var,
            values=["super_admin", "admin", "manager", "supervisor", "user", "guest"],
            width=28,
            state="readonly"
        )
        role_combo.grid(row=row, column=1, pady=5, sticky=tk.W)
        row += 1
        
        # Privilege level dropdown
        tk.Label(form, text="Privilege Level", font=("Arial", 10), bg="white").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.privilege_var = tk.StringVar(value=user.privilege_level)
        privilege_combo = ttk.Combobox(
            form,
            textvariable=self.privilege_var,
            values=["full_access", "read_write", "read_only", "limited", "none"],
            width=28,
            state="readonly"
        )
        privilege_combo.grid(row=row, column=1, pady=5, sticky=tk.W)
        row += 1
        
        # Status dropdown
        tk.Label(form, text="Status", font=("Arial", 10), bg="white").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.status_var = tk.StringVar(value=user.status)
        status_combo = ttk.Combobox(
            form,
            textvariable=self.status_var,
            values=["active", "suspended", "locked", "pending_activation", "deactivated"],
            width=28,
            state="readonly"
        )
        status_combo.grid(row=row, column=1, pady=5, sticky=tk.W)
        row += 1
        
        # Buttons
        button_frame = tk.Frame(form, bg="white")
        button_frame.grid(row=row, column=0, columnspan=2, pady=20)
        
        tk.Button(
            button_frame,
            text="Save Changes",
            command=self.save_changes,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Cancel",
            command=self.dialog.destroy,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
    
    def save_changes(self):
        """Save user changes"""
        try:
            updates = {
                'username': self.entries['username'].get().strip(),
                'email': self.entries['email'].get().strip(),
                'role': self.role_var.get(),
                'privilege_level': self.privilege_var.get(),
                'status': self.status_var.get(),
                'company': self.entries['company'].get().strip() or None,
                'position': self.entries['position'].get().strip() or None,
                'department': self.entries['department'].get().strip() or None,
                'employee_id': self.entries['employee_id'].get().strip() or None,
                'phone_number': self.entries['phone_number'].get().strip() or None,
            }
            
            if self.admin.update_user(self.user.user_id, self.modified_by, **updates):
                self.result = True
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to update user")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update user: {e}")


class PasswordChangeDialog:
    """Dialog for changing user password"""
    
    def __init__(self, parent, username: str):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Change Password: {username}")
        self.dialog.geometry("400x200")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        form = tk.Frame(self.dialog, bg="white", padx=20, pady=20)
        form.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(
            form,
            text=f"Change password for: {username}",
            font=("Arial", 12, "bold"),
            bg="white"
        ).pack(pady=10)
        
        tk.Label(form, text="New Password:", bg="white").pack(anchor=tk.W)
        self.password_entry = tk.Entry(form, show="*", width=30)
        self.password_entry.pack(pady=5)
        
        tk.Label(form, text="Confirm Password:", bg="white").pack(anchor=tk.W)
        self.confirm_entry = tk.Entry(form, show="*", width=30)
        self.confirm_entry.pack(pady=5)
        
        button_frame = tk.Frame(form, bg="white")
        button_frame.pack(pady=20)
        
        tk.Button(
            button_frame,
            text="Change Password",
            command=self.change_password,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Cancel",
            command=self.dialog.destroy,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
    
    def change_password(self):
        """Validate and save password"""
        password = self.password_entry.get()
        confirm = self.confirm_entry.get()
        
        if not password:
            messagebox.showerror("Error", "Password cannot be empty")
            return
        
        if password != confirm:
            messagebox.showerror("Error", "Passwords do not match")
            return
        
        self.result = password
        self.dialog.destroy()


if __name__ == "__main__":
    """Run the User Administration GUI"""
    try:
        app = UserAdminGUI()
        app.run()
    except Exception as e:
        logger.error(f"Failed to start User Administration GUI: {e}")
        print(f"Error: {e}")

