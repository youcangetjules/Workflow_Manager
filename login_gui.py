"""
Login GUI Module
================

This module provides the login interface and user authentication GUI
for the Degrow Workflow Manager system.

Author: Workflow Manager System
Version: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess
from datetime import datetime
from typing import Optional, Callable
from user_manager import UserManager, create_default_admin_user
from db_config import get_database_connection
from svg_banner_renderer import create_banner_from_svg
from background_renderer import BackgroundRenderer


class LoginWindow:
    """Login window for user authentication"""
    
    def __init__(self, parent=None, on_success: Callable = None):
        """
        Initialize login window
        
        Args:
            parent: Parent window
            on_success: Callback function when login is successful
        """
        self.parent = parent
        self.on_success = on_success
        self.user_manager = UserManager()
        self.background_photo = None  # Store background reference
        
        # Create login window - Full screen like the mockup
        self.login_window = tk.Toplevel(parent) if parent else tk.Tk()
        self.login_window.title("Workflow Manager Service Login")
        
        # Get screen dimensions and cap window at 1920x1080
        screen_width = self.login_window.winfo_screenwidth()
        screen_height = self.login_window.winfo_screenheight()
        target_width = min(1920, screen_width)
        target_height = min(1080, screen_height)
        
        # Regenerate composite background on launch
        try:
            base_dir = os.path.dirname(__file__)
            combined_path = os.path.join(base_dir, "combined.png")
            overlay_script = os.path.join(base_dir, "backgroundoverlay.py")
            
            # Delete old combined background if it exists
            if os.path.exists(combined_path):
                try:
                    os.remove(combined_path)
                except Exception as rm_err:
                    print(f"Warning: Could not remove existing combined.png: {rm_err}")
            
            # Run overlay generator (blocking)
            if os.path.exists(overlay_script):
                subprocess.run([sys.executable, overlay_script], check=True)
            else:
                print(f"Warning: overlay script not found: {overlay_script}")
        except Exception as gen_err:
            print(f"Error generating combined background: {gen_err}")

        bg_renderer = BackgroundRenderer(r"C:\Lumen\Workflow Manager\combined.png")
        bg_renderer.apply_background_to_window(self.login_window, target_width, target_height)
        # Center window at capped size
        x = max(0, (screen_width - target_width) // 2)
        y = max(0, (screen_height - target_height) // 2)
        self.login_window.geometry(f"{target_width}x{target_height}+{x}+{y}")
        # Cap maximum size
        self.login_window.maxsize(1920, 1080)
        self.login_window.resizable(True, True)
        
        # Configure window
        self.login_window.configure(bg='#1a1a1a')  # Dark background
        
        # Center window
        self.login_window.transient(parent) if parent else None
        self.login_window.grab_set()
        
        # Make window modal
        if parent:
            self.login_window.transient(parent)
            self.login_window.grab_set()
        
        # Initialize GUI
        self.create_widgets()
        # Create and update resolution box in lower-right
        self.create_resolution_box()
        self.login_window.bind('<Configure>', lambda e: self.update_resolution_box())
        # Create DB status indicator bottom-left
        self.create_db_status_indicator()
        self.update_db_status_indicator()
        
        # Create default admin if no users exist
        create_default_admin_user(self.user_manager)
        
        # Focus on username entry
        self.username_entry.focus()
        
        # Bind Enter key to login
        self.login_window.bind('<Return>', lambda e: self.login())
    
    def create_gradient_background(self, canvas):
        """Create a gradient background on the canvas"""
        # Get canvas dimensions
        canvas.update_idletasks()
        width = canvas.winfo_width()
        height = canvas.winfo_height()
        
        if width <= 1 or height <= 1:
            # Canvas not ready yet, schedule for later
            canvas.after(100, lambda: self.create_gradient_background(canvas))
            return
        
        # Clear canvas first
        canvas.delete("all")
        
        # Create a simple gradient background
        canvas.configure(bg='#1a1a1a')
        
        # Create radial gradient effect
        center_x, center_y = width // 2, height // 2
        max_radius = max(width, height) // 2
        
        # Create gradient from center to edges
        for radius in range(max_radius, 0, -20):
            # Calculate color based on radius
            intensity = int(255 * (1 - radius / max_radius))
            # Create blue-gray gradient similar to the SVG
            r = max(0, min(255, intensity // 3))
            g = max(0, min(255, intensity // 2))
            b = max(0, min(255, intensity // 2))
            color = f"#{r:02x}{g:02x}{b:02x}"
            
            # Draw circle
            x1 = center_x - radius
            y1 = center_y - radius
            x2 = center_x + radius
            y2 = center_y + radius
            
            canvas.create_oval(x1, y1, x2, y2, fill=color, outline="")
        
        print(f"Background created: {width}x{height}")
    
    def create_simple_gradient(self, canvas):
        """Create a simple gradient background on canvas"""
        def draw_gradient():
            canvas.update_idletasks()
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            
            if width <= 1 or height <= 1:
                canvas.after(100, draw_gradient)
                return
            
            # Clear canvas
            canvas.delete("all")
            
            # Create radial gradient effect
            center_x, center_y = width // 2, height // 2
            max_radius = max(width, height) // 2
            
            # Create gradient from center to edges
            for radius in range(max_radius, 0, -20):
                # Calculate color based on radius
                intensity = int(255 * (1 - radius / max_radius))
                # Create more visible blue-gray gradient
                r = max(0, min(255, intensity // 2))
                g = max(0, min(255, intensity // 2))
                b = max(0, min(255, intensity))
                color = f"#{r:02x}{g:02x}{b:02x}"
                
                # Draw circle
                x1 = center_x - radius
                y1 = center_y - radius
                x2 = center_x + radius
                y2 = center_y + radius
                
                canvas.create_oval(x1, y1, x2, y2, fill=color, outline="")
        
        # Start drawing after canvas is ready
        canvas.after(100, draw_gradient)
    
    def create_widgets(self):
        """Create login window widgets matching the JPEG design"""
        # Set window background to dark color
        self.login_window.configure(bg='#1a1a1a')
        
        # Apply PNG/SVG background to the window
        try:
            screen_width = self.login_window.winfo_screenwidth()
            screen_height = self.login_window.winfo_screenheight()
            # Use the pre-composited background
            self._bg_renderer = BackgroundRenderer("combined.png")
            self._bg_label = self._bg_renderer.apply_background_to_window(self.login_window, screen_width, screen_height)
        except Exception as e:
            print(f"Failed to apply background image, using solid background: {e}")

        # Main container sized to content (do not cover entire window)
        main_container = tk.Frame(self.login_window, bg='#1a1a1a', highlightthickness=0)
        main_container.pack()  # no fill/expand so background image remains visible
        
        # Top banner removed per request to avoid horizontal bars
        
        # No separate mid banner label; combined.png already includes it
        
        # Login inputs as individual items (no container pane)
        self.username_var = tk.StringVar()
        self.username_entry = tk.Entry(self.login_window,
                                      textvariable=self.username_var,
                                      font=('Arial', 12),
                                      width=18,
                                      relief='solid',
                                      bd=1,
                                      bg='#f5f5f5',
                                      fg='#000000')
        self.username_entry.place(relx=0.43, rely=0.58, anchor='center')

        self.password_var = tk.StringVar()
        self.password_entry = tk.Entry(self.login_window,
                                      textvariable=self.password_var,
                                      show="*",
                                      font=('Arial', 12),
                                      width=18,
                                      relief='solid',
                                      bd=1,
                                      bg='#f5f5f5',
                                      fg='#000000')
        self.password_entry.place(relx=0.57, rely=0.58, anchor='center')

        # Image button using greenbutton.png
        try:
            from PIL import Image, ImageTk
            btn_img = Image.open(r"C:\Lumen\Workflow Manager\greenbutton.png").convert('RGBA')
            # Optional: scale button to a sensible height
            target_h = 40
            scale = target_h / btn_img.height if btn_img.height else 1
            btn_img = btn_img.resize((max(1, int(btn_img.width * scale)), target_h), Image.LANCZOS)
            self._login_btn_photo = ImageTk.PhotoImage(btn_img)
            self.login_button = tk.Button(self.login_window,
                                          image=self._login_btn_photo,
                                          command=self.login,
                                          borderwidth=0,
                                          highlightthickness=0,
                                          bg='#1a1a1a',
                                          activebackground='#1a1a1a')
        except Exception:
            # Fallback to text button if image missing
            self.login_button = tk.Button(self.login_window,
                                          text="Login",
                                          command=self.login,
                                          font=('Arial', 12, 'bold'),
                                          fg='white',
                                          bg='#00aa00',
                                          relief='raised',
                                          bd=2)
        self.login_button.place(relx=0.68, rely=0.58, anchor='center')

        # Status label for login feedback (small, no pane)
        self.status_label = tk.Label(self.login_window,
                                     text="",
                                     font=('Arial', 10),
                                     fg='red',
                                     bg='#1a1a1a')
        self.status_label.place(relx=0.5, rely=0.62, anchor='n')
    
    def create_resolution_box(self):
        """Create a small box in the lower-right showing detected resolution."""
        # Container frame overlayed on the window (no fill, fixed corner placement)
        self.resolution_box = tk.Frame(self.login_window,
                                       bg='#2a2a2a',
                                       relief='ridge',
                                       bd=1,
                                       highlightthickness=0)
        # Content
        self.resolution_title = tk.Label(self.resolution_box,
                                         text="Resolution:",
                                         font=('Arial', 8, 'bold'),
                                         fg='#cccccc',
                                         bg='#2a2a2a')
        self.resolution_title.pack(anchor='w', padx=8, pady=(6, 0))

        self.resolution_label = tk.Label(self.resolution_box,
                                         text="",
                                         font=('Arial', 9),
                                         fg='white',
                                         bg='#2a2a2a')
        self.resolution_label.pack(anchor='w', padx=8, pady=(0, 6))

        # Place near bottom-right with small margin
        self.resolution_box.place(relx=1.0, rely=1.0, x=-16, y=-16, anchor='se')

    def create_db_status_indicator(self):
        """Create a small bottom-left DB connection status label."""
        self.db_status_label = tk.Label(self.login_window,
                                        text="SQL: Checking...",
                                        font=('Arial', 9),
                                        fg='#cccccc',
                                        bg='#1a1a1a')
        self.db_status_label.place(relx=0.0, rely=1.0, x=10, y=-10, anchor='sw')

    def update_db_status_indicator(self):
        """Check DB connectivity and update the status label color/text."""
        status_text = "SQL: Not connected"
        status_color = 'red'
        try:
            db = get_database_connection()
            conn = db.connect()
            if conn:
                status_text = "SQL: Connected"
                status_color = '#00aa00'
            # Close if possible
            try:
                db.close()
            except Exception:
                pass
        except Exception:
            pass
        self.db_status_label.config(text=status_text, fg=status_color)

    def update_resolution_box(self):
        """Update resolution text based on current window size."""
        try:
            width = self.login_window.winfo_width()
            height = self.login_window.winfo_height()
            # Fallback to screen size if window size not yet computed
            if width <= 1 or height <= 1:
                width = self.login_window.winfo_screenwidth()
                height = self.login_window.winfo_screenheight()
            self.resolution_label.config(text=f"{width}x{height}")
        except Exception:
            pass
    
    def login(self):
        """Handle login attempt"""
        username = self.username_var.get().strip()
        password = self.password_var.get()
        
        if not username or not password:
            self.status_label.config(text="Please enter both username and password", fg="red")
            return
        
        # Clear status
        self.status_label.config(text="")
        
        # Disable login button during authentication
        self.login_button.config(state="disabled", bg='#666666')
        self.login_window.update()
        
        try:
            # Authenticate user
            user = self.user_manager.authenticate_user(username, password)
            
            if user:
                # Login successful
                self.status_label.config(text="Login successful!", fg="green")
                self.login_window.update()
                
                # Wait a moment then close
                self.login_window.after(1000, self.login_successful)
            else:
                # Login failed
                self.status_label.config(text="Invalid username or password", fg="red")
                self.login_button.config(state="normal", bg='#0c9ed9')
                
        except Exception as e:
            self.status_label.config(text=f"Login error: {str(e)}", fg="red")
            self.login_button.config(state="normal", bg='#0c9ed9')
    
    def login_successful(self):
        """Handle successful login"""
        if self.on_success:
            self.on_success(self.user_manager)
        
        self.login_window.destroy()
    
    def cancel(self):
        """Handle cancel button"""
        self.login_window.destroy()
    
    def show_help(self):
        """Show help dialog"""
        help_text = """
Login Help
==========

Username: Enter your username or email address
Password: Enter your password

Default Admin Account:
- Username: admin
- Password: admin123

If you're having trouble logging in:
1. Check your username and password
2. Contact your system administrator
3. Ensure your account is active

For security:
- Change the default password after first login
- Use strong passwords
- Log out when finished
        """
        
        help_window = tk.Toplevel(self.login_window)
        help_window.title("Login Help")
        help_window.geometry("500x400")
        help_window.configure(bg='#1a1a1a')
        help_window.transient(self.login_window)
        help_window.grab_set()
        
        # Center the help window
        help_window.geometry("+%d+%d" % (
            self.login_window.winfo_rootx() + 50,
            self.login_window.winfo_rooty() + 50
        ))
        
        # Main frame
        main_frame = tk.Frame(help_window, bg='#1a1a1a')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(main_frame, 
                              text="Login Help",
                              font=('Segoe UI', 16, 'bold'),
                              fg='white',
                              bg='#1a1a1a')
        title_label.pack(pady=(0, 20))
        
        # Text widget with custom styling
        text_frame = tk.Frame(main_frame, bg='#2a2a2a', relief='sunken', bd=1)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        text_widget = tk.Text(text_frame, 
                             wrap=tk.WORD, 
                             padx=15, 
                             pady=15,
                             font=('Segoe UI', 10),
                             fg='white',
                             bg='#2a2a2a',
                             relief='flat',
                             bd=0)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
        
        # Close button
        close_button = tk.Button(main_frame, 
                                text="Close", 
                                command=help_window.destroy,
                                font=('Segoe UI', 10, 'bold'),
                                fg='white',
                                bg='#0c9ed9',
                                relief='raised',
                                bd=2,
                                width=15,
                                height=2)
        close_button.pack(pady=(10, 0))


class UserManagementWindow:
    """User management window for administrators"""
    
    def __init__(self, parent, user_manager: UserManager):
        """
        Initialize user management window
        
        Args:
            parent: Parent window
            user_manager: UserManager instance
        """
        self.parent = parent
        self.user_manager = user_manager
        
        # Create window
        self.window = tk.Toplevel(parent)
        self.window.title("User Management")
        self.window.geometry("800x600")
        self.window.transient(parent)
        self.window.grab_set()
        
        self.create_widgets()
        self.load_users()
    
    def create_widgets(self):
        """Create user management widgets"""
        # Main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="User Management", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(buttons_frame, text="Add User", 
                  command=self.add_user).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Edit User", 
                  command=self.edit_user).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Deactivate User", 
                  command=self.deactivate_user).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Refresh", 
                  command=self.load_users).pack(side=tk.LEFT, padx=(0, 5))
        
        # Users treeview
        columns = ('ID', 'Username', 'Email', 'Role', 'Active', 'Last Login')
        self.users_tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        for col in columns:
            self.users_tree.heading(col, text=col)
            self.users_tree.column(col, width=120)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.users_tree.yview)
        self.users_tree.configure(yscrollcommand=scrollbar.set)
        
        self.users_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Activity log frame
        log_frame = ttk.LabelFrame(main_frame, text="Recent Activity", padding="5")
        log_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.activity_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        activity_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", 
                                         command=self.activity_text.yview)
        self.activity_text.configure(yscrollcommand=activity_scrollbar.set)
        
        self.activity_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        activity_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Close button
        ttk.Button(main_frame, text="Close", 
                  command=self.window.destroy).pack(pady=(10, 0))
    
    def load_users(self):
        """Load users into treeview"""
        # Clear existing items
        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        
        # Load users
        users = self.user_manager.get_all_users()
        for user in users:
            last_login = user.last_login.strftime('%Y-%m-%d %H:%M') if user.last_login else 'Never'
            self.users_tree.insert('', 'end', values=(
                user.user_id,
                user.username,
                user.email,
                user.role,
                'Yes' if user.is_active else 'No',
                last_login
            ))
        
        # Load activity log
        self.load_activity_log()
    
    def load_activity_log(self):
        """Load recent activity log"""
        self.activity_text.delete(1.0, tk.END)
        
        activities = self.user_manager.get_user_activity_log(limit=20)
        for activity in activities:
            timestamp = activity['timestamp'].strftime('%Y-%m-%d %H:%M:%S')
            self.activity_text.insert(tk.END, 
                f"{timestamp} - {activity['username']}: {activity['activity_type']}\n")
    
    def add_user(self):
        """Add new user dialog"""
        self.user_dialog("Add User")
    
    def edit_user(self):
        """Edit selected user"""
        selection = self.users_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a user to edit")
            return
        
        item = self.users_tree.item(selection[0])
        user_id = item['values'][0]
        
        # Get user details
        users = self.user_manager.get_all_users()
        user = next((u for u in users if u.user_id == user_id), None)
        
        if user:
            self.user_dialog("Edit User", user)
    
    def user_dialog(self, title: str, user: Optional = None):
        """User dialog for adding/editing users"""
        dialog = tk.Toplevel(self.window)
        dialog.title(title)
        dialog.geometry("400x300")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Username
        ttk.Label(form_frame, text="Username:").pack(anchor=tk.W, pady=(0, 5))
        username_var = tk.StringVar(value=user.username if user else "")
        username_entry = ttk.Entry(form_frame, textvariable=username_var, width=30)
        username_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Email
        ttk.Label(form_frame, text="Email:").pack(anchor=tk.W, pady=(0, 5))
        email_var = tk.StringVar(value=user.email if user else "")
        email_entry = ttk.Entry(form_frame, textvariable=email_var, width=30)
        email_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Role
        ttk.Label(form_frame, text="Role:").pack(anchor=tk.W, pady=(0, 5))
        role_var = tk.StringVar(value=user.role if user else "user")
        role_combo = ttk.Combobox(form_frame, textvariable=role_var, 
                                 values=["user", "manager", "admin"], state="readonly")
        role_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Password (only for new users)
        if not user:
            ttk.Label(form_frame, text="Password:").pack(anchor=tk.W, pady=(0, 5))
            password_var = tk.StringVar()
            password_entry = ttk.Entry(form_frame, textvariable=password_var, 
                                     show="*", width=30)
            password_entry.pack(fill=tk.X, pady=(0, 10))
        else:
            password_var = None
        
        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_user():
            username = username_var.get().strip()
            email = email_var.get().strip()
            role = role_var.get()
            
            if not username or not email:
                messagebox.showerror("Error", "Username and email are required")
                return
            
            if user:  # Editing existing user
                # Update user role
                if self.user_manager.update_user_role(user.user_id, role):
                    messagebox.showinfo("Success", "User updated successfully")
                    dialog.destroy()
                    self.load_users()
                else:
                    messagebox.showerror("Error", "Failed to update user")
            else:  # Adding new user
                password = password_var.get() if password_var else ""
                if not password:
                    messagebox.showerror("Error", "Password is required for new users")
                    return
                
                if self.user_manager.create_user(username, email, password, role):
                    messagebox.showinfo("Success", "User created successfully")
                    dialog.destroy()
                    self.load_users()
                else:
                    messagebox.showerror("Error", "Failed to create user")
        
        ttk.Button(button_frame, text="Save", command=save_user).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT)
    
    def deactivate_user(self):
        """Deactivate selected user"""
        selection = self.users_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a user to deactivate")
            return
        
        item = self.users_tree.item(selection[0])
        user_id = item['values'][0]
        username = item['values'][1]
        
        if messagebox.askyesno("Confirm", f"Deactivate user '{username}'?"):
            if self.user_manager.deactivate_user(user_id):
                messagebox.showinfo("Success", "User deactivated successfully")
                self.load_users()
            else:
                messagebox.showerror("Error", "Failed to deactivate user")


def show_login_window(parent=None, on_success: Callable = None) -> Optional[UserManager]:
    """
    Show login window and return UserManager on success
    
    Args:
        parent: Parent window
        on_success: Callback function when login is successful
        
    Returns:
        UserManager instance if login successful, None otherwise
    """
    user_manager = None
    
    def on_login_success(manager):
        nonlocal user_manager
        user_manager = manager
    
    login_window = LoginWindow(parent, on_login_success)
    login_window.login_window.wait_window()
    
    return user_manager


if __name__ == "__main__":
    # Test login window
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    user_manager = show_login_window()
    if user_manager:
        print(f"Logged in as: {user_manager.get_current_user().username}")
        # Show user management window
        UserManagementWindow(root, user_manager)
        root.mainloop()
    else:
        print("Login cancelled")
        root.destroy()
