#!/usr/bin/env python3
"""
Degrow Workflow Manager - Main Application
==========================================

This is the main entry point for the Degrow Workflow Manager application.
It handles user authentication and launches the appropriate interface.

Author: Workflow Manager System
Version: 1.0.0
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from question_sheet_gui import QuestionSheetGUI
from question_sheet_console import QuestionSheetConsole
from user_manager import UserManager, create_default_admin_user
from login_gui import show_login_window


def main():
    """Main application entry point"""
    try:
        # Check if running in console mode
        if len(sys.argv) > 1 and sys.argv[1] == '--console':
            run_console_mode()
        else:
            run_gui_mode()
            
    except KeyboardInterrupt:
        print("\nApplication interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"Application error: {e}")
        sys.exit(1)


def run_gui_mode():
    """Run the GUI application"""
    try:
        # Initialize user manager and create default admin if needed
        user_manager = UserManager()
        create_default_admin_user(user_manager)
        
        # Show login window
        user_manager = show_login_window()
        if not user_manager:
            print("Login cancelled. Exiting application.")
            return
        
        # Launch main GUI application
        app = QuestionSheetGUI(user_manager=user_manager)
        app.run()
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start application: {e}")
        sys.exit(1)


def run_console_mode():
    """Run the console application"""
    try:
        # Initialize user manager
        user_manager = UserManager()
        create_default_admin_user(user_manager)
        
        # For console mode, we'll use a simple authentication
        print("Degrow Workflow Manager - Console Mode")
        print("=" * 40)
        
        username = input("Username: ")
        password = input("Password: ")
        
        user = user_manager.authenticate_user(username, password)
        if not user:
            print("Authentication failed. Exiting.")
            return
        
        print(f"Welcome, {user.username}!")
        
        # Launch console application
        console_app = QuestionSheetConsole()
        console_app.run()
        
    except Exception as e:
        print(f"Console application error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
