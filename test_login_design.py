#!/usr/bin/env python3
"""
Test script to demonstrate the new login page design
"""

import tkinter as tk
from login_gui import LoginWindow

def main():
    """Test the new login design"""
    print("Starting Workflow Manager Login Test...")
    print("The login window should now display with the new design matching the mockup.")
    print("Features:")
    print("- Full screen layout (1920x1080)")
    print("- Dark theme with blue accents")
    print("- Professional styling with Segoe UI font")
    print("- 'Workflow Manager Service Login:' title")
    print("- 'discover. structure. collaborate. manage. improve.' tagline")
    print("- Centered login form with username and password fields")
    print("- System information and confidentiality notice")
    print("- Help and Cancel buttons")
    print("\nPress Ctrl+C to exit or close the window to cancel login.")
    
    try:
        # Create and show login window
        root = tk.Tk()
        root.withdraw()  # Hide main window
        
        def on_login_success(user_manager):
            print(f"\nLogin successful! User: {user_manager.get_current_user().username}")
            print("Login window will close in 1 second...")
        
        login_window = LoginWindow(parent=None, on_success=on_login_success)
        login_window.login_window.mainloop()
        
    except KeyboardInterrupt:
        print("\nTest cancelled by user")
    except Exception as e:
        print(f"Error during test: {e}")

if __name__ == "__main__":
    main()
