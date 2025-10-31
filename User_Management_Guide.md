# User Management Guide for Degrow Workflow Manager

This guide covers the user management and authentication features of the Degrow Workflow Manager system.

## Overview

The system now includes comprehensive user management with:
- User authentication and login
- Role-based access control
- User activity tracking
- Session management
- Password security

## User Roles

### Admin
- Full system access
- User management capabilities
- Database administration
- All workflow operations

### Manager
- Workflow management
- Milestone and subtask editing
- Data export capabilities
- Limited user management

### User
- Basic workflow operations
- View and edit entries
- Limited administrative access

## Getting Started

### First Time Setup

1. **Start the Application**
   ```bash
   python question_sheet_gui.py
   ```

2. **Default Admin Account**
   - Username: `admin`
   - Password: `admin123`
   - **Important**: Change this password immediately after first login

3. **Login Process**
   - Enter username or email
   - Enter password
   - Click "Login" or press Enter

### Changing Default Password

1. Login with admin account
2. Go to **User** → **Change Password**
3. Enter current password
4. Enter new password (minimum 6 characters)
5. Confirm new password
6. Click "Change Password"

## User Management (Admin Only)

### Adding New Users

1. Login as admin
2. Go to **User** → **User Management**
3. Click "Add User"
4. Fill in user details:
   - Username (required)
   - Email (required)
   - Password (required for new users)
   - Role (user, manager, admin)
5. Click "Save"

### Editing Users

1. Open User Management
2. Select user from list
3. Click "Edit User"
4. Modify details as needed
5. Click "Save"

### Deactivating Users

1. Open User Management
2. Select user from list
3. Click "Deactivate User"
4. Confirm deactivation

## Security Features

### Password Security
- Passwords are hashed using PBKDF2 with SHA-256
- Salt is generated for each password
- Minimum 6 characters required
- Failed login attempts are tracked

### Account Lockout
- 5 failed login attempts locks account for 15 minutes
- Lockout resets on successful login
- Admin can unlock accounts

### Session Management
- Sessions are tracked in database
- Automatic logout on application close
- Session timeout (configurable)
- IP address tracking

### Activity Logging
- All user actions are logged
- Includes timestamp and IP address
- Activity types tracked:
  - Login/Logout
  - Create/Update/Delete entries
  - Database operations
  - User management actions

## Database Schema

### Users Table
```sql
CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL UNIQUE,
    email VARCHAR(255) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    salt VARCHAR(32) NOT NULL,
    role ENUM('admin', 'manager', 'user') DEFAULT 'user',
    is_active BOOLEAN DEFAULT TRUE,
    created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
    last_login DATETIME NULL,
    failed_login_attempts INT DEFAULT 0,
    locked_until DATETIME NULL
);
```

### User Sessions Table
```sql
CREATE TABLE user_sessions (
    id INT AUTO_INCREMENT PRIMARY KEY,
    session_id VARCHAR(64) NOT NULL UNIQUE,
    user_id INT NOT NULL,
    login_time DATETIME DEFAULT CURRENT_TIMESTAMP,
    last_activity DATETIME DEFAULT CURRENT_TIMESTAMP,
    ip_address VARCHAR(45),
    is_active BOOLEAN DEFAULT TRUE,
    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
);
```

### Activity Log Table
```sql
CREATE TABLE user_activity_log (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    session_id VARCHAR(64),
    activity_type VARCHAR(50) NOT NULL,
    activity_description TEXT,
    ip_address VARCHAR(45),
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
);
```

## Configuration

### User Management Settings

The user management system uses the same `config.json` file as the database configuration:

```json
{
  "database": {
    "type": "mysql",
    "host": "localhost",
    "port": 3306,
    "username": "workflow_user",
    "password": "your_password",
    "database_name": "degrow_workflow"
  },
  "user_management": {
    "session_timeout_minutes": 480,
    "max_failed_attempts": 5,
    "lockout_duration_minutes": 15,
    "password_min_length": 6
  }
}
```

## Troubleshooting

### Common Issues

**"Access Denied" Error**
- Check if user has required role
- Verify user account is active
- Contact administrator

**Login Failed**
- Check username and password
- Verify account is not locked
- Wait for lockout period to expire

**User Management Not Available**
- Must be logged in as admin
- Check user permissions
- Verify database connection

### Password Issues

**Forgot Password**
- Contact administrator to reset
- Admin can change password in User Management

**Password Not Working**
- Check caps lock
- Verify correct username
- Try resetting password

### Session Issues

**Session Expired**
- Log in again
- Check session timeout settings
- Verify system clock

**Multiple Sessions**
- Only one active session per user
- New login deactivates old session

## Best Practices

### For Administrators

1. **Change Default Password**
   - Immediately after installation
   - Use strong, unique password
   - Regular password updates

2. **User Management**
   - Create users with appropriate roles
   - Regular review of user accounts
   - Deactivate unused accounts

3. **Security Monitoring**
   - Review activity logs regularly
   - Monitor failed login attempts
   - Check for suspicious activity

### For Users

1. **Password Security**
   - Use strong passwords
   - Don't share credentials
   - Change password regularly

2. **Session Management**
   - Log out when finished
   - Don't leave sessions open
   - Report suspicious activity

3. **Data Access**
   - Only access needed data
   - Follow company policies
   - Report security concerns

## API Reference

### UserManager Class

```python
from user_manager import UserManager

# Initialize user manager
user_manager = UserManager()

# Authenticate user
user = user_manager.authenticate_user(username, password)

# Check permissions
if user_manager.has_permission('admin'):
    # Admin-only code

# Log activity
user_manager.log_activity(user_id, 'action', 'description')

# Logout
user_manager.logout()
```

### Login GUI

```python
from login_gui import show_login_window

# Show login window
user_manager = show_login_window(parent_window, on_success_callback)
```

## Support

For additional help with user management:

1. Check this guide for common solutions
2. Review activity logs for error details
3. Contact your system administrator
4. Refer to the main application documentation

## Security Considerations

- All passwords are securely hashed
- Sessions are properly managed
- Activity is logged for audit purposes
- Role-based access prevents unauthorized access
- Account lockout prevents brute force attacks

Remember to keep your credentials secure and report any security concerns to your administrator.
