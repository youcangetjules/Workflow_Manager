# User Administration Module Guide

## ⚠️ ADMINISTRATIVE ACCESS ONLY ⚠️

**Version:** 2.0.0  
**Date:** October 31, 2025  
**Author:** Workflow Manager System

---

## Table of Contents

1. [Overview](#overview)
2. [Security Notice](#security-notice)
3. [Features](#features)
4. [Database Schema](#database-schema)
5. [Installation & Setup](#installation--setup)
6. [Using the GUI](#using-the-gui)
7. [Programmatic Usage](#programmatic-usage)
8. [User Management](#user-management)
9. [Password Security](#password-security)
10. [Privileges & Roles](#privileges--roles)
11. [Audit Logging](#audit-logging)
12. [Troubleshooting](#troubleshooting)

---

## Overview

The User Administration Module is a comprehensive system for managing users, passwords, privileges, and access control within the Workflow Manager application. It provides enterprise-grade user management capabilities with a focus on security, auditability, and ease of administration.

### Key Components

- **user_admin.py**: Core administrative module (backend)
- **user_admin_gui.py**: Graphical administrative interface
- **MySQL Database**: Secure storage for user data

### Important

⚠️ **This module is NOT accessible from main.py by design**. It is a standalone administrative tool that should only be accessible to system administrators.

---

## Security Notice

### Access Control

1. **Administrative Access Only**: This system requires administrator-level authentication
2. **Not Part of Main Application**: Deliberately isolated from the main application workflow
3. **Audit Trail**: All administrative actions are logged with timestamps and admin identifiers
4. **Password Protection**: All passwords are hashed using PBKDF2-HMAC-SHA256 with 200,000 iterations

### Best Practices

- ✅ Keep the super admin credentials secure
- ✅ Change default passwords immediately
- ✅ Limit administrator access to trusted personnel only
- ✅ Regularly review audit logs
- ✅ Use strong passwords that meet complexity requirements
- ❌ Never share administrator credentials
- ❌ Never run this on a publicly accessible machine without proper security measures

---

## Features

### Comprehensive User Management

- ✅ **Create Users**: Add new users with complete profile information
- ✅ **Edit Users**: Modify user details, roles, and privileges
- ✅ **Suspend/Activate**: Temporarily or permanently control user access
- ✅ **Delete Users**: Remove users from the system (with confirmation)
- ✅ **Password Management**: Secure password hashing and change functionality
- ✅ **Search & Filter**: Find users by status, role, company, or search terms

### Enhanced User Fields

- Username (unique)
- Email address (unique)
- Role (super_admin, admin, manager, supervisor, user, guest)
- Privilege level (full_access, read_write, read_only, limited, none)
- Status (active, suspended, locked, pending_activation, deactivated)
- **Company** (organization/company name)
- **Position** (job title/position)
- Department
- Employee ID (unique identifier)
- Phone number
- Manager ID (hierarchical management structure)
- Creation date (automatic)
- Last login tracking
- Suspension information (date, reason, end date)

### Security Features

- **Password Hashing**: PBKDF2-HMAC-SHA256 with 200,000 iterations
- **Password History**: Track password changes, prevent reuse
- **Account Locking**: Automatic lock after failed login attempts
- **Password Requirements**: 
  - Minimum 8 characters
  - At least one uppercase letter
  - At least one lowercase letter
  - At least one digit
  - At least one special character
- **Session Management**: Track active sessions and force logout
- **Granular Privileges**: Custom privilege system beyond basic roles

### Audit & Compliance

- **Complete Audit Log**: Track all administrative actions
- **Change Tracking**: Record old and new values for updates
- **Action Attribution**: Know which admin performed which action
- **Timestamp Tracking**: Accurate datetime for all events
- **IP Address Logging**: Track where actions originated
- **Searchable History**: Filter and review audit logs by user or action type

### Statistics & Reporting

- Total user count
- Active sessions
- New users in last 30 days
- User distribution by status
- User distribution by role
- User distribution by company

---

## Database Schema

### Main Tables

#### `admin_users`

Primary user table with enhanced fields:

| Field | Type | Description |
|-------|------|-------------|
| id | INT (Primary Key) | Auto-incrementing user ID |
| username | VARCHAR(50) UNIQUE | Unique username |
| email | VARCHAR(255) UNIQUE | Email address |
| password_hash | VARCHAR(255) | Hashed password |
| salt | VARCHAR(64) | Password salt |
| role | ENUM | User role (super_admin, admin, manager, etc.) |
| privilege_level | ENUM | Access level (full_access, read_write, etc.) |
| status | ENUM | Account status |
| **company** | VARCHAR(255) | Company/organization |
| **position** | VARCHAR(255) | Job title/position |
| department | VARCHAR(255) | Department name |
| employee_id | VARCHAR(50) UNIQUE | Employee identifier |
| manager_id | INT | Foreign key to manager |
| phone_number | VARCHAR(20) | Contact number |
| created_date | DATETIME | Account creation date |
| last_login | DATETIME | Last successful login |
| last_password_change | DATETIME | Last password change |
| suspension_date | DATETIME | When account was suspended |
| suspension_reason | TEXT | Reason for suspension |
| suspension_end_date | DATETIME | When suspension ends |
| suspended_by | INT | Admin who suspended account |
| failed_login_attempts | INT | Counter for failed logins |
| locked_until | DATETIME | Account lock expiration |
| password_reset_token | VARCHAR(255) | Password reset token |
| password_reset_expires | DATETIME | Token expiration |
| created_by | INT | Admin who created account |
| modified_by | INT | Admin who last modified |
| modified_date | DATETIME | Last modification date |
| notes | TEXT | Administrative notes |

#### `user_privileges`

Granular privilege management:

| Field | Type | Description |
|-------|------|-------------|
| id | INT (Primary Key) | Privilege ID |
| user_id | INT | Foreign key to admin_users |
| privilege_name | VARCHAR(100) | Name of privilege |
| privilege_value | BOOLEAN | Granted/revoked |
| granted_by | INT | Admin who granted privilege |
| granted_date | DATETIME | When granted |
| expires_date | DATETIME | Optional expiration |

#### `user_audit_log`

Complete audit trail:

| Field | Type | Description |
|-------|------|-------------|
| id | INT (Primary Key) | Log entry ID |
| user_id | INT | User being acted upon |
| action_type | VARCHAR(50) | Type of action |
| action_description | TEXT | Detailed description |
| performed_by | INT | Admin who performed action |
| ip_address | VARCHAR(45) | IP address |
| timestamp | DATETIME | When action occurred |
| old_values | JSON | Previous values (for updates) |
| new_values | JSON | New values (for updates) |

#### `password_history`

Password change tracking:

| Field | Type | Description |
|-------|------|-------------|
| id | INT (Primary Key) | History entry ID |
| user_id | INT | User ID |
| password_hash | VARCHAR(255) | Historical password hash |
| salt | VARCHAR(64) | Historical salt |
| changed_date | DATETIME | When password was changed |

#### `admin_user_sessions`

Session tracking:

| Field | Type | Description |
|-------|------|-------------|
| id | INT (Primary Key) | Session ID |
| session_id | VARCHAR(255) UNIQUE | Session identifier |
| user_id | INT | User ID |
| login_time | DATETIME | Login timestamp |
| last_activity | DATETIME | Last activity timestamp |
| ip_address | VARCHAR(45) | IP address |
| user_agent | TEXT | Browser/client info |
| is_active | BOOLEAN | Active status |
| logout_time | DATETIME | Logout timestamp |

---

## Installation & Setup

### Prerequisites

- Python 3.11 or later
- MySQL 8.0 or later (configured and running)
- Required Python packages: `tkinter`, `mysql-connector-python`

### Step 1: Ensure MySQL is Configured

Make sure your `config.json` is properly configured for MySQL:

```json
{
  "database": {
    "type": "mysql",
    "host": "localhost",
    "port": 3306,
    "username": "your_username",
    "password": "your_password",
    "database_name": "voice_workflow_manager",
    "charset": "utf8mb4",
    "autocommit": true
  }
}
```

### Step 2: Initialize the Database Schema

Run the user administration module to create the necessary tables:

```bash
python user_admin.py
```

This will:
- Create all required database tables
- Create indexes for performance
- Set up the super admin account (if it doesn't exist)

**Default Super Admin Credentials:**
- Username: `superadmin`
- Password: `SuperAdmin123!`

⚠️ **CHANGE THESE IMMEDIATELY IN PRODUCTION!**

### Step 3: Verify Installation

Check that the following tables were created:
- `admin_users`
- `user_privileges`
- `user_audit_log`
- `password_history`
- `admin_user_sessions`

---

## Using the GUI

### Launching the GUI

```bash
python user_admin_gui.py
```

### Authentication

1. The GUI will prompt for administrator credentials
2. Only users with `super_admin` or `admin` roles can access
3. Enter your username and password

### Main Interface

The GUI has four main tabs:

#### 1. User Management Tab

**Features:**
- Search and filter users
- View user list in a table
- Create new users
- Edit existing users
- Change passwords
- Suspend/activate users
- Delete users

**Actions:**
- **New User**: Opens dialog to create a new user
- **Refresh**: Reload user list
- **View Details**: Show complete user information
- **Edit User**: Modify user details
- **Change Password**: Update user password
- **Suspend User**: Temporarily disable account
- **Activate User**: Enable account
- **Delete User**: Permanently remove user (requires confirmation)

**Search & Filter:**
- Search by username, email, or company
- Filter by status (active, suspended, etc.)
- Filter by role (admin, user, etc.)

**Context Menu:**
Right-click on any user for quick actions

#### 2. User Details Tab

**Features:**
- Complete user profile view
- All user fields displayed
- Automatic navigation when selecting a user

**Information Displayed:**
- Basic info (username, email, ID)
- Role and privileges
- Organization info (company, position, department)
- Contact info (phone, employee ID)
- Status and dates
- Suspension information (if applicable)

#### 3. Statistics Tab

**Features:**
- Overview of user metrics
- Distribution charts
- Recent activity tracking

**Statistics Shown:**
- Total users
- Active sessions
- New users (last 30 days)
- Users by status
- Users by role
- Users by company

#### 4. Audit Log Tab

**Features:**
- Complete history of administrative actions
- Filter by user
- Search by action type

**Log Information:**
- Action ID
- User affected
- Action type (created, updated, suspended, etc.)
- Description
- Admin who performed action
- IP address
- Timestamp

---

## Programmatic Usage

### Importing the Module

```python
from user_admin import UserAdministration, AdminUser
```

### Creating an Instance

```python
admin = UserAdministration(config_file="config.json")
```

### Creating a User

```python
user_id = admin.create_user(
    username="john.doe",
    email="john.doe@company.com",
    password="SecurePass123!",
    role="user",
    privilege_level="read_write",
    company="Acme Corporation",
    position="Software Engineer",
    department="Engineering",
    employee_id="EMP-12345",
    phone_number="+1-555-0123",
    auto_activate=True,
    created_by=1  # Admin user ID
)

if user_id:
    print(f"User created with ID: {user_id}")
```

### Getting a User

```python
# By ID
user = admin.get_user_by_id(1)

# By username
user = admin.get_user_by_username("john.doe")

# By email
user = admin.get_user_by_email("john.doe@company.com")

# All users
all_users = admin.get_all_users()

# Filtered users
active_users = admin.get_users_by_status("active")
company_users = admin.get_users_by_company("Acme Corporation")
```

### Updating a User

```python
success = admin.update_user(
    user_id=1,
    modified_by=1,  # Admin user ID
    position="Senior Software Engineer",
    department="Engineering - Platform",
    phone_number="+1-555-0124"
)
```

### Changing Password

```python
success = admin.change_password(
    user_id=1,
    new_password="NewSecurePass123!",
    changed_by=1  # Admin user ID
)
```

### Suspending a User

```python
success = admin.suspend_user(
    user_id=1,
    reason="Policy violation - pending investigation",
    suspended_by=1,  # Admin user ID
    suspension_end_date=datetime.now() + timedelta(days=7)
)
```

### Activating a User

```python
success = admin.activate_user(
    user_id=1,
    activated_by=1  # Admin user ID
)
```

### Deleting a User

```python
success = admin.delete_user(
    user_id=1,
    deleted_by=1  # Admin user ID
)
```

### Privilege Management

```python
# Grant a privilege
success = admin.grant_privilege(
    user_id=1,
    privilege_name="export_data",
    granted_by=1,
    expires_date=datetime.now() + timedelta(days=90)
)

# Revoke a privilege
success = admin.revoke_privilege(
    user_id=1,
    privilege_name="export_data",
    revoked_by=1
)

# Get user privileges
privileges = admin.get_user_privileges(user_id=1)
```

### Audit Log

```python
# Get all audit logs
logs = admin.get_audit_log(limit=100)

# Get logs for specific user
user_logs = admin.get_audit_log(user_id=1, limit=50)
```

### Statistics

```python
stats = admin.get_user_statistics()

print(f"Total users: {stats['total_users']}")
print(f"Active sessions: {stats['active_sessions']}")
print(f"Recently created: {stats['recently_created']}")
print(f"By status: {stats['by_status']}")
print(f"By role: {stats['by_role']}")
print(f"By company: {stats['by_company']}")
```

---

## User Management

### Creating Users

**Required Fields:**
- Username (3+ characters, unique)
- Email (valid format, unique)
- Password (meets complexity requirements)

**Optional Fields:**
- Company
- Position
- Department
- Employee ID
- Phone number
- Manager ID
- Role (default: user)
- Privilege level (default: limited)
- Auto-activate (default: False)

**Best Practices:**
- Use meaningful usernames (e.g., first.last)
- Require corporate email addresses
- Assign appropriate roles based on job function
- Set privilege levels based on least privilege principle
- Document user creation in notes field

### User Roles

| Role | Description | Typical Use |
|------|-------------|-------------|
| `super_admin` | Full system access | System administrators |
| `admin` | Administrative access | Department heads |
| `manager` | Management functions | Team managers |
| `supervisor` | Supervisory access | Team leads |
| `user` | Standard user | Regular employees |
| `guest` | Limited access | Temporary/external users |

### User Status

| Status | Description |
|--------|-------------|
| `active` | Account is active and accessible |
| `suspended` | Temporarily disabled (with reason and optional end date) |
| `locked` | Automatically locked after failed logins |
| `pending_activation` | Awaiting admin activation |
| `deactivated` | Permanently disabled |

### Privilege Levels

| Level | Description |
|-------|-------------|
| `full_access` | Complete system access |
| `read_write` | Can read and modify data |
| `read_only` | View-only access |
| `limited` | Restricted access to specific features |
| `none` | No access (essentially disabled) |

---

## Password Security

### Password Requirements

By default, passwords must meet these criteria:
- Minimum 8 characters
- At least one uppercase letter (A-Z)
- At least one lowercase letter (a-z)
- At least one digit (0-9)
- At least one special character (!@#$%^&*(),.?":{}|<>)

### Hashing Algorithm

Passwords are hashed using **PBKDF2-HMAC-SHA256** with:
- 200,000 iterations (industry standard for security)
- Unique 64-character salt per user
- 256-bit output

### Password History

The system tracks password history to:
- Prevent password reuse
- Maintain audit trail of changes
- Support security compliance requirements

### Failed Login Protection

- After 5 failed login attempts, account is automatically locked
- Lock duration: 15 minutes
- Reset counter upon successful login
- All attempts are logged

### Best Practices

✅ **DO:**
- Use unique passwords for each account
- Change default passwords immediately
- Enforce regular password changes
- Use password managers
- Educate users on password security

❌ **DON'T:**
- Share passwords
- Write passwords down
- Use common passwords
- Reuse old passwords
- Store passwords in plain text

---

## Privileges & Roles

### Role Hierarchy

```
super_admin (highest)
    └─ admin
        └─ manager
            └─ supervisor
                └─ user
                    └─ guest (lowest)
```

### Granular Privileges

Beyond basic roles, the system supports custom privileges:

```python
# Example privileges
admin.grant_privilege(user_id, "create_workflow")
admin.grant_privilege(user_id, "approve_requests")
admin.grant_privilege(user_id, "export_data")
admin.grant_privilege(user_id, "manage_team")
```

### Permission Checking

```python
# Check if user has permission
user = admin.get_user_by_id(user_id)

# Role-based check
if user.role in ['admin', 'super_admin']:
    # Allow administrative action
    pass

# Privilege-based check
privileges = admin.get_user_privileges(user_id)
if any(p['name'] == 'export_data' and p['value'] for p in privileges):
    # Allow data export
    pass
```

---

## Audit Logging

### What is Logged

Every administrative action is logged, including:
- User creation
- User updates
- Password changes
- Account suspensions
- Account activations
- Account deletions
- Privilege grants/revokes
- All failed operations

### Log Entry Details

Each log entry contains:
- **User ID**: User being acted upon
- **Action Type**: Type of operation
- **Description**: Human-readable description
- **Performed By**: Admin user ID
- **IP Address**: Source IP
- **Timestamp**: When action occurred
- **Old Values**: Previous data (for updates)
- **New Values**: New data (for updates)

### Accessing Logs

**Via GUI:**
- Navigate to "Audit Log" tab
- Filter by user ID if needed
- View complete history

**Programmatically:**
```python
# Get all logs
logs = admin.get_audit_log(limit=100)

# Get logs for specific user
user_logs = admin.get_audit_log(user_id=1, limit=50)

# Process logs
for log in logs:
    print(f"{log['timestamp']}: {log['username']} - {log['action_type']}")
    print(f"  Performed by: {log['performed_by_name']}")
    print(f"  Description: {log['description']}")
```

### Log Retention

- Logs are retained indefinitely in the database
- Consider implementing archival for very old logs
- Regular backups are recommended
- Comply with your organization's data retention policies

---

## Troubleshooting

### Common Issues

#### 1. Cannot Connect to Database

**Error:** `Failed to connect to database`

**Solutions:**
- Verify MySQL is running
- Check `config.json` settings
- Verify network connectivity
- Check firewall settings
- Verify user has database permissions

#### 2. Super Admin Login Fails

**Problem:** Cannot login with default credentials

**Solutions:**
- Run `python user_admin.py` to create super admin
- Check if superadmin user exists in database
- Verify password hasn't been changed
- Check user status (should be 'active')

```sql
-- Check superadmin user
SELECT id, username, status, role FROM admin_users WHERE username = 'superadmin';
```

#### 3. Password Requirements Too Strict

**Problem:** Cannot create user due to password requirements

**Solution:** Modify password requirements in `user_admin.py`:

```python
# In UserAdministration.__init__
self.min_password_length = 8  # Change as needed
self.require_uppercase = True  # Set to False to disable
self.require_lowercase = True
self.require_digit = True
self.require_special = True
```

#### 4. GUI Won't Start

**Error:** Authentication fails or GUI crashes

**Solutions:**
- Verify tkinter is installed
- Check Python version (3.11+)
- Verify database tables exist
- Check logs in `user_admin.log`
- Run with `python user_admin_gui.py` from command line to see errors

#### 5. User Creation Fails

**Problem:** Cannot create new user

**Possible Causes:**
- Username already exists
- Email already exists
- Employee ID already exists (if provided)
- Password doesn't meet requirements
- Database connection issues

**Debug:**
```python
# Check if username exists
user = admin.get_user_by_username("problematic_username")
if user:
    print("Username already exists")

# Check logs
with open('user_admin.log', 'r') as f:
    print(f.read())
```

#### 6. Table Creation Errors

**Error:** `Table already exists` or schema mismatch

**Solution:**
```sql
-- Drop and recreate tables (WARNING: DELETES ALL DATA)
DROP TABLE IF EXISTS admin_user_sessions;
DROP TABLE IF EXISTS password_history;
DROP TABLE IF EXISTS user_audit_log;
DROP TABLE IF EXISTS user_privileges;
DROP TABLE IF EXISTS admin_users;

-- Then run: python user_admin.py
```

### Logging

All operations are logged to:
- **File:** `user_admin.log`
- **Console:** Standard output

Check logs for detailed error messages:
```bash
tail -f user_admin.log
```

### Database Verification

```sql
-- Check table structure
SHOW TABLES;

-- Check admin_users table
DESCRIBE admin_users;

-- Check user count
SELECT COUNT(*) FROM admin_users;

-- Check super admin
SELECT * FROM admin_users WHERE role = 'super_admin';

-- Check recent audit logs
SELECT * FROM user_audit_log ORDER BY timestamp DESC LIMIT 10;
```

---

## Security Considerations

### Deployment

1. **Access Control:**
   - Restrict physical/network access to admin interface
   - Use VPN or firewall rules to limit access
   - Never expose on public internet without proper security

2. **Authentication:**
   - Change default passwords immediately
   - Use strong, unique passwords
   - Consider implementing 2FA (future enhancement)

3. **Database Security:**
   - Use strong MySQL root password
   - Create dedicated MySQL user with limited privileges
   - Enable SSL for MySQL connections
   - Regular database backups

4. **Audit & Monitoring:**
   - Regularly review audit logs
   - Monitor for suspicious activity
   - Set up alerts for critical actions
   - Backup audit logs separately

5. **Update & Maintenance:**
   - Keep Python and MySQL updated
   - Review and update security settings
   - Test backups regularly
   - Document all procedures

### Compliance

This module supports compliance with:
- **GDPR**: Data access controls, audit trails
- **SOC 2**: Access logging, user management
- **ISO 27001**: User lifecycle management, authentication
- **HIPAA**: Access controls, audit logs (if applicable)

Consult with your compliance officer to ensure proper configuration.

---

## Support & Maintenance

### Backup Recommendations

1. **Database Backup:**
```bash
mysqldump -u username -p voice_workflow_manager admin_users user_privileges user_audit_log password_history admin_user_sessions > user_admin_backup_$(date +%Y%m%d).sql
```

2. **Automated Backups:**
- Schedule daily backups
- Rotate backup files
- Store offsite
- Test restore procedures

### Maintenance Tasks

**Daily:**
- Monitor audit logs
- Check for suspended accounts needing review

**Weekly:**
- Review user access levels
- Check for inactive accounts
- Verify backup integrity

**Monthly:**
- User access audit
- Password policy review
- Update documentation
- Security review

### Getting Help

**Log Files:**
- `user_admin.log` - All administrative operations
- `workflow_manager.log` - General application logs

**Database Query:**
```sql
-- Get recent errors from logs
SELECT * FROM user_audit_log 
WHERE action_description LIKE '%error%' 
ORDER BY timestamp DESC 
LIMIT 20;
```

---

## Appendix

### SQL Queries for Common Tasks

**Find all admins:**
```sql
SELECT username, email, role, status 
FROM admin_users 
WHERE role IN ('admin', 'super_admin');
```

**Find suspended users:**
```sql
SELECT username, email, suspension_reason, suspension_date 
FROM admin_users 
WHERE status = 'suspended';
```

**Find users by company:**
```sql
SELECT username, email, position 
FROM admin_users 
WHERE company = 'Acme Corporation';
```

**Get audit summary:**
```sql
SELECT action_type, COUNT(*) as count 
FROM user_audit_log 
GROUP BY action_type 
ORDER BY count DESC;
```

### Configuration Options

**Password Requirements:**
Located in `UserAdministration.__init__()`:
```python
self.min_password_length = 8
self.require_uppercase = True
self.require_lowercase = True
self.require_digit = True
self.require_special = True
```

**Failed Login Settings:**
Located in `user_manager.py` (if integrating):
```python
MAX_FAILED_ATTEMPTS = 5
LOCK_DURATION_MINUTES = 15
```

---

## Changelog

### Version 2.0.0 (2025-10-31)

**New Features:**
- Complete user administration system
- Enhanced user fields (company, position, department, employee_id)
- Comprehensive GUI interface
- Granular privilege system
- Full audit logging with JSON support
- Password history tracking
- Session management
- Statistics and reporting
- Search and filter capabilities

**Security Enhancements:**
- PBKDF2-HMAC-SHA256 with 200,000 iterations
- Password complexity requirements
- Account locking after failed attempts
- Comprehensive audit trail
- IP address logging

**Database Schema:**
- `admin_users` table with all enhanced fields
- `user_privileges` for granular permissions
- `user_audit_log` with JSON support
- `password_history` for change tracking
- `admin_user_sessions` for session management

---

## License

This User Administration Module is part of the Workflow Manager System.  
All rights reserved.

---

## Contact

For support, feature requests, or bug reports, contact your system administrator.

**Remember:** This is a powerful administrative tool. Use responsibly and always follow security best practices.

