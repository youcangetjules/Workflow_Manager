# User Administration Module - Quick Start

## ⚠️ ADMINISTRATIVE ACCESS ONLY ⚠️

This is a standalone administrative module for managing users in the Workflow Manager system. It is **NOT** accessible from `main.py` by design.

---

## Quick Setup

### 1. Install/Verify MySQL

Ensure MySQL is running and configured in `config.json`:

```json
{
  "database": {
    "type": "mysql",
    "host": "localhost",
    "port": 3306,
    "username": "your_username",
    "password": "your_password",
    "database_name": "voice_workflow_manager"
  }
}
```

### 2. Initialize Database

Run once to create tables and super admin:

```bash
python user_admin.py
```

**Default Super Admin:**
- Username: `superadmin`
- Password: `SuperAdmin123!`

⚠️ **CHANGE THIS IMMEDIATELY!**

### 3. Launch GUI

```bash
python user_admin_gui.py
```

---

## Features

✅ **Complete User Management**
- Create, edit, delete users
- Suspend/activate accounts
- Change passwords securely

✅ **Enhanced User Fields**
- Company name
- Position/job title
- Department
- Employee ID
- Phone number
- Manager hierarchy

✅ **Security**
- Hashed passwords (PBKDF2-HMAC-SHA256)
- Role-based access control
- Privilege management
- Account suspension with reasons

✅ **Audit & Compliance**
- Complete audit trail
- All actions logged with timestamps
- Track who did what and when

✅ **Statistics**
- User distribution reports
- Active session tracking
- Company/role breakdowns

---

## User Roles

| Role | Access Level |
|------|--------------|
| `super_admin` | Full system access |
| `admin` | Administrative functions |
| `manager` | Management capabilities |
| `supervisor` | Supervisory access |
| `user` | Standard access |
| `guest` | Limited access |

---

## Common Tasks

### Create a New User

**Via GUI:**
1. Launch `user_admin_gui.py`
2. Go to "User Management" tab
3. Click "➕ New User"
4. Fill in required fields (username, email, password)
5. Set role and privilege level
6. Fill in company, position, etc.
7. Click "Create User"

**Programmatically:**
```python
from user_admin import UserAdministration

admin = UserAdministration()
user_id = admin.create_user(
    username="john.doe",
    email="john.doe@company.com",
    password="SecurePass123!",
    role="user",
    company="Acme Corporation",
    position="Software Engineer",
    auto_activate=True,
    created_by=1  # Your admin user ID
)
```

### Suspend a User

```python
admin.suspend_user(
    user_id=5,
    reason="Policy violation - pending review",
    suspended_by=1
)
```

### Change Password

```python
admin.change_password(
    user_id=5,
    new_password="NewSecurePass123!",
    changed_by=1
)
```

### View Statistics

```python
stats = admin.get_user_statistics()
print(f"Total users: {stats['total_users']}")
print(f"Active sessions: {stats['active_sessions']}")
```

---

## Database Tables

The module creates these tables:

1. **admin_users** - Main user table with all fields
2. **user_privileges** - Granular privilege management
3. **user_audit_log** - Complete audit trail
4. **password_history** - Password change tracking
5. **admin_user_sessions** - Session management

---

## Security Features

✅ Password hashing with 200,000 PBKDF2 iterations  
✅ Unique salt per user  
✅ Password complexity requirements  
✅ Account locking after failed logins  
✅ Complete audit logging  
✅ IP address tracking  
✅ Session management  

---

## Files

| File | Purpose |
|------|---------|
| `user_admin.py` | Core administrative module (backend) |
| `user_admin_gui.py` | Graphical interface |
| `User_Administration_Guide.md` | Comprehensive documentation |
| `USER_ADMIN_README.md` | This quick-start guide |

---

## Important Notes

⚠️ **Standalone Module**
- NOT imported by `main.py`
- Separate from main application
- Administrative access only

⚠️ **Security**
- Change default passwords
- Restrict access to trusted admins
- Regular audit log reviews
- Keep credentials secure

⚠️ **Backup**
- Regular database backups
- Test restore procedures
- Keep offsite copies

---

## Troubleshooting

### Cannot connect to database
- Verify MySQL is running
- Check `config.json` settings
- Verify database exists
- Check user permissions

### GUI won't start
- Verify Python 3.11+
- Check tkinter is installed
- Run from command line to see errors
- Check `user_admin.log`

### Super admin login fails
- Run `python user_admin.py` to create
- Check user exists in database
- Verify status is 'active'

---

## Support

**Documentation:** See `User_Administration_Guide.md` for complete documentation

**Logs:** Check `user_admin.log` for detailed operation logs

**Database:** Direct MySQL access for troubleshooting

---

## Quick Reference

### Password Requirements
- Minimum 8 characters
- 1+ uppercase letter
- 1+ lowercase letter
- 1+ digit
- 1+ special character

### User Status Values
- `active` - Account is active
- `suspended` - Temporarily disabled
- `locked` - Auto-locked after failed logins
- `pending_activation` - Awaiting activation
- `deactivated` - Permanently disabled

### Privilege Levels
- `full_access` - Complete access
- `read_write` - Can modify data
- `read_only` - View only
- `limited` - Restricted features
- `none` - No access

---

**Version:** 2.0.0  
**Last Updated:** October 31, 2025

For detailed information, see **User_Administration_Guide.md**

