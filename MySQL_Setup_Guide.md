# MySQL Setup Guide for Degrow Workflow Manager

This guide will help you set up MySQL database for the Degrow Workflow Manager system.

## Prerequisites

- MySQL Server 8.0 or later installed and running
- Python 3.11 or later
- All Python dependencies installed (see requirements.txt)

## Step 1: Install MySQL Server

### Windows
1. Download MySQL Community Server from https://dev.mysql.com/downloads/mysql/
2. Run the installer and follow the setup wizard
3. Remember the root password you set during installation
4. Ensure MySQL service is running

### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install mysql-server
sudo mysql_secure_installation
```

### macOS
```bash
brew install mysql
brew services start mysql
mysql_secure_installation
```

## Step 2: Create Database and User

1. Connect to MySQL as root:
```bash
mysql -u root -p
```

2. Create the database:
```sql
CREATE DATABASE degrow_workflow CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

3. Create a dedicated user (optional but recommended):
```sql
CREATE USER 'workflow_user'@'localhost' IDENTIFIED BY 'your_secure_password';
GRANT ALL PRIVILEGES ON degrow_workflow.* TO 'workflow_user'@'localhost';
FLUSH PRIVILEGES;
```

4. Exit MySQL:
```sql
EXIT;
```

## Step 3: Configure the Application

1. Copy the example configuration file:
```bash
cp config.example.json config.json
```

2. Edit `config.json` with your MySQL settings:
```json
{
  "database": {
    "type": "mysql",
    "host": "localhost",
    "port": 3306,
    "username": "workflow_user",
    "password": "your_secure_password",
    "database_name": "degrow_workflow",
    "charset": "utf8mb4",
    "autocommit": true
  },
  "backup": {
    "enabled": false,
    "frequency": "daily",
    "location": "./backups"
  },
  "logging": {
    "level": "INFO",
    "file": "workflow_manager.log"
  }
}
```

## Step 4: Test the Connection

1. Run the database connection test:
```bash
python db_config.py
```

2. If successful, you should see: "Database connection test successful!"

## Step 5: Migrate Existing Data (Optional)

If you have existing SQLite data to migrate:

1. Ensure your SQLite database files are accessible
2. Run the migration script:
```bash
python migrate_sqlite_to_mysql.py
```

3. Follow the prompts and check the migration summary

## Step 6: Run the Application

1. Start the GUI application:
```bash
python question_sheet_gui.py
```

2. Or start the console application:
```bash
python question_sheet_console.py
```

## Troubleshooting

### Connection Issues

**Error: "Access denied for user"**
- Check username and password in config.json
- Ensure the user has proper privileges on the database

**Error: "Can't connect to MySQL server"**
- Verify MySQL service is running
- Check host and port settings
- Ensure firewall allows connections

**Error: "Unknown database"**
- Verify the database exists
- Check database name in config.json

### Performance Issues

**Slow queries:**
- Ensure proper indexes are created
- Check MySQL configuration
- Monitor query performance

**Memory usage:**
- Adjust MySQL buffer pool size
- Monitor connection limits

## Backup and Restore

### Creating Backups

The application includes built-in backup functionality:
1. Open the Database Manager from the main application
2. Go to the "Backup & Restore" tab
3. Select backup location and create backup

### Manual Backup

```bash
mysqldump -u workflow_user -p degrow_workflow > backup_$(date +%Y%m%d_%H%M%S).sql
```

### Restoring from Backup

```bash
mysql -u workflow_user -p degrow_workflow < backup_file.sql
```

## Security Considerations

1. **Use strong passwords** for database users
2. **Limit user privileges** to only what's needed
3. **Enable SSL** for remote connections
4. **Regular backups** of both data and configuration
5. **Keep MySQL updated** with security patches

## Configuration Options

### Database Settings
- `host`: MySQL server hostname (default: localhost)
- `port`: MySQL server port (default: 3306)
- `username`: Database username
- `password`: Database password
- `database_name`: Name of the database to use
- `charset`: Character set (default: utf8mb4)
- `autocommit`: Enable auto-commit (default: true)

### Backup Settings
- `enabled`: Enable automatic backups
- `frequency`: How often to backup (daily, weekly, monthly)
- `location`: Directory to store backups

### Logging Settings
- `level`: Log level (DEBUG, INFO, WARNING, ERROR)
- `file`: Log file path

## Support

If you encounter issues:

1. Check the log files for error messages
2. Verify MySQL server status and configuration
3. Test database connection manually
4. Review this guide for common solutions

For additional help, refer to the main application documentation or contact your system administrator.
