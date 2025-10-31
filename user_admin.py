"""
User Administration Module
==========================

Comprehensive user administration system for managing users, passwords, privileges,
suspension status, company and position information.

WARNING: This module is for administrative purposes only and should NOT be imported 
or accessed from main.py. Access to this module should be restricted to system 
administrators only.

Author: Workflow Manager System
Version: 2.0.0
Date: 2025-10-31
"""

import hashlib
import secrets
import re
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Tuple, Any
import logging
from enum import Enum
from db_config import get_database_connection, DatabaseConnection

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('user_admin.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class UserRole(Enum):
    """User role enumeration"""
    SUPER_ADMIN = 'super_admin'
    ADMIN = 'admin'
    MANAGER = 'manager'
    SUPERVISOR = 'supervisor'
    USER = 'user'
    GUEST = 'guest'


class PrivilegeLevel(Enum):
    """Privilege level enumeration"""
    FULL_ACCESS = 'full_access'
    READ_WRITE = 'read_write'
    READ_ONLY = 'read_only'
    LIMITED = 'limited'
    NONE = 'none'


class UserStatus(Enum):
    """User status enumeration"""
    ACTIVE = 'active'
    SUSPENDED = 'suspended'
    LOCKED = 'locked'
    PENDING_ACTIVATION = 'pending_activation'
    DEACTIVATED = 'deactivated'


class AdminUser:
    """Enhanced User data class for administration"""
    
    def __init__(self, user_id: int, username: str, email: str, 
                 role: str, privilege_level: str, status: str,
                 company: str = None, position: str = None,
                 department: str = None, manager_id: int = None,
                 created_date: datetime = None, last_login: datetime = None,
                 suspension_date: datetime = None, suspension_reason: str = None,
                 suspension_end_date: datetime = None, created_by: int = None,
                 phone_number: str = None, employee_id: str = None):
        
        self.user_id = user_id
        self.username = username
        self.email = email
        self.role = role
        self.privilege_level = privilege_level
        self.status = status
        self.company = company
        self.position = position
        self.department = department
        self.manager_id = manager_id
        self.created_date = created_date or datetime.now()
        self.last_login = last_login
        self.suspension_date = suspension_date
        self.suspension_reason = suspension_reason
        self.suspension_end_date = suspension_end_date
        self.created_by = created_by
        self.phone_number = phone_number
        self.employee_id = employee_id
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert user object to dictionary"""
        return {
            'user_id': self.user_id,
            'username': self.username,
            'email': self.email,
            'role': self.role,
            'privilege_level': self.privilege_level,
            'status': self.status,
            'company': self.company,
            'position': self.position,
            'department': self.department,
            'manager_id': self.manager_id,
            'created_date': self.created_date.isoformat() if self.created_date else None,
            'last_login': self.last_login.isoformat() if self.last_login else None,
            'suspension_date': self.suspension_date.isoformat() if self.suspension_date else None,
            'suspension_reason': self.suspension_reason,
            'suspension_end_date': self.suspension_end_date.isoformat() if self.suspension_end_date else None,
            'phone_number': self.phone_number,
            'employee_id': self.employee_id
        }


class UserAdministration:
    """
    Comprehensive User Administration System
    
    This class provides complete user management functionality including:
    - User creation, modification, and deletion
    - Password management with secure hashing
    - Privilege and role management
    - User suspension and activation
    - Company and position tracking
    - Comprehensive audit logging
    """
    
    def __init__(self, config_file: str = "config.json"):
        """
        Initialize User Administration System
        
        Args:
            config_file: Path to database configuration file
        """
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        self.admin_user = None  # The admin using this module
        
        # Password requirements
        self.min_password_length = 8
        self.require_uppercase = True
        self.require_lowercase = True
        self.require_digit = True
        self.require_special = True
        
        # Initialize database schema
        self._initialize_admin_schema()
        
        logger.info("User Administration System initialized")
    
    def _initialize_admin_schema(self):
        """Create or update the enhanced user administration schema"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            is_mysql = self._is_mysql(conn)
            
            if is_mysql:
                # MySQL schema with enhanced fields
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS admin_users (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        username VARCHAR(50) NOT NULL UNIQUE,
                        email VARCHAR(255) NOT NULL UNIQUE,
                        password_hash VARCHAR(255) NOT NULL,
                        salt VARCHAR(64) NOT NULL,
                        role ENUM('super_admin', 'admin', 'manager', 'supervisor', 'user', 'guest') DEFAULT 'user',
                        privilege_level ENUM('full_access', 'read_write', 'read_only', 'limited', 'none') DEFAULT 'limited',
                        status ENUM('active', 'suspended', 'locked', 'pending_activation', 'deactivated') DEFAULT 'pending_activation',
                        
                        -- Organization fields
                        company VARCHAR(255),
                        position VARCHAR(255),
                        department VARCHAR(255),
                        employee_id VARCHAR(50) UNIQUE,
                        manager_id INT,
                        
                        -- Contact information
                        phone_number VARCHAR(20),
                        
                        -- Date tracking
                        created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                        last_login DATETIME NULL,
                        last_password_change DATETIME DEFAULT CURRENT_TIMESTAMP,
                        
                        -- Suspension tracking
                        suspension_date DATETIME NULL,
                        suspension_reason TEXT,
                        suspension_end_date DATETIME NULL,
                        suspended_by INT,
                        
                        -- Security fields
                        failed_login_attempts INT DEFAULT 0,
                        locked_until DATETIME NULL,
                        password_reset_token VARCHAR(255),
                        password_reset_expires DATETIME NULL,
                        
                        -- Audit fields
                        created_by INT,
                        modified_by INT,
                        modified_date DATETIME NULL,
                        
                        -- Notes
                        notes TEXT,
                        
                        FOREIGN KEY (manager_id) REFERENCES admin_users (id) ON DELETE SET NULL,
                        FOREIGN KEY (created_by) REFERENCES admin_users (id) ON DELETE SET NULL,
                        FOREIGN KEY (suspended_by) REFERENCES admin_users (id) ON DELETE SET NULL,
                        INDEX idx_username (username),
                        INDEX idx_email (email),
                        INDEX idx_company (company),
                        INDEX idx_status (status),
                        INDEX idx_role (role),
                        INDEX idx_employee_id (employee_id)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
                
                # User privileges table for granular permissions
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_privileges (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        user_id INT NOT NULL,
                        privilege_name VARCHAR(100) NOT NULL,
                        privilege_value BOOLEAN DEFAULT TRUE,
                        granted_by INT,
                        granted_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                        expires_date DATETIME NULL,
                        
                        FOREIGN KEY (user_id) REFERENCES admin_users (id) ON DELETE CASCADE,
                        FOREIGN KEY (granted_by) REFERENCES admin_users (id) ON DELETE SET NULL,
                        UNIQUE KEY unique_user_privilege (user_id, privilege_name),
                        INDEX idx_user_privilege (user_id, privilege_name)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
                
                # User audit log
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_audit_log (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        user_id INT NOT NULL,
                        action_type VARCHAR(50) NOT NULL,
                        action_description TEXT,
                        performed_by INT,
                        ip_address VARCHAR(45),
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        old_values JSON,
                        new_values JSON,
                        
                        FOREIGN KEY (user_id) REFERENCES admin_users (id) ON DELETE CASCADE,
                        FOREIGN KEY (performed_by) REFERENCES admin_users (id) ON DELETE SET NULL,
                        INDEX idx_user_audit (user_id),
                        INDEX idx_action_type (action_type),
                        INDEX idx_timestamp (timestamp),
                        INDEX idx_performed_by (performed_by)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
                
                # Password history table (prevent password reuse)
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS password_history (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        user_id INT NOT NULL,
                        password_hash VARCHAR(255) NOT NULL,
                        salt VARCHAR(64) NOT NULL,
                        changed_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                        
                        FOREIGN KEY (user_id) REFERENCES admin_users (id) ON DELETE CASCADE,
                        INDEX idx_user_password_history (user_id, changed_date)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
                
                # User sessions table
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS admin_user_sessions (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        session_id VARCHAR(255) NOT NULL UNIQUE,
                        user_id INT NOT NULL,
                        login_time DATETIME DEFAULT CURRENT_TIMESTAMP,
                        last_activity DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ip_address VARCHAR(45),
                        user_agent TEXT,
                        is_active BOOLEAN DEFAULT TRUE,
                        logout_time DATETIME NULL,
                        
                        FOREIGN KEY (user_id) REFERENCES admin_users (id) ON DELETE CASCADE,
                        INDEX idx_session_user (user_id),
                        INDEX idx_session_active (is_active),
                        INDEX idx_session_id (session_id)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
            
            else:
                # SQLite schema (simplified)
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS admin_users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT NOT NULL UNIQUE,
                        email TEXT NOT NULL UNIQUE,
                        password_hash TEXT NOT NULL,
                        salt TEXT NOT NULL,
                        role TEXT DEFAULT 'user' CHECK(role IN ('super_admin', 'admin', 'manager', 'supervisor', 'user', 'guest')),
                        privilege_level TEXT DEFAULT 'limited' CHECK(privilege_level IN ('full_access', 'read_write', 'read_only', 'limited', 'none')),
                        status TEXT DEFAULT 'pending_activation' CHECK(status IN ('active', 'suspended', 'locked', 'pending_activation', 'deactivated')),
                        company TEXT,
                        position TEXT,
                        department TEXT,
                        employee_id TEXT UNIQUE,
                        manager_id INTEGER,
                        phone_number TEXT,
                        created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                        last_login TEXT,
                        last_password_change TEXT DEFAULT CURRENT_TIMESTAMP,
                        suspension_date TEXT,
                        suspension_reason TEXT,
                        suspension_end_date TEXT,
                        suspended_by INTEGER,
                        failed_login_attempts INTEGER DEFAULT 0,
                        locked_until TEXT,
                        password_reset_token TEXT,
                        password_reset_expires TEXT,
                        created_by INTEGER,
                        modified_by INTEGER,
                        modified_date TEXT,
                        notes TEXT
                    )
                """)
                
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_audit_log (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER NOT NULL,
                        action_type TEXT NOT NULL,
                        action_description TEXT,
                        performed_by INTEGER,
                        ip_address TEXT,
                        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
                        old_values TEXT,
                        new_values TEXT
                    )
                """)
            
            conn.commit()
            cursor.close()
            conn.close()
            
            logger.info("User administration schema initialized successfully")
            
        except Exception as e:
            logger.error(f"Error initializing admin schema: {e}")
            raise
    
    def _is_mysql(self, conn) -> bool:
        """Check if connection is MySQL"""
        return hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
    
    def _get_placeholder(self) -> str:
        """Get the correct SQL placeholder for the database type"""
        try:
            conn = self.db_conn.connect()
            is_mysql = self._is_mysql(conn)
            conn.close()
            return '%s' if is_mysql else '?'
        except:
            return '?'
    
    def _hash_password(self, password: str, salt: str = None) -> Tuple[str, str]:
        """
        Hash password using PBKDF2-HMAC-SHA256
        
        Args:
            password: Plain text password
            salt: Salt (generated if None)
            
        Returns:
            Tuple of (hashed_password, salt)
        """
        if salt is None:
            salt = secrets.token_hex(32)
        
        # Use PBKDF2 with 200,000 iterations for enhanced security
        password_hash = hashlib.pbkdf2_hmac(
            'sha256',
            password.encode('utf-8'),
            salt.encode('utf-8'),
            200000
        )
        return password_hash.hex(), salt
    
    def _validate_password(self, password: str) -> Tuple[bool, str]:
        """
        Validate password against security requirements
        
        Args:
            password: Password to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if len(password) < self.min_password_length:
            return False, f"Password must be at least {self.min_password_length} characters"
        
        if self.require_uppercase and not re.search(r'[A-Z]', password):
            return False, "Password must contain at least one uppercase letter"
        
        if self.require_lowercase and not re.search(r'[a-z]', password):
            return False, "Password must contain at least one lowercase letter"
        
        if self.require_digit and not re.search(r'\d', password):
            return False, "Password must contain at least one digit"
        
        if self.require_special and not re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
            return False, "Password must contain at least one special character"
        
        return True, ""
    
    def _validate_email(self, email: str) -> bool:
        """Validate email format"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def _log_audit(self, user_id: int, action_type: str, description: str,
                   performed_by: int = None, old_values: Dict = None,
                   new_values: Dict = None, ip_address: str = None):
        """
        Log administrative action to audit log
        
        Args:
            user_id: User being modified
            action_type: Type of action performed
            description: Description of the action
            performed_by: Admin user who performed the action
            old_values: Previous values (for updates)
            new_values: New values (for updates)
            ip_address: IP address of the action
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            is_mysql = self._is_mysql(conn)
            placeholder = self._get_placeholder()
            
            if is_mysql:
                import json
                cursor.execute(f"""
                    INSERT INTO user_audit_log 
                    (user_id, action_type, action_description, performed_by, ip_address, old_values, new_values)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
                """, (user_id, action_type, description, performed_by, ip_address,
                      json.dumps(old_values) if old_values else None,
                      json.dumps(new_values) if new_values else None))
            else:
                cursor.execute(f"""
                    INSERT INTO user_audit_log 
                    (user_id, action_type, action_description, performed_by, ip_address, old_values, new_values)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
                """, (user_id, action_type, description, performed_by, ip_address,
                      str(old_values) if old_values else None,
                      str(new_values) if new_values else None))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            logger.info(f"Audit log: {action_type} for user {user_id} by admin {performed_by}")
            
        except Exception as e:
            logger.error(f"Error logging audit: {e}")
    
    def create_user(self, username: str, email: str, password: str,
                   role: str = 'user', privilege_level: str = 'limited',
                   company: str = None, position: str = None,
                   department: str = None, employee_id: str = None,
                   phone_number: str = None, manager_id: int = None,
                   created_by: int = None, auto_activate: bool = False) -> Optional[int]:
        """
        Create a new user with comprehensive details
        
        Args:
            username: Unique username
            email: Email address
            password: Plain text password
            role: User role
            privilege_level: Privilege level
            company: Company name
            position: Job position/title
            department: Department name
            employee_id: Employee ID
            phone_number: Contact phone number
            manager_id: Manager's user ID
            created_by: Admin user ID who created this user
            auto_activate: Whether to activate user immediately
            
        Returns:
            User ID if successful, None otherwise
        """
        try:
            # Validate inputs
            if not username or len(username) < 3:
                logger.error("Username must be at least 3 characters")
                return None
            
            if not self._validate_email(email):
                logger.error("Invalid email format")
                return None
            
            is_valid, error_msg = self._validate_password(password)
            if not is_valid:
                logger.error(f"Password validation failed: {error_msg}")
                return None
            
            # Check if user already exists
            if self.get_user_by_username(username):
                logger.error(f"Username '{username}' already exists")
                return None
            
            if self.get_user_by_email(email):
                logger.error(f"Email '{email}' already exists")
                return None
            
            # Hash password
            password_hash, salt = self._hash_password(password)
            
            # Create user
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            status = 'active' if auto_activate else 'pending_activation'
            
            cursor.execute(f"""
                INSERT INTO admin_users 
                (username, email, password_hash, salt, role, privilege_level, status,
                 company, position, department, employee_id, phone_number, manager_id, created_by)
                VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder},
                        {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder},
                        {placeholder}, {placeholder}, {placeholder}, {placeholder})
            """, (username, email, password_hash, salt, role, privilege_level, status,
                  company, position, department, employee_id, phone_number, manager_id, created_by))
            
            user_id = cursor.lastrowid
            
            # Store password in history
            cursor.execute(f"""
                INSERT INTO password_history (user_id, password_hash, salt)
                VALUES ({placeholder}, {placeholder}, {placeholder})
            """, (user_id, password_hash, salt))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='USER_CREATED',
                description=f"User '{username}' created with role '{role}'",
                performed_by=created_by,
                new_values={
                    'username': username,
                    'email': email,
                    'role': role,
                    'company': company,
                    'position': position
                }
            )
            
            logger.info(f"User '{username}' (ID: {user_id}) created successfully")
            return user_id
            
        except Exception as e:
            logger.error(f"Error creating user: {e}")
            return None
    
    def get_user_by_id(self, user_id: int) -> Optional[AdminUser]:
        """Get user by ID"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, role, privilege_level, status,
                       company, position, department, manager_id, created_date, last_login,
                       suspension_date, suspension_reason, suspension_end_date, created_by,
                       phone_number, employee_id
                FROM admin_users WHERE id = {placeholder}
            """, (user_id,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return AdminUser(*result)
            return None
            
        except Exception as e:
            logger.error(f"Error getting user by ID: {e}")
            return None
    
    def get_user_by_username(self, username: str) -> Optional[AdminUser]:
        """Get user by username"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, role, privilege_level, status,
                       company, position, department, manager_id, created_date, last_login,
                       suspension_date, suspension_reason, suspension_end_date, created_by,
                       phone_number, employee_id
                FROM admin_users WHERE username = {placeholder}
            """, (username,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return AdminUser(*result)
            return None
            
        except Exception as e:
            logger.error(f"Error getting user by username: {e}")
            return None
    
    def get_user_by_email(self, email: str) -> Optional[AdminUser]:
        """Get user by email"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, role, privilege_level, status,
                       company, position, department, manager_id, created_date, last_login,
                       suspension_date, suspension_reason, suspension_end_date, created_by,
                       phone_number, employee_id
                FROM admin_users WHERE email = {placeholder}
            """, (email,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return AdminUser(*result)
            return None
            
        except Exception as e:
            logger.error(f"Error getting user by email: {e}")
            return None
    
    def get_all_users(self, status_filter: str = None, 
                     company_filter: str = None,
                     role_filter: str = None) -> List[AdminUser]:
        """
        Get all users with optional filters
        
        Args:
            status_filter: Filter by status
            company_filter: Filter by company
            role_filter: Filter by role
            
        Returns:
            List of AdminUser objects
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            query = """
                SELECT id, username, email, role, privilege_level, status,
                       company, position, department, manager_id, created_date, last_login,
                       suspension_date, suspension_reason, suspension_end_date, created_by,
                       phone_number, employee_id
                FROM admin_users WHERE 1=1
            """
            params = []
            placeholder = self._get_placeholder()
            
            if status_filter:
                query += f" AND status = {placeholder}"
                params.append(status_filter)
            
            if company_filter:
                query += f" AND company = {placeholder}"
                params.append(company_filter)
            
            if role_filter:
                query += f" AND role = {placeholder}"
                params.append(role_filter)
            
            query += " ORDER BY username"
            
            cursor.execute(query, tuple(params))
            
            users = []
            for row in cursor.fetchall():
                users.append(AdminUser(*row))
            
            cursor.close()
            conn.close()
            
            return users
            
        except Exception as e:
            logger.error(f"Error getting all users: {e}")
            return []
    
    def update_user(self, user_id: int, modified_by: int = None, **kwargs) -> bool:
        """
        Update user fields
        
        Args:
            user_id: User ID to update
            modified_by: Admin user ID performing the update
            **kwargs: Fields to update (username, email, role, privilege_level, 
                     company, position, department, etc.)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get current user data for audit
            old_user = self.get_user_by_id(user_id)
            if not old_user:
                logger.error(f"User {user_id} not found")
                return False
            
            # Build update query
            valid_fields = ['username', 'email', 'role', 'privilege_level', 'status',
                          'company', 'position', 'department', 'phone_number',
                          'employee_id', 'manager_id', 'notes']
            
            update_fields = []
            params = []
            placeholder = self._get_placeholder()
            
            for field, value in kwargs.items():
                if field in valid_fields:
                    update_fields.append(f"{field} = {placeholder}")
                    params.append(value)
            
            if not update_fields:
                logger.warning("No valid fields to update")
                return False
            
            # Add modified_by and modified_date
            is_mysql = True
            try:
                conn = self.db_conn.connect()
                is_mysql = self._is_mysql(conn)
            except:
                pass
            
            if is_mysql:
                update_fields.append(f"modified_by = {placeholder}")
                update_fields.append("modified_date = NOW()")
            else:
                update_fields.append(f"modified_by = {placeholder}")
                update_fields.append("modified_date = datetime('now')")
            
            params.append(modified_by)
            params.append(user_id)
            
            # Execute update
            cursor = conn.cursor()
            query = f"""
                UPDATE admin_users 
                SET {', '.join(update_fields)}
                WHERE id = {placeholder}
            """
            
            cursor.execute(query, tuple(params))
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='USER_UPDATED',
                description=f"User '{old_user.username}' updated",
                performed_by=modified_by,
                old_values=old_user.to_dict(),
                new_values=kwargs
            )
            
            logger.info(f"User {user_id} updated successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error updating user: {e}")
            return False
    
    def change_password(self, user_id: int, new_password: str,
                       changed_by: int = None) -> bool:
        """
        Change user password
        
        Args:
            user_id: User ID
            new_password: New password
            changed_by: Admin user ID performing the change
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Validate password
            is_valid, error_msg = self._validate_password(new_password)
            if not is_valid:
                logger.error(f"Password validation failed: {error_msg}")
                return False
            
            # Hash new password
            password_hash, salt = self._hash_password(new_password)
            
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            is_mysql = self._is_mysql(conn)
            
            # Update password
            if is_mysql:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET password_hash = {placeholder}, salt = {placeholder}, 
                        last_password_change = NOW(),
                        modified_by = {placeholder}, modified_date = NOW()
                    WHERE id = {placeholder}
                """, (password_hash, salt, changed_by, user_id))
            else:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET password_hash = {placeholder}, salt = {placeholder},
                        last_password_change = datetime('now'),
                        modified_by = {placeholder}, modified_date = datetime('now')
                    WHERE id = {placeholder}
                """, (password_hash, salt, changed_by, user_id))
            
            # Store in password history
            cursor.execute(f"""
                INSERT INTO password_history (user_id, password_hash, salt)
                VALUES ({placeholder}, {placeholder}, {placeholder})
            """, (user_id, password_hash, salt))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='PASSWORD_CHANGED',
                description=f"Password changed for user {user_id}",
                performed_by=changed_by
            )
            
            logger.info(f"Password changed for user {user_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error changing password: {e}")
            return False
    
    def suspend_user(self, user_id: int, reason: str,
                    suspended_by: int = None,
                    suspension_end_date: datetime = None) -> bool:
        """
        Suspend user account
        
        Args:
            user_id: User ID to suspend
            reason: Reason for suspension
            suspended_by: Admin user ID performing the suspension
            suspension_end_date: Optional end date for suspension
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            is_mysql = self._is_mysql(conn)
            
            if is_mysql:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'suspended',
                        suspension_date = NOW(),
                        suspension_reason = {placeholder},
                        suspension_end_date = {placeholder},
                        suspended_by = {placeholder},
                        modified_by = {placeholder},
                        modified_date = NOW()
                    WHERE id = {placeholder}
                """, (reason, suspension_end_date, suspended_by, suspended_by, user_id))
            else:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'suspended',
                        suspension_date = datetime('now'),
                        suspension_reason = {placeholder},
                        suspension_end_date = {placeholder},
                        suspended_by = {placeholder},
                        modified_by = {placeholder},
                        modified_date = datetime('now')
                    WHERE id = {placeholder}
                """, (reason, suspension_end_date, suspended_by, suspended_by, user_id))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='USER_SUSPENDED',
                description=f"User suspended: {reason}",
                performed_by=suspended_by,
                new_values={'reason': reason, 'end_date': str(suspension_end_date)}
            )
            
            logger.info(f"User {user_id} suspended by {suspended_by}")
            return True
            
        except Exception as e:
            logger.error(f"Error suspending user: {e}")
            return False
    
    def activate_user(self, user_id: int, activated_by: int = None) -> bool:
        """
        Activate or reactivate user account
        
        Args:
            user_id: User ID to activate
            activated_by: Admin user ID performing the activation
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            is_mysql = self._is_mysql(conn)
            
            if is_mysql:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'active',
                        suspension_date = NULL,
                        suspension_reason = NULL,
                        suspension_end_date = NULL,
                        suspended_by = NULL,
                        modified_by = {placeholder},
                        modified_date = NOW()
                    WHERE id = {placeholder}
                """, (activated_by, user_id))
            else:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'active',
                        suspension_date = NULL,
                        suspension_reason = NULL,
                        suspension_end_date = NULL,
                        suspended_by = NULL,
                        modified_by = {placeholder},
                        modified_date = datetime('now')
                    WHERE id = {placeholder}
                """, (activated_by, user_id))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='USER_ACTIVATED',
                description=f"User activated",
                performed_by=activated_by
            )
            
            logger.info(f"User {user_id} activated by {activated_by}")
            return True
            
        except Exception as e:
            logger.error(f"Error activating user: {e}")
            return False
    
    def deactivate_user(self, user_id: int, deactivated_by: int = None) -> bool:
        """
        Permanently deactivate user account
        
        Args:
            user_id: User ID to deactivate
            deactivated_by: Admin user ID performing the deactivation
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            is_mysql = self._is_mysql(conn)
            
            if is_mysql:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'deactivated',
                        modified_by = {placeholder},
                        modified_date = NOW()
                    WHERE id = {placeholder}
                """, (deactivated_by, user_id))
            else:
                cursor.execute(f"""
                    UPDATE admin_users 
                    SET status = 'deactivated',
                        modified_by = {placeholder},
                        modified_date = datetime('now')
                    WHERE id = {placeholder}
                """, (deactivated_by, user_id))
            
            # Deactivate all sessions
            cursor.execute(f"""
                UPDATE admin_user_sessions 
                SET is_active = FALSE 
                WHERE user_id = {placeholder}
            """, (user_id,))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='USER_DEACTIVATED',
                description=f"User permanently deactivated",
                performed_by=deactivated_by
            )
            
            logger.info(f"User {user_id} deactivated by {deactivated_by}")
            return True
            
        except Exception as e:
            logger.error(f"Error deactivating user: {e}")
            return False
    
    def delete_user(self, user_id: int, deleted_by: int = None) -> bool:
        """
        Delete user (use with caution - prefer deactivation)
        
        Args:
            user_id: User ID to delete
            deleted_by: Admin user ID performing the deletion
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Get user data for audit
            user = self.get_user_by_id(user_id)
            if not user:
                return False
            
            # Log audit before deletion
            self._log_audit(
                user_id=user_id,
                action_type='USER_DELETED',
                description=f"User '{user.username}' deleted",
                performed_by=deleted_by,
                old_values=user.to_dict()
            )
            
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                DELETE FROM admin_users WHERE id = {placeholder}
            """, (user_id,))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            logger.warning(f"User {user_id} deleted by {deleted_by}")
            return True
            
        except Exception as e:
            logger.error(f"Error deleting user: {e}")
            return False
    
    def get_users_by_company(self, company: str) -> List[AdminUser]:
        """Get all users in a specific company"""
        return self.get_all_users(company_filter=company)
    
    def get_users_by_status(self, status: str) -> List[AdminUser]:
        """Get all users with a specific status"""
        return self.get_all_users(status_filter=status)
    
    def get_suspended_users(self) -> List[AdminUser]:
        """Get all suspended users"""
        return self.get_users_by_status('suspended')
    
    def get_audit_log(self, user_id: int = None, limit: int = 100) -> List[Dict]:
        """
        Get audit log entries
        
        Args:
            user_id: Filter by user ID (None for all)
            limit: Maximum number of entries
            
        Returns:
            List of audit log entries
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            
            if user_id:
                query = f"""
                    SELECT al.id, al.user_id, u.username, al.action_type, 
                           al.action_description, al.performed_by, p.username as performed_by_name,
                           al.ip_address, al.timestamp
                    FROM user_audit_log al
                    JOIN admin_users u ON al.user_id = u.id
                    LEFT JOIN admin_users p ON al.performed_by = p.id
                    WHERE al.user_id = {placeholder}
                    ORDER BY al.timestamp DESC
                    LIMIT {placeholder}
                """
                cursor.execute(query, (user_id, limit))
            else:
                query = f"""
                    SELECT al.id, al.user_id, u.username, al.action_type,
                           al.action_description, al.performed_by, p.username as performed_by_name,
                           al.ip_address, al.timestamp
                    FROM user_audit_log al
                    JOIN admin_users u ON al.user_id = u.id
                    LEFT JOIN admin_users p ON al.performed_by = p.id
                    ORDER BY al.timestamp DESC
                    LIMIT {placeholder}
                """
                cursor.execute(query, (limit,))
            
            logs = []
            for row in cursor.fetchall():
                logs.append({
                    'id': row[0],
                    'user_id': row[1],
                    'username': row[2],
                    'action_type': row[3],
                    'description': row[4],
                    'performed_by': row[5],
                    'performed_by_name': row[6],
                    'ip_address': row[7],
                    'timestamp': row[8]
                })
            
            cursor.close()
            conn.close()
            
            return logs
            
        except Exception as e:
            logger.error(f"Error getting audit log: {e}")
            return []
    
    def get_user_statistics(self) -> Dict[str, Any]:
        """
        Get comprehensive user statistics
        
        Returns:
            Dictionary with various statistics
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            stats = {}
            
            # Total users
            cursor.execute("SELECT COUNT(*) FROM admin_users")
            stats['total_users'] = cursor.fetchone()[0]
            
            # Users by status
            cursor.execute("""
                SELECT status, COUNT(*) 
                FROM admin_users 
                GROUP BY status
            """)
            stats['by_status'] = {row[0]: row[1] for row in cursor.fetchall()}
            
            # Users by role
            cursor.execute("""
                SELECT role, COUNT(*) 
                FROM admin_users 
                GROUP BY role
            """)
            stats['by_role'] = {row[0]: row[1] for row in cursor.fetchall()}
            
            # Users by company
            cursor.execute("""
                SELECT company, COUNT(*) 
                FROM admin_users 
                WHERE company IS NOT NULL
                GROUP BY company
            """)
            stats['by_company'] = {row[0]: row[1] for row in cursor.fetchall()}
            
            # Recently created users (last 30 days)
            is_mysql = self._is_mysql(conn)
            if is_mysql:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM admin_users 
                    WHERE created_date >= DATE_SUB(NOW(), INTERVAL 30 DAY)
                """)
            else:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM admin_users 
                    WHERE created_date >= datetime('now', '-30 days')
                """)
            stats['recently_created'] = cursor.fetchone()[0]
            
            # Active sessions
            cursor.execute("""
                SELECT COUNT(*) 
                FROM admin_user_sessions 
                WHERE is_active = TRUE
            """)
            stats['active_sessions'] = cursor.fetchone()[0]
            
            cursor.close()
            conn.close()
            
            return stats
            
        except Exception as e:
            logger.error(f"Error getting user statistics: {e}")
            return {}
    
    def grant_privilege(self, user_id: int, privilege_name: str,
                       granted_by: int = None, expires_date: datetime = None) -> bool:
        """
        Grant a specific privilege to a user
        
        Args:
            user_id: User ID
            privilege_name: Name of privilege to grant
            granted_by: Admin user ID granting the privilege
            expires_date: Optional expiration date
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            
            # Check if privilege already exists
            cursor.execute(f"""
                SELECT id FROM user_privileges 
                WHERE user_id = {placeholder} AND privilege_name = {placeholder}
            """, (user_id, privilege_name))
            
            if cursor.fetchone():
                # Update existing privilege
                cursor.execute(f"""
                    UPDATE user_privileges 
                    SET privilege_value = TRUE, granted_by = {placeholder},
                        granted_date = CURRENT_TIMESTAMP, expires_date = {placeholder}
                    WHERE user_id = {placeholder} AND privilege_name = {placeholder}
                """, (granted_by, expires_date, user_id, privilege_name))
            else:
                # Insert new privilege
                cursor.execute(f"""
                    INSERT INTO user_privileges 
                    (user_id, privilege_name, privilege_value, granted_by, expires_date)
                    VALUES ({placeholder}, {placeholder}, TRUE, {placeholder}, {placeholder})
                """, (user_id, privilege_name, granted_by, expires_date))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='PRIVILEGE_GRANTED',
                description=f"Privilege '{privilege_name}' granted",
                performed_by=granted_by
            )
            
            logger.info(f"Privilege '{privilege_name}' granted to user {user_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error granting privilege: {e}")
            return False
    
    def revoke_privilege(self, user_id: int, privilege_name: str,
                        revoked_by: int = None) -> bool:
        """
        Revoke a specific privilege from a user
        
        Args:
            user_id: User ID
            privilege_name: Name of privilege to revoke
            revoked_by: Admin user ID revoking the privilege
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                UPDATE user_privileges 
                SET privilege_value = FALSE 
                WHERE user_id = {placeholder} AND privilege_name = {placeholder}
            """, (user_id, privilege_name))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Log audit
            self._log_audit(
                user_id=user_id,
                action_type='PRIVILEGE_REVOKED',
                description=f"Privilege '{privilege_name}' revoked",
                performed_by=revoked_by
            )
            
            logger.info(f"Privilege '{privilege_name}' revoked from user {user_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error revoking privilege: {e}")
            return False
    
    def get_user_privileges(self, user_id: int) -> List[Dict[str, Any]]:
        """
        Get all privileges for a user
        
        Args:
            user_id: User ID
            
        Returns:
            List of privilege dictionaries
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT privilege_name, privilege_value, granted_date, expires_date
                FROM user_privileges
                WHERE user_id = {placeholder}
                ORDER BY privilege_name
            """, (user_id,))
            
            privileges = []
            for row in cursor.fetchall():
                privileges.append({
                    'name': row[0],
                    'value': row[1],
                    'granted_date': row[2],
                    'expires_date': row[3]
                })
            
            cursor.close()
            conn.close()
            
            return privileges
            
        except Exception as e:
            logger.error(f"Error getting user privileges: {e}")
            return []


# Standalone utility functions
def create_super_admin(username: str = 'superadmin', 
                      email: str = 'superadmin@workflow.local',
                      password: str = 'SuperAdmin123!') -> bool:
    """
    Create a super admin user for initial setup
    
    Args:
        username: Super admin username
        email: Super admin email
        password: Super admin password
        
    Returns:
        True if successful, False otherwise
    """
    try:
        admin = UserAdministration()
        
        # Check if super admin already exists
        if admin.get_user_by_username(username):
            logger.info("Super admin already exists")
            return True
        
        user_id = admin.create_user(
            username=username,
            email=email,
            password=password,
            role='super_admin',
            privilege_level='full_access',
            auto_activate=True
        )
        
        if user_id:
            logger.info(f"Super admin created with ID: {user_id}")
            return True
        else:
            logger.error("Failed to create super admin")
            return False
            
    except Exception as e:
        logger.error(f"Error creating super admin: {e}")
        return False


if __name__ == "__main__":
    """
    Main entry point for user administration CLI
    """
    import sys
    
    print("=" * 60)
    print("  USER ADMINISTRATION MODULE")
    print("  WARNING: Administrative Access Only")
    print("=" * 60)
    print()
    
    # Create super admin if it doesn't exist
    print("Checking for super admin account...")
    if create_super_admin():
        print("✓ Super admin account ready")
        print("  Username: superadmin")
        print("  Password: SuperAdmin123!")
        print()
        print("Please change the default password immediately!")
    else:
        print("✗ Failed to create super admin")
        sys.exit(1)
    
    print()
    print("User Administration Module initialized successfully.")
    print("This module provides comprehensive user management capabilities.")
    print()
    print("For GUI administration, run: python user_admin_gui.py")

