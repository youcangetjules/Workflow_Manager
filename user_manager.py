"""
User Management Module
=====================

This module handles user authentication, session management, and user tracking
for the Degrow Workflow Manager system.

Author: Workflow Manager System
Version: 1.0.0
"""

import hashlib
import secrets
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Tuple
import json
import os
from db_config import get_database_connection


class User:
    """User data class"""
    def __init__(self, user_id: int, username: str, email: str, role: str, 
                 is_active: bool = True, created_date: datetime = None, 
                 last_login: datetime = None):
        self.user_id = user_id
        self.username = username
        self.email = email
        self.role = role
        self.is_active = is_active
        self.created_date = created_date or datetime.now()
        self.last_login = last_login


class UserSession:
    """User session data class"""
    def __init__(self, session_id: str, user_id: int, login_time: datetime, 
                 last_activity: datetime, ip_address: str = None):
        self.session_id = session_id
        self.user_id = user_id
        self.login_time = login_time
        self.last_activity = last_activity
        self.ip_address = ip_address


class UserManager:
    """User management and authentication system"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        Initialize User Manager
        
        Args:
            config_file: Path to database configuration file
        """
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        self.current_user: Optional[User] = None
        self.current_session: Optional[UserSession] = None
        
        # Initialize database tables
        self._create_user_tables()
    
    def _create_user_tables(self):
        """Create user management tables"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Check if this is MySQL or SQLite
            is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
            
            if is_mysql:
                # MySQL syntax
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS users (
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
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
            else:
                # SQLite syntax
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT NOT NULL UNIQUE,
                        email TEXT NOT NULL UNIQUE,
                        password_hash TEXT NOT NULL,
                        salt TEXT NOT NULL,
                        role TEXT DEFAULT 'user' CHECK(role IN ('admin', 'manager', 'user')),
                        is_active BOOLEAN DEFAULT 1,
                        created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                        last_login TEXT NULL,
                        failed_login_attempts INTEGER DEFAULT 0,
                        locked_until TEXT NULL
                    )
                """)
            
            # Create user_sessions table
            if is_mysql:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_sessions (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        session_id VARCHAR(64) NOT NULL UNIQUE,
                        user_id INT NOT NULL,
                        login_time DATETIME DEFAULT CURRENT_TIMESTAMP,
                        last_activity DATETIME DEFAULT CURRENT_TIMESTAMP,
                        ip_address VARCHAR(45),
                        is_active BOOLEAN DEFAULT TRUE,
                        FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
            else:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_sessions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        session_id TEXT NOT NULL UNIQUE,
                        user_id INTEGER NOT NULL,
                        login_time TEXT DEFAULT CURRENT_TIMESTAMP,
                        last_activity TEXT DEFAULT CURRENT_TIMESTAMP,
                        ip_address TEXT,
                        is_active BOOLEAN DEFAULT 1,
                        FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                    )
                """)
            
            # Create user_activity_log table
            if is_mysql:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_activity_log (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        user_id INT NOT NULL,
                        session_id VARCHAR(64),
                        activity_type VARCHAR(50) NOT NULL,
                        activity_description TEXT,
                        ip_address VARCHAR(45),
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """)
            else:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS user_activity_log (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER NOT NULL,
                        session_id TEXT,
                        activity_type TEXT NOT NULL,
                        activity_description TEXT,
                        ip_address TEXT,
                        timestamp TEXT DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                    )
                """)
            
            # Create indexes for performance
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sessions_user_id ON user_sessions(user_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sessions_active ON user_sessions(is_active)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_user_id ON user_activity_log(user_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_timestamp ON user_activity_log(timestamp)")
            
            conn.commit()
            cursor.close()
            conn.close()
            
            print("User management tables created successfully")
            
        except Exception as e:
            print(f"Error creating user tables: {e}")
    
    def _get_placeholder(self):
        """Get the correct placeholder for the database type"""
        try:
            conn = self.db_conn.connect()
            is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
            conn.close()
            return '%s' if is_mysql else '?'
        except:
            return '?'  # Default to SQLite placeholder
    
    def _hash_password(self, password: str, salt: str = None) -> Tuple[str, str]:
        """
        Hash password with salt
        
        Args:
            password: Plain text password
            salt: Salt string (generated if None)
            
        Returns:
            Tuple of (hashed_password, salt)
        """
        if salt is None:
            salt = secrets.token_hex(16)
        
        # Use PBKDF2 for password hashing
        password_hash = hashlib.pbkdf2_hmac('sha256', 
                                          password.encode('utf-8'), 
                                          salt.encode('utf-8'), 
                                          100000)
        return password_hash.hex(), salt
    
    def create_user(self, username: str, email: str, password: str, 
                   role: str = 'user') -> bool:
        """
        Create a new user
        
        Args:
            username: Username
            email: Email address
            password: Plain text password
            role: User role (admin, manager, user)
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if user already exists
            if self.get_user_by_username(username) or self.get_user_by_email(email):
                return False
            
            # Hash password
            password_hash, salt = self._hash_password(password)
            
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                INSERT INTO users (username, email, password_hash, salt, role)
                VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
            """, (username, email, password_hash, salt, role))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            print(f"User {username} created successfully")
            return True
            
        except Exception as e:
            print(f"Error creating user: {e}")
            return False
    
    def authenticate_user(self, username: str, password: str, 
                         ip_address: str = None) -> Optional[User]:
        """
        Authenticate user with username and password
        
        Args:
            username: Username or email
            password: Plain text password
            ip_address: Client IP address
            
        Returns:
            User object if successful, None otherwise
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Check if user is locked
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, password_hash, salt, role, is_active,
                       failed_login_attempts, locked_until
                FROM users 
                WHERE (username = {placeholder} OR email = {placeholder}) AND is_active = TRUE
            """, (username, username))
            
            result = cursor.fetchone()
            if not result:
                cursor.close()
                conn.close()
                return None
            
            user_id, db_username, email, password_hash, salt, role, is_active, \
            failed_attempts, locked_until = result
            
            # Check if account is locked
            if locked_until and datetime.now() < locked_until:
                cursor.close()
                conn.close()
                return None
            
            # Verify password
            hashed_password, _ = self._hash_password(password, salt)
            if hashed_password != password_hash:
                # Increment failed login attempts
                placeholder = self._get_placeholder()
                is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
                
                if is_mysql:
                    cursor.execute(f"""
                        UPDATE users 
                        SET failed_login_attempts = failed_login_attempts + 1,
                            locked_until = CASE 
                                WHEN failed_login_attempts >= 4 THEN DATE_ADD(NOW(), INTERVAL 15 MINUTE)
                                ELSE locked_until
                            END
                        WHERE id = {placeholder}
                    """, (user_id,))
                else:
                    # SQLite version
                    cursor.execute(f"""
                        UPDATE users 
                        SET failed_login_attempts = failed_login_attempts + 1,
                            locked_until = CASE 
                                WHEN failed_login_attempts >= 4 THEN datetime('now', '+15 minutes')
                                ELSE locked_until
                            END
                        WHERE id = {placeholder}
                    """, (user_id,))
                conn.commit()
                cursor.close()
                conn.close()
                return None
            
            # Reset failed login attempts on successful login
            placeholder = self._get_placeholder()
            is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
            
            if is_mysql:
                cursor.execute(f"""
                    UPDATE users 
                    SET failed_login_attempts = 0, 
                        locked_until = NULL,
                        last_login = NOW()
                    WHERE id = {placeholder}
                """, (user_id,))
            else:
                # SQLite version
                cursor.execute(f"""
                    UPDATE users 
                    SET failed_login_attempts = 0, 
                        locked_until = NULL,
                        last_login = datetime('now')
                    WHERE id = {placeholder}
                """, (user_id,))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Create user object
            user = User(user_id, db_username, email, role, is_active)
            user.last_login = datetime.now()
            
            # Create session
            self._create_session(user, ip_address)
            
            # Log successful login
            self.log_activity(user.user_id, 'login', f'User {username} logged in', ip_address)
            
            return user
            
        except Exception as e:
            print(f"Error authenticating user: {e}")
            return None
    
    def _create_session(self, user: User, ip_address: str = None) -> str:
        """
        Create a new user session
        
        Args:
            user: User object
            ip_address: Client IP address
            
        Returns:
            Session ID
        """
        session_id = secrets.token_urlsafe(32)
        now = datetime.now()
        
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Deactivate old sessions for this user
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                UPDATE user_sessions 
                SET is_active = FALSE 
                WHERE user_id = {placeholder} AND is_active = TRUE
            """, (user.user_id,))
            
            # Create new session
            is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
            
            if is_mysql:
                cursor.execute(f"""
                    INSERT INTO user_sessions (session_id, user_id, login_time, last_activity, ip_address)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
                """, (session_id, user.user_id, now, now, ip_address))
            else:
                # SQLite version - convert datetime to string
                now_str = now.strftime('%Y-%m-%d %H:%M:%S')
                cursor.execute(f"""
                    INSERT INTO user_sessions (session_id, user_id, login_time, last_activity, ip_address)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
                """, (session_id, user.user_id, now_str, now_str, ip_address))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            # Store current session
            self.current_user = user
            self.current_session = UserSession(session_id, user.user_id, now, now, ip_address)
            
            return session_id
            
        except Exception as e:
            print(f"Error creating session: {e}")
            return None
    
    def logout(self):
        """Logout current user"""
        if self.current_session:
            try:
                conn = self.db_conn.connect()
                cursor = conn.cursor()
                
                # Deactivate session
                placeholder = self._get_placeholder()
                cursor.execute(f"""
                    UPDATE user_sessions 
                    SET is_active = FALSE 
                    WHERE session_id = {placeholder}
                """, (self.current_session.session_id,))
                
                # Log logout activity
                self.log_activity(self.current_user.user_id, 'logout', 
                                f'User {self.current_user.username} logged out')
                
                conn.commit()
                cursor.close()
                conn.close()
                
            except Exception as e:
                print(f"Error during logout: {e}")
        
        self.current_user = None
        self.current_session = None
    
    def log_activity(self, user_id: int, activity_type: str, 
                    description: str, ip_address: str = None):
        """
        Log user activity
        
        Args:
            user_id: User ID
            activity_type: Type of activity
            description: Activity description
            ip_address: Client IP address
        """
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            session_id = self.current_session.session_id if self.current_session else None
            
            placeholder = self._get_placeholder()
            is_mysql = hasattr(conn, 'server_version') or 'mysql' in str(type(conn)).lower()
            
            if is_mysql:
                cursor.execute(f"""
                    INSERT INTO user_activity_log (user_id, session_id, activity_type, 
                                                 activity_description, ip_address)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
                """, (user_id, session_id, activity_type, description, ip_address))
            else:
                # SQLite version - use current timestamp
                cursor.execute(f"""
                    INSERT INTO user_activity_log (user_id, session_id, activity_type, 
                                                 activity_description, ip_address, timestamp)
                    VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, datetime('now'))
                """, (user_id, session_id, activity_type, description, ip_address))
            
            conn.commit()
            cursor.close()
            conn.close()
            
        except Exception as e:
            print(f"Error logging activity: {e}")
    
    def get_user_by_username(self, username: str) -> Optional[User]:
        """Get user by username"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, role, is_active, created_date, last_login
                FROM users WHERE username = {placeholder}
            """, (username,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return User(*result)
            return None
            
        except Exception as e:
            print(f"Error getting user by username: {e}")
            return None
    
    def get_user_by_email(self, email: str) -> Optional[User]:
        """Get user by email"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                SELECT id, username, email, role, is_active, created_date, last_login
                FROM users WHERE email = {placeholder}
            """, (email,))
            
            result = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if result:
                return User(*result)
            return None
            
        except Exception as e:
            print(f"Error getting user by email: {e}")
            return None
    
    def get_all_users(self) -> List[User]:
        """Get all users"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT id, username, email, role, is_active, created_date, last_login
                FROM users ORDER BY username
            """)
            
            users = []
            for row in cursor.fetchall():
                users.append(User(*row))
            
            cursor.close()
            conn.close()
            
            return users
            
        except Exception as e:
            print(f"Error getting all users: {e}")
            return []
    
    def update_user_role(self, user_id: int, new_role: str) -> bool:
        """Update user role"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                UPDATE users SET role = {placeholder} WHERE id = {placeholder}
            """, (new_role, user_id))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            return True
            
        except Exception as e:
            print(f"Error updating user role: {e}")
            return False
    
    def deactivate_user(self, user_id: int) -> bool:
        """Deactivate user account"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            placeholder = self._get_placeholder()
            cursor.execute(f"""
                UPDATE users SET is_active = FALSE WHERE id = {placeholder}
            """, (user_id,))
            
            conn.commit()
            cursor.close()
            conn.close()
            
            return True
            
        except Exception as e:
            print(f"Error deactivating user: {e}")
            return False
    
    def get_user_activity_log(self, user_id: int = None, 
                             limit: int = 100) -> List[Dict]:
        """Get user activity log"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            if user_id:
                placeholder = self._get_placeholder()
                cursor.execute(f"""
                    SELECT u.username, al.activity_type, al.activity_description, 
                           al.ip_address, al.timestamp
                    FROM user_activity_log al
                    JOIN users u ON al.user_id = u.id
                    WHERE al.user_id = {placeholder}
                    ORDER BY al.timestamp DESC
                    LIMIT {placeholder}
                """, (user_id, limit))
            else:
                placeholder = self._get_placeholder()
                cursor.execute(f"""
                    SELECT u.username, al.activity_type, al.activity_description, 
                           al.ip_address, al.timestamp
                    FROM user_activity_log al
                    JOIN users u ON al.user_id = u.id
                    ORDER BY al.timestamp DESC
                    LIMIT {placeholder}
                """, (limit,))
            
            activities = []
            for row in cursor.fetchall():
                activities.append({
                    'username': row[0],
                    'activity_type': row[1],
                    'description': row[2],
                    'ip_address': row[3],
                    'timestamp': row[4]
                })
            
            cursor.close()
            conn.close()
            
            return activities
            
        except Exception as e:
            print(f"Error getting activity log: {e}")
            return []
    
    def is_authenticated(self) -> bool:
        """Check if user is currently authenticated"""
        return self.current_user is not None and self.current_session is not None
    
    def get_current_user(self) -> Optional[User]:
        """Get current authenticated user"""
        return self.current_user
    
    def has_permission(self, required_role: str) -> bool:
        """Check if current user has required permission"""
        if not self.current_user:
            return False
        
        role_hierarchy = {'user': 1, 'manager': 2, 'admin': 3}
        current_level = role_hierarchy.get(self.current_user.role, 0)
        required_level = role_hierarchy.get(required_role, 0)
        
        return current_level >= required_level


def create_default_admin_user(user_manager: UserManager) -> bool:
    """Create default admin user if no users exist"""
    try:
        users = user_manager.get_all_users()
        if not users:
            return user_manager.create_user('admin', 'admin@workflow.local', 'admin123', 'admin')
        return True
    except Exception as e:
        print(f"Error creating default admin user: {e}")
        return False
