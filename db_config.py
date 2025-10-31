"""
Database Configuration Module
============================

This module handles database configuration and connection management for the workflow system.
Supports both MySQL and SQLite databases with configuration file support.

Author: Workflow Manager System
Version: 1.0.0
"""

import json
import os
import sqlite3
import mysql.connector
from mysql.connector import Error
from typing import Optional, Dict, Any, Union
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DatabaseConfig:
    """Database configuration manager"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        Initialize database configuration
        
        Args:
            config_file: Path to configuration file
        """
        self.config_file = config_file
        self.config = self._load_config()
        self.db_type = self.config.get('database', {}).get('type', 'sqlite')
        
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    logger.info(f"Configuration loaded from {self.config_file}")
                    return config
            else:
                logger.warning(f"Configuration file {self.config_file} not found. Using defaults.")
                return self._get_default_config()
        except Exception as e:
            logger.error(f"Error loading configuration: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration"""
        return {
            "database": {
                "type": "sqlite",
                "host": "localhost",
                "port": 3306,
                "username": "root",
                "password": "",
                "database_name": "degrow_workflow",
                "charset": "utf8mb4",
                "autocommit": True
            },
            "backup": {
                "enabled": False,
                "frequency": "daily",
                "location": "./backups"
            },
            "logging": {
                "level": "INFO",
                "file": "workflow_manager.log"
            }
        }
    
    def get_database_config(self) -> Dict[str, Any]:
        """Get database configuration"""
        return self.config.get('database', {})
    
    def get_backup_config(self) -> Dict[str, Any]:
        """Get backup configuration"""
        return self.config.get('backup', {})
    
    def get_logging_config(self) -> Dict[str, Any]:
        """Get logging configuration"""
        return self.config.get('logging', {})


class DatabaseConnection:
    """Database connection manager"""
    
    def __init__(self, config: DatabaseConfig):
        """
        Initialize database connection manager
        
        Args:
            config: Database configuration instance
        """
        self.config = config
        self.db_config = config.get_database_config()
        self.db_type = self.db_config.get('type', 'sqlite')
        self.connection = None
        
    def connect(self) -> Union[sqlite3.Connection, mysql.connector.MySQLConnection]:
        """
        Establish database connection
        
        Returns:
            Database connection object
        """
        try:
            if self.db_type.lower() == 'mysql':
                return self._connect_mysql()
            else:
                return self._connect_sqlite()
        except Exception as e:
            logger.error(f"Failed to connect to database: {e}")
            raise
    
    def _connect_mysql(self) -> mysql.connector.MySQLConnection:
        """Connect to MySQL database"""
        try:
            connection = mysql.connector.connect(
                host=self.db_config.get('host', 'localhost'),
                port=self.db_config.get('port', 3306),
                user=self.db_config.get('username', 'root'),
                password=self.db_config.get('password', ''),
                database=self.db_config.get('database_name', 'degrow_workflow'),
                charset=self.db_config.get('charset', 'utf8mb4'),
                autocommit=self.db_config.get('autocommit', True)
            )
            
            if connection.is_connected():
                logger.info("Successfully connected to MySQL database")
                return connection
            else:
                raise Error("Failed to establish MySQL connection")
                
        except Error as e:
            logger.error(f"MySQL connection error: {e}")
            raise
    
    def _connect_sqlite(self) -> sqlite3.Connection:
        """Connect to SQLite database (fallback)"""
        try:
            # For backward compatibility, use the old path logic
            db_path = self.db_config.get('database_name', 'degrow_workflow.db')
            if not db_path.endswith('.db'):
                db_path += '.db'
            
            connection = sqlite3.connect(db_path)
            logger.info(f"Successfully connected to SQLite database: {db_path}")
            return connection
            
        except Exception as e:
            logger.error(f"SQLite connection error: {e}")
            raise
    
    def close(self):
        """Close database connection"""
        if self.connection:
            try:
                if hasattr(self.connection, 'close'):
                    self.connection.close()
                    logger.info("Database connection closed")
            except Exception as e:
                logger.error(f"Error closing database connection: {e}")
    
    def get_cursor(self):
        """Get database cursor"""
        if not self.connection:
            self.connection = self.connect()
        return self.connection.cursor()
    
    def execute_query(self, query: str, params: tuple = None) -> list:
        """
        Execute a SELECT query
        
        Args:
            query: SQL query string
            params: Query parameters
            
        Returns:
            Query results
        """
        try:
            cursor = self.get_cursor()
            cursor.execute(query, params or ())
            results = cursor.fetchall()
            cursor.close()
            return results
        except Exception as e:
            logger.error(f"Query execution error: {e}")
            raise
    
    def execute_update(self, query: str, params: tuple = None) -> int:
        """
        Execute an INSERT/UPDATE/DELETE query
        
        Args:
            query: SQL query string
            params: Query parameters
            
        Returns:
            Number of affected rows
        """
        try:
            cursor = self.get_cursor()
            cursor.execute(query, params or ())
            affected_rows = cursor.rowcount
            if not self.db_config.get('autocommit', True):
                self.connection.commit()
            cursor.close()
            return affected_rows
        except Exception as e:
            logger.error(f"Update execution error: {e}")
            if not self.db_config.get('autocommit', True):
                self.connection.rollback()
            raise


def get_database_connection(config_file: str = "config.json") -> DatabaseConnection:
    """
    Get a database connection instance
    
    Args:
        config_file: Path to configuration file
        
    Returns:
        DatabaseConnection instance
    """
    config = DatabaseConfig(config_file)
    return DatabaseConnection(config)


def create_mysql_database(config_file: str = "config.json") -> bool:
    """
    Create MySQL database if it doesn't exist
    
    Args:
        config_file: Path to configuration file
        
    Returns:
        True if successful, False otherwise
    """
    try:
        config = DatabaseConfig(config_file)
        db_config = config.get_database_config()
        
        if db_config.get('type', 'sqlite').lower() != 'mysql':
            logger.info("Not a MySQL configuration, skipping database creation")
            return True
        
        # Connect without specifying database
        connection = mysql.connector.connect(
            host=db_config.get('host', 'localhost'),
            port=db_config.get('port', 3306),
            user=db_config.get('username', 'root'),
            password=db_config.get('password', ''),
            charset=db_config.get('charset', 'utf8mb4')
        )
        
        cursor = connection.cursor()
        database_name = db_config.get('database_name', 'degrow_workflow')
        
        # Create database if it doesn't exist
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{database_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        logger.info(f"MySQL database '{database_name}' created or verified")
        
        cursor.close()
        connection.close()
        return True
        
    except Error as e:
        logger.error(f"Error creating MySQL database: {e}")
        return False
    except Exception as e:
        logger.error(f"Unexpected error creating MySQL database: {e}")
        return False


if __name__ == "__main__":
    # Test database connection
    try:
        db_conn = get_database_connection()
        connection = db_conn.connect()
        print("Database connection test successful!")
        db_conn.close()
    except Exception as e:
        print(f"Database connection test failed: {e}")
