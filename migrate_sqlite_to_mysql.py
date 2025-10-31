#!/usr/bin/env python3
"""
SQLite to MySQL Migration Script
===============================

This script migrates data from existing SQLite databases to MySQL.
It handles the workflow_entries, milestones, subtasks, and stakeholders tables.

Author: Workflow Manager System
Version: 1.0.0
"""

import sqlite3
import mysql.connector
from mysql.connector import Error
import os
import sys
from datetime import datetime
from typing import Dict, List, Any
import json

# Add current directory to path to import db_config
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from db_config import get_database_connection, create_mysql_database


class SQLiteToMySQLMigrator:
    """Handles migration from SQLite to MySQL"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        Initialize migrator
        
        Args:
            config_file: Path to MySQL configuration file
        """
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        self.sqlite_connections = {}
        self.migration_stats = {
            'workflow_entries': 0,
            'milestones': 0,
            'subtasks': 0,
            'stakeholders': 0,
            'date_entries': 0,
            'errors': []
        }
    
    def connect_sqlite_databases(self):
        """Connect to existing SQLite databases"""
        try:
            # Try to find existing SQLite databases
            possible_paths = [
                "TestDatabase.db",
                "TestDatabase_milestones.db",
                r"C:\BASHFlowSandbox\TestDatabase.db",
                r"C:\BASHFlowSandbox\TestDatabase_milestones.db"
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    conn = sqlite3.connect(path)
                    self.sqlite_connections[path] = conn
                    print(f"Connected to SQLite database: {path}")
            
            if not self.sqlite_connections:
                print("No SQLite databases found. Migration will create empty MySQL tables.")
                return True
                
            return True
            
        except Exception as e:
            print(f"Error connecting to SQLite databases: {e}")
            return False
    
    def create_mysql_tables(self):
        """Create MySQL tables if they don't exist"""
        try:
            # Create MySQL database
            create_mysql_database(self.config_file)
            
            # Connect to MySQL
            mysql_conn = self.db_conn.connect()
            cursor = mysql_conn.cursor()
            
            # Create workflow_entries table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS workflow_entries (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    state VARCHAR(50) NOT NULL,
                    clli VARCHAR(20),
                    host_wire_centre VARCHAR(100),
                    lata VARCHAR(10),
                    equipment_type VARCHAR(50),
                    current_milestone VARCHAR(100) NOT NULL,
                    milestone_subtask VARCHAR(100),
                    status VARCHAR(50) NOT NULL,
                    planned_start VARCHAR(20),
                    actual_start VARCHAR(20),
                    duration VARCHAR(20),
                    planned_end VARCHAR(20),
                    actual_end VARCHAR(20),
                    milestone_date VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_update DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create milestones table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS milestones (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL UNIQUE,
                    description TEXT,
                    order_index INT DEFAULT 0,
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create subtasks table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS subtasks (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    milestone_id INT NOT NULL,
                    name VARCHAR(100) NOT NULL,
                    description TEXT,
                    criticality VARCHAR(50) DEFAULT 'Should be Complete',
                    order_index INT DEFAULT 0,
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL,
                    FOREIGN KEY (milestone_id) REFERENCES milestones (id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create stakeholders table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS stakeholders (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    home_state VARCHAR(50) NOT NULL,
                    role VARCHAR(100),
                    team VARCHAR(100),
                    email VARCHAR(255),
                    phone VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            # Create date_entries table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS date_entries (
                    clli VARCHAR(20) PRIMARY KEY,
                    planned_start VARCHAR(20),
                    actual_start VARCHAR(20),
                    planned_end VARCHAR(20),
                    actual_end VARCHAR(20),
                    created_date DATETIME NOT NULL,
                    last_updated DATETIME NOT NULL
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            mysql_conn.commit()
            cursor.close()
            mysql_conn.close()
            
            print("MySQL tables created successfully")
            return True
            
        except Exception as e:
            print(f"Error creating MySQL tables: {e}")
            self.migration_stats['errors'].append(f"MySQL table creation: {e}")
            return False
    
    def migrate_workflow_entries(self):
        """Migrate workflow entries from SQLite to MySQL"""
        try:
            # Find SQLite database with workflow_entries table
            sqlite_conn = None
            for path, conn in self.sqlite_connections.items():
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='workflow_entries'")
                if cursor.fetchone():
                    sqlite_conn = conn
                    break
                cursor.close()
            
            if not sqlite_conn:
                print("No workflow_entries table found in SQLite databases")
                return True
            
            # Get data from SQLite
            cursor = sqlite_conn.cursor()
            cursor.execute("SELECT * FROM workflow_entries")
            rows = cursor.fetchall()
            cursor.close()
            
            if not rows:
                print("No workflow entries to migrate")
                return True
            
            # Get column names
            cursor = sqlite_conn.cursor()
            cursor.execute("PRAGMA table_info(workflow_entries)")
            columns = [col[1] for col in cursor.fetchall()]
            cursor.close()
            
            # Connect to MySQL
            mysql_conn = self.db_conn.connect()
            mysql_cursor = mysql_conn.cursor()
            
            # Insert data into MySQL
            for row in rows:
                try:
                    # Convert row to dict
                    row_dict = dict(zip(columns, row))
                    
                    # Prepare data for MySQL
                    data = (
                        row_dict.get('state', ''),
                        row_dict.get('clli', ''),
                        row_dict.get('host_wire_centre', ''),
                        row_dict.get('lata', ''),
                        row_dict.get('equipment_type', ''),
                        row_dict.get('current_milestone', ''),
                        row_dict.get('milestone_subtask', ''),
                        row_dict.get('status', ''),
                        row_dict.get('planned_start', ''),
                        row_dict.get('actual_start', ''),
                        row_dict.get('duration', ''),
                        row_dict.get('planned_end', ''),
                        row_dict.get('actual_end', ''),
                        row_dict.get('milestone_date', ''),
                        self._convert_datetime(row_dict.get('created_date', '')),
                        self._convert_datetime(row_dict.get('last_update', ''))
                    )
                    
                    mysql_cursor.execute('''
                        INSERT INTO workflow_entries 
                        (state, clli, host_wire_centre, lata, equipment_type, current_milestone, 
                         milestone_subtask, status, planned_start, actual_start, duration, 
                         planned_end, actual_end, milestone_date, created_date, last_update)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ''', data)
                    
                    self.migration_stats['workflow_entries'] += 1
                    
                except Exception as e:
                    error_msg = f"Error migrating workflow entry {row_dict.get('id', 'unknown')}: {e}"
                    print(error_msg)
                    self.migration_stats['errors'].append(error_msg)
            
            mysql_conn.commit()
            mysql_cursor.close()
            mysql_conn.close()
            
            print(f"Migrated {self.migration_stats['workflow_entries']} workflow entries")
            return True
            
        except Exception as e:
            print(f"Error migrating workflow entries: {e}")
            self.migration_stats['errors'].append(f"Workflow entries migration: {e}")
            return False
    
    def migrate_milestones(self):
        """Migrate milestones from SQLite to MySQL"""
        try:
            # Find SQLite database with milestones table
            sqlite_conn = None
            for path, conn in self.sqlite_connections.items():
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='milestones'")
                if cursor.fetchone():
                    sqlite_conn = conn
                    break
                cursor.close()
            
            if not sqlite_conn:
                print("No milestones table found in SQLite databases")
                return True
            
            # Get data from SQLite
            cursor = sqlite_conn.cursor()
            cursor.execute("SELECT * FROM milestones")
            rows = cursor.fetchall()
            cursor.close()
            
            if not rows:
                print("No milestones to migrate")
                return True
            
            # Get column names
            cursor = sqlite_conn.cursor()
            cursor.execute("PRAGMA table_info(milestones)")
            columns = [col[1] for col in cursor.fetchall()]
            cursor.close()
            
            # Connect to MySQL
            mysql_conn = self.db_conn.connect()
            mysql_cursor = mysql_conn.cursor()
            
            # Insert data into MySQL
            for row in rows:
                try:
                    row_dict = dict(zip(columns, row))
                    
                    data = (
                        row_dict.get('name', ''),
                        row_dict.get('description', ''),
                        row_dict.get('order_index', 0),
                        self._convert_datetime(row_dict.get('created_date', '')),
                        self._convert_datetime(row_dict.get('last_updated', ''))
                    )
                    
                    mysql_cursor.execute('''
                        INSERT INTO milestones (name, description, order_index, created_date, last_updated)
                        VALUES (%s, %s, %s, %s, %s)
                    ''', data)
                    
                    self.migration_stats['milestones'] += 1
                    
                except Exception as e:
                    error_msg = f"Error migrating milestone {row_dict.get('id', 'unknown')}: {e}"
                    print(error_msg)
                    self.migration_stats['errors'].append(error_msg)
            
            mysql_conn.commit()
            mysql_cursor.close()
            mysql_conn.close()
            
            print(f"Migrated {self.migration_stats['milestones']} milestones")
            return True
            
        except Exception as e:
            print(f"Error migrating milestones: {e}")
            self.migration_stats['errors'].append(f"Milestones migration: {e}")
            return False
    
    def migrate_subtasks(self):
        """Migrate subtasks from SQLite to MySQL"""
        try:
            # Find SQLite database with subtasks table
            sqlite_conn = None
            for path, conn in self.sqlite_connections.items():
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='subtasks'")
                if cursor.fetchone():
                    sqlite_conn = conn
                    break
                cursor.close()
            
            if not sqlite_conn:
                print("No subtasks table found in SQLite databases")
                return True
            
            # Get data from SQLite
            cursor = sqlite_conn.cursor()
            cursor.execute("SELECT * FROM subtasks")
            rows = cursor.fetchall()
            cursor.close()
            
            if not rows:
                print("No subtasks to migrate")
                return True
            
            # Get column names
            cursor = sqlite_conn.cursor()
            cursor.execute("PRAGMA table_info(subtasks)")
            columns = [col[1] for col in cursor.fetchall()]
            cursor.close()
            
            # Connect to MySQL
            mysql_conn = self.db_conn.connect()
            mysql_cursor = mysql_conn.cursor()
            
            # Insert data into MySQL
            for row in rows:
                try:
                    row_dict = dict(zip(columns, row))
                    
                    data = (
                        row_dict.get('milestone_id', 0),
                        row_dict.get('name', ''),
                        row_dict.get('description', ''),
                        row_dict.get('criticality', 'Should be Complete'),
                        row_dict.get('order_index', 0),
                        self._convert_datetime(row_dict.get('created_date', '')),
                        self._convert_datetime(row_dict.get('last_updated', ''))
                    )
                    
                    mysql_cursor.execute('''
                        INSERT INTO subtasks (milestone_id, name, description, criticality, order_index, created_date, last_updated)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ''', data)
                    
                    self.migration_stats['subtasks'] += 1
                    
                except Exception as e:
                    error_msg = f"Error migrating subtask {row_dict.get('id', 'unknown')}: {e}"
                    print(error_msg)
                    self.migration_stats['errors'].append(error_msg)
            
            mysql_conn.commit()
            mysql_cursor.close()
            mysql_conn.close()
            
            print(f"Migrated {self.migration_stats['subtasks']} subtasks")
            return True
            
        except Exception as e:
            print(f"Error migrating subtasks: {e}")
            self.migration_stats['errors'].append(f"Subtasks migration: {e}")
            return False
    
    def migrate_date_entries(self):
        """Migrate date entries from SQLite to MySQL"""
        try:
            # Find SQLite database with date_entries table
            sqlite_conn = None
            for path, conn in self.sqlite_connections.items():
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='date_entries'")
                if cursor.fetchone():
                    sqlite_conn = conn
                    break
                cursor.close()
            
            if not sqlite_conn:
                print("No date_entries table found in SQLite databases")
                return True
            
            # Get data from SQLite
            cursor = sqlite_conn.cursor()
            cursor.execute("SELECT * FROM date_entries")
            rows = cursor.fetchall()
            cursor.close()
            
            if not rows:
                print("No date entries to migrate")
                return True
            
            # Get column names
            cursor = sqlite_conn.cursor()
            cursor.execute("PRAGMA table_info(date_entries)")
            columns = [col[1] for col in cursor.fetchall()]
            cursor.close()
            
            # Connect to MySQL
            mysql_conn = self.db_conn.connect()
            mysql_cursor = mysql_conn.cursor()
            
            # Insert data into MySQL
            for row in rows:
                try:
                    row_dict = dict(zip(columns, row))
                    
                    data = (
                        row_dict.get('clli', ''),
                        row_dict.get('planned_start', ''),
                        row_dict.get('actual_start', ''),
                        row_dict.get('planned_end', ''),
                        row_dict.get('actual_end', ''),
                        self._convert_datetime(row_dict.get('created_date', '')),
                        self._convert_datetime(row_dict.get('last_updated', ''))
                    )
                    
                    mysql_cursor.execute('''
                        INSERT INTO date_entries (clli, planned_start, actual_start, planned_end, actual_end, created_date, last_updated)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ''', data)
                    
                    self.migration_stats['date_entries'] += 1
                    
                except Exception as e:
                    error_msg = f"Error migrating date entry {row_dict.get('clli', 'unknown')}: {e}"
                    print(error_msg)
                    self.migration_stats['errors'].append(error_msg)
            
            mysql_conn.commit()
            mysql_cursor.close()
            mysql_conn.close()
            
            print(f"Migrated {self.migration_stats['date_entries']} date entries")
            return True
            
        except Exception as e:
            print(f"Error migrating date entries: {e}")
            self.migration_stats['errors'].append(f"Date entries migration: {e}")
            return False
    
    def _convert_datetime(self, datetime_str):
        """Convert datetime string to MySQL format"""
        if not datetime_str:
            return datetime.now()
        
        try:
            # Try to parse various datetime formats
            if isinstance(datetime_str, str):
                # Handle different formats
                for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y']:
                    try:
                        return datetime.strptime(datetime_str, fmt)
                    except ValueError:
                        continue
            
            return datetime.now()
        except:
            return datetime.now()
    
    def run_migration(self):
        """Run the complete migration process"""
        print("Starting SQLite to MySQL migration...")
        print("=" * 50)
        
        # Step 1: Connect to SQLite databases
        if not self.connect_sqlite_databases():
            print("Failed to connect to SQLite databases")
            return False
        
        # Step 2: Create MySQL tables
        if not self.create_mysql_tables():
            print("Failed to create MySQL tables")
            return False
        
        # Step 3: Migrate data
        print("\nMigrating data...")
        self.migrate_workflow_entries()
        self.migrate_milestones()
        self.migrate_subtasks()
        self.migrate_date_entries()
        
        # Step 4: Close SQLite connections
        for conn in self.sqlite_connections.values():
            conn.close()
        
        # Step 5: Print migration summary
        print("\n" + "=" * 50)
        print("MIGRATION SUMMARY")
        print("=" * 50)
        print(f"Workflow entries: {self.migration_stats['workflow_entries']}")
        print(f"Milestones: {self.migration_stats['milestones']}")
        print(f"Subtasks: {self.migration_stats['subtasks']}")
        print(f"Date entries: {self.migration_stats['date_entries']}")
        print(f"Errors: {len(self.migration_stats['errors'])}")
        
        if self.migration_stats['errors']:
            print("\nErrors encountered:")
            for error in self.migration_stats['errors']:
                print(f"  - {error}")
        
        print("\nMigration completed!")
        return True


def main():
    """Main function"""
    print("SQLite to MySQL Migration Tool")
    print("=============================")
    
    # Check if config file exists
    config_file = "config.json"
    if not os.path.exists(config_file):
        print(f"Configuration file {config_file} not found!")
        print("Please create a config.json file based on config.example.json")
        return False
    
    # Run migration
    migrator = SQLiteToMySQLMigrator(config_file)
    success = migrator.run_migration()
    
    if success:
        print("\nMigration completed successfully!")
        print("You can now use the application with MySQL database.")
    else:
        print("\nMigration failed. Please check the errors above.")
    
    return success


if __name__ == "__main__":
    main()
