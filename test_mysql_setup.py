#!/usr/bin/env python3
"""
MySQL Setup Test Script
======================

This script tests the MySQL database setup for the Degrow Workflow Manager.
It verifies connection, creates tables, and performs basic operations.

Author: Workflow Manager System
Version: 1.0.0
"""

import sys
import os
from datetime import datetime

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from db_config import get_database_connection, create_mysql_database


def test_database_connection():
    """Test basic database connection"""
    print("Testing database connection...")
    try:
        db_conn = get_database_connection()
        connection = db_conn.connect()
        print("✓ Database connection successful")
        connection.close()
        return True
    except Exception as e:
        print(f"✗ Database connection failed: {e}")
        return False


def test_database_creation():
    """Test database creation"""
    print("Testing database creation...")
    try:
        success = create_mysql_database()
        if success:
            print("✓ Database creation successful")
            return True
        else:
            print("✗ Database creation failed")
            return False
    except Exception as e:
        print(f"✗ Database creation error: {e}")
        return False


def test_table_creation():
    """Test table creation"""
    print("Testing table creation...")
    try:
        db_conn = get_database_connection()
        conn = db_conn.connect()
        cursor = conn.cursor()
        
        # Test creating a simple table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_table (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(100),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        
        print("✓ Table creation successful")
        
        # Clean up test table
        cursor.execute("DROP TABLE IF EXISTS test_table")
        conn.commit()
        
        cursor.close()
        conn.close()
        return True
        
    except Exception as e:
        print(f"✗ Table creation failed: {e}")
        return False


def test_crud_operations():
    """Test basic CRUD operations"""
    print("Testing CRUD operations...")
    try:
        db_conn = get_database_connection()
        conn = db_conn.connect()
        cursor = conn.cursor()
        
        # Create test table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_crud (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(100),
                value INT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        
        # Test INSERT
        cursor.execute("""
            INSERT INTO test_crud (name, value) VALUES (%s, %s)
        """, ("Test Item", 42))
        
        # Test SELECT
        cursor.execute("SELECT * FROM test_crud WHERE name = %s", ("Test Item",))
        result = cursor.fetchone()
        
        if result and result[1] == "Test Item":
            print("✓ INSERT and SELECT operations successful")
        else:
            print("✗ INSERT or SELECT operations failed")
            return False
        
        # Test UPDATE
        cursor.execute("""
            UPDATE test_crud SET value = %s WHERE name = %s
        """, (100, "Test Item"))
        
        cursor.execute("SELECT value FROM test_crud WHERE name = %s", ("Test Item",))
        result = cursor.fetchone()
        
        if result and result[0] == 100:
            print("✓ UPDATE operation successful")
        else:
            print("✗ UPDATE operation failed")
            return False
        
        # Test DELETE
        cursor.execute("DELETE FROM test_crud WHERE name = %s", ("Test Item",))
        
        cursor.execute("SELECT COUNT(*) FROM test_crud WHERE name = %s", ("Test Item",))
        count = cursor.fetchone()[0]
        
        if count == 0:
            print("✓ DELETE operation successful")
        else:
            print("✗ DELETE operation failed")
            return False
        
        # Clean up
        cursor.execute("DROP TABLE test_crud")
        conn.commit()
        
        cursor.close()
        conn.close()
        return True
        
    except Exception as e:
        print(f"✗ CRUD operations failed: {e}")
        return False


def test_workflow_tables():
    """Test creating workflow-specific tables"""
    print("Testing workflow table creation...")
    try:
        db_conn = get_database_connection()
        conn = db_conn.connect()
        cursor = conn.cursor()
        
        # Create workflow_entries table
        cursor.execute("""
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
        """)
        
        # Create milestones table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS milestones (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(100) NOT NULL UNIQUE,
                description TEXT,
                order_index INT DEFAULT 0,
                created_date DATETIME NOT NULL,
                last_updated DATETIME NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        
        # Create subtasks table
        cursor.execute("""
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
        """)
        
        # Create stakeholders table
        cursor.execute("""
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
        """)
        
        # Create date_entries table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS date_entries (
                clli VARCHAR(20) PRIMARY KEY,
                planned_start VARCHAR(20),
                actual_start VARCHAR(20),
                planned_end VARCHAR(20),
                actual_end VARCHAR(20),
                created_date DATETIME NOT NULL,
                last_updated DATETIME NOT NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """)
        
        print("✓ Workflow tables created successfully")
        
        # Test inserting sample data
        now = datetime.now()
        
        # Insert sample milestone
        cursor.execute("""
            INSERT INTO milestones (name, description, order_index, created_date, last_updated)
            VALUES (%s, %s, %s, %s, %s)
        """, ("Test Milestone", "A test milestone", 1, now, now))
        
        # Get the milestone ID
        cursor.execute("SELECT id FROM milestones WHERE name = %s", ("Test Milestone",))
        milestone_id = cursor.fetchone()[0]
        
        # Insert sample subtask
        cursor.execute("""
            INSERT INTO subtasks (milestone_id, name, description, criticality, order_index, created_date, last_updated)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (milestone_id, "Test Subtask", "A test subtask", "Should be Complete", 1, now, now))
        
        # Insert sample workflow entry
        cursor.execute("""
            INSERT INTO workflow_entries 
            (state, clli, host_wire_centre, lata, equipment_type, current_milestone, 
             milestone_subtask, status, created_date, last_update)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, ("Test State", "TEST1234", "Test Center", "123", "Switch", 
              "Test Milestone", "Test Subtask", "In Progress", now, now))
        
        print("✓ Sample data inserted successfully")
        
        # Clean up sample data
        cursor.execute("DELETE FROM workflow_entries WHERE clli = %s", ("TEST1234",))
        cursor.execute("DELETE FROM subtasks WHERE name = %s", ("Test Subtask",))
        cursor.execute("DELETE FROM milestones WHERE name = %s", ("Test Milestone",))
        
        conn.commit()
        cursor.close()
        conn.close()
        
        return True
        
    except Exception as e:
        print(f"✗ Workflow table creation failed: {e}")
        return False


def main():
    """Run all tests"""
    print("MySQL Setup Test for Degrow Workflow Manager")
    print("=" * 50)
    
    tests = [
        ("Database Connection", test_database_connection),
        ("Database Creation", test_database_creation),
        ("Table Creation", test_table_creation),
        ("CRUD Operations", test_crud_operations),
        ("Workflow Tables", test_workflow_tables)
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"\n{test_name}:")
        if test_func():
            passed += 1
        else:
            print(f"  {test_name} failed!")
    
    print("\n" + "=" * 50)
    print(f"Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("✓ All tests passed! MySQL setup is working correctly.")
        return True
    else:
        print("✗ Some tests failed. Please check the configuration and try again.")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
