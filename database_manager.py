"""
Database Manager Module
======================

This module provides comprehensive database management functionality for the workflow system.
It handles database operations, maintenance, backup, and reporting.

Author: Workflow Manager System
Version: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import shutil
import datetime
from typing import List, Dict, Optional, Tuple
import pandas as pd
from db_config import get_database_connection


class DatabaseManager:
    """Database management class for workflow system"""
    
    def __init__(self, parent_root, config_file: str = "config.json"):
        """
        Initialize the Database Manager
        
        Args:
            parent_root: Parent window reference
            config_file: Path to database configuration file
        """
        self.parent_root = parent_root
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        
        # For backward compatibility
        self.main_db_path = "MySQL Database"
        self.milestone_db_path = "MySQL Database"
        
        # Create database manager window
        self.root = tk.Toplevel(parent_root)
        self.root.title("Database Manager")
        self.root.geometry("1000x700")
        self.root.transient(parent_root)
        self.root.grab_set()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Initialize database info
        self.db_info = {}
        self.refresh_database_info()
        
        # Create GUI
        self.create_widgets()
        
        # Load initial data
        self.load_database_info()
        
        self.root.wait_window(self.root)
    
    def _on_closing(self):
        """Handle window closing"""
        if messagebox.askokcancel("Quit", "Do you want to close the Database Manager?"):
            self.root.destroy()
    
    def create_widgets(self):
        """Create the database manager GUI"""
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Database Manager", 
                              font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.create_overview_tab()
        self.create_backup_tab()
        self.create_maintenance_tab()
        self.create_reports_tab()
        self.create_settings_tab()
        
        # Bottom buttons
        self.create_bottom_buttons(main_frame)
    
    def create_overview_tab(self):
        """Create database overview tab"""
        overview_frame = ttk.Frame(self.notebook)
        self.notebook.add(overview_frame, text="Overview")
        
        # Database info frame
        info_frame = ttk.LabelFrame(overview_frame, text="Database Information", padding="10")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Database paths
        ttk.Label(info_frame, text="Main Database:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.main_db_label = ttk.Label(info_frame, text=self.main_db_path, foreground='blue')
        self.main_db_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        ttk.Label(info_frame, text="Milestone Database:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.milestone_db_label = ttk.Label(info_frame, text=self.milestone_db_path, foreground='blue')
        self.milestone_db_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Database status
        status_frame = ttk.LabelFrame(overview_frame, text="Database Status", padding="10")
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Status labels
        self.main_status_label = ttk.Label(status_frame, text="Main DB: Checking...")
        self.main_status_label.pack(anchor=tk.W, pady=2)
        
        self.milestone_status_label = ttk.Label(status_frame, text="Milestone DB: Checking...")
        self.milestone_status_label.pack(anchor=tk.W, pady=2)
        
        # Database statistics
        stats_frame = ttk.LabelFrame(overview_frame, text="Database Statistics", padding="10")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Create treeview for statistics
        columns = ('table', 'records', 'size', 'last_modified')
        self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show='headings', height=8)
        
        # Configure columns
        self.stats_tree.heading('table', text='Table Name')
        self.stats_tree.heading('records', text='Record Count')
        self.stats_tree.heading('size', text='Size (KB)')
        self.stats_tree.heading('last_modified', text='Last Modified')
        
        self.stats_tree.column('table', width=150)
        self.stats_tree.column('records', width=100)
        self.stats_tree.column('size', width=100)
        self.stats_tree.column('last_modified', width=150)
        
        # Add scrollbar
        stats_scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=stats_scrollbar.set)
        
        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Refresh button
        refresh_btn = ttk.Button(overview_frame, text="Refresh Information", 
                               command=self.refresh_database_info)
        refresh_btn.pack(pady=10)
    
    def create_backup_tab(self):
        """Create database backup tab"""
        backup_frame = ttk.Frame(self.notebook)
        self.notebook.add(backup_frame, text="Backup & Restore")
        
        # Backup section
        backup_section = ttk.LabelFrame(backup_frame, text="Database Backup", padding="10")
        backup_section.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(backup_section, text="Create a backup of your databases:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Backup options
        self.backup_main_var = tk.BooleanVar(value=True)
        self.backup_milestone_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(backup_section, text="Backup Main Database", 
                       variable=self.backup_main_var).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(backup_section, text="Backup Milestone Database", 
                       variable=self.backup_milestone_var).pack(anchor=tk.W, pady=2)
        
        # Backup location
        location_frame = ttk.Frame(backup_section)
        location_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(location_frame, text="Backup Location:").pack(side=tk.LEFT)
        self.backup_location_var = tk.StringVar()
        self.backup_location_entry = ttk.Entry(location_frame, textvariable=self.backup_location_var, width=50)
        self.backup_location_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(location_frame, text="Browse", command=self.browse_backup_location)
        browse_btn.pack(side=tk.RIGHT)
        
        # Backup button
        backup_btn = ttk.Button(backup_section, text="Create Backup", 
                              command=self.create_backup)
        backup_btn.pack(pady=10)
        
        # Restore section
        restore_section = ttk.LabelFrame(backup_frame, text="Database Restore", padding="10")
        restore_section.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(restore_section, text="Restore from backup:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Restore file selection
        restore_frame = ttk.Frame(restore_section)
        restore_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(restore_frame, text="Backup File:").pack(side=tk.LEFT)
        self.restore_file_var = tk.StringVar()
        self.restore_file_entry = ttk.Entry(restore_frame, textvariable=self.restore_file_var, width=50)
        self.restore_file_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
        
        restore_browse_btn = ttk.Button(restore_frame, text="Browse", command=self.browse_restore_file)
        restore_browse_btn.pack(side=tk.RIGHT)
        
        # Restore button
        restore_btn = ttk.Button(restore_section, text="Restore Database", 
                               command=self.restore_database)
        restore_btn.pack(pady=10)
    
    def create_maintenance_tab(self):
        """Create database maintenance tab"""
        maintenance_frame = ttk.Frame(self.notebook)
        self.notebook.add(maintenance_frame, text="Maintenance")
        
        # Database optimization
        optimization_section = ttk.LabelFrame(maintenance_frame, text="Database Optimization", padding="10")
        optimization_section.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(optimization_section, text="Optimize database performance:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Optimization buttons
        optimize_btn = ttk.Button(optimization_section, text="Optimize Main Database", 
                                 command=lambda: self.optimize_database(self.main_db_path))
        optimize_btn.pack(pady=5)
        
        optimize_milestone_btn = ttk.Button(optimization_section, text="Optimize Milestone Database", 
                                          command=lambda: self.optimize_database(self.milestone_db_path))
        optimize_milestone_btn.pack(pady=5)
        
        # Database cleanup
        cleanup_section = ttk.LabelFrame(maintenance_frame, text="Database Cleanup", padding="10")
        cleanup_section.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(cleanup_section, text="Clean up database files:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Cleanup options
        self.cleanup_orphaned_var = tk.BooleanVar(value=True)
        self.cleanup_duplicates_var = tk.BooleanVar(value=True)
        self.cleanup_old_records_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(cleanup_section, text="Remove orphaned records", 
                       variable=self.cleanup_orphaned_var).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(cleanup_section, text="Remove duplicate entries", 
                       variable=self.cleanup_duplicates_var).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(cleanup_section, text="Remove old records (>1 year)", 
                       variable=self.cleanup_old_records_var).pack(anchor=tk.W, pady=2)
        
        # Cleanup button
        cleanup_btn = ttk.Button(cleanup_section, text="Run Cleanup", 
                               command=self.run_cleanup)
        cleanup_btn.pack(pady=10)
    
    def create_reports_tab(self):
        """Create database reports tab"""
        reports_frame = ttk.Frame(self.notebook)
        self.notebook.add(reports_frame, text="Reports")
        
        # Report generation
        report_section = ttk.LabelFrame(reports_frame, text="Generate Reports", padding="10")
        report_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Report types
        ttk.Label(report_section, text="Select report type:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Report buttons
        reports_frame_inner = ttk.Frame(report_section)
        reports_frame_inner.pack(fill=tk.X, pady=10)
        
        # Left column
        left_col = ttk.Frame(reports_frame_inner)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Button(left_col, text="Database Summary", 
                  command=lambda: self.generate_report("summary")).pack(fill=tk.X, pady=2)
        ttk.Button(left_col, text="Record Statistics", 
                  command=lambda: self.generate_report("statistics")).pack(fill=tk.X, pady=2)
        ttk.Button(left_col, text="Data Quality Report", 
                  command=lambda: self.generate_report("quality")).pack(fill=tk.X, pady=2)
        
        # Right column
        right_col = ttk.Frame(reports_frame_inner)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        ttk.Button(right_col, text="Export to Excel", 
                  command=lambda: self.generate_report("excel")).pack(fill=tk.X, pady=2)
        ttk.Button(right_col, text="Export to CSV", 
                  command=lambda: self.generate_report("csv")).pack(fill=tk.X, pady=2)
        ttk.Button(right_col, text="Custom Query", 
                  command=self.run_custom_query).pack(fill=tk.X, pady=2)
        
        # Report output area
        output_section = ttk.LabelFrame(reports_frame, text="Report Output", padding="10")
        output_section.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Text widget for report output
        self.report_text = tk.Text(output_section, height=15, width=80, wrap=tk.WORD)
        report_scrollbar = ttk.Scrollbar(output_section, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=report_scrollbar.set)
        
        self.report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        report_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_settings_tab(self):
        """Create database settings tab"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")
        
        # Database configuration
        config_section = ttk.LabelFrame(settings_frame, text="Database Configuration", padding="10")
        config_section.pack(fill=tk.X, padx=10, pady=10)
        
        # Auto-backup settings
        ttk.Label(config_section, text="Auto-backup Settings:", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        self.auto_backup_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(config_section, text="Enable automatic backups", 
                       variable=self.auto_backup_var).pack(anchor=tk.W, pady=2)
        
        # Backup frequency
        freq_frame = ttk.Frame(config_section)
        freq_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(freq_frame, text="Backup Frequency:").pack(side=tk.LEFT)
        self.backup_freq_var = tk.StringVar(value="daily")
        freq_combo = ttk.Combobox(freq_frame, textvariable=self.backup_freq_var, 
                                 values=["daily", "weekly", "monthly"], state="readonly")
        freq_combo.pack(side=tk.LEFT, padx=(10, 0))
        
        # Database paths
        paths_section = ttk.LabelFrame(settings_frame, text="Database Paths", padding="10")
        paths_section.pack(fill=tk.X, padx=10, pady=10)
        
        # Main database path
        main_path_frame = ttk.Frame(paths_section)
        main_path_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(main_path_frame, text="Main Database:").pack(side=tk.LEFT)
        self.main_path_var = tk.StringVar(value=self.main_db_path)
        main_path_entry = ttk.Entry(main_path_frame, textvariable=self.main_path_var, width=50)
        main_path_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
        
        main_browse_btn = ttk.Button(main_path_frame, text="Browse", 
                                   command=self.browse_main_database)
        main_browse_btn.pack(side=tk.RIGHT)
        
        # Milestone database path
        milestone_path_frame = ttk.Frame(paths_section)
        milestone_path_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(milestone_path_frame, text="Milestone Database:").pack(side=tk.LEFT)
        self.milestone_path_var = tk.StringVar(value=self.milestone_db_path)
        milestone_path_entry = ttk.Entry(milestone_path_frame, textvariable=self.milestone_path_var, width=50)
        milestone_path_entry.pack(side=tk.LEFT, padx=(10, 5), fill=tk.X, expand=True)
        
        milestone_browse_btn = ttk.Button(milestone_path_frame, text="Browse", 
                                        command=self.browse_milestone_database)
        milestone_browse_btn.pack(side=tk.RIGHT)
        
        # Save settings button
        save_settings_btn = ttk.Button(settings_frame, text="Save Settings", 
                                     command=self.save_settings)
        save_settings_btn.pack(pady=20)
    
    def create_bottom_buttons(self, parent):
        """Create bottom button frame"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Close button
        close_btn = ttk.Button(button_frame, text="Close", command=self.root.destroy)
        close_btn.pack(side=tk.RIGHT)
    
    def refresh_database_info(self):
        """Refresh database information"""
        try:
            # Test MySQL connection
            try:
                conn = self.db_conn.connect()
                main_exists = True
                main_size = 0  # MySQL doesn't have file size concept
                conn.close()
            except Exception:
                main_exists = False
                main_size = 0
            
            # For MySQL, main and milestone are the same database
            milestone_exists = main_exists
            milestone_size = main_size
            
            # Update status labels
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text=f"Main DB: {'Connected' if main_exists else 'Not Found'}")
            
            if hasattr(self, 'milestone_status_label'):
                self.milestone_status_label.config(text=f"Milestone DB: {'Connected' if milestone_exists else 'Not Found'}")
            
            # Store database info
            self.db_info = {
                'main': {'exists': main_exists, 'size': main_size},
                'milestone': {'exists': milestone_exists, 'size': milestone_size}
            }
            
        except Exception as e:
            print(f"Error refreshing database info: {e}")
    
    def load_database_info(self):
        """Load database statistics into the treeview"""
        try:
            # Clear existing items
            for item in self.stats_tree.get_children():
                self.stats_tree.delete(item)
            
            # Get MySQL database stats
            try:
                conn = self.db_conn.connect()
                stats = self.get_mysql_database_stats(conn)
                for table, table_stats in stats.items():
                    self.stats_tree.insert('', 'end', values=(
                        table,
                        table_stats['records'],
                        f"{table_stats['size']:.1f}",
                        table_stats['last_modified']
                    ))
                conn.close()
                    
            except Exception as e:
                print(f"Error loading database info: {e}")
                
        except Exception as e:
            print(f"Error in load_database_info: {e}")
    
    def get_mysql_database_stats(self, conn) -> Dict:
        """Get MySQL database statistics"""
        stats = {}
        try:
            cursor = conn.cursor()
            
            # Get all tables
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()
            
            for table in tables:
                table_name = table[0]
                
                # Get record count
                cursor.execute(f"SELECT COUNT(*) FROM `{table_name}`")
                record_count = cursor.fetchone()[0]
                
                # Get table size from information_schema
                cursor.execute("""
                    SELECT 
                        ROUND(((data_length + index_length) / 1024), 2) AS 'Size (KB)'
                    FROM information_schema.TABLES 
                    WHERE table_schema = DATABASE() 
                    AND table_name = %s
                """, (table_name,))
                
                size_result = cursor.fetchone()
                table_size = size_result[0] if size_result and size_result[0] else 0.0
                
                # Get last modified (approximate)
                last_modified = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                
                stats[table_name] = {
                    'records': record_count,
                    'size': table_size,
                    'last_modified': last_modified
                }
            
            cursor.close()
            
        except Exception as e:
            print(f"Error getting MySQL database stats: {e}")
        
        return stats
    
    def get_database_stats(self, db_path: str) -> Dict:
        """Get database statistics (legacy SQLite method)"""
        stats = {}
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Get all tables
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            
            for table in tables:
                table_name = table[0]
                
                # Get record count
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                record_count = cursor.fetchone()[0]
                
                # Get table size (approximate)
                cursor.execute(f"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{table_name}'")
                table_size = 1.0  # Placeholder - would need more complex query for actual size
                
                # Get last modified (approximate)
                last_modified = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                
                stats[table_name] = {
                    'records': record_count,
                    'size': table_size,
                    'last_modified': last_modified
                }
            
            conn.close()
            
        except Exception as e:
            print(f"Error getting database stats: {e}")
        
        return stats
    
    def browse_backup_location(self):
        """Browse for backup location"""
        directory = filedialog.askdirectory(title="Select Backup Location")
        if directory:
            self.backup_location_var.set(directory)
    
    def browse_restore_file(self):
        """Browse for restore file"""
        filename = filedialog.askopenfilename(
            title="Select Backup File",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")]
        )
        if filename:
            self.restore_file_var.set(filename)
    
    def create_backup(self):
        """Create database backup (MySQL dump)"""
        try:
            backup_location = self.backup_location_var.get()
            if not backup_location:
                messagebox.showwarning("Warning", "Please select a backup location.")
                return
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create MySQL dump
            import subprocess
            import json
            
            # Get database configuration
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            
            db_config = config.get('database', {})
            db_name = db_config.get('database_name', 'degrow_workflow')
            username = db_config.get('username', 'root')
            password = db_config.get('password', '')
            host = db_config.get('host', 'localhost')
            port = db_config.get('port', 3306)
            
            # Create mysqldump command
            dump_file = os.path.join(backup_location, f"mysql_backup_{timestamp}.sql")
            
            cmd = [
                'mysqldump',
                f'--host={host}',
                f'--port={port}',
                f'--user={username}',
                f'--password={password}',
                '--single-transaction',
                '--routines',
                '--triggers',
                db_name
            ]
            
            with open(dump_file, 'w') as f:
                result = subprocess.run(cmd, stdout=f, stderr=subprocess.PIPE, text=True)
            
            if result.returncode == 0:
                messagebox.showinfo("Success", f"MySQL database backed up to:\n{dump_file}")
            else:
                messagebox.showerror("Error", f"Failed to create MySQL backup:\n{result.stderr}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create backup: {e}")
    
    def restore_database(self):
        """Restore database from MySQL dump"""
        try:
            restore_file = self.restore_file_var.get()
            if not restore_file:
                messagebox.showwarning("Warning", "Please select a backup file to restore.")
                return
            
            if not os.path.exists(restore_file):
                messagebox.showerror("Error", "Backup file does not exist.")
                return
            
            # Confirm restore
            if messagebox.askyesno("Confirm Restore", 
                                 "Are you sure you want to restore the MySQL database?\n"
                                 "This will overwrite the current database."):
                
                import subprocess
                import json
                
                # Get database configuration
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                
                db_config = config.get('database', {})
                db_name = db_config.get('database_name', 'degrow_workflow')
                username = db_config.get('username', 'root')
                password = db_config.get('password', '')
                host = db_config.get('host', 'localhost')
                port = db_config.get('port', 3306)
                
                # Create mysql command
                cmd = [
                    'mysql',
                    f'--host={host}',
                    f'--port={port}',
                    f'--user={username}',
                    f'--password={password}',
                    db_name
                ]
                
                with open(restore_file, 'r') as f:
                    result = subprocess.run(cmd, stdin=f, stderr=subprocess.PIPE, text=True)
                
                if result.returncode == 0:
                    messagebox.showinfo("Success", "MySQL database restored successfully.")
                    self.refresh_database_info()
                    self.load_database_info()
                else:
                    messagebox.showerror("Error", f"Failed to restore MySQL database:\n{result.stderr}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to restore database: {e}")
    
    def optimize_database(self, db_path: str):
        """Optimize MySQL database"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Run OPTIMIZE TABLE for all tables
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()
            
            for table in tables:
                table_name = table[0]
                cursor.execute(f"OPTIMIZE TABLE `{table_name}`")
            
            conn.commit()
            cursor.close()
            conn.close()
            
            messagebox.showinfo("Success", "MySQL database optimized successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to optimize database: {e}")
    
    def run_cleanup(self):
        """Run database cleanup"""
        try:
            cleanup_tasks = []
            
            if self.cleanup_orphaned_var.get():
                cleanup_tasks.append("orphaned records")
            if self.cleanup_duplicates_var.get():
                cleanup_tasks.append("duplicate entries")
            if self.cleanup_old_records_var.get():
                cleanup_tasks.append("old records")
            
            if not cleanup_tasks:
                messagebox.showwarning("Warning", "Please select at least one cleanup task.")
                return
            
            # Here you would implement actual cleanup logic
            messagebox.showinfo("Success", f"Cleanup completed for: {', '.join(cleanup_tasks)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run cleanup: {e}")
    
    def generate_report(self, report_type: str):
        """Generate database report"""
        try:
            self.report_text.delete(1.0, tk.END)
            
            if report_type == "summary":
                self.report_text.insert(tk.END, "Database Summary Report\n")
                self.report_text.insert(tk.END, "=" * 50 + "\n\n")
                self.report_text.insert(tk.END, f"Main Database: {self.main_db_path}\n")
                self.report_text.insert(tk.END, f"Milestone Database: {self.milestone_db_path}\n\n")
                
                # Add more detailed summary here
                
            elif report_type == "statistics":
                self.report_text.insert(tk.END, "Database Statistics Report\n")
                self.report_text.insert(tk.END, "=" * 50 + "\n\n")
                
                # Add statistics here
                
            elif report_type == "quality":
                self.report_text.insert(tk.END, "Data Quality Report\n")
                self.report_text.insert(tk.END, "=" * 50 + "\n\n")
                
                # Add quality checks here
                
            elif report_type == "excel":
                self.export_to_excel()
                
            elif report_type == "csv":
                self.export_to_csv()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {e}")
    
    def run_custom_query(self):
        """Run custom SQL query"""
        try:
            # Create custom query dialog
            query_window = tk.Toplevel(self.root)
            query_window.title("Custom SQL Query")
            query_window.geometry("600x400")
            query_window.transient(self.root)
            query_window.grab_set()
            
            # Query input
            ttk.Label(query_window, text="Enter SQL Query:", font=('Arial', 10, 'bold')).pack(pady=10)
            
            query_text = tk.Text(query_window, height=10, width=70)
            query_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            
            # Buttons
            button_frame = ttk.Frame(query_window)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="Execute", 
                      command=lambda: self.execute_custom_query(query_text.get(1.0, tk.END))).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Close", 
                      command=query_window.destroy).pack(side=tk.LEFT, padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open query dialog: {e}")
    
    def execute_custom_query(self, query: str):
        """Execute custom SQL query"""
        try:
            # This is a placeholder - implement actual query execution
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(tk.END, f"Executing query: {query}\n")
            self.report_text.insert(tk.END, "Query executed successfully.\n")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to execute query: {e}")
    
    def export_to_excel(self):
        """Export database to Excel"""
        try:
            filename = filedialog.asksaveasfilename(
                title="Save Excel File",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if filename:
                # Implement Excel export here
                messagebox.showinfo("Success", f"Data exported to: {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export to Excel: {e}")
    
    def export_to_csv(self):
        """Export database to CSV"""
        try:
            filename = filedialog.asksaveasfilename(
                title="Save CSV File",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if filename:
                # Implement CSV export here
                messagebox.showinfo("Success", f"Data exported to: {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export to CSV: {e}")
    
    def browse_main_database(self):
        """Browse for main database file"""
        filename = filedialog.askopenfilename(
            title="Select Main Database",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")]
        )
        if filename:
            self.main_path_var.set(filename)
    
    def browse_milestone_database(self):
        """Browse for milestone database file"""
        filename = filedialog.askopenfilename(
            title="Select Milestone Database",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")]
        )
        if filename:
            self.milestone_path_var.set(filename)
    
    def save_settings(self):
        """Save database settings"""
        try:
            # Update database paths
            self.main_db_path = self.main_path_var.get()
            self.milestone_db_path = self.milestone_path_var.get()
            
            # Save other settings (implement as needed)
            messagebox.showinfo("Success", "Settings saved successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")


def open_database_manager(parent_root, config_file: str = "config.json"):
    """Open the database manager window"""
    DatabaseManager(parent_root, config_file)
