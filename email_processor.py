#!/usr/bin/env python3
"""
BASH Flow Management - Email Processor
Python backend for VBA email processing
"""

import json
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass, asdict
from db_config import get_database_connection

@dataclass
class ValidationResult:
    is_valid: bool
    record_type: str
    clli_number: str
    ms_number: str
    event_number: str
    record_date: str
    description: str
    status: str
    error_message: str

@dataclass
class DatabaseRecord:
    subject: str
    record_type: str
    clli_number: str
    ms_number: str
    event_number: str
    record_date: str
    description: str
    status: str
    created_date: str

class EmailProcessor:
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.db_conn = get_database_connection(config_file)
        self.patterns = {
            'CLLI': r'^(CLLI)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$',
            'MS': r'^(MS)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$',
            'Event': r'^(Event)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$'
        }
    
    def validate_subject_line(self, subject: str) -> ValidationResult:
        """Validate subject line against patterns"""
        try:
            for record_type, pattern in self.patterns.items():
                match = re.match(pattern, subject.strip())
                if match:
                    # Extract status from description for events
                    description = match.group(4)
                    status = "Active"
                    if record_type == 'Event':
                        status = self._extract_event_status(description)
                    
                    return ValidationResult(
                        is_valid=True,
                        record_type=match.group(1),
                        clli_number=match.group(2) if record_type == 'CLLI' else '',
                        ms_number=match.group(2) if record_type == 'MS' else '',
                        event_number=match.group(2) if record_type == 'Event' else '',
                        record_date=match.group(3),
                        description=description,
                        status=status,
                        error_message=''
                    )
            
            return ValidationResult(
                is_valid=False,
                record_type='',
                clli_number='',
                ms_number='',
                event_number='',
                record_date='',
                description='',
                status='',
                error_message='Subject line does not match any preset format'
            )
        except Exception as e:
            return ValidationResult(
                is_valid=False,
                record_type='',
                clli_number='',
                ms_number='',
                event_number='',
                record_date='',
                description='',
                status='',
                error_message=f'Validation error: {str(e)}'
            )
    
    def _extract_event_status(self, description: str) -> str:
        """Extract status from event description"""
        desc_lower = description.lower()
        if 'completed' in desc_lower:
            return 'Completed'
        elif 'blocked' in desc_lower:
            return 'Blocked'
        elif 'pending' in desc_lower:
            return 'Pending'
        else:
            return 'Active'
    
    def create_database_record(self, subject: str, validation: ValidationResult) -> DatabaseRecord:
        """Create database record from validation result"""
        return DatabaseRecord(
            subject=subject.strip(),
            record_type=validation.record_type,
            clli_number=validation.clli_number,
            ms_number=validation.ms_number,
            event_number=validation.event_number,
            record_date=validation.record_date,
            description=validation.description,
            status=validation.status,
            created_date=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
    
    def save_to_database(self, record: DatabaseRecord) -> bool:
        """Save record to database"""
        try:
            conn = self.db_conn.connect()
            cursor = conn.cursor()
            
            # Create table if it doesn't exist
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS BASHFlowSandbox (
                    ID INT AUTO_INCREMENT PRIMARY KEY,
                    Subject VARCHAR(255),
                    RecordType VARCHAR(50),
                    CLLINumber VARCHAR(20),
                    MSNumber VARCHAR(20),
                    EventNumber VARCHAR(20),
                    RecordDate VARCHAR(20),
                    Description TEXT,
                    Status VARCHAR(50),
                    CreatedDate DATETIME
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """)
            
            cursor.execute("""
                INSERT INTO BASHFlowSandbox 
                (Subject, RecordType, CLLINumber, MSNumber, EventNumber, 
                 RecordDate, Description, Status, CreatedDate)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                record.subject, record.record_type, record.clli_number,
                record.ms_number, record.event_number, record.record_date,
                record.description, record.status, record.created_date
            ))
            
            conn.commit()
            cursor.close()
            conn.close()
            return True
        except Exception as e:
            print(f"Database error: {e}", file=sys.stderr)
            return False
    
    def process_email(self, subject: str) -> Dict:
        """Main processing function - validates and saves email"""
        try:
            # Validate subject line
            validation = self.validate_subject_line(subject)
            
            if not validation.is_valid:
                return {
                    'success': False,
                    'error': validation.error_message,
                    'validation': asdict(validation)
                }
            
            # Create database record
            record = self.create_database_record(subject, validation)
            
            # Save to database
            if self.save_to_database(record):
                return {
                    'success': True,
                    'message': 'Email processed successfully',
                    'validation': asdict(validation),
                    'record': asdict(record)
                }
            else:
                return {
                    'success': False,
                    'error': 'Failed to save to database',
                    'validation': asdict(validation)
                }
        except Exception as e:
            return {
                'success': False,
                'error': f'Processing error: {str(e)}',
                'validation': None
            }

def main():
    """Main function for command-line usage"""
    if len(sys.argv) < 2:
        print(json.dumps({
            'success': False,
            'error': 'No subject line provided'
        }))
        return
    
    subject = sys.argv[1]
    processor = EmailProcessor()
    result = processor.process_email(subject)
    
    # Output JSON result for VBA to consume
    print(json.dumps(result, indent=2))

if __name__ == "__main__":
    main()
