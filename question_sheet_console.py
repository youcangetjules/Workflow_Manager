#!/usr/bin/env python3
"""
BASH Flow Management - Question Sheet Console
Python console application that matches BASHQuestionSheet.vba functionality
"""

import json
import sqlite3
import sys
from datetime import datetime, date
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict
import os

@dataclass
class QuestionSheetEntry:
    state: str
    stage: str
    milestone_start_date: date
    status: str
    created_date: datetime

class QuestionSheetConsole:
    def __init__(self, db_path: str = r"C:\BASHFlowSandbox\TestDatabase.accdb"):
        self.db_path = db_path
        self.states = self._initialize_states()
        self.stages = self._initialize_stages()
        self.statuses = self._initialize_statuses()
        self.milestone_dates = self._initialize_milestone_dates()
    
    def _initialize_states(self) -> List[str]:
        """Initialize list of US states"""
        return [
            "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
            "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
            "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
            "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi",
            "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
            "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma",
            "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
            "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
            "West Virginia", "Wisconsin", "Wyoming"
        ]
    
    def _initialize_stages(self) -> List[str]:
        """Initialize list of project stages"""
        return [
            "Planning", "Design", "Development", "Testing", 
            "Deployment", "Maintenance", "Review", "Completion"
        ]
    
    def _initialize_statuses(self) -> List[str]:
        """Initialize list of status options"""
        return ["To Do", "In Progress", "Blocked", "Done"]
    
    def _initialize_milestone_dates(self) -> Dict[str, date]:
        """Initialize milestone dates for each stage"""
        return {
            "Planning": date(2024, 1, 1),
            "Design": date(2024, 1, 15),
            "Development": date(2024, 2, 1),
            "Testing": date(2024, 3, 1),
            "Deployment": date(2024, 4, 1),
            "Maintenance": date(2024, 5, 1),
            "Review": date(2024, 6, 1),
            "Completion": date(2024, 7, 1)
        }
    
    def show_welcome(self):
        """Display welcome message"""
        print("=" * 60)
        print("BASH Flow Management - Question Sheet Console")
        print("=" * 60)
        print()
    
    def get_state_selection(self) -> Optional[str]:
        """Get state selection from user"""
        print("Please select a state:")
        print()
        
        # Display states in columns
        for i, state in enumerate(self.states, 1):
            print(f"{i:2d}. {state:<15}", end="")
            if i % 3 == 0:  # New line every 3 states
                print()
        
        if len(self.states) % 3 != 0:
            print()  # Final newline if needed
        
        print()
        print("Enter the number (1-50) or state name:")
        
        while True:
            try:
                user_input = input("> ").strip()
                if not user_input:
                    return None
                
                # Check if input is a number
                if user_input.isdigit():
                    state_num = int(user_input)
                    if 1 <= state_num <= len(self.states):
                        return self.states[state_num - 1]
                    else:
                        print(f"Please enter a number between 1 and {len(self.states)}")
                        continue
                
                # Check if input matches a state name (case insensitive)
                for state in self.states:
                    if user_input.lower() == state.lower():
                        return state
                
                print("Invalid state. Please try again.")
                
            except (ValueError, KeyboardInterrupt):
                return None
    
    def get_stage_selection(self) -> Optional[str]:
        """Get stage selection from user"""
        print("\nPlease select a stage:")
        print()
        
        for i, stage in enumerate(self.stages, 1):
            print(f"{i}. {stage}")
        
        print()
        print("Enter the number (1-8) or stage name:")
        
        while True:
            try:
                user_input = input("> ").strip()
                if not user_input:
                    return None
                
                # Check if input is a number
                if user_input.isdigit():
                    stage_num = int(user_input)
                    if 1 <= stage_num <= len(self.stages):
                        return self.stages[stage_num - 1]
                    else:
                        print(f"Please enter a number between 1 and {len(self.stages)}")
                        continue
                
                # Check if input matches a stage name (case insensitive)
                for stage in self.stages:
                    if user_input.lower() == stage.lower():
                        return stage
                
                print("Invalid stage. Please try again.")
                
            except (ValueError, KeyboardInterrupt):
                return None
    
    def get_status_selection(self) -> Optional[str]:
        """Get status selection from user"""
        print("\nPlease select a status:")
        print()
        
        for i, status in enumerate(self.statuses, 1):
            print(f"{i}. {status}")
        
        print()
        print("Enter the number (1-4) or status name:")
        
        while True:
            try:
                user_input = input("> ").strip()
                if not user_input:
                    return None
                
                # Check if input is a number
                if user_input.isdigit():
                    status_num = int(user_input)
                    if 1 <= status_num <= len(self.statuses):
                        return self.statuses[status_num - 1]
                    else:
                        print(f"Please enter a number between 1 and {len(self.statuses)}")
                        continue
                
                # Check if input matches a status name (case insensitive)
                for status in self.statuses:
                    if user_input.lower() == status.lower():
                        return status
                
                print("Invalid status. Please try again.")
                
            except (ValueError, KeyboardInterrupt):
                return None
    
    def get_milestone_date(self, stage: str) -> date:
        """Get milestone date based on stage"""
        return self.milestone_dates.get(stage, date.today())
    
    def create_entry(self, state: str, stage: str, status: str) -> QuestionSheetEntry:
        """Create a question sheet entry"""
        return QuestionSheetEntry(
            state=state,
            stage=stage,
            milestone_start_date=self.get_milestone_date(stage),
            status=status,
            created_date=datetime.now()
        )
    
    def save_entry(self, entry: QuestionSheetEntry) -> bool:
        """Save entry to database"""
        try:
            # For Access database, we'll use a simplified approach
            # In production, you'd use pyodbc for Access
            conn = sqlite3.connect(self.db_path.replace('.accdb', '.db'))
            cursor = conn.cursor()
            
            # Create table if it doesn't exist
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS BASHFlowSandbox (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    State TEXT,
                    Type TEXT,
                    Date TEXT,
                    Description TEXT,
                    Status TEXT,
                    DateCreated TEXT
                )
            """)
            
            # Insert record
            cursor.execute("""
                INSERT INTO BASHFlowSandbox 
                (State, Type, Date, Description, Status, DateCreated)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                entry.state,
                entry.stage,
                entry.milestone_start_date.strftime('%Y-%m-%d'),
                'Question Sheet Entry',
                entry.status,
                entry.created_date.strftime('%Y-%m-%d %H:%M:%S')
            ))
            
            conn.commit()
            conn.close()
            return True
            
        except Exception as e:
            print(f"Database error: {e}")
            return False
    
    def show_entry_summary(self, entry: QuestionSheetEntry):
        """Show entry summary"""
        print("\n" + "=" * 50)
        print("ENTRY SUMMARY")
        print("=" * 50)
        print(f"State: {entry.state}")
        print(f"Stage: {entry.stage}")
        print(f"Milestone Start Date: {entry.milestone_start_date.strftime('%m/%d/%Y')}")
        print(f"Status: {entry.status}")
        print(f"Created: {entry.created_date.strftime('%m/%d/%Y %H:%M:%S')}")
        print("=" * 50)
    
    def run_question_interface(self):
        """Main question interface loop"""
        self.show_welcome()
        
        while True:
            try:
                print("\n" + "-" * 40)
                print("NEW ENTRY")
                print("-" * 40)
                
                # Get state selection
                state = self.get_state_selection()
                if not state:
                    print("State selection cancelled.")
                    break
                
                # Get stage selection
                stage = self.get_stage_selection()
                if not stage:
                    print("Stage selection cancelled.")
                    break
                
                # Get status selection
                status = self.get_status_selection()
                if not status:
                    print("Status selection cancelled.")
                    break
                
                # Create entry
                entry = self.create_entry(state, stage, status)
                
                # Show summary
                self.show_entry_summary(entry)
                
                # Save entry
                if self.save_entry(entry):
                    print("\n✅ Question sheet entry saved successfully!")
                else:
                    print("\n❌ Error saving question sheet entry.")
                
                # Ask if user wants to continue
                print("\nWould you like to create another entry?")
                continue_choice = input("Enter 'y' for yes, any other key to exit: ").strip().lower()
                
                if continue_choice != 'y':
                    break
                    
            except KeyboardInterrupt:
                print("\n\nExiting...")
                break
            except Exception as e:
                print(f"\nError: {e}")
                print("Please try again.")
        
        print("\nThank you for using BASH Flow Management Question Sheet!")

def main():
    """Main function"""
    try:
        console = QuestionSheetConsole()
        console.run_question_interface()
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
