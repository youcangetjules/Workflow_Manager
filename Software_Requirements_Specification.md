# Software Requirements Specification (SRS)
## Degrow Workflow Manager System

**Document Version:** 1.0  
**Date:** January 2025  
**Prepared by:** Development Team  
**Project:** Degrow Workflow Manager  

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [Overall Description](#2-overall-description)
3. [System Features](#3-system-features)
4. [External Interface Requirements](#4-external-interface-requirements)
5. [Non-Functional Requirements](#5-non-functional-requirements)
6. [Other Requirements](#6-other-requirements)
7. [Appendices](#7-appendices)

---

## 1. Introduction

### 1.1 Purpose
This Software Requirements Specification (SRS) document describes the functional and non-functional requirements for the Degrow Workflow Manager System. The system is designed to manage workflow processes, track milestones, handle email processing, and provide comprehensive database management capabilities for telecommunications and network infrastructure projects.

### 1.2 Scope
The Degrow Workflow Manager System is a comprehensive workflow management solution that includes:
- Question Sheet Management (GUI and Console interfaces)
- Milestone and Subtask Management
- Database Management and Administration
- Email Processing and Validation
- Data Visualization and Reporting
- Integration with VBA/Excel workflows

### 1.3 Definitions, Acronyms, and Abbreviations
- **CLLI**: Common Language Location Identifier
- **LATA**: Local Access and Transport Area
- **MS**: Milestone
- **GUI**: Graphical User Interface
- **SRS**: Software Requirements Specification
- **VBA**: Visual Basic for Applications
- **API**: Application Programming Interface

### 1.4 References
- Python 3.11+ Documentation
- Tkinter GUI Framework Documentation
- MySQLCommunity Edition Database Documentation
- Matplotlib Visualization Library
- Pandas Data Analysis Library

### 1.5 Overview
This document is organized into seven main sections covering system introduction, overall description, detailed features, interface requirements, non-functional requirements, and additional specifications.

---

## 2. Overall Description

### 2.1 Product Perspective
The Degrow Workflow Manager System is a standalone desktop application built using Python with the following key components:

- **Main Application**: `question_sheet_gui.py` - Primary GUI interface
- **Console Application**: `question_sheet_console.py` - Command-line interface
- **Database Manager**: `database_manager.py` - Database administration
- **Email Processor**: `email_processor.py` - Email validation and processing
- **Milestone Editor**: `milestone_editor.py` - Milestone and subtask management

### 2.2 Product Functions
The system provides the following major functions:

1. **Workflow Management**
   - Create and manage question sheet entries
   - Track project states and stages
   - Monitor milestone progress
   - Manage subtasks and dependencies

2. **Data Management**
   - Database operations (CRUD)
   - Data validation and integrity
   - Backup and restore functionality
   - Data export/import capabilities

3. **Email Processing**
   - Subject line validation
   - Pattern matching for CLLI, MS, and Event records
   - Automated database record creation
   - Status extraction and classification

4. **Visualization and Reporting**
   - Interactive charts and graphs
   - Data export to Excel/CSV
   - Custom query execution
   - Statistical analysis

5. **Administration**
   - User interface management
   - Database maintenance
   - System configuration
   - Error logging and debugging

### 2.3 User Classes and Characteristics
- **Primary Users**: Workflow managers, project coordinators
- **Secondary Users**: Database administrators, system administrators
- **End Users**: Field technicians, project team members

### 2.4 Operating Environment
- **Operating System**: Windows 10/11
- **Python Version**: 3.11 or higher
- **Database**: MySQLCommunity Edition 
- **Dependencies**: See requirements.txt for complete list

### 2.5 Design and Implementation Constraints
- Must maintain compatibility with existing VBA/Excel workflows
- Database schema must support legacy Access database format
- GUI must be responsive and user-friendly
- System must handle large datasets efficiently

### 2.6 Assumptions and Dependencies
- Users have basic computer literacy
- Database files are accessible and properly configured
- Network connectivity for email processing
- Sufficient system resources for data processing

---

## 3. System Features

### 3.1 Question Sheet Management

#### 3.1.1 Description
The system provides both GUI and console interfaces for managing question sheet entries, allowing users to create, view, edit, and track workflow entries.

#### 3.1.2 Functional Requirements

**FR-001: Question Sheet Entry Creation**
- The system SHALL allow users to create new question sheet entries
- The system SHALL validate all required fields before saving
- The system SHALL automatically generate timestamps for created entries
- The system SHALL support state selection from predefined list
- The system SHALL support stage selection from predefined list
- The system SHALL support status selection from predefined list

**FR-002: Data Validation**
- The system SHALL validate CLLI numbers against standard format
- The system SHALL validate LATA codes
- The system SHALL validate equipment types
- The system SHALL ensure data integrity before database operations

**FR-003: Search and Filter**
- The system SHALL provide search functionality across all fields
- The system SHALL support filtering by state, stage, status, and date ranges
- The system SHALL provide column visibility toggle options
- The system SHALL support sorting by any column

**FR-004: Data Display**
- The system SHALL display entries in a tabular format
- The system SHALL show all relevant entry details
- The system SHALL provide real-time updates
- The system SHALL support pagination for large datasets

### 3.2 Milestone and Subtask Management

#### 3.2.1 Description
The system provides comprehensive milestone and subtask management capabilities with hierarchical organization and dependency tracking.

#### 3.2.2 Functional Requirements

**FR-005: Milestone Management**
- The system SHALL allow creation of new milestones
- The system SHALL support milestone editing and deletion
- The system SHALL provide milestone reordering functionality
- The system SHALL track milestone duration and stakeholder teams
- The system SHALL support duration calculation methods (milestone duration vs. sum of subtasks)

**FR-006: Subtask Management**
- The system SHALL allow creation of subtasks under milestones
- The system SHALL support subtask editing and deletion
- The system SHALL track subtask prerequisites and dependencies
- The system SHALL support subtask reordering within milestones
- The system SHALL calculate milestone duration from subtask durations

**FR-007: Criticality Management**
- The system SHALL assign criticality levels to subtasks
- The system SHALL support three criticality levels: "Must be complete", "Should be Complete", "Does not block"
- The system SHALL allow editing of criticality assignments
- The system SHALL provide visual indicators for criticality levels

**FR-007A: Stakeholder Management**
- The system SHALL allow creation of new stakeholders with Name, Home State, Role, Team, Email, and Phone fields
- The system SHALL provide a dropdown selection for Home State with all 48 continental US states
- The system SHALL support stakeholder editing and deletion
- The system SHALL validate email format for stakeholder entries
- The system SHALL display stakeholders in a tabular format with all relevant fields
- The system SHALL support stakeholder reordering and management

### 3.3 Database Management

#### 3.3.1 Description
The system provides comprehensive database administration capabilities including backup, restore, optimization, and reporting.

#### 3.3.2 Functional Requirements

**FR-008: Database Operations**
- The system SHALL support database connection management
- The system SHALL provide database status monitoring
- The system SHALL display database statistics and metrics
- The system SHALL support database optimization (VACUUM operations)

**FR-009: Backup and Restore**
- The system SHALL create database backups with timestamps
- The system SHALL support selective backup (main database, milestone database, or both)
- The system SHALL provide restore functionality from backup files
- The system SHALL validate backup file integrity before restore

**FR-010: Data Maintenance**
- The system SHALL provide database cleanup functionality
- The system SHALL remove orphaned records
- The system SHALL identify and remove duplicate entries
- The system SHALL support archiving of old records

**FR-011: Reporting**
- The system SHALL generate database summary reports
- The system SHALL provide record statistics
- The system SHALL create data quality reports
- The system SHALL support custom SQL query execution
- The system SHALL export data to Excel and CSV formats

### 3.4 Email Processing

#### 3.4.1 Description
The system processes and validates email subject lines to extract structured data and create database records automatically.

#### 3.4.2 Functional Requirements

**FR-012: Email Validation**
- The system SHALL validate email subject lines against predefined patterns
- The system SHALL support CLLI record format: `CLLI-XXXX-YYYY-MM-DD-Description`
- The system SHALL support MS record format: `MS-XXXX-YYYY-MM-DD-Description`
- The system SHALL support Event record format: `Event-XXXX-YYYY-MM-DD-Description`
- The system SHALL extract status information from event descriptions

**FR-013: Data Extraction**
- The system SHALL extract record type from subject line
- The system SHALL extract identification numbers (CLLI, MS, Event)
- The system SHALL extract record dates
- The system SHALL extract descriptions
- The system SHALL determine record status automatically

**FR-014: Database Integration**
- The system SHALL create database records from validated emails
- The system SHALL maintain data integrity during email processing
- The system SHALL provide error handling for invalid email formats
- The system SHALL log processing results

### 3.5 Data Visualization and Reporting

#### 3.5.1 Description
The system provides comprehensive data visualization and reporting capabilities using matplotlib and other visualization tools.

#### 3.5.2 Functional Requirements

**FR-015: Chart Generation**
- The system SHALL generate burndown charts for project tracking
- The system SHALL create milestone progress charts
- The system SHALL provide status distribution charts
- The system SHALL support interactive chart features

**FR-016: Data Export**
- The system SHALL export data to Excel format (.xlsx)
- The system SHALL export data to CSV format
- The system SHALL support custom data filtering for exports
- The system SHALL maintain data formatting during export

**FR-017: Report Generation**
- The system SHALL generate summary reports
- The system SHALL create detailed analysis reports
- The system SHALL provide custom report templates
- The system SHALL support scheduled report generation

### 3.6 User Interface Management

#### 3.6.1 Description
The system provides both graphical and console user interfaces with comprehensive functionality.

#### 3.6.2 Functional Requirements

**FR-018: GUI Interface**
- The system SHALL provide a modern, responsive GUI using Tkinter
- The system SHALL support window resizing and layout management
- The system SHALL provide intuitive navigation and menu systems
- The system SHALL support keyboard shortcuts and accessibility features

**FR-019: Console Interface**
- The system SHALL provide a command-line interface for batch operations
- The system SHALL support interactive console sessions
- The system SHALL provide clear error messages and help text
- The system SHALL support script automation

**FR-020: Integration**
- The system SHALL integrate with VBA/Excel workflows
- The system SHALL support Python-VBA bridge functionality
- The system SHALL maintain compatibility with existing systems
- The system SHALL provide API endpoints for external integration

---

## 4. External Interface Requirements

### 4.1 User Interfaces

#### 4.1.1 GUI Interface
- **Technology**: Tkinter (Python standard library)
- **Layout**: Responsive design with tabbed interface
- **Navigation**: Menu bar, toolbar, and context menus
- **Data Display**: TreeView widgets with sorting and filtering
- **Forms**: Input dialogs and data entry forms
- **Charts**: Matplotlib integration for data visualization

#### 4.1.2 Console Interface
- **Technology**: Python standard input/output
- **Format**: Text-based interactive prompts
- **Navigation**: Numbered menu selections and text input
- **Output**: Formatted text reports and summaries
- **Error Handling**: Clear error messages and recovery options

### 4.2 Hardware Interfaces
- **Minimum Requirements**:
  - CPU: Intel Core i3 or equivalent
  - RAM: 4GB minimum, 8GB recommended
  - Storage: 500MB for application and data
  - Display: 1024x768 minimum resolution
- **Recommended Requirements**:
  - CPU: Intel Core i5 or equivalent
  - RAM: 8GB or more
  - Storage: 2GB for application and data
  - Display: 1920x1080 or higher resolution

### 4.3 Software Interfaces
- **Operating System**: Windows 10/11
- **Python Runtime**: Python 3.11 or higher
- **Database**: SQLite 3.x, Microsoft Access (legacy)
- **Dependencies**: See requirements.txt for complete list
- **Integration**: VBA/Excel compatibility layer

### 4.4 Communications Interfaces
- **Email Processing**: SMTP/POP3 for email integration
- **File System**: Local file system access for data storage
- **Network**: Optional network connectivity for updates and reporting

---

## 5. Non-Functional Requirements

### 5.1 Performance Requirements

#### 5.1.1 Response Time
- **GUI Operations**: All GUI operations SHALL complete within 2 seconds
- **Database Queries**: Simple queries SHALL complete within 1 second
- **Complex Queries**: Complex queries SHALL complete within 10 seconds
- **Data Export**: Export operations SHALL complete within 30 seconds for datasets up to 10,000 records

#### 5.1.2 Throughput
- **Data Processing**: System SHALL process up to 1,000 records per minute
- **Concurrent Users**: System SHALL support single-user operation
- **Database Operations**: System SHALL handle up to 100,000 records efficiently

#### 5.1.3 Resource Utilization
- **Memory Usage**: System SHALL use no more than 512MB RAM during normal operation
- **CPU Usage**: System SHALL not exceed 50% CPU utilization during normal operation
- **Disk Space**: System SHALL require no more than 1GB disk space for installation

### 5.2 Security Requirements

#### 5.2.1 Data Protection
- **Data Encryption**: Sensitive data SHALL be encrypted at rest
- **Access Control**: System SHALL implement user authentication
- **Data Integrity**: System SHALL validate data integrity before database operations
- **Audit Trail**: System SHALL maintain logs of all data modifications

#### 5.2.2 Error Handling
- **Input Validation**: System SHALL validate all user inputs
- **Error Recovery**: System SHALL provide graceful error recovery
- **Data Backup**: System SHALL maintain automatic data backups
- **Exception Handling**: System SHALL handle all exceptions gracefully

### 5.3 Reliability Requirements

#### 5.3.1 Availability
- **Uptime**: System SHALL maintain 99% availability during business hours
- **Recovery Time**: System SHALL recover from failures within 5 minutes
- **Data Loss**: System SHALL prevent data loss through backup mechanisms

#### 5.3.2 Fault Tolerance
- **Error Detection**: System SHALL detect and report errors immediately
- **Graceful Degradation**: System SHALL continue operating with reduced functionality during partial failures
- **Data Consistency**: System SHALL maintain data consistency during concurrent operations

### 5.4 Usability Requirements

#### 5.4.1 User Experience
- **Learning Curve**: New users SHALL be able to perform basic operations within 30 minutes
- **Help System**: System SHALL provide comprehensive help documentation
- **Error Messages**: System SHALL provide clear, actionable error messages
- **User Interface**: System SHALL follow standard Windows UI conventions

#### 5.4.2 Accessibility
- **Keyboard Navigation**: System SHALL support full keyboard navigation
- **Screen Reader**: System SHALL be compatible with screen readers
- **High Contrast**: System SHALL support high contrast display modes
- **Font Scaling**: System SHALL support font size adjustments

### 5.5 Maintainability Requirements

#### 5.5.1 Code Quality
- **Documentation**: All code SHALL be properly documented
- **Modularity**: System SHALL be designed with modular architecture
- **Testing**: System SHALL include comprehensive unit tests
- **Version Control**: System SHALL use version control for all code

#### 5.5.2 Extensibility
- **Plugin Architecture**: System SHALL support plugin extensions
- **API Design**: System SHALL provide well-defined APIs
- **Configuration**: System SHALL support external configuration files
- **Customization**: System SHALL allow user customization of interface elements

---

## 6. Other Requirements

### 6.1 Regulatory Requirements
- **Data Privacy**: System SHALL comply with applicable data privacy regulations
- **Data Retention**: System SHALL support configurable data retention policies
- **Audit Compliance**: System SHALL maintain audit trails for compliance requirements

### 6.2 Legal Requirements
- **Software Licensing**: System SHALL use only properly licensed software components
- **Intellectual Property**: System SHALL respect intellectual property rights
- **Export Control**: System SHALL comply with applicable export control regulations

### 6.3 Standards Compliance
- **Coding Standards**: System SHALL follow Python PEP 8 coding standards
- **Database Standards**: System SHALL follow SQL standards for database operations
- **UI Standards**: System SHALL follow Windows UI design guidelines

### 6.4 Installation and Deployment
- **Installation**: System SHALL provide automated installation process
- **Configuration**: System SHALL support configuration through setup wizard
- **Updates**: System SHALL support automated updates
- **Uninstallation**: System SHALL provide clean uninstallation process

---

## 7. Appendices

### 7.1 Dependencies
The system requires the following Python packages (as specified in requirements.txt):
- pandas>=1.3.0
- openpyxl>=3.0.0
- matplotlib>=3.5.0
- numpy>=1.21.0
- Pillow>=8.0.0
- cairosvg>=2.5.0

### 7.2 Database Schema
The system uses the following main database tables:
- **BASHFlowSandbox**: Main workflow entries
- **Milestones**: Milestone definitions and tracking
- **Subtasks**: Subtask definitions and dependencies
- **Stakeholders**: Stakeholder information (Name, Home State, Role, Team, Email, Phone)
- **EmailTemplates**: Email template definitions

### 7.3 Configuration Files
The system uses the following configuration files:
- **requirements.txt**: Python package dependencies
- **CustomRibbon.xml**: Excel ribbon customization
- **Database configuration**: SQLite database settings

### 7.4 Error Codes
The system uses the following error code categories:
- **1000-1999**: Database errors
- **2000-2999**: Validation errors
- **3000-3999**: File I/O errors
- **4000-4999**: Network errors
- **5000-5999**: System errors

### 7.5 Glossary
- **CLLI**: Common Language Location Identifier - a standardized code for telecommunications locations
- **LATA**: Local Access and Transport Area - a geographic area for telecommunications regulation
- **Milestone**: A significant point in a project timeline
- **Subtask**: A smaller task that contributes to a milestone
- **Workflow**: A sequence of processes or tasks
- **Stakeholder**: A person or organization with interest in the project

---

**Document End**

*This Software Requirements Specification document provides a comprehensive overview of the Degrow Workflow Manager System requirements. For technical implementation details, refer to the system design documents and code documentation.*
