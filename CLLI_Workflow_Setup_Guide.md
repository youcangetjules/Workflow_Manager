# CLLI Workflow Manager Setup Guide

## Overview
This VBA script automates CLLI (Common Language Location Identifier) workflow tasks by processing emails with specific subject line formats and populating them into a database for project tracking.

## Features
- ✅ Automatic email processing from Outlook inbox
- ✅ Subject line format validation using regex patterns
- ✅ CLLI project tracking with Milestone and Event management
- ✅ Database integration with error handling
- ✅ Comprehensive logging system
- ✅ Smart status detection for Events

## Subject Line Formats

The script recognizes these preset formats:

| Type | Format | Example |
|------|--------|---------|
| **CLLI** | `CLLI-####-YYYY-MM-DD-Description` | `CLLI-1234-2024-01-15-Network Infrastructure` |
| **Milestone** | `MS-####-YYYY-MM-DD-Description` | `MS-5678-2024-01-15-Phase 1 Completion` |
| **Event** | `Event-####-YYYY-MM-DD-Description` | `Event-9012-2024-01-15-Task completed successfully` |

## Event Status Types

Events automatically detect status based on keywords in the description:

| Status | Keywords | Example |
|--------|----------|---------|
| **Completed** | "completed" | `Event-3001-2024-01-17-Task completed successfully` |
| **Blocked** | "blocked" | `Event-3002-2024-01-18-Implementation blocked by vendor` |
| **Delayed** | "delayed" | `Event-3003-2024-01-19-Timeline delayed due to resources` |
| **Follow-up Required** | "follow-up", "followup" | `Event-3004-2024-01-20-Follow-up required with client` |
| **Urgency 1** | "urgency 1", "urgent 1" | `Event-3005-2024-01-21-Urgency 1 critical issue` |
| **Urgency 2** | "urgency 2", "urgent 2" | `Event-3006-2024-01-22-Urgency 2 performance issue` |
| **Urgency 3** | "urgency 3", "urgent 3" | `Event-3007-2024-01-23-Urgency 3 minor issue` |

## Setup Instructions

### 1. Database Setup
1. Create an Access database or use existing database
2. Update the `DB_CONNECTION_STRING` in the VBA script:
   ```vba
   Private Const DB_CONNECTION_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\Your\Database.accdb;"
   ```
3. Run the `CreateDatabaseTable()` subroutine once to create the required table

### 2. File Paths Configuration
Update these constants in the script:
```vba
Private Const DB_CONNECTION_STRING As String = "Your database path"
Private Const LOG_FILE_PATH As String = "Your log file path"
```

### 3. Outlook Setup
- Ensure Microsoft Outlook is installed and configured
- The script will access the default inbox folder

### 4. VBA Environment Setup
1. Open Excel or Word
2. Press `Alt + F11` to open VBA Editor
3. Insert a new module
4. Copy the entire `BASHWorkflowManager.bas` code into the module
5. Save the file as a macro-enabled document

## Usage

### Running the Workflow
1. Open the VBA editor (`Alt + F11`)
2. Run the main subroutine: `ProcessWorkflowEmails()`

### Testing and Utilities
- `TestDatabaseConnection()` - Test database connectivity
- `ShowFormatExamples()` - Display valid subject line formats
- `CreateDatabaseTable()` - Set up database table (run once)
- `GenerateSampleData()` - Generate sample subject lines for testing

## Database Schema

The script creates/uses this table structure:

| Column | Type | Description |
|--------|------|-------------|
| ID | AutoNumber | Primary key |
| Subject | Text | Original email subject |
| Type | Text | CLLI, MS, or Event |
| CLLINumber | Text | 4-digit CLLI identifier |
| MSNumber | Text | 4-digit Milestone identifier |
| EventNumber | Text | 4-digit Event identifier |
| Date | Text | YYYY-MM-DD format |
| Description | Text | Task/project description |
| Status | Text | Auto-detected status for Events |
| CreatedDate | DateTime | When record was created |

## Sample Subject Lines

Here are examples of valid subject lines for testing:

```
CLLI-1001-2024-01-15-Network Infrastructure Project
MS-2001-2024-01-16-Phase 1 Network Design Complete
Event-3001-2024-01-17-Task completed successfully
Event-3002-2024-01-18-Implementation blocked by vendor delay
Event-3003-2024-01-19-Timeline delayed due to resource constraints
Event-3004-2024-01-20-Follow-up required with client approval
Event-3005-2024-01-21-Urgency 1 critical security issue identified
MS-2002-2024-01-22-Phase 2 Implementation Milestone
Event-3006-2024-01-23-Urgency 2 performance issue detected
CLLI-1002-2024-01-24-Customer Portal Development
```

## Customization Options

### Adding New Event Status Types
1. Add new status constant:
   ```vba
   Private Const STATUS_CUSTOM As String = "Custom Status"
   ```
2. Update the `ExtractEventStatus()` function to detect new keywords
3. Add the new keyword detection logic

### Modifying CLLI Number Formats
1. Update the regex patterns in `ValidateSubjectLine()` function
2. Modify the pattern to match your specific CLLI format requirements
3. Update validation logic accordingly

### Changing Database Fields
1. Update the column constants at the top of the script
2. Modify the `CreateDatabaseTable()` SQL statement
3. Update the `InsertIntoDatabase()` function

## Troubleshooting

### Common Issues

1. **Database Connection Error**
   - Verify database path in `DB_CONNECTION_STRING`
   - Ensure database file exists and is accessible
   - Check OLEDB provider installation

2. **Subject Line Validation Fails**
   - Use `ShowFormatExamples()` to verify correct format
   - Check for special characters in description
   - Ensure date format is YYYY-MM-DD
   - Verify CLLI/MS/Event prefix is correct

3. **Event Status Not Detected**
   - Check that status keywords are in the description
   - Keywords are case-insensitive
   - Use exact keywords: "completed", "blocked", "delayed", etc.

4. **Outlook Access Error**
   - Ensure Outlook is running
   - Check Outlook security settings for VBA access
   - Verify Outlook profile is configured

### Log File Location
Check the log file for detailed error messages and processing information:
- Default location: `C:\Path\To\Your\WorkflowLog.txt`
- Contains timestamps and detailed error descriptions

## Security Considerations

- SQL injection protection through string escaping
- Error handling prevents script crashes
- Logging provides audit trail
- Database connection uses parameterized approach

## Performance Notes

- Processes up to 100 emails per run (configurable)
- Sorts emails by received time (newest first)
- Includes connection cleanup and error recovery
- Logging is lightweight and non-blocking

## Workflow Process

1. **Email Processing**: Script reads emails from Outlook inbox
2. **Format Validation**: Validates subject line against CLLI/MS/Event patterns
3. **Status Detection**: For Events, automatically detects status from description
4. **Database Storage**: Inserts validated data into structured database
5. **Logging**: Records all processing activities and errors

## Support

For issues or customization requests:
1. Check the log file for specific error messages
2. Verify all configuration constants are set correctly
3. Test database connection separately
4. Ensure subject lines match the exact format requirements
5. Use `GenerateSampleData()` to test with known valid formats
