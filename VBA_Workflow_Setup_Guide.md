# VBA Workflow Manager Setup Guide

## Overview
This VBA script automates workflow tasks by processing emails with specific subject line formats and populating them into a database.

## Features
- ✅ Automatic email processing from Outlook inbox
- ✅ Subject line format validation using regex patterns
- ✅ Database integration with error handling
- ✅ Comprehensive logging system
- ✅ Multiple task type support (PROJ, TASK, MEET, ISSUE)

## Subject Line Formats

The script recognizes these preset formats:

| Type | Format | Example |
|------|--------|---------|
| **Project** | `PROJ-####-YYYY-MM-DD-Description` | `PROJ-1234-2024-01-15-Website Development` |
| **Task** | `TASK-####-YYYY-MM-DD-Description` | `TASK-5678-2024-01-15-Review Documentation` |
| **Meeting** | `MEET-####-YYYY-MM-DD-Description` | `MEET-9012-2024-01-15-Weekly Standup` |
| **Issue** | `ISSUE-####-YYYY-MM-DD-Description` | `ISSUE-3456-2024-01-15-Bug in Login System` |

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

### Testing
- `TestDatabaseConnection()` - Test database connectivity
- `ShowFormatExamples()` - Display valid subject line formats
- `CreateDatabaseTable()` - Set up database table (run once)

## Database Schema

The script creates/uses this table structure:

| Column | Type | Description |
|--------|------|-------------|
| ID | AutoNumber | Primary key |
| Subject | Text | Original email subject |
| Type | Text | PROJ, TASK, MEET, or ISSUE |
| Number | Text | 4-digit identifier |
| Date | Text | YYYY-MM-DD format |
| Description | Text | Task description |
| Status | Text | Default: "New" |
| CreatedDate | DateTime | When record was created |

## Customization Options

### Adding New Subject Line Formats
1. Add new format constant:
   ```vba
   Private Const FORMAT_CUSTOM As String = "CUSTOM-####-YYYY-MM-DD-Description"
   ```
2. Add validation pattern in `ValidateSubjectLine()` function
3. Update the regex pattern matching logic

### Modifying Database Fields
1. Update the `TABLE_NAME` and column constants
2. Modify the `CreateDatabaseTable()` SQL statement
3. Update the `InsertIntoDatabase()` function

### Changing Processing Logic
- Modify `ProcessEmailItem()` for different email handling
- Update `CreateDatabaseRecord()` for different data mapping
- Customize logging in `LogMessage()` function

## Troubleshooting

### Common Issues

1. **Database Connection Error**
   - Verify database path in `DB_CONNECTION_STRING`
   - Ensure database file exists and is accessible
   - Check OLEDB provider installation

2. **Outlook Access Error**
   - Ensure Outlook is running
   - Check Outlook security settings for VBA access
   - Verify Outlook profile is configured

3. **Subject Line Validation Fails**
   - Use `ShowFormatExamples()` to verify correct format
   - Check for special characters in description
   - Ensure date format is YYYY-MM-DD

4. **Permission Errors**
   - Run Excel/Word as Administrator
   - Check file permissions for log and database files
   - Verify VBA macro security settings

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

## Support

For issues or customization requests:
1. Check the log file for specific error messages
2. Verify all configuration constants are set correctly
3. Test database connection separately
4. Ensure subject lines match the exact format requirements
