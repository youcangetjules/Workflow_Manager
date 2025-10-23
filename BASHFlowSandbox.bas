Attribute VB_Name = "BASHFlowSandbox"
'===============================================================================
' BASH Flow Management - SANDBOX VERSION
' Isolated testing environment that won't interfere with existing VBA code
'===============================================================================

Option Explicit

'===============================================================================
' Data Structures - Must be defined first
'===============================================================================

' Validation Result Class
Public Type validationResult
    IsValid As Boolean
    RecordType As String
    CLLINumber As String
    MSNumber As String
    EventNumber As String
    RecordDate As String
    Description As String
    Status As String
    ErrorMessage As String
End Type

' Database Record Class
Public Type DatabaseRecord
    State As String
    RecordType As String
    CLLINumber As String
    MSNumber As String
    EventNumber As String
    RecordDate As String
    Description As String
    Status As String
    CreatedDate As Date
End Type

'===============================================================================
' Configuration Constants
'===============================================================================

' SANDBOX Configuration Constants - All prefixed with SANDBOX_
Private Const SANDBOX_DB_CONNECTION_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
Private Const SANDBOX_LOG_FILE_PATH As String = "C:\BASHFlowSandbox\SandboxLog.txt"
Private Const SANDBOX_BACKUP_FOLDER As String = "C:\BASHFlowSandbox\Backups\"

' Subject Line Format Patterns (same as main version)
Private Const FORMAT_CLLI As String = "CLLI-####-YYYY-MM-DD-Description"
Private Const FORMAT_MS As String = "MS-####-YYYY-MM-DD-Description"
Private Const FORMAT_EVENT As String = "Event-####-YYYY-MM-DD-Description"

' Event Status Types
Private Const STATUS_COMPLETED As String = "Completed"
Private Const STATUS_BLOCKED As String = "Blocked"
Private Const STATUS_DELAYED As String = "Delayed"
Private Const STATUS_FOLLOWUP As String = "Follow-up required"
Private Const STATUS_URGENCY1 As String = "Urgency 1"
Private Const STATUS_URGENCY2 As String = "Urgency 2"
Private Const STATUS_URGENCY3 As String = "Urgency 3"

' SANDBOX Database Table Structure - Uses different table name
Private Const SANDBOX_TABLE_NAME As String = "BASHFlowSandbox"
Private Const COLUMN_ID As String = "ID"
Private Const COLUMN_STATE As String = "[State]"
Private Const COLUMN_CLLI_NUMBER As String = "[CLLI Number]"
Private Const COLUMN_DATE As String = "[Date]"
Private Const COLUMN_TYPE As String = "[Type]"
Private Const COLUMN_MS_NUMBER As String = "[MS Number]"
Private Const COLUMN_EVENT_NUMBER As String = "[Event]"
Private Const COLUMN_DESCRIPTION As String = "[Description]"
Private Const COLUMN_STATUS As String = "[Status]"
Private Const COLUMN_CREATED As String = "[Date Created]"

' Global ribbon object for sandbox
Private sandboxRibbonUI As IRibbonUI

'===============================================================================
' SANDBOX Setup Functions
'===============================================================================

' Simple test to verify VBA is working
Sub SimpleTest()
    MsgBox "VBA is working! This is a simple test.", vbInformation, "VBA Test"
End Sub

' Quick reset function (alternative to ResetSandbox)
Sub QuickReset()
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists("C:\BASHFlowSandbox") Then
        fso.DeleteFolder "C:\BASHFlowSandbox", True
        MsgBox "✅ Sandbox folder deleted. Run InitializeSandbox() to create fresh sandbox.", vbInformation
    Else
        MsgBox "ℹ️ No sandbox folder found to delete.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Error: " & Err.Description, vbCritical
End Sub

' Simple database setup (run this after InitializeSandbox)
Sub SetupDatabase()
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    Dim dbPath As String
    
    On Error GoTo ErrorHandler
    
    MsgBox "Setting up database table structure...", vbInformation, "Database Setup"
    
    ' Get database path
    dbPath = Replace(SANDBOX_DB_CONNECTION_STRING, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=", "")
    
    ' Connect to existing database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Create table with correct structure in specified order
    sql = "CREATE TABLE " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_ID & " AUTOINCREMENT PRIMARY KEY, " & _
          COLUMN_STATE & " TEXT(10), " & _
          COLUMN_CLLI_NUMBER & " TEXT(10), " & _
          COLUMN_DATE & " TEXT(10) NOT NULL, " & _
          COLUMN_TYPE & " TEXT(10) NOT NULL, " & _
          COLUMN_MS_NUMBER & " TEXT(10), " & _
          COLUMN_EVENT_NUMBER & " TEXT(10), " & _
          COLUMN_DESCRIPTION & " TEXT(255), " & _
          COLUMN_STATUS & " TEXT(20), " & _
          COLUMN_CREATED & " DATETIME DEFAULT NOW())"
    
    cmd.CommandText = sql
    cmd.Execute
    
    ' Insert sample record
    sql = "INSERT INTO " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_STATE & ", " & COLUMN_CLLI_NUMBER & ", " & COLUMN_DATE & ", " & COLUMN_TYPE & ", " & _
          COLUMN_DESCRIPTION & ", " & COLUMN_STATUS & ") VALUES " & _
          "('TX', '1001', '2024-01-15', 'CLLI', 'Network Infrastructure Project (SANDBOX)', 'Active')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    MsgBox "✅ Database table created successfully!" & vbCrLf & _
           "Table: " & SANDBOX_TABLE_NAME & vbCrLf & _
           "Sample record inserted." & vbCrLf & vbCrLf & _
           "You can now test the workflow functions!", vbInformation, "Database Ready"
    
    LogSandboxMessage "Database table created successfully with sample data"
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    MsgBox "❌ Error setting up database: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Database Error"
    LogSandboxMessage "Error setting up database: " & Err.Description
End Sub

' Fix database table structure
Sub FixDatabaseTable()
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    MsgBox "Fixing database table structure...", vbInformation, "Database Fix"
    
    ' Connect to existing database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Drop existing table if it exists (Access doesn't support IF EXISTS)
    On Error Resume Next
    sql = "DROP TABLE " & SANDBOX_TABLE_NAME
    cmd.CommandText = sql
    cmd.Execute
    On Error GoTo ErrorHandler
    
    ' Create table with correct structure in specified order
    sql = "CREATE TABLE " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_ID & " AUTOINCREMENT PRIMARY KEY, " & _
          COLUMN_STATE & " TEXT(10), " & _
          COLUMN_CLLI_NUMBER & " TEXT(10), " & _
          COLUMN_DATE & " TEXT(10) NOT NULL, " & _
          COLUMN_TYPE & " TEXT(10) NOT NULL, " & _
          COLUMN_MS_NUMBER & " TEXT(10), " & _
          COLUMN_EVENT_NUMBER & " TEXT(10), " & _
          COLUMN_DESCRIPTION & " TEXT(255), " & _
          COLUMN_STATUS & " TEXT(20), " & _
          COLUMN_CREATED & " DATETIME DEFAULT NOW())"
    
    cmd.CommandText = sql
    cmd.Execute
    
    ' Insert sample data
    sql = "INSERT INTO " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_STATE & ", " & COLUMN_CLLI_NUMBER & ", " & COLUMN_DATE & ", " & COLUMN_TYPE & ", " & _
          COLUMN_DESCRIPTION & ", " & COLUMN_STATUS & ") VALUES " & _
          "('TX', '1001', '2024-01-15', 'CLLI', 'Network Infrastructure Project (SANDBOX)', 'Active')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    MsgBox "✅ Database table structure fixed successfully!" & vbCrLf & _
           "Table: " & SANDBOX_TABLE_NAME & vbCrLf & _
           "Sample record inserted.", vbInformation, "Database Fixed"
    
    LogSandboxMessage "Database table structure fixed successfully"
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    MsgBox "❌ Error fixing database: " & Err.Description, vbCritical
    LogSandboxMessage "Error fixing database: " & Err.Description
End Sub

' Reset sandbox to clean state (removes existing files)
Sub ResetSandbox()
    Dim fso As Object
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will delete all existing sandbox files and start fresh." & vbCrLf & vbCrLf & _
                     "This will remove:" & vbCrLf & _
                     "• C:\BASHFlowSandbox\ folder and all contents" & vbCrLf & _
                     "• TestDatabase.accdb" & vbCrLf & _
                     "• SandboxLog.txt" & vbCrLf & _
                     "• All backup files" & vbCrLf & vbCrLf & _
                     "Are you sure you want to proceed?", vbYesNo + vbExclamation, "Reset Sandbox")
    
    If response = vbYes Then
        On Error GoTo ErrorHandler
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Delete sandbox folder and all contents
        If fso.FolderExists("C:\BASHFlowSandbox") Then
            fso.DeleteFolder "C:\BASHFlowSandbox", True
            MsgBox "✅ Sandbox reset completed. Run InitializeSandbox() to create fresh sandbox.", vbInformation
        Else
            MsgBox "ℹ️ No sandbox found to reset.", vbInformation
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Error resetting sandbox: " & Err.Description & vbCrLf & _
           "You may need to manually delete C:\BASHFlowSandbox", vbCritical
End Sub

' Simple test function to verify sandbox setup
Sub TestSandboxSetup()
    Dim fso As Object
    Dim message As String
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    message = "BASH Flow Sandbox Status:" & vbCrLf & vbCrLf
    
    ' Check main folder
    If fso.FolderExists("C:\BASHFlowSandbox") Then
        message = message & "✅ Main folder: C:\BASHFlowSandbox exists" & vbCrLf
    Else
        message = message & "❌ Main folder: C:\BASHFlowSandbox missing" & vbCrLf
    End If
    
    ' Check backup folder
    If fso.FolderExists(SANDBOX_BACKUP_FOLDER) Then
        message = message & "✅ Backup folder: " & SANDBOX_BACKUP_FOLDER & " exists" & vbCrLf
    Else
        message = message & "❌ Backup folder: " & SANDBOX_BACKUP_FOLDER & " missing" & vbCrLf
    End If
    
    ' Check database file
    If fso.FileExists("C:\BASHFlowSandbox\TestDatabase.accdb") Then
        message = message & "✅ Database: TestDatabase.accdb exists" & vbCrLf
    Else
        message = message & "❌ Database: TestDatabase.accdb missing" & vbCrLf
    End If
    
    ' Check log file
    If fso.FileExists(SANDBOX_LOG_FILE_PATH) Then
        message = message & "✅ Log file: SandboxLog.txt exists" & vbCrLf
    Else
        message = message & "❌ Log file: SandboxLog.txt missing" & vbCrLf
    End If
    
    message = message & vbCrLf & "Run InitializeSandbox() if any items are missing."
    
    MsgBox message, vbInformation, "Sandbox Status Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking sandbox status: " & Err.Description, vbCritical
End Sub

' Initialize sandbox environment
Sub InitializeSandbox()
    Dim fso As Object
    Dim sandboxFolder As Object
    Dim logFile As Object
    Dim timestamp As String
    
    ' Show immediate feedback
    MsgBox "Starting sandbox initialization...", vbInformation, "Sandbox Init"
    
    On Error GoTo ErrorHandler
    
    ' Create timestamp for logging
    timestamp = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    ' Show progress
    MsgBox "Step 1: Creating folder structure...", vbInformation, "Progress"
    
    ' Create sandbox folder structure
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create main sandbox folder
    If Not fso.FolderExists("C:\BASHFlowSandbox") Then
        Set sandboxFolder = fso.CreateFolder("C:\BASHFlowSandbox")
        MsgBox "✅ Created main folder: C:\BASHFlowSandbox", vbInformation, "Progress"
    Else
        MsgBox "ℹ️ Main folder already exists: C:\BASHFlowSandbox", vbInformation, "Progress"
    End If
    
    ' Create backups folder
    If Not fso.FolderExists(SANDBOX_BACKUP_FOLDER) Then
        Set sandboxFolder = fso.CreateFolder(SANDBOX_BACKUP_FOLDER)
        MsgBox "✅ Created backup folder: " & SANDBOX_BACKUP_FOLDER, vbInformation, "Progress"
    Else
        MsgBox "ℹ️ Backup folder already exists: " & SANDBOX_BACKUP_FOLDER, vbInformation, "Progress"
    End If
    
    ' Show progress
    MsgBox "Step 2: Creating log file...", vbInformation, "Progress"
    
    ' Create initial log file
    Set logFile = fso.CreateTextFile(SANDBOX_LOG_FILE_PATH, True)
    logFile.WriteLine "[SANDBOX] " & timestamp & " - Initializing BASH Flow Management Sandbox..."
    logFile.WriteLine "[SANDBOX] " & timestamp & " - Created sandbox folder: C:\BASHFlowSandbox"
    logFile.WriteLine "[SANDBOX] " & timestamp & " - Created backup folder: " & SANDBOX_BACKUP_FOLDER
    logFile.Close
    
    MsgBox "✅ Log file created: " & SANDBOX_LOG_FILE_PATH, vbInformation, "Progress"
    
    ' Show progress
    MsgBox "Step 3: Creating database...", vbInformation, "Progress"
    
    ' Create test database
    CreateSandboxDatabase
    
    MsgBox "✅ Database creation attempted", vbInformation, "Progress"
    
    ' Show progress
    MsgBox "Step 4: Creating sample data...", vbInformation, "Progress"
    
    ' Create sample data
    CreateSandboxSampleData
    
    MsgBox "✅ Sample data created", vbInformation, "Progress"
    
    ' Log completion
    LogSandboxMessage "BASH Flow Management Sandbox initialized successfully!"
    
    MsgBox "🎉 BASH Flow Management Sandbox initialized successfully!" & vbCrLf & _
           "Sandbox Location: C:\BASHFlowSandbox" & vbCrLf & _
           "Database: TestDatabase.accdb" & vbCrLf & _
           "Log File: SandboxLog.txt" & vbCrLf & vbCrLf & _
           "You can now test the sandbox functions!", vbInformation, "Sandbox Ready"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Error initializing sandbox: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: " & Err.Source, vbCritical, "Sandbox Error"
End Sub

' Create sandbox database
Private Sub CreateSandboxDatabase()
    Dim fso As Object
    Dim dbPath As String
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Get database path
    dbPath = Replace(SANDBOX_DB_CONNECTION_STRING, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=", "")
    
    ' Check if database already exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dbPath) Then
        MsgBox "ℹ️ Database already exists: " & dbPath & vbCrLf & _
               "Skipping database creation and proceeding with setup.", vbInformation, "Database Check"
        LogSandboxMessage "Database already exists, skipping creation"
        Exit Sub
    End If
    
    ' Try to create database using ADOX (if available)
    On Error Resume Next
    Set conn = CreateObject("ADOX.Catalog")
    If Err.Number <> 0 Then
        MsgBox "⚠️ ADOX not available, creating CSV database instead", vbInformation, "Database Creation"
        ' Create a CSV file as database placeholder
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set conn = fso.CreateTextFile(dbPath, True)
        conn.WriteLine "ID,Subject,Type,CLLINumber,MSNumber,EventNumber,Date,Description,Status,CreatedDate"
        conn.WriteLine "1,Sample Record,CLLI,1001,,,2024-01-15,Sample CLLI Project,Active," & Format(Now(), "yyyy-mm-dd hh:nn:ss")
        conn.Close
        LogSandboxMessage "Created CSV database: " & dbPath
        MsgBox "✅ CSV database created successfully: " & dbPath, vbInformation, "CSV Database Created"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Create new Access database using ADOX
    conn.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    Set conn = Nothing
    
    ' Connect to the new database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Create sandbox table with simplified SQL
    sql = "CREATE TABLE " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_ID & " AUTOINCREMENT PRIMARY KEY, " & _
          COLUMN_STATE & " TEXT(10), " & _
          COLUMN_CLLI_NUMBER & " TEXT(10), " & _
          COLUMN_DATE & " TEXT(10) NOT NULL, " & _
          COLUMN_TYPE & " TEXT(10) NOT NULL, " & _
          COLUMN_MS_NUMBER & " TEXT(10), " & _
          COLUMN_EVENT_NUMBER & " TEXT(10), " & _
          COLUMN_DESCRIPTION & " TEXT(255), " & _
          COLUMN_STATUS & " TEXT(20), " & _
          COLUMN_CREATED & " DATETIME DEFAULT NOW())"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    LogSandboxMessage "Sandbox database and table created successfully"
    MsgBox "✅ Database created successfully: " & dbPath, vbInformation, "Database Created"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Error creating database: " & Err.Description & vbCrLf & vbCrLf & _
           "This might be due to:" & vbCrLf & _
           "• Missing Access Database Engine" & vbCrLf & _
           "• Permission issues" & vbCrLf & _
           "• ADOX not available" & vbCrLf & vbCrLf & _
           "The sandbox will continue with limited functionality.", vbExclamation, "Database Error"
    LogSandboxMessage "Error creating sandbox database: " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
End Sub

' Create sample data for testing
Private Sub CreateSandboxSampleData()
    Dim sampleData As Variant
    Dim i As Integer
    
    sampleData = Array( _
        "CLLI-1001-2024-01-15-Network Infrastructure Project (SANDBOX)", _
        "MS-2001-2024-01-16-Phase 1 Network Design Complete (SANDBOX)", _
        "Event-3001-2024-01-17-Task completed successfully (SANDBOX)", _
        "Event-3002-2024-01-18-Implementation blocked by vendor delay (SANDBOX)", _
        "Event-3003-2024-01-19-Timeline delayed due to resource constraints (SANDBOX)", _
        "Event-3004-2024-01-20-Follow-up required with client approval (SANDBOX)", _
        "Event-3005-2024-01-21-Urgency 1 critical security issue identified (SANDBOX)", _
        "MS-2002-2024-01-22-Phase 2 Implementation Milestone (SANDBOX)", _
        "Event-3006-2024-01-23-Urgency 2 performance issue detected (SANDBOX)", _
        "CLLI-1002-2024-01-24-Customer Portal Development (SANDBOX)" _
    )
    
    LogSandboxMessage "Sample data created for testing:"
    For i = LBound(sampleData) To UBound(sampleData)
        LogSandboxMessage "Sample " & (i + 1) & ": " & sampleData(i)
    Next i
End Sub

'===============================================================================
' SANDBOX Main Processing Functions
'===============================================================================

' Process workflow emails in sandbox mode
Sub ProcessSandboxWorkflowEmails()
    Dim outlookApp As Object
    Dim namespace As Object
    Dim inbox As Object
    Dim Items As Object
    Dim item As Object
    Dim i As Long
    Dim processedCount As Long
    Dim errorCount As Long
    
    On Error GoTo ErrorHandler
    
    LogSandboxMessage "Starting SANDBOX workflow email processing..."
    
    ' Initialize Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    Set namespace = outlookApp.GetNamespace("MAPI")
    Set inbox = namespace.GetDefaultFolder(6) ' olFolderInbox
    Set Items = inbox.Items
    
    ' Sort items by received time (newest first)
    Items.Sort "[ReceivedTime]", True
    
    ' Process each email (limited to 10 for sandbox)
    For i = 1 To Items.Count
        Set item = Items(i)
        
        ' Check if item is an email
        If item.Class = 43 Then ' olMail
            If ProcessSandboxEmailItem(item) Then
                processedCount = processedCount + 1
            Else
                errorCount = errorCount + 1
            End If
        End If
        
        ' Limit sandbox processing to 10 emails
        If i > 10 Then Exit For
    Next i
    
    LogSandboxMessage "SANDBOX processing complete. Processed: " & processedCount & ", Errors: " & errorCount
    
    ' Cleanup
    Set item = Nothing
    Set Items = Nothing
    Set inbox = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
    
    MsgBox "SANDBOX Workflow processing complete!" & vbCrLf & _
           "Processed: " & processedCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & _
           "Check sandbox log for details.", vbInformation, "Sandbox Results"
    
    Exit Sub
    
ErrorHandler:
    LogSandboxMessage "Error in ProcessSandboxWorkflowEmails: " & Err.Description
    MsgBox "Sandbox Error: " & Err.Description, vbCritical
End Sub

' Process individual email item in sandbox
Private Function ProcessSandboxEmailItem(item As Object) As Boolean
    Dim subjectLine As String
    Dim validationResult As validationResult
    Dim dbRecord As DatabaseRecord
    
    On Error GoTo ErrorHandler
    
    subjectLine = Trim(item.Subject)
    
    ' Skip if no subject
    If Len(subjectLine) = 0 Then
        LogSandboxMessage "Skipping email with no subject line"
        ProcessSandboxEmailItem = True
        Exit Function
    End If
    
    ' Validate subject line format
    validationResult = ValidateSubjectLine(subjectLine)
    
    If validationResult.IsValid Then
        ' Create database record
        dbRecord = CreateDatabaseRecord(item, validationResult)
        
        ' Insert into sandbox database
        If InsertIntoSandboxDatabase(dbRecord) Then
            LogSandboxMessage "Successfully processed (SANDBOX): " & subjectLine
            ProcessSandboxEmailItem = True
        Else
            LogSandboxMessage "Failed to insert into sandbox database: " & subjectLine
            ProcessSandboxEmailItem = False
        End If
    Else
        LogSandboxMessage "Invalid subject line format (SANDBOX): " & subjectLine & " - " & validationResult.ErrorMessage
        ProcessSandboxEmailItem = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogSandboxMessage "Error processing sandbox email: " & Err.Description
    ProcessSandboxEmailItem = False
End Function

'===============================================================================
' SANDBOX Database Functions
'===============================================================================

' Insert record into sandbox database
Private Function InsertIntoSandboxDatabase(record As DatabaseRecord) As Boolean
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Create connection to sandbox database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    ' Create command
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Build SQL insert statement for sandbox table
    sql = "INSERT INTO " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_STATE & ", " & _
          COLUMN_CLLI_NUMBER & ", " & _
          COLUMN_DATE & ", " & _
          COLUMN_TYPE & ", " & _
          COLUMN_MS_NUMBER & ", " & _
          COLUMN_EVENT_NUMBER & ", " & _
          COLUMN_DESCRIPTION & ", " & _
          COLUMN_STATUS & ") VALUES (" & _
          "'" & EscapeString(record.State) & "', " & _
          "'" & EscapeString(record.CLLINumber) & "', " & _
          "'" & EscapeString(record.RecordDate) & "', " & _
          "'" & EscapeString(record.RecordType) & "', " & _
          "'" & EscapeString(record.MSNumber) & "', " & _
          "'" & EscapeString(record.EventNumber) & "', " & _
          "'" & EscapeString(record.Description) & "', " & _
          "'" & EscapeString(record.Status) & "')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    InsertIntoSandboxDatabase = True
    
    Exit Function
    
ErrorHandler:
    LogSandboxMessage "Sandbox database error: " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    InsertIntoSandboxDatabase = False
End Function

' Test sandbox database connection
Sub TestSandboxDatabaseConnection()
    Dim conn As Object
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    LogSandboxMessage "Sandbox database connection test: SUCCESS"
    MsgBox "Sandbox database connection successful!" & vbCrLf & _
           "Database: " & Replace(SANDBOX_DB_CONNECTION_STRING, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=", ""), vbInformation
    conn.Close
    
    Exit Sub
    
ErrorHandler:
    LogSandboxMessage "Sandbox database connection test: FAILED - " & Err.Description
    MsgBox "Sandbox database connection failed: " & Err.Description, vbCritical
End Sub

' Show sandbox workflow status
Sub ShowSandboxWorkflowStatus()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim statusMessage As String
    Dim totalCount As Long
    Dim clliCount As Long
    Dim msCount As Long
    Dim eventCount As Long
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Get counts from sandbox table
    sql = "SELECT COUNT(*) AS TotalCount FROM " & SANDBOX_TABLE_NAME
    rs.Open sql, conn
    totalCount = rs("TotalCount")
    rs.Close
    
    sql = "SELECT COUNT(*) AS CLLICount FROM " & SANDBOX_TABLE_NAME & " WHERE " & COLUMN_TYPE & " = 'CLLI'"
    rs.Open sql, conn
    clliCount = rs("CLLICount")
    rs.Close
    
    sql = "SELECT COUNT(*) AS MSCount FROM " & SANDBOX_TABLE_NAME & " WHERE " & COLUMN_TYPE & " = 'MS'"
    rs.Open sql, conn
    msCount = rs("MSCount")
    rs.Close
    
    sql = "SELECT COUNT(*) AS EventCount FROM " & SANDBOX_TABLE_NAME & " WHERE " & COLUMN_TYPE & " = 'Event'"
    rs.Open sql, conn
    eventCount = rs("EventCount")
    rs.Close
    
    conn.Close
    
    statusMessage = "BASH Flow Management SANDBOX Status:" & vbCrLf & vbCrLf & _
                   "Total Records: " & totalCount & vbCrLf & _
                   "CLLI Projects: " & clliCount & vbCrLf & _
                   "Milestones: " & msCount & vbCrLf & _
                   "Events: " & eventCount & vbCrLf & vbCrLf & _
                   "Database: SANDBOX (Test)" & vbCrLf & _
                   "Last Updated: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    LogSandboxMessage "Sandbox status displayed - Total: " & totalCount
    MsgBox statusMessage, vbInformation, "Sandbox Status"
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    LogSandboxMessage "Error retrieving sandbox status: " & Err.Description
    MsgBox "Error retrieving sandbox status: " & Err.Description, vbCritical
End Sub

'===============================================================================
' SANDBOX Utility Functions
'===============================================================================

' Log messages to sandbox log file
Private Sub LogSandboxMessage(message As String)
    Dim fso As Object
    Dim logFile As Object
    Dim timestamp As String
    
    On Error Resume Next
    
    timestamp = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(SANDBOX_LOG_FILE_PATH) Then
        Set logFile = fso.OpenTextFile(SANDBOX_LOG_FILE_PATH, 8, True) ' 8 = ForAppending
    Else
        Set logFile = fso.CreateTextFile(SANDBOX_LOG_FILE_PATH, True)
    End If
    
    logFile.WriteLine "[SANDBOX] " & timestamp & " - " & message
    logFile.Close
End Sub

' View sandbox log file
Sub ViewSandboxLogFile()
    Dim fso As Object
    Dim logFile As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(SANDBOX_LOG_FILE_PATH) Then
        Set logFile = fso.OpenTextFile(SANDBOX_LOG_FILE_PATH, 1, False) ' 1 = ForReading
        MsgBox logFile.ReadAll, vbInformation, "Sandbox Log File"
        logFile.Close
    Else
        MsgBox "Sandbox log file not found at: " & SANDBOX_LOG_FILE_PATH, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening sandbox log file: " & Err.Description, vbCritical
End Sub

' Export sandbox data to Excel
Sub ExportSandboxData()
    Dim conn As Object
    Dim rs As Object
    Dim excelApp As Object
    Dim workbook As Object
    Dim worksheet As Object
    Dim sql As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Create Excel application
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    Set workbook = excelApp.Workbooks.Add
    Set worksheet = workbook.Worksheets(1)
    
    ' Set up headers
    worksheet.Cells(1, 1).Value = "ID"
    worksheet.Cells(1, 2).Value = "State"
    worksheet.Cells(1, 3).Value = "CLLI Number"
    worksheet.Cells(1, 4).Value = "Date"
    worksheet.Cells(1, 5).Value = "Type"
    worksheet.Cells(1, 6).Value = "MS Number"
    worksheet.Cells(1, 7).Value = "Event"
    worksheet.Cells(1, 8).Value = "Description"
    worksheet.Cells(1, 9).Value = "Status"
    worksheet.Cells(1, 10).Value = "Date Created"
    
    ' Format headers
    With worksheet.Range("A1:J1")
        .Font.bold = True
        .Interior.Color = RGB(255, 200, 200) ' Light red for sandbox
    End With
    
    ' Add title
    worksheet.Cells(1, 1).Value = "BASH Flow Management - SANDBOX DATA"
    With worksheet.Range("A1:J1")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.bold = True
        .Font.Size = 14
    End With
    
    ' Connect to sandbox database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open SANDBOX_DB_CONNECTION_STRING
    
    Set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM " & SANDBOX_TABLE_NAME & " ORDER BY CreatedDate DESC"
    rs.Open sql, conn
    
    ' Export data starting from row 3
    i = 3
    Do While Not rs.EOF
        worksheet.Cells(i, 1).Value = rs(COLUMN_ID)
        worksheet.Cells(i, 2).Value = rs(COLUMN_STATE)
        worksheet.Cells(i, 3).Value = rs(COLUMN_CLLI_NUMBER)
        worksheet.Cells(i, 4).Value = rs(COLUMN_DATE)
        worksheet.Cells(i, 5).Value = rs(COLUMN_TYPE)
        worksheet.Cells(i, 6).Value = rs(COLUMN_MS_NUMBER)
        worksheet.Cells(i, 7).Value = rs(COLUMN_EVENT_NUMBER)
        worksheet.Cells(i, 8).Value = rs(COLUMN_DESCRIPTION)
        worksheet.Cells(i, 9).Value = rs(COLUMN_STATUS)
        worksheet.Cells(i, 10).Value = rs(COLUMN_CREATED)
        rs.MoveNext
        i = i + 1
    Loop
    
    ' Auto-fit columns
    worksheet.Columns.AutoFit
    
    ' Close database connection
    rs.Close
    conn.Close
    
    LogSandboxMessage "Sandbox data exported to Excel successfully"
    MsgBox "Sandbox workflow data exported successfully to Excel!" & vbCrLf & _
           "File is marked as SANDBOX data for easy identification.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    LogSandboxMessage "Error exporting sandbox data: " & Err.Description
    MsgBox "Error exporting sandbox data: " & Err.Description, vbCritical
End Sub

'===============================================================================
' SANDBOX Cleanup Functions
'===============================================================================

' Clean up sandbox environment
Sub CleanupSandbox()
    Dim fso As Object
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will permanently delete the BASH Flow Management sandbox environment." & vbCrLf & vbCrLf & _
                     "This includes:" & vbCrLf & _
                     "• C:\BASHFlowSandbox\ folder" & vbCrLf & _
                     "• TestDatabase.accdb" & vbCrLf & _
                     "• SandboxLog.txt" & vbCrLf & _
                     "• All backup files" & vbCrLf & vbCrLf & _
                     "Are you sure you want to proceed?", vbYesNo + vbExclamation, "Cleanup Sandbox")
    
    If response = vbYes Then
        On Error GoTo ErrorHandler
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Delete sandbox folder and all contents
        If fso.FolderExists("C:\BASHFlowSandbox") Then
            fso.DeleteFolder "C:\BASHFlowSandbox", True
            LogSandboxMessage "Sandbox cleanup completed successfully"
            MsgBox "BASH Flow Management sandbox has been completely removed.", vbInformation
        Else
            MsgBox "Sandbox folder not found. Nothing to clean up.", vbInformation
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    LogSandboxMessage "Error during sandbox cleanup: " & Err.Description
    MsgBox "Error during cleanup: " & Err.Description & vbCrLf & _
           "You may need to manually delete C:\BASHFlowSandbox", vbCritical
End Sub

' Backup sandbox data before cleanup
Sub BackupSandboxData()
    Dim fso As Object
    Dim backupPath As String
    Dim timestamp As String
    
    On Error GoTo ErrorHandler
    
    timestamp = Format(Now(), "yyyy-mm-dd_hh-nn-ss")
    backupPath = SANDBOX_BACKUP_FOLDER & "SandboxBackup_" & timestamp & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create backup folder
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
    End If
    
    ' Copy database
    If fso.FileExists("C:\BASHFlowSandbox\TestDatabase.accdb") Then
        fso.CopyFile "C:\BASHFlowSandbox\TestDatabase.accdb", backupPath & "TestDatabase.accdb"
    End If
    
    ' Copy log file
    If fso.FileExists("C:\BASHFlowSandbox\SandboxLog.txt") Then
        fso.CopyFile "C:\BASHFlowSandbox\SandboxLog.txt", backupPath & "SandboxLog.txt"
    End If
    
    LogSandboxMessage "Sandbox data backed up to: " & backupPath
    MsgBox "Sandbox data backed up successfully to:" & vbCrLf & backupPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogSandboxMessage "Error backing up sandbox data: " & Err.Description
    MsgBox "Error backing up sandbox data: " & Err.Description, vbCritical
End Sub

'===============================================================================
' Shared Functions (copied from main version)
'===============================================================================

' Validate Subject Line Against Preset Formats (same as main version)
Private Function ValidateSubjectLine(Subject As String) As validationResult
    Dim result As validationResult
    Dim regex As Object
    Dim pattern As String
    Dim matches As Object
    
    On Error GoTo ErrorHandler
    
    ' Initialize result as empty ValidationResult
    result.IsValid = False
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Check for CLLI format: CLLI-####-YYYY-MM-DD-Description
    pattern = "^(CLLI)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$"
    regex.pattern = pattern
    Set matches = regex.Execute(Subject)
    
    If matches.Count > 0 Then
        result.IsValid = True
        result.RecordType = matches(0).SubMatches(0)
        result.CLLINumber = matches(0).SubMatches(1)
        result.RecordDate = matches(0).SubMatches(2)
        result.Description = matches(0).SubMatches(3)
        result.Status = "Active"
        ValidateSubjectLine = result
        Exit Function
    End If
    
    ' Check for Milestone format: MS-####-YYYY-MM-DD-Description
    pattern = "^(MS)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$"
    regex.pattern = pattern
    Set matches = regex.Execute(Subject)
    
    If matches.Count > 0 Then
        result.IsValid = True
        result.RecordType = matches(0).SubMatches(0)
        result.MSNumber = matches(0).SubMatches(1)
        result.RecordDate = matches(0).SubMatches(2)
        result.Description = matches(0).SubMatches(3)
        result.Status = "Active"
        ValidateSubjectLine = result
        Exit Function
    End If
    
    ' Check for Event format: Event-####-YYYY-MM-DD-Description
    pattern = "^(Event)-(\d{4})-(\d{4}-\d{2}-\d{2})-(.+)$"
    regex.pattern = pattern
    Set matches = regex.Execute(Subject)
    
    If matches.Count > 0 Then
        result.IsValid = True
        result.RecordType = matches(0).SubMatches(0)
        result.EventNumber = matches(0).SubMatches(1)
        result.RecordDate = matches(0).SubMatches(2)
        result.Description = matches(0).SubMatches(3)
        result.Status = ExtractEventStatus(result.Description)
        ValidateSubjectLine = result
        Exit Function
    End If
    
    ' If no pattern matches
    result.IsValid = False
    result.ErrorMessage = "Subject line does not match any preset format"
    ValidateSubjectLine = result
    
    Exit Function
    
ErrorHandler:
    result.IsValid = False
    result.ErrorMessage = "Validation error: " & Err.Description
    ValidateSubjectLine = result
End Function

' Extract Event Status from Description (same as main version)
Private Function ExtractEventStatus(Description As String) As String
    Dim lowerDesc As String
    lowerDesc = LCase(Trim(Description))
    
    If InStr(lowerDesc, "completed") > 0 Then
        ExtractEventStatus = STATUS_COMPLETED
    ElseIf InStr(lowerDesc, "blocked") > 0 Then
        ExtractEventStatus = STATUS_BLOCKED
    ElseIf InStr(lowerDesc, "delayed") > 0 Then
        ExtractEventStatus = STATUS_DELAYED
    ElseIf InStr(lowerDesc, "follow-up") > 0 Or InStr(lowerDesc, "followup") > 0 Then
        ExtractEventStatus = STATUS_FOLLOWUP
    ElseIf InStr(lowerDesc, "urgency 1") > 0 Or InStr(lowerDesc, "urgent 1") > 0 Then
        ExtractEventStatus = STATUS_URGENCY1
    ElseIf InStr(lowerDesc, "urgency 2") > 0 Or InStr(lowerDesc, "urgent 2") > 0 Then
        ExtractEventStatus = STATUS_URGENCY2
    ElseIf InStr(lowerDesc, "urgency 3") > 0 Or InStr(lowerDesc, "urgent 3") > 0 Then
        ExtractEventStatus = STATUS_URGENCY3
    Else
        ExtractEventStatus = "New Event"
    End If
End Function

' Create Database Record from Email and Validation Results (same as main version)
Private Function CreateDatabaseRecord(item As Object, validation As validationResult) As DatabaseRecord
    Dim record As DatabaseRecord
    
    record.State = "TX" ' Default state - could be extracted from subject line or set as needed
    record.RecordType = validation.RecordType
    record.CLLINumber = validation.CLLINumber
    record.MSNumber = validation.MSNumber
    record.EventNumber = validation.EventNumber
    record.RecordDate = validation.RecordDate
    record.Description = validation.Description
    record.Status = validation.Status
    record.CreatedDate = Now()
    
    CreateDatabaseRecord = record
End Function

' Escape SQL strings to prevent injection (same as main version)
Private Function EscapeString(inputStr As String) As String
    EscapeString = Replace(inputStr, "'", "''")
End Function

'===============================================================================
' End of BASH Flow Management Sandbox
'===============================================================================
