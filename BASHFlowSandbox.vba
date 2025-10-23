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
    
    ' Insert sample records
    sql = "INSERT INTO " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_STATE & ", " & COLUMN_CLLI_NUMBER & ", " & COLUMN_DATE & ", " & COLUMN_TYPE & ", " & _
          COLUMN_DESCRIPTION & ", " & COLUMN_STATUS & ") VALUES " & _
          "('TX', '1001', '2024-01-15', 'CLLI', 'Network Infrastructure Project (SANDBOX)', 'Active')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    sql = "INSERT INTO " & SANDBOX_TABLE_NAME & " (" & _
          COLUMN_STATE & ", " & COLUMN_MS_NUMBER & ", " & COLUMN_DATE & ", " & COLUMN_TYPE & ", " & _
          COLUMN_DESCRIPTION & ", " & COLUMN_STATUS & ") VALUES " & _
          "('TX', '2001', '2024-01-16', 'MS', 'Phase 1 Network Design Complete (SANDBOX)', 'Completed')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    MsgBox "✅ Database table fixed successfully!" & vbCrLf & _
           "Table: " & SANDBOX_TABLE_NAME & vbCrLf & _
           "Sample records inserted." & vbCrLf & vbCrLf & _
           "The table structure has been corrected.", vbInformation, "Database Fixed"
    
    LogSandboxMessage "Database table fixed successfully with sample data"
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    MsgBox "❌ Error fixing database: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Database Error"
    LogSandboxMessage "Error fixing database: " & Err.Description
End Sub

'===============================================================================
' SANDBOX Logging Functions
'===============================================================================

' Log message to sandbox log file
Private Sub LogSandboxMessage(message As String)
    Dim fso As Object
    Dim logFile As Object
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(SANDBOX_LOG_FILE_PATH) Then
        Set logFile = fso.OpenTextFile(SANDBOX_LOG_FILE_PATH, 8, True) ' Append mode
    Else
        Set logFile = fso.CreateTextFile(SANDBOX_LOG_FILE_PATH, True) ' Create new
    End If
    
    logFile.WriteLine "[SANDBOX] " & Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & message
    logFile.Close
    
    On Error GoTo 0
End Sub

'===============================================================================
' SANDBOX Utility Functions
'===============================================================================

' Check if sandbox is initialized
Private Function IsSandboxInitialized() As Boolean
    Dim fso As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if sandbox folder exists
    If Not fso.FolderExists("C:\BASHFlowSandbox") Then
        IsSandboxInitialized = False
        Exit Function
    End If
    
    ' Check if database exists
    If Not fso.FileExists("C:\BASHFlowSandbox\TestDatabase.accdb") Then
        IsSandboxInitialized = False
        Exit Function
    End If
    
    IsSandboxInitialized = True
    
    Exit Function
    
ErrorHandler:
    IsSandboxInitialized = False
End Function

'===============================================================================
' SANDBOX Test Functions
'===============================================================================

' Test sandbox functionality
Sub TestSandboxFunctionality()
    On Error GoTo ErrorHandler
    
    If IsSandboxInitialized() Then
        MsgBox "✅ Sandbox is properly initialized and ready for testing!", vbInformation, "Sandbox Test"
        LogSandboxMessage "Sandbox functionality test passed"
    Else
        MsgBox "❌ Sandbox is not properly initialized. Please run InitializeSandbox() first.", vbExclamation, "Sandbox Test"
        LogSandboxMessage "Sandbox functionality test failed - not initialized"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Error testing sandbox: " & Err.Description, vbCritical, "Sandbox Test Error"
    LogSandboxMessage "Error testing sandbox: " & Err.Description
End Sub

'===============================================================================
' SANDBOX Ribbon Integration
'===============================================================================

' Initialize sandbox ribbon
Public Sub InitializeSandboxRibbon(ribbon As IRibbonUI)
    Set sandboxRibbonUI = ribbon
End Sub

' Sandbox ribbon callback
Public Sub ShowSandboxRibbon(control As IRibbonControl)
    TestSandboxFunctionality
End Sub
