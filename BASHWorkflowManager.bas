Attribute VB_Name = "BASHWorkflowManager"
'===============================================================================
' Workflow Manager VBA Script
' Automates workflow tasks based on email subject lines and dates
' Populates validated data into database for CLLI project tracking
'===============================================================================

Option Explicit

' Configuration Constants
Private Const DB_CONNECTION_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\Your\Database.accdb;"
Private Const LOG_FILE_PATH As String = "C:\Path\To\Your\WorkflowLog.txt"

' Subject Line Format Patterns
Private Const FORMAT_CLLI As String = "CLLI-####-YYYY-MM-DD-Description"           ' CLLI (Project) format
Private Const FORMAT_MS As String = "MS-####-YYYY-MM-DD-Description"              ' Milestone format
Private Const FORMAT_EVENT As String = "Event-####-YYYY-MM-DD-Description"        ' Event format

' Event Status Types
Private Const STATUS_COMPLETED As String = "Completed"
Private Const STATUS_BLOCKED As String = "Blocked"
Private Const STATUS_DELAYED As String = "Delayed"
Private Const STATUS_FOLLOWUP As String = "Follow-up required"
Private Const STATUS_URGENCY1 As String = "Urgency 1"
Private Const STATUS_URGENCY2 As String = "Urgency 2"
Private Const STATUS_URGENCY3 As String = "Urgency 3"

' Database Table Structure
Private Const TABLE_NAME As String = "CLLIWorkflow"
Private Const COLUMN_ID As String = "ID"
Private Const COLUMN_SUBJECT As String = "Subject"
Private Const COLUMN_TYPE As String = "Type"
Private Const COLUMN_CLLI_NUMBER As String = "CLLINumber"
Private Const COLUMN_MS_NUMBER As String = "MSNumber"
Private Const COLUMN_EVENT_NUMBER As String = "EventNumber"
Private Const COLUMN_DATE As String = "Date"
Private Const COLUMN_DESCRIPTION As String = "Description"
Private Const COLUMN_STATUS As String = "Status"
Private Const COLUMN_CREATED As String = "CreatedDate"

'===============================================================================
' Main Subroutine - Entry Point
'===============================================================================
Sub ProcessWorkflowEmails()
    Dim outlookApp As Object
    Dim namespace As Object
    Dim inbox As Object
    Dim Items As Object
    Dim item As Object
    Dim i As Long
    Dim processedCount As Long
    Dim errorCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Initialize Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    Set namespace = outlookApp.GetNamespace("MAPI")
    Set inbox = namespace.GetDefaultFolder(6) ' olFolderInbox
    Set Items = inbox.Items
    
    ' Sort items by received time (newest first)
    Items.Sort "[ReceivedTime]", True
    
    LogMessage "Starting CLLI workflow email processing..."
    
    ' Process each email
    For i = 1 To Items.Count
        Set item = Items(i)
        
        ' Check if item is an email
        If item.Class = 43 Then ' olMail
            If ProcessEmailItem(item) Then
                processedCount = processedCount + 1
            Else
                errorCount = errorCount + 1
            End If
        End If
        
        ' Prevent processing too many emails at once
        If i > 100 Then Exit For
    Next i
    
    LogMessage "Processing complete. Processed: " & processedCount & ", Errors: " & errorCount
    
    ' Cleanup
    Set item = Nothing
    Set Items = Nothing
    Set inbox = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
    
    MsgBox "CLLI Workflow processing complete!" & vbCrLf & _
           "Processed: " & processedCount & vbCrLf & _
           "Errors: " & errorCount, vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogMessage "Error in ProcessWorkflowEmails: " & Err.Description
    MsgBox "Error occurred: " & Err.Description, vbCritical
End Sub

'===============================================================================
' Process Individual Email Item
'===============================================================================
Private Function ProcessEmailItem(item As Object) As Boolean
    Dim subjectLine As String
    Dim validationResult As validationResult
    Dim dbRecord As DatabaseRecord
    
    On Error GoTo ErrorHandler
    
    subjectLine = Trim(item.Subject)
    
    ' Skip if no subject
    If Len(subjectLine) = 0 Then
        LogMessage "Skipping email with no subject line"
        ProcessEmailItem = True
        Exit Function
    End If
    
    ' Validate subject line format
    validationResult = ValidateSubjectLine(subjectLine)
    
    If validationResult.IsValid Then
        ' Create database record
        dbRecord = CreateDatabaseRecord(item, validationResult)
        
        ' Insert into database
        If InsertIntoDatabase(dbRecord) Then
            LogMessage "Successfully processed: " & subjectLine
            ProcessEmailItem = True
        Else
            LogMessage "Failed to insert into database: " & subjectLine
            ProcessEmailItem = False
        End If
    Else
        LogMessage "Invalid subject line format: " & subjectLine & " - " & validationResult.ErrorMessage
        ProcessEmailItem = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Error processing email: " & Err.Description
    ProcessEmailItem = False
End Function

'===============================================================================
' Validate Subject Line Against Preset Formats
'===============================================================================
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
        result.Type = matches(0).SubMatches(0)
        result.CLLINumber = matches(0).SubMatches(1)
        result.Date = matches(0).SubMatches(2)
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
        result.Type = matches(0).SubMatches(0)
        result.MSNumber = matches(0).SubMatches(1)
        result.Date = matches(0).SubMatches(2)
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
        result.Type = matches(0).SubMatches(0)
        result.EventNumber = matches(0).SubMatches(1)
        result.Date = matches(0).SubMatches(2)
        result.Description = matches(0).SubMatches(3)
        ' For events, we need to check if the description contains a status
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

'===============================================================================
' Extract Event Status from Description
'===============================================================================
Private Function ExtractEventStatus(Description As String) As String
    Dim lowerDesc As String
    lowerDesc = LCase(Trim(Description))
    
    ' Check for status keywords in description
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

'===============================================================================
' Create Database Record from Email and Validation Results
'===============================================================================
Private Function CreateDatabaseRecord(item As Object, validation As validationResult) As DatabaseRecord
    Dim record As DatabaseRecord
    
    ' Initialize record as empty DatabaseRecord
    record.Subject = Trim(item.Subject)
    record.Type = validation.Type
    record.CLLINumber = validation.CLLINumber
    record.MSNumber = validation.MSNumber
    record.EventNumber = validation.EventNumber
    record.Date = validation.Date
    record.Description = validation.Description
    record.Status = validation.Status
    record.CreatedDate = Now()
    
    CreateDatabaseRecord = record
End Function

'===============================================================================
' Insert Record into Database
'===============================================================================
Private Function InsertIntoDatabase(record As DatabaseRecord) As Boolean
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Create connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    
    ' Create command
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Build SQL insert statement
    sql = "INSERT INTO " & TABLE_NAME & " (" & _
          COLUMN_SUBJECT & ", " & _
          COLUMN_TYPE & ", " & _
          COLUMN_CLLI_NUMBER & ", " & _
          COLUMN_MS_NUMBER & ", " & _
          COLUMN_EVENT_NUMBER & ", " & _
          COLUMN_DATE & ", " & _
          COLUMN_DESCRIPTION & ", " & _
          COLUMN_STATUS & ", " & _
          COLUMN_CREATED & ") VALUES (" & _
          "'" & EscapeString(record.Subject) & "', " & _
          "'" & EscapeString(record.Type) & "', " & _
          "'" & EscapeString(record.CLLINumber) & "', " & _
          "'" & EscapeString(record.MSNumber) & "', " & _
          "'" & EscapeString(record.EventNumber) & "', " & _
          "'" & EscapeString(record.Date) & "', " & _
          "'" & EscapeString(record.Description) & "', " & _
          "'" & EscapeString(record.Status) & "', " & _
          "'" & Format(record.CreatedDate, "yyyy-mm-dd hh:nn:ss") & "')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    InsertIntoDatabase = True
    
    Exit Function
    
ErrorHandler:
    LogMessage "Database error: " & Err.Description
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    InsertIntoDatabase = False
End Function

'===============================================================================
' Utility Functions
'===============================================================================

' Escape SQL strings to prevent injection
Private Function EscapeString(inputStr As String) As String
    EscapeString = Replace(inputStr, "'", "''")
End Function

' Log messages to file
Private Sub LogMessage(message As String)
    Dim fso As Object
    Dim logFile As Object
    Dim timestamp As String
    
    On Error Resume Next
    
    timestamp = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(LOG_FILE_PATH) Then
        Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 8, True) ' 8 = ForAppending
    Else
        Set logFile = fso.CreateTextFile(LOG_FILE_PATH, True)
    End If
    
    logFile.WriteLine timestamp & " - " & message
    logFile.Close
End Sub

'===============================================================================
' Data Structures
'===============================================================================

' Validation Result Class
Private Class ValidationResult
    Public IsValid As Boolean
    Public Type As String
    Public CLLINumber As String
    Public MSNumber As String
    Public EventNumber As String
    Public Date As String
    Public Description As String
    Public Status As String
    Public ErrorMessage As String
End Class

' Database Record Class
Private Class DatabaseRecord
    Public Subject As String
    Public Type As String
    Public CLLINumber As String
    Public MSNumber As String
    Public EventNumber As String
    Public Date As String
    Public Description As String
    Public Status As String
    Public CreatedDate As Date
End Class

'===============================================================================
' Ribbon Integration Functions
'===============================================================================

' Global ribbon object
Private ribbonUI As IRibbonUI

' Ribbon callback for onLoad event
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub

' Get status label for ribbon
Sub GetStatusLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "BASH Flow Management - Ready"
End Sub

' Ribbon button callbacks
Sub ProcessWorkflowEmailsRibbon(control As IRibbonControl)
    ProcessWorkflowEmails
    If Not ribbonUI Is Nothing Then
        ribbonUI.Invalidate
    End If
End Sub

Sub ShowFormatExamplesRibbon(control As IRibbonControl)
    ShowFormatExamples
End Sub

Sub TestDatabaseConnectionRibbon(control As IRibbonControl)
    TestDatabaseConnection
End Sub

Sub CreateDatabaseTableRibbon(control As IRibbonControl)
    CreateDatabaseTable
End Sub

Sub GenerateSampleDataRibbon(control As IRibbonControl)
    GenerateSampleData
End Sub

Sub ViewLogFile(control As IRibbonControl)
    Dim fso As Object
    Dim logFile As Object
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(LOG_FILE_PATH) Then
        Set logFile = fso.OpenTextFile(LOG_FILE_PATH, 1, False) ' 1 = ForReading
        MsgBox logFile.ReadAll, vbInformation, "Workflow Log File"
        logFile.Close
    Else
        MsgBox "Log file not found at: " & LOG_FILE_PATH, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening log file: " & Err.Description, vbCritical
End Sub

Sub ShowWorkflowStatus(control As IRibbonControl)
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
    conn.Open DB_CONNECTION_STRING
    
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Get total count
    sql = "SELECT COUNT(*) AS TotalCount FROM " & TABLE_NAME
    rs.Open sql, conn
    totalCount = rs("TotalCount")
    rs.Close
    
    ' Get CLLI count
    sql = "SELECT COUNT(*) AS CLLICount FROM " & TABLE_NAME & " WHERE Type = 'CLLI'"
    rs.Open sql, conn
    clliCount = rs("CLLICount")
    rs.Close
    
    ' Get MS count
    sql = "SELECT COUNT(*) AS MSCount FROM " & TABLE_NAME & " WHERE Type = 'MS'"
    rs.Open sql, conn
    msCount = rs("MSCount")
    rs.Close
    
    ' Get Event count
    sql = "SELECT COUNT(*) AS EventCount FROM " & TABLE_NAME & " WHERE Type = 'Event'"
    rs.Open sql, conn
    eventCount = rs("EventCount")
    rs.Close
    
    conn.Close
    
    statusMessage = "BASH Flow Management Status:" & vbCrLf & vbCrLf & _
                   "Total Records: " & totalCount & vbCrLf & _
                   "CLLI Projects: " & clliCount & vbCrLf & _
                   "Milestones: " & msCount & vbCrLf & _
                   "Events: " & eventCount & vbCrLf & vbCrLf & _
                   "Last Updated: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    MsgBox statusMessage, vbInformation, "Workflow Status"
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    MsgBox "Error retrieving status: " & Err.Description, vbCritical
End Sub

Sub ExportWorkflowData(control As IRibbonControl)
    Dim conn As Object
    Dim rs As Object
    Dim excelApp As Object
    Dim workbook As Object
    Dim worksheet As Object
    Dim sql As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    ' Create Excel application
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    Set workbook = excelApp.Workbooks.Add
    Set worksheet = workbook.Worksheets(1)
    
    ' Set up headers
    worksheet.Cells(1, 1).Value = "ID"
    worksheet.Cells(1, 2).Value = "Subject"
    worksheet.Cells(1, 3).Value = "Type"
    worksheet.Cells(1, 4).Value = "CLLI Number"
    worksheet.Cells(1, 5).Value = "MS Number"
    worksheet.Cells(1, 6).Value = "Event Number"
    worksheet.Cells(1, 7).Value = "Date"
    worksheet.Cells(1, 8).Value = "Description"
    worksheet.Cells(1, 9).Value = "Status"
    worksheet.Cells(1, 10).Value = "Created Date"
    
    ' Format headers
    With worksheet.Range("A1:J1")
        .Font.bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Connect to database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    
    Set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM " & TABLE_NAME & " ORDER BY CreatedDate DESC"
    rs.Open sql, conn
    
    ' Export data
    i = 2
    Do While Not rs.EOF
        worksheet.Cells(i, 1).Value = rs(COLUMN_ID)
        worksheet.Cells(i, 2).Value = rs(COLUMN_SUBJECT)
        worksheet.Cells(i, 3).Value = rs(COLUMN_TYPE)
        worksheet.Cells(i, 4).Value = rs(COLUMN_CLLI_NUMBER)
        worksheet.Cells(i, 5).Value = rs(COLUMN_MS_NUMBER)
        worksheet.Cells(i, 6).Value = rs(COLUMN_EVENT_NUMBER)
        worksheet.Cells(i, 7).Value = rs(COLUMN_DATE)
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
    
    MsgBox "Workflow data exported successfully to Excel!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    MsgBox "Error exporting data: " & Err.Description, vbCritical
End Sub

'===============================================================================
' Configuration and Setup Functions
'===============================================================================

' Test database connection
Sub TestDatabaseConnection()
    Dim conn As Object
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    
    MsgBox "Database connection successful!", vbInformation
    conn.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Database connection failed: " & Err.Description, vbCritical
End Sub

' Show subject line format examples
Sub ShowFormatExamples()
    Dim message As String
    
    message = "Valid Subject Line Formats for CLLI Workflow:" & vbCrLf & vbCrLf & _
              "CLLI (Project): CLLI-1234-2024-01-15-Network Infrastructure" & vbCrLf & _
              "Milestone: MS-5678-2024-01-15-Phase 1 Completion" & vbCrLf & _
              "Event (Completed): Event-9012-2024-01-15-Task completed successfully" & vbCrLf & _
              "Event (Blocked): Event-3456-2024-01-15-Implementation blocked by dependency" & vbCrLf & _
              "Event (Delayed): Event-7890-2024-01-15-Timeline delayed due to resource constraints" & vbCrLf & _
              "Event (Follow-up): Event-2468-2024-01-15-Follow-up required with client" & vbCrLf & _
              "Event (Urgency): Event-1357-2024-01-15-Urgency 1 critical issue identified" & vbCrLf & vbCrLf & _
              "Format: TYPE-NUMBER-YYYY-MM-DD-Description" & vbCrLf & _
              "• TYPE: CLLI, MS, or Event" & vbCrLf & _
              "• NUMBER: 4-digit number" & vbCrLf & _
              "• DATE: YYYY-MM-DD format" & vbCrLf & _
              "• DESCRIPTION: Any text (for Events, include status keywords)" & vbCrLf & vbCrLf & _
              "Event Status Keywords:" & vbCrLf & _
              "• Completed, Blocked, Delayed" & vbCrLf & _
              "• Follow-up, Followup" & vbCrLf & _
              "• Urgency 1/2/3, Urgent 1/2/3"
    
    MsgBox message, vbInformation, "CLLI Workflow Format Guide"
End Sub

' Create database table (run once for setup)
Sub CreateDatabaseTable()
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    ' Drop table if exists (be careful in production!)
    sql = "DROP TABLE IF EXISTS " & TABLE_NAME
    cmd.CommandText = sql
    On Error Resume Next
    cmd.Execute
    On Error GoTo ErrorHandler
    
    ' Create table
    sql = "CREATE TABLE " & TABLE_NAME & " (" & _
          COLUMN_ID & " AUTOINCREMENT PRIMARY KEY, " & _
          COLUMN_SUBJECT & " TEXT(255) NOT NULL, " & _
          COLUMN_TYPE & " TEXT(10) NOT NULL, " & _
          COLUMN_CLLI_NUMBER & " TEXT(10), " & _
          COLUMN_MS_NUMBER & " TEXT(10), " & _
          COLUMN_EVENT_NUMBER & " TEXT(10), " & _
          COLUMN_DATE & " TEXT(10) NOT NULL, " & _
          COLUMN_DESCRIPTION & " TEXT(255), " & _
          COLUMN_STATUS & " TEXT(20) DEFAULT 'Active', " & _
          COLUMN_CREATED & " DATETIME DEFAULT NOW())"
    
    cmd.CommandText = sql
    cmd.Execute
    
    MsgBox "CLLI Workflow database table created successfully!", vbInformation
    conn.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating table: " & Err.Description, vbCritical
End Sub

' Generate sample data for testing
Sub GenerateSampleData()
    Dim sampleSubjects As Variant
    Dim i As Integer
    
    sampleSubjects = Array( _
        "CLLI-1001-2024-01-15-Network Infrastructure Project", _
        "MS-2001-2024-01-16-Phase 1 Network Design Complete", _
        "Event-3001-2024-01-17-Task completed successfully", _
        "Event-3002-2024-01-18-Implementation blocked by vendor delay", _
        "Event-3003-2024-01-19-Timeline delayed due to resource constraints", _
        "Event-3004-2024-01-20-Follow-up required with client approval", _
        "Event-3005-2024-01-21-Urgency 1 critical security issue identified", _
        "MS-2002-2024-01-22-Phase 2 Implementation Milestone", _
        "Event-3006-2024-01-23-Urgency 2 performance issue detected", _
        "CLLI-1002-2024-01-24-Customer Portal Development" _
    )
    
    For i = LBound(sampleSubjects) To UBound(sampleSubjects)
        Debug.Print sampleSubjects(i)
    Next i
    
    MsgBox "Sample data generated. Check Immediate window for examples.", vbInformation
End Sub


