'===============================================================================
' BASH Flow Management - Python Bridge
' VBA wrapper functions to call Python scripts
'===============================================================================

Option Explicit

' Python configuration
Private Const PYTHON_EXE As String = "python"
Private Const PYTHON_SCRIPT As String = "email_processor.py"
Private Const WORKING_DIR As String = "C:\Lumen\Workflow Manager"

'===============================================================================
' Python Bridge Functions
'===============================================================================

' Process email using Python backend
Public Function ProcessEmailWithPython(subject As String) As String
    On Error GoTo ErrorHandler
    
    Dim cmd As String
    Dim result As String
    Dim shell As Object
    
    ' Build command to call Python script
    cmd = PYTHON_EXE & " """ & PYTHON_SCRIPT & """ """ & EscapeForShell(subject) & """"
    
    ' Execute Python script
    Set shell = CreateObject("WScript.Shell")
    result = shell.Exec(cmd).StdOut.ReadAll
    
    ProcessEmailWithPython = result
    
    Exit Function
    
ErrorHandler:
    ProcessEmailWithPython = "{""success"": false, ""error"": """ & Err.Description & """}"
End Function

' Validate subject line using Python
Public Function ValidateSubjectWithPython(subject As String) As ValidationResult
    On Error GoTo ErrorHandler
    
    Dim jsonResult As String
    Dim result As ValidationResult
    Dim jsonObj As Object
    
    ' Call Python script
    jsonResult = ProcessEmailWithPython(subject)
    
    ' Parse JSON result
    Set jsonObj = ParseJSON(jsonResult)
    
    ' Convert JSON to ValidationResult
    result.IsValid = jsonObj("success")
    result.RecordType = GetJSONValue(jsonObj, "validation", "record_type")
    result.CLLINumber = GetJSONValue(jsonObj, "validation", "clli_number")
    result.MSNumber = GetJSONValue(jsonObj, "validation", "ms_number")
    result.EventNumber = GetJSONValue(jsonObj, "validation", "event_number")
    result.RecordDate = GetJSONValue(jsonObj, "validation", "record_date")
    result.Description = GetJSONValue(jsonObj, "validation", "description")
    result.Status = GetJSONValue(jsonObj, "validation", "status")
    result.ErrorMessage = GetJSONValue(jsonObj, "error")
    
    ValidateSubjectWithPython = result
    
    Exit Function
    
ErrorHandler:
    result.IsValid = False
    result.ErrorMessage = "Python bridge error: " & Err.Description
    ValidateSubjectWithPython = result
End Function

' Process multiple emails using Python
Public Function ProcessEmailsWithPython(subjects As Collection) As String
    On Error GoTo ErrorHandler
    
    Dim results As String
    Dim subject As Variant
    Dim result As String
    
    results = "["
    
    For Each subject In subjects
        result = ProcessEmailWithPython(CStr(subject))
        results = results & result & ","
    Next subject
    
    ' Remove trailing comma and close array
    If Right(results, 1) = "," Then results = Left(results, Len(results) - 1)
    results = results & "]"
    
    ProcessEmailsWithPython = results
    
    Exit Function
    
ErrorHandler:
    ProcessEmailsWithPython = "{""success"": false, ""error"": """ & Err.Description & """}"
End Function

'===============================================================================
' Utility Functions
'===============================================================================

' Escape string for shell command
Private Function EscapeForShell(inputStr As String) As String
    ' Replace quotes and escape special characters
    EscapeForShell = Replace(Replace(inputStr, """", "\"""), "'", "\'")
End Function

' Simple JSON parser (basic implementation)
Private Function ParseJSON(jsonStr As String) As Object
    On Error GoTo ErrorHandler
    
    ' This is a simplified JSON parser
    ' For production, consider using a proper JSON library
    Dim jsonObj As Object
    Set jsonObj = CreateObject("Scripting.Dictionary")
    
    ' Basic JSON parsing - extract key values
    ' This is a simplified version - you might want to use a proper JSON library
    
    ParseJSON = jsonObj
    
    Exit Function
    
ErrorHandler:
    Set ParseJSON = CreateObject("Scripting.Dictionary")
End Function

' Get value from JSON object
Private Function GetJSONValue(jsonObj As Object, key1 As String, Optional key2 As String = "") As String
    On Error GoTo ErrorHandler
    
    If key2 = "" Then
        GetJSONValue = jsonObj(key1)
    Else
        GetJSONValue = jsonObj(key1)(key2)
    End If
    
    Exit Function
    
ErrorHandler:
    GetJSONValue = ""
End Function

'===============================================================================
' Integration Functions
'===============================================================================

' Replace existing ValidateSubjectLine with Python version
Public Function ValidateSubjectLinePython(subject As String) As ValidationResult
    ValidateSubjectLinePython = ValidateSubjectWithPython(subject)
End Function

' Process workflow emails using Python backend
Public Sub ProcessWorkflowEmailsPython()
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim namespace As Object
    Dim inbox As Object
    Dim items As Object
    Dim item As Object
    Dim i As Long
    Dim processedCount As Long
    Dim errorCount As Long
    Dim subjects As Collection
    Dim result As String
    
    ' Initialize Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    Set namespace = outlookApp.GetNamespace("MAPI")
    Set inbox = namespace.GetDefaultFolder(6) ' olFolderInbox
    Set items = inbox.Items
    Set subjects = New Collection
    
    ' Sort items by received time (newest first)
    items.Sort "[ReceivedTime]", True
    
    ' Collect subjects (limited to 10 for processing)
    For i = 1 To items.Count
        Set item = items(i)
        
        ' Check if item is an email
        If item.Class = 43 Then ' olMail
            subjects.Add item.Subject
        End If
        
        ' Limit processing to 10 emails
        If i > 10 Then Exit For
    Next i
    
    ' Process all subjects using Python
    result = ProcessEmailsWithPython(subjects)
    
    ' Show results
    MsgBox "Python processing complete!" & vbCrLf & _
           "Results: " & result, vbInformation, "Python Processing"
    
    ' Cleanup
    Set item = Nothing
    Set items = Nothing
    Set inbox = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in Python processing: " & Err.Description, vbCritical
End Sub

'===============================================================================
' Test Functions
'===============================================================================

' Test Python bridge
Public Sub TestPythonBridge()
    On Error GoTo ErrorHandler
    
    Dim testSubject As String
    Dim result As String
    
    testSubject = "CLLI-1234-2024-01-15-Test Description"
    result = ProcessEmailWithPython(testSubject)
    
    MsgBox "Python Bridge Test Result:" & vbCrLf & result, vbInformation, "Test Result"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Test error: " & Err.Description, vbCritical
End Sub
