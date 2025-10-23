Attribute VB_Name = "CheckDatabase"
'===============================================================================
' Database Structure Checker
' Run this to see what's actually in the database
'===============================================================================

Sub CheckDatabaseStructure()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim i As Integer
    Dim message As String
    
    On Error GoTo ErrorHandler
    
    ' Connect to database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    
    ' Check if table exists and get its structure
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Get table schema
    sql = "SELECT * FROM BASHFlowSandbox WHERE 1=0" ' Returns schema only
    rs.Open sql, conn
    
    message = "Database Table Structure:" & vbCrLf & vbCrLf
    
    ' Show field information
    For i = 0 To rs.Fields.Count - 1
        message = message & "Field " & (i + 1) & ": " & rs.Fields(i).Name & vbCrLf & _
                  "  Type: " & rs.Fields(i).Type & vbCrLf & _
                  "  Size: " & rs.Fields(i).DefinedSize & vbCrLf & _
                  "  Nullable: " & rs.Fields(i).Attributes & vbCrLf & vbCrLf
    Next i
    
    rs.Close
    
    ' Try to get record count
    sql = "SELECT COUNT(*) AS RecordCount FROM BASHFlowSandbox"
    rs.Open sql, conn
    message = message & "Total Records: " & rs("RecordCount") & vbCrLf & vbCrLf
    rs.Close
    
    ' Try to get a sample record
    sql = "SELECT TOP 1 * FROM BASHFlowSandbox"
    rs.Open sql, conn
    
    If Not rs.EOF Then
        message = message & "Sample Record:" & vbCrLf
        For i = 0 To rs.Fields.Count - 1
            message = message & rs.Fields(i).Name & ": " & rs.Fields(i).Value & vbCrLf
        Next i
    Else
        message = message & "No records found in table." & vbCrLf
    End If
    
    rs.Close
    conn.Close
    
    MsgBox message, vbInformation, "Database Structure Check"
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    
    MsgBox "Error checking database: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Database Check Error"
End Sub

' Alternative: Check what tables exist in the database
Sub ListDatabaseTables()
    Dim conn As Object
    Dim rs As Object
    Dim message As String
    
    On Error GoTo ErrorHandler
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    
    Set rs = conn.OpenSchema(20) ' adSchemaTables
    
    message = "Tables in Database:" & vbCrLf & vbCrLf
    
    Do While Not rs.EOF
        If rs("TABLE_TYPE") = "TABLE" Then
            message = message & "Table: " & rs("TABLE_NAME") & vbCrLf & _
                      "Type: " & rs("TABLE_TYPE") & vbCrLf & vbCrLf
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    
    MsgBox message, vbInformation, "Database Tables"
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    
    MsgBox "Error listing tables: " & Err.Description, vbCritical
End Sub
