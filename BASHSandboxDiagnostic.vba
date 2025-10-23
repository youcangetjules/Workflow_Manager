'===============================================================================
' BASH Flow Sandbox Diagnostic Tool
' Comprehensive logging to help diagnose sandbox initialization issues
'===============================================================================

Option Explicit

'===============================================================================
' Diagnostic Functions
'===============================================================================

' Main diagnostic function - run this first
Sub RunSandboxDiagnostic()
    Dim logPath As String
    Dim fso As Object
    Dim logFile As Object
    
    On Error GoTo ErrorHandler
    
    ' Set up logging
    logPath = "C:\Lumen\Workflow Manager\SandboxDiagnostic.log"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create or open log file
    If fso.FileExists(logPath) Then
        Set logFile = fso.OpenTextFile(logPath, 8, True) ' Append mode
    Else
        Set logFile = fso.CreateTextFile(logPath, True) ' Create new
    End If
    
    ' Write header
    logFile.WriteLine "==============================================================================="
    logFile.WriteLine "BASH Flow Sandbox Diagnostic Report"
    logFile.WriteLine "Generated: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    logFile.WriteLine "==============================================================================="
    
    ' Run all diagnostic checks
    Call CheckSystemRequirements(logFile)
    Call CheckOfficeInstallation(logFile)
    Call CheckVBAEnvironment(logFile)
    Call TestDatabaseEngine(logFile)
    Call TestFileSystemAccess(logFile)
    Call TestSandboxInitialization(logFile)
    
    ' Close log file
    logFile.Close
    
    ' Show completion message
    MsgBox "Diagnostic complete! Check the log file at:" & vbCrLf & logPath & vbCrLf & vbCrLf & _
           "This log contains detailed information about what's preventing the sandbox from working.", _
           vbInformation, "Diagnostic Complete"
    
    ' Open log file for review
    Shell "notepad.exe """ & logPath & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    If Not logFile Is Nothing Then
        logFile.WriteLine "CRITICAL ERROR in diagnostic: " & Err.Description
        logFile.Close
    End If
    MsgBox "Error running diagnostic: " & Err.Description, vbCritical
End Sub

' Check system requirements
Private Sub CheckSystemRequirements(logFile As Object)
    Dim fso As Object
    Dim wsh As Object
    
    logFile.WriteLine vbCrLf & "=== SYSTEM REQUIREMENTS CHECK ==="
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    ' Check Windows version
    logFile.WriteLine "Windows Version: " & wsh.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
    logFile.WriteLine "Windows Build: " & wsh.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild")
    
    ' Check available disk space
    Dim drive As Object
    Set drive = fso.GetDrive("C:")
    logFile.WriteLine "C: Drive Free Space: " & Format(drive.FreeSpace / (1024 * 1024 * 1024), "0.00") & " GB"
    
    ' Check if we can create folders in C:\
    On Error Resume Next
    Dim testFolder As Object
    Set testFolder = fso.CreateFolder("C:\SandboxTest_" & Format(Now(), "yyyymmdd_hhnnss"))
    If Err.Number = 0 Then
        logFile.WriteLine "C:\ folder access: SUCCESS - Can create folders"
        testFolder.Delete
    Else
        logFile.WriteLine "C:\ folder access: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    logFile.WriteLine "System Requirements Check: COMPLETE"
End Sub

' Check Office installation
Private Sub CheckOfficeInstallation(logFile As Object)
    Dim wsh As Object
    Dim outlookPath As String
    
    logFile.WriteLine vbCrLf & "=== OFFICE INSTALLATION CHECK ==="
    
    Set wsh = CreateObject("WScript.Shell")
    
    ' Check Outlook installation
    On Error Resume Next
    outlookPath = wsh.RegRead("HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\InstallPath")
    If Err.Number = 0 Then
        logFile.WriteLine "Office Install Path: " & outlookPath
        logFile.WriteLine "Office Platform: " & wsh.RegRead("HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\Platform")
    Else
        logFile.WriteLine "Office Install Path: Not found in ClickToRun registry"
    End If
    On Error GoTo 0
    
    ' Check if Outlook is running
    Dim outlookApp As Object
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If Err.Number = 0 Then
        logFile.WriteLine "Outlook Status: RUNNING - Version " & outlookApp.Version
    Else
        logFile.WriteLine "Outlook Status: NOT RUNNING or not accessible"
    End If
    On Error GoTo 0
    
    ' Check VBA project access
    On Error Resume Next
    Dim vbaProject As Object
    Set vbaProject = Application.VBE.VBProjects
    If Err.Number = 0 Then
        logFile.WriteLine "VBA Project Access: SUCCESS - Can access VBA environment"
    Else
        logFile.WriteLine "VBA Project Access: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    logFile.WriteLine "Office Installation Check: COMPLETE"
End Sub

' Check VBA environment
Private Sub CheckVBAEnvironment(logFile As Object)
    logFile.WriteLine vbCrLf & "=== VBA ENVIRONMENT CHECK ==="
    
    ' Check if we can create objects
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        logFile.WriteLine "FileSystemObject: SUCCESS - Can create FSO object"
    Else
        logFile.WriteLine "FileSystemObject: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    ' Check ADODB connection capability
    Dim conn As Object
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    If Err.Number = 0 Then
        logFile.WriteLine "ADODB Connection: SUCCESS - Can create connection object"
        conn = Nothing
    Else
        logFile.WriteLine "ADODB Connection: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    ' Check ADOX Catalog capability
    Dim catalog As Object
    On Error Resume Next
    Set catalog = CreateObject("ADOX.Catalog")
    If Err.Number = 0 Then
        logFile.WriteLine "ADOX Catalog: SUCCESS - Can create catalog object"
        catalog = Nothing
    Else
        logFile.WriteLine "ADOX Catalog: FAILED - " & Err.Description
        logFile.WriteLine "  This is likely why sandbox initialization fails!"
    End If
    On Error GoTo 0
    
    logFile.WriteLine "VBA Environment Check: COMPLETE"
End Sub

' Test database engine
Private Sub TestDatabaseEngine(logFile As Object)
    logFile.WriteLine vbCrLf & "=== DATABASE ENGINE TEST ==="
    
    ' Test connection strings
    Dim connStrings As Variant
    connStrings = Array( _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\TestDB.accdb;", _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\TestDB.mdb;", _
        "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\TestDB.accdb;" _
    )
    
    Dim i As Integer
    For i = LBound(connStrings) To UBound(connStrings)
        Dim conn As Object
        On Error Resume Next
        Set conn = CreateObject("ADODB.Connection")
        conn.Open connStrings(i)
        If Err.Number = 0 Then
            logFile.WriteLine "Connection String " & (i + 1) & ": SUCCESS - " & connStrings(i)
            conn.Close
        Else
            logFile.WriteLine "Connection String " & (i + 1) & ": FAILED - " & Err.Description
        End If
        On Error GoTo 0
        conn = Nothing
    Next i
    
    logFile.WriteLine "Database Engine Test: COMPLETE"
End Sub

' Test file system access
Private Sub TestFileSystemAccess(logFile As Object)
    logFile.WriteLine vbCrLf & "=== FILE SYSTEM ACCESS TEST ==="
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Test creating folders in different locations
    Dim testPaths As Variant
    testPaths = Array( _
        "C:\BASHFlowSandbox", _
        "C:\Lumen\Workflow Manager\TestFolder", _
        Environ("TEMP") & "\BASHFlowTest" _
    )
    
    Dim i As Integer
    For i = LBound(testPaths) To UBound(testPaths)
        On Error Resume Next
        Dim testFolder As Object
        Set testFolder = fso.CreateFolder(testPaths(i))
        If Err.Number = 0 Then
            logFile.WriteLine "Create Folder " & (i + 1) & ": SUCCESS - " & testPaths(i)
            testFolder.Delete
        Else
            logFile.WriteLine "Create Folder " & (i + 1) & ": FAILED - " & Err.Description
        End If
        On Error GoTo 0
    Next i
    
    ' Test creating files
    On Error Resume Next
    Dim testFile As Object
    Set testFile = fso.CreateTextFile("C:\Lumen\Workflow Manager\TestLog.txt", True)
    If Err.Number = 0 Then
        logFile.WriteLine "Create File: SUCCESS - Can create files in project directory"
        testFile.Close
        fso.DeleteFile "C:\Lumen\Workflow Manager\TestLog.txt"
    Else
        logFile.WriteLine "Create File: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    logFile.WriteLine "File System Access Test: COMPLETE"
End Sub

' Test sandbox initialization step by step
Private Sub TestSandboxInitialization(logFile As Object)
    logFile.WriteLine vbCrLf & "=== SANDBOX INITIALIZATION TEST ==="
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Step 1: Create main folder
    logFile.WriteLine "Step 1: Creating C:\BASHFlowSandbox folder..."
    On Error Resume Next
    Dim mainFolder As Object
    Set mainFolder = fso.CreateFolder("C:\BASHFlowSandbox")
    If Err.Number = 0 Then
        logFile.WriteLine "  Result: SUCCESS"
    Else
        logFile.WriteLine "  Result: FAILED - " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Step 2: Create backups folder
    logFile.WriteLine "Step 2: Creating backups folder..."
    On Error Resume Next
    Dim backupFolder As Object
    Set backupFolder = fso.CreateFolder("C:\BASHFlowSandbox\Backups")
    If Err.Number = 0 Then
        logFile.WriteLine "  Result: SUCCESS"
    Else
        logFile.WriteLine "  Result: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    ' Step 3: Test database creation
    logFile.WriteLine "Step 3: Testing database creation..."
    On Error Resume Next
    Dim catalog As Object
    Set catalog = CreateObject("ADOX.Catalog")
    catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    If Err.Number = 0 Then
        logFile.WriteLine "  Result: SUCCESS - Database created"
        catalog = Nothing
    Else
        logFile.WriteLine "  Result: FAILED - " & Err.Description
        logFile.WriteLine "  This is the main issue! Database engine not available."
    End If
    On Error GoTo 0
    
    ' Step 4: Test log file creation
    logFile.WriteLine "Step 4: Testing log file creation..."
    On Error Resume Next
    Dim logFile As Object
    Set logFile = fso.CreateTextFile("C:\BASHFlowSandbox\SandboxLog.txt", True)
    If Err.Number = 0 Then
        logFile.WriteLine "  Result: SUCCESS - Log file created"
        logFile.WriteLine "[SANDBOX] " & Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - Test log entry"
        logFile.Close
    Else
        logFile.WriteLine "  Result: FAILED - " & Err.Description
    End If
    On Error GoTo 0
    
    logFile.WriteLine "Sandbox Initialization Test: COMPLETE"
End Sub

'===============================================================================
' Quick Test Functions
'===============================================================================

' Quick test - run this for immediate feedback
Sub QuickSandboxTest()
    Dim result As String
    result = ""
    
    ' Test 1: File System Access
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder "C:\BASHFlowSandbox"
    If Err.Number = 0 Then
        result = result & "✓ File System Access: OK" & vbCrLf
        fso.DeleteFolder "C:\BASHFlowSandbox"
    Else
        result = result & "✗ File System Access: FAILED - " & Err.Description & vbCrLf
    End If
    On Error GoTo 0
    
    ' Test 2: ADOX Catalog (Database Creation)
    On Error Resume Next
    Dim catalog As Object
    Set catalog = CreateObject("ADOX.Catalog")
    If Err.Number = 0 Then
        result = result & "✓ ADOX Catalog: OK" & vbCrLf
    Else
        result = result & "✗ ADOX Catalog: FAILED - " & Err.Description & vbCrLf
    End If
    On Error GoTo 0
    
    ' Test 3: ADODB Connection
    On Error Resume Next
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    If Err.Number = 0 Then
        result = result & "✓ ADODB Connection: OK" & vbCrLf
    Else
        result = result & "✗ ADODB Connection: FAILED - " & Err.Description & vbCrLf
    End If
    On Error GoTo 0
    
    ' Show results
    MsgBox "Quick Sandbox Test Results:" & vbCrLf & vbCrLf & result & vbCrLf & _
           "If any tests failed, that's why the sandbox won't initialize." & vbCrLf & _
           "Run RunSandboxDiagnostic() for detailed analysis.", vbInformation
End Sub

' View existing log file
Sub ViewDiagnosticLog()
    Dim logPath As String
    Dim fso As Object
    
    logPath = "C:\Lumen\Workflow Manager\SandboxDiagnostic.log"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(logPath) Then
        Shell "notepad.exe """ & logPath & """", vbNormalFocus
    Else
        MsgBox "No diagnostic log file found. Run RunSandboxDiagnostic() first.", vbExclamation
    End If
End Sub
