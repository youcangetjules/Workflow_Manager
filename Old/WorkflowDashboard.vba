'===============================================================================
' BASH Flow Management Dashboard
' Created for Outlook VBA Environment
' Provides a comprehensive entry dashboard for workflow management
'===============================================================================

Option Explicit

' Dashboard Form Controls
Private dashboardForm As Object
Private dashboardRibbonUI As IRibbonUI

' Dashboard Statistics
Private Type DashboardStats
    TotalRecords As Long
    CLLIRecords As Long
    MSRecords As Long
    EventRecords As Long
    PendingRecords As Long
    Urgency1Records As Long
    Urgency2Records As Long
    Urgency3Records As Long
    LastUpdated As String
End Type

' Dashboard Constants
Private Const DASHBOARD_TITLE As String = "BASH Flow Management Dashboard"
Private Const REFRESH_INTERVAL As Long = 30 ' seconds

'===============================================================================
' MAIN DASHBOARD FUNCTIONS
'===============================================================================

' Initialize the dashboard
Public Sub ShowWorkflowDashboard()
    On Error GoTo ErrorHandler
    
    ' Check if sandbox is available
    If Not IsSandboxAvailable() Then
        MsgBox "Sandbox environment not available. Please run BASH_Flow_Sandbox.InitializeSandbox() first.", vbExclamation, "Dashboard Error"
        Exit Sub
    End If
    
    ' Show the simple dashboard (which is the current working implementation)
    ShowSimpleDashboard
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

' Create the dashboard form
Private Sub CreateDashboardForm()
    On Error GoTo ErrorHandler
    
    ' Create a simple message-based dashboard instead of complex form
    ' This avoids ActiveX control creation issues that can occur in Outlook VBA
    ' The simple dashboard provides all necessary functionality without form complexity
    ShowSimpleDashboard
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating dashboard form: " & Err.Description, vbCritical, "Form Creation Error"
End Sub

' Show simple dashboard using message boxes and input dialogs
Private Sub ShowSimpleDashboard()
    On Error GoTo ErrorHandler
    
    Dim stats As DashboardStats
    Dim dashboardMsg As String
    Dim userChoice As String
    
    ' Get dashboard statistics
    stats = GetDashboardStatistics
    
    ' Build dashboard message
    dashboardMsg = "=== BASH FLOW MANAGEMENT DASHBOARD ===" & vbCrLf & vbCrLf
    dashboardMsg = dashboardMsg & "WORKFLOW STATISTICS:" & vbCrLf
    dashboardMsg = dashboardMsg & "• Total Records: " & stats.TotalRecords & vbCrLf
    dashboardMsg = dashboardMsg & "• CLLI Records: " & stats.CLLIRecords & vbCrLf
    dashboardMsg = dashboardMsg & "• MS Records: " & stats.MSRecords & vbCrLf
    dashboardMsg = dashboardMsg & "• Event Records: " & stats.EventRecords & vbCrLf
    dashboardMsg = dashboardMsg & "• Pending Records: " & stats.PendingRecords & vbCrLf & vbCrLf
    dashboardMsg = dashboardMsg & "URGENCY STATUS:" & vbCrLf
    dashboardMsg = dashboardMsg & "• Urgency 1: " & stats.Urgency1Records & vbCrLf
    dashboardMsg = dashboardMsg & "• Urgency 2: " & stats.Urgency2Records & vbCrLf
    dashboardMsg = dashboardMsg & "• Urgency 3: " & stats.Urgency3Records & vbCrLf & vbCrLf
    dashboardMsg = dashboardMsg & "Last Updated: " & stats.LastUpdated & vbCrLf & vbCrLf
    dashboardMsg = dashboardMsg & "QUICK ACTIONS:" & vbCrLf
    dashboardMsg = dashboardMsg & "1 - Refresh Data" & vbCrLf
    dashboardMsg = dashboardMsg & "2 - Process New Items" & vbCrLf
    dashboardMsg = dashboardMsg & "3 - View All Records" & vbCrLf
    dashboardMsg = dashboardMsg & "4 - Export Data" & vbCrLf
    dashboardMsg = dashboardMsg & "5 - Settings" & vbCrLf
    dashboardMsg = dashboardMsg & "0 - Close Dashboard" & vbCrLf & vbCrLf
    dashboardMsg = dashboardMsg & "Enter your choice (0-5):"
    
    ' Show dashboard and get user choice
    userChoice = InputBox(dashboardMsg, DASHBOARD_TITLE, "1")
    
    ' Process user choice
    Select Case userChoice
        Case "1"
            RefreshDashboardData
        Case "2"
            ProcessNewWorkflowItems
        Case "3"
            ViewAllRecords
        Case "4"
            ExportDashboardData
        Case "5"
            ShowDashboardSettings
        Case "0"
            ' Close dashboard - no action needed
            Exit Sub
        Case ""
            ' User cancelled - no action needed
            Exit Sub
        Case Else
            MsgBox "Invalid choice. Please try again.", vbExclamation, "Invalid Selection"
            ShowSimpleDashboard ' Show dashboard again
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing simple dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

'===============================================================================
' SIMPLE DASHBOARD HELPERS
'===============================================================================

' These functions are no longer needed as we use a simple message-based dashboard
' instead of complex ActiveX form controls to avoid ActiveX failures

'===============================================================================
' DASHBOARD DATA FUNCTIONS
'===============================================================================

' Load dashboard data (now handled in ShowSimpleDashboard)
Private Sub LoadDashboardData()
    On Error GoTo ErrorHandler
    
    ' For simple dashboard, we refresh the data by showing the dashboard again
    ' This ensures fresh statistics are displayed
    ShowSimpleDashboard
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading dashboard data: " & Err.Description, vbCritical, "Data Loading Error"
End Sub

' Get dashboard statistics
Private Function GetDashboardStatistics() As DashboardStats
    On Error GoTo ErrorHandler
    
    Dim stats As DashboardStats
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    
    ' Initialize connection
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Open connection to sandbox database (auto-detect ACE provider)
    Call OpenSandboxConnection(conn)
    
    ' Get total records
    sql = "SELECT COUNT(*) AS TotalCount FROM BASHFlowSandbox"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.TotalRecords = rs("TotalCount")
    rs.Close
    
    ' Get CLLI records
    sql = "SELECT COUNT(*) AS CLLICount FROM BASHFlowSandbox WHERE [Type] = 'CLLI'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.CLLIRecords = rs("CLLICount")
    rs.Close
    
    ' Get MS records
    sql = "SELECT COUNT(*) AS MSCount FROM BASHFlowSandbox WHERE [Type] = 'MS'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.MSRecords = rs("MSCount")
    rs.Close
    
    ' Get Event records
    sql = "SELECT COUNT(*) AS EventCount FROM BASHFlowSandbox WHERE [Type] = 'Event'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.EventRecords = rs("EventCount")
    rs.Close
    
    ' Get pending records
    sql = "SELECT COUNT(*) AS PendingCount FROM BASHFlowSandbox WHERE [Status] LIKE '%Pending%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.PendingRecords = rs("PendingCount")
    rs.Close
    
    ' Get urgency records
    sql = "SELECT COUNT(*) AS Urgency1Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 1%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency1Records = rs("Urgency1Count")
    rs.Close
    
    sql = "SELECT COUNT(*) AS Urgency2Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 2%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency2Records = rs("Urgency2Count")
    rs.Close
    
    sql = "SELECT COUNT(*) AS Urgency3Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 3%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency3Records = rs("Urgency3Count")
    rs.Close
    
    ' Set last updated time
    stats.LastUpdated = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Close connection
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    
    GetDashboardStatistics = stats
    
    Exit Function
    
ErrorHandler:
    stats.LastUpdated = "Error: " & Err.Description
    GetDashboardStatistics = stats
    
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
End Function

' Load recent activity (now integrated into simple dashboard display)
Private Sub LoadRecentActivity()
    ' This function is no longer needed as recent activity is shown
    ' in the simple dashboard message format
End Sub

' Open ADODB connection to the sandbox database, trying ACE 16.0 then 12.0
Private Sub OpenSandboxConnection(ByRef conn As Object)
    On Error GoTo Try12
    Dim connStr As String
    connStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    conn.Open connStr
    Exit Sub
Try12:
    On Error GoTo TryJet
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    Exit Sub
TryJet:
    On Error GoTo Fail
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BASHFlowSandbox\TestDatabase.mdb;"
    Exit Sub
Fail:
    Err.Raise vbObjectError + 5100, "OpenSandboxConnection", "Database connection failed. Please ensure:" & vbCrLf & vbCrLf & _
        "1. Microsoft Access Database Engine (ACE OLE DB) is installed" & vbCrLf & _
        "2. The sandbox database exists at C:\BASHFlowSandbox\TestDatabase.accdb" & vbCrLf & _
        "3. You have proper file permissions" & vbCrLf & vbCrLf & _
        "Run BASH_Flow_Sandbox.InitializeSandbox() to create the database."
End Sub

'===============================================================================
' DASHBOARD EVENT HANDLERS
'===============================================================================

' Refresh dashboard data
Public Sub RefreshDashboardData()
    On Error GoTo ErrorHandler
    
    LoadDashboardData
    MsgBox "Dashboard data refreshed successfully!", vbInformation, "Dashboard Refresh"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical, "Refresh Error"
End Sub

' Process new workflow items
Public Sub ProcessNewWorkflowItems()
    On Error GoTo ErrorHandler
    
    ' Call the existing process function from BASH_Flow_Sandbox.vba
    BASH_Flow_Sandbox.ProcessSandboxWorkflowEmails
    
    ' Refresh dashboard after processing
    LoadDashboardData
    
    MsgBox "New workflow items processed successfully!", vbInformation, "Process Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing new items: " & Err.Description, vbCritical, "Processing Error"
End Sub

' View all records
Public Sub ViewAllRecords()
    On Error GoTo ErrorHandler
    
    ' Call the existing view function from BASH_Flow_Sandbox.vba
    BASH_Flow_Sandbox.ShowSandboxWorkflowStatus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error viewing records: " & Err.Description, vbCritical, "View Error"
End Sub

' Export dashboard data
Public Sub ExportDashboardData()
    On Error GoTo ErrorHandler
    
    ' Call the existing export function from BASH_Flow_Sandbox.vba
    BASH_Flow_Sandbox.ExportSandboxData
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vbCritical, "Export Error"
End Sub

' Show dashboard settings
Public Sub ShowDashboardSettings()
    On Error GoTo ErrorHandler
    
    Dim settingsMsg As String
    settingsMsg = "Dashboard Settings:" & vbCrLf & vbCrLf
    settingsMsg = settingsMsg & "• Refresh Interval: " & REFRESH_INTERVAL & " seconds" & vbCrLf
    settingsMsg = settingsMsg & "• Data Source: C:\BASHFlowSandbox\TestDatabase.accdb" & vbCrLf
    settingsMsg = settingsMsg & "• Auto-refresh: Disabled" & vbCrLf & vbCrLf
    settingsMsg = settingsMsg & "To modify settings, edit the constants in WorkflowDashboard.vba"
    
    MsgBox settingsMsg, vbInformation, "Dashboard Settings"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing settings: " & Err.Description, vbCritical, "Settings Error"
End Sub

' Close dashboard
Public Sub CloseDashboard()
    On Error GoTo ErrorHandler
    
    ' For simple dashboard, we just exit - no form to close
    MsgBox "Dashboard closed.", vbInformation, "Dashboard"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error closing dashboard: " & Err.Description, vbCritical, "Close Error"
End Sub

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

' Check if sandbox is available
Private Function IsSandboxAvailable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if sandbox folder exists
    If Not fso.FolderExists("C:\BASHFlowSandbox") Then
        IsSandboxAvailable = False
        Exit Function
    End If
    
    ' Check if database exists
    If Not fso.FileExists("C:\BASHFlowSandbox\TestDatabase.accdb") Then
        IsSandboxAvailable = False
        Exit Function
    End If
    
    ' Check if database is accessible
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Call OpenSandboxConnection(conn)
    If conn.State = 1 Then
        conn.Close
        IsSandboxAvailable = True
    Else
        IsSandboxAvailable = False
    End If
    
    Exit Function
    
ErrorHandler:
    IsSandboxAvailable = False
End Function

'===============================================================================
' RIBBON INTEGRATION
'===============================================================================

' Initialize dashboard ribbon
Public Sub InitializeDashboardRibbon(ribbon As IRibbonUI)
    Set dashboardRibbonUI = ribbon
End Sub

' Dashboard ribbon callback
Public Sub ShowDashboardRibbon(control As IRibbonControl)
    ShowWorkflowDashboard
End Sub

' Refresh ribbon callback
Public Sub RefreshDashboardRibbon(control As IRibbonControl)
    ' For simple dashboard, we just refresh the data
    RefreshDashboardData
End Sub

' Ribbon callbacks for other functions
Public Sub ProcessNewWorkflowItemsRibbon(control As IRibbonControl)
    ProcessNewWorkflowItems
End Sub

Public Sub ViewAllRecordsRibbon(control As IRibbonControl)
    ViewAllRecords
End Sub

Public Sub ExportDashboardDataRibbon(control As IRibbonControl)
    ExportDashboardData
End Sub

Public Sub InitializeSandboxRibbon(control As IRibbonControl)
    BASH_Flow_Sandbox.InitializeSandbox
End Sub

Public Sub FixDatabaseTableRibbon(control As IRibbonControl)
    BASH_Flow_Sandbox.FixDatabaseTable
End Sub

Public Sub CheckDatabaseStructureRibbon(control As IRibbonControl)
    BASHCheckDatabase.CheckDatabaseStructure
End Sub

Public Sub RunSandboxDiagnosticsRibbon(control As IRibbonControl)
    BASHSandboxDiagnostic.RunSandboxDiagnostic
End Sub

Public Sub ViewSandboxLogsRibbon(control As IRibbonControl)
    ViewSandboxLogs
End Sub

'===============================================================================
' ADDITIONAL UTILITY FUNCTIONS
'===============================================================================

' View sandbox logs
Private Sub ViewSandboxLogs()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim logFile As Object
    Dim logContent As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists("C:\BASHFlowSandbox\SandboxLog.txt") Then
        Set logFile = fso.OpenTextFile("C:\BASHFlowSandbox\SandboxLog.txt", 1)
        logContent = logFile.ReadAll
        logFile.Close
        
        ' Show log content in a message box (truncated if too long)
        If Len(logContent) > 3000 Then
            logContent = Left(logContent, 3000) & vbCrLf & "... (truncated)"
        End If
        
        MsgBox logContent, vbInformation, "Sandbox Log"
    Else
        MsgBox "Log file not found at C:\BASHFlowSandbox\SandboxLog.txt", vbExclamation, "Log Not Found"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error viewing logs: " & Err.Description, vbCritical, "Log View Error"
End Sub

' Check database structure
Private Sub CheckDatabaseStructure()
    On Error GoTo ErrorHandler
    
    ' This function should call the CheckDatabase.vba module
    ' For now, we'll create a simple implementation
    MsgBox "Database structure check would be performed here. Please run CheckDatabaseStructure() from BASHCheckDatabase.bas for detailed analysis.", vbInformation, "Database Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking database structure: " & Err.Description, vbCritical, "Database Check Error"
End Sub

' Run sandbox diagnostics
Private Sub RunSandboxDiagnostics()
    On Error GoTo ErrorHandler
    
    ' This function should call the SandboxDiagnostic.vba module
    ' For now, we'll create a simple implementation
    MsgBox "System diagnostics would be performed here. Please run RunSandboxDiagnostic() from BASHSandboxDiagnostic.bas for detailed analysis.", vbInformation, "System Diagnostics"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error running diagnostics: " & Err.Description, vbCritical, "Diagnostics Error"
End Sub
