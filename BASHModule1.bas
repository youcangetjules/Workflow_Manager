Attribute VB_Name = "BASHModule1"
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
        MsgBox "Sandbox environment not available. Please run InitializeSandbox() first.", vbExclamation, "Dashboard Error"
        Exit Sub
    End If
    
    ' Create and show dashboard form
    CreateDashboardForm
    LoadDashboardData
    dashboardForm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

' Create the dashboard form
Private Sub CreateDashboardForm()
    On Error GoTo ErrorHandler
    
    ' Create user form (MSForms UserForm)
    On Error Resume Next
    Set dashboardForm = CreateObject("Forms.UserForm.1")
    If dashboardForm Is Nothing Then
        ' Fallback to VB runtime factory
        Set dashboardForm = VBA.UserForms.Add()
    End If
    On Error GoTo ErrorHandler

    ' Verify MSForms is available for dynamic control creation
    If Not IsMsFormsAvailable() Then
        MsgBox "Microsoft Forms 2.0 controls are not available. Please enable the 'Microsoft Forms 2.0 Object Library' reference and retry.", vbCritical, "Forms Library Missing"
        Exit Sub
    End If
    
    ' Set form properties
    With dashboardForm
        .caption = DASHBOARD_TITLE
        .width = 800
        .height = 600
        .StartUpPosition = 2 ' Center on screen
        .BackColor = RGB(240, 248, 255)
        .BorderStyle = 3 ' Fixed Dialog
    End With
    
    ' Add controls to form
    AddDashboardControls
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating dashboard form: " & Err.Description, vbCritical, "Form Creation Error"
End Sub

' Add controls to the dashboard form
Private Sub AddDashboardControls()
    On Error GoTo ErrorHandler
    
    ' Title Label
    AddLabel "lblTitle", "BASH Flow Management Dashboard", 20, 20, 760, 30, True, 14
    
    ' Statistics Frame
    AddFrame "fraStats", "Workflow Statistics", 20, 70, 760, 200
    
    ' Record Counts
    AddLabel "lblTotalRecords", "Total Records: 0", 40, 100, 150, 20
    AddLabel "lblCLLIRecords", "CLLI Records: 0", 40, 125, 150, 20
    AddLabel "lblMSRecords", "MS Records: 0", 40, 150, 150, 20
    AddLabel "lblEventRecords", "Event Records: 0", 40, 175, 150, 20
    AddLabel "lblPendingRecords", "Pending Records: 0", 40, 200, 150, 20
    
    ' Status Counts
    AddLabel "lblUrgency1", "Urgency 1: 0", 220, 100, 150, 20
    AddLabel "lblUrgency2", "Urgency 2: 0", 220, 125, 150, 20
    AddLabel "lblUrgency3", "Urgency 3: 0", 220, 150, 150, 20
    AddLabel "lblLastUpdated", "Last Updated: Never", 220, 175, 150, 20
    
    ' Recent Activity Frame
    AddFrame "fraRecent", "Recent Activity", 20, 290, 760, 150
    
    ' Recent Activity List
    AddListBox "lstRecent", 40, 320, 720, 100
    
    ' Control Buttons Frame
    AddFrame "fraControls", "Quick Actions", 20, 460, 760, 100
    
    ' Buttons
    AddButton "btnRefresh", "Refresh Data", 40, 490, 100, 30
    AddButton "btnProcessNew", "Process New Items", 160, 490, 120, 30
    AddButton "btnViewAll", "View All Records", 300, 490, 120, 30
    AddButton "btnExport", "Export Data", 440, 490, 100, 30
    AddButton "btnSettings", "Settings", 560, 490, 100, 30
    AddButton "btnClose", "Close", 680, 490, 80, 30
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding dashboard controls: " & Err.Description, vbCritical, "Control Creation Error"
End Sub

'===============================================================================
' CONTROL CREATION HELPERS
'===============================================================================

Private Function IsMsFormsAvailable() As Boolean
    On Error GoTo Failed
    Dim tmp As Object
    Set tmp = CreateObject("Forms.Label.1")
    IsMsFormsAvailable = Not tmp Is Nothing
    Exit Function
Failed:
    IsMsFormsAvailable = False
End Function

' Add a label control
Private Sub AddLabel(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer, Optional bold As Boolean = False, Optional fontSize As Integer = 9)
    On Error GoTo ErrorHandler
    
    Dim lbl As Object
    Set lbl = dashboardForm.Controls.Add("Forms.Label.1", controlName)
    
    With lbl
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
        .Font.Size = fontSize
        .Font.bold = bold
        .BackStyle = 0 ' Transparent
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding label " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

' Add a frame control
Private Sub AddFrame(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim fra As Object
    Set fra = dashboardForm.Controls.Add("Forms.Frame.1", controlName)
    
    With fra
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
        .BackColor = RGB(255, 255, 255)
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding frame " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

' Add a button control
Private Sub AddButton(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    Set btn = dashboardForm.Controls.Add("Forms.CommandButton.1", controlName)
    
    With btn
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
        .BackColor = RGB(70, 130, 180)
        .ForeColor = RGB(255, 255, 255)
        .Font.bold = True
    End With
    
    ' Add click event handler
    If controlName = "btnRefresh" Then
        btn.OnClick = "RefreshDashboardData"
    ElseIf controlName = "btnProcessNew" Then
        btn.OnClick = "ProcessNewWorkflowItems"
    ElseIf controlName = "btnViewAll" Then
        btn.OnClick = "ViewAllRecords"
    ElseIf controlName = "btnExport" Then
        btn.OnClick = "ExportDashboardData"
    ElseIf controlName = "btnSettings" Then
        btn.OnClick = "ShowDashboardSettings"
    ElseIf controlName = "btnClose" Then
        btn.OnClick = "CloseDashboard"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding button " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

' Add a listbox control
Private Sub AddListBox(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim lst As Object
    Set lst = dashboardForm.Controls.Add("Forms.ListBox.1", controlName)
    
    With lst
        .left = left
        .top = top
        .width = width
        .height = height
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding listbox " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

'===============================================================================
' DASHBOARD DATA FUNCTIONS
'===============================================================================

' Load dashboard data
Private Sub LoadDashboardData()
    On Error GoTo ErrorHandler
    
    Dim stats As DashboardStats
    stats = GetDashboardStatistics
    
    ' Update statistics labels
    dashboardForm.Controls("lblTotalRecords").caption = "Total Records: " & stats.TotalRecords
    dashboardForm.Controls("lblCLLIRecords").caption = "CLLI Records: " & stats.CLLIRecords
    dashboardForm.Controls("lblMSRecords").caption = "MS Records: " & stats.MSRecords
    dashboardForm.Controls("lblEventRecords").caption = "Event Records: " & stats.EventRecords
    dashboardForm.Controls("lblPendingRecords").caption = "Pending Records: " & stats.PendingRecords
    dashboardForm.Controls("lblUrgency1").caption = "Urgency 1: " & stats.Urgency1Records
    dashboardForm.Controls("lblUrgency2").caption = "Urgency 2: " & stats.Urgency2Records
    dashboardForm.Controls("lblUrgency3").caption = "Urgency 3: " & stats.Urgency3Records
    dashboardForm.Controls("lblLastUpdated").caption = "Last Updated: " & stats.LastUpdated
    
    ' Load recent activity
    LoadRecentActivity
    
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
    If rs.State = 1 Then rs.Close
    
    ' Get CLLI records
    sql = "SELECT COUNT(*) AS CLLICount FROM BASHFlowSandbox WHERE [Type] = 'CLLI'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.CLLIRecords = rs("CLLICount")
    If rs.State = 1 Then rs.Close
    
    ' Get MS records
    sql = "SELECT COUNT(*) AS MSCount FROM BASHFlowSandbox WHERE [Type] = 'MS'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.MSRecords = rs("MSCount")
    If rs.State = 1 Then rs.Close
    
    ' Get Event records
    sql = "SELECT COUNT(*) AS EventCount FROM BASHFlowSandbox WHERE [Type] = 'Event'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.EventRecords = rs("EventCount")
    If rs.State = 1 Then rs.Close
    
    ' Get pending records
    sql = "SELECT COUNT(*) AS PendingCount FROM BASHFlowSandbox WHERE [Status] LIKE '%Pending%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.PendingRecords = rs("PendingCount")
    If rs.State = 1 Then rs.Close
    
    ' Get urgency records
    sql = "SELECT COUNT(*) AS Urgency1Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 1%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency1Records = rs("Urgency1Count")
    If rs.State = 1 Then rs.Close
    
    sql = "SELECT COUNT(*) AS Urgency2Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 2%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency2Records = rs("Urgency2Count")
    If rs.State = 1 Then rs.Close
    
    sql = "SELECT COUNT(*) AS Urgency3Count FROM BASHFlowSandbox WHERE [Status] LIKE '%Urgency 3%'"
    rs.Open sql, conn
    If rs.State = 1 And Not rs.EOF Then stats.Urgency3Records = rs("Urgency3Count")
    If rs.State = 1 Then rs.Close
    
    ' Set last updated time
    stats.LastUpdated = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    ' Close connection
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    
    GetDashboardStatistics = stats
    
    Exit Function
    
ErrorHandler:
    stats.LastUpdated = "Error: " & Err.Description
    GetDashboardStatistics = stats
    
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Function

' Load recent activity
Private Sub LoadRecentActivity()
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim lstRecent As Object
    
    Set lstRecent = dashboardForm.Controls("lstRecent")
    lstRecent.Clear
    
    ' Initialize connection
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Open connection to sandbox database (auto-detect ACE provider)
    Call OpenSandboxConnection(conn)
    
    ' Get recent records (last 10)
    sql = "SELECT TOP 10 [Date Created], [Type], [CLLI Number], [Status] FROM BASHFlowSandbox ORDER BY [Date Created] DESC"
    rs.Open sql, conn
    
    Do While Not rs.EOF
        lstRecent.AddItem Format(rs("[Date Created]"), "mm/dd hh:mm") & " - " & _
                         rs("[Type]") & " - " & _
                         rs("[CLLI Number]") & " - " & _
                         rs("[Status]")
        rs.MoveNext
    Loop
    
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    
    Exit Sub
    
ErrorHandler:
    lstRecent.AddItem "Error loading recent activity: " & Err.Description
    
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub

' Open ADODB connection to the sandbox database, trying ACE 16.0 then 12.0
Private Sub OpenSandboxConnection(ByRef conn As Object)
    On Error GoTo Try12
    Dim connStr As String
    connStr = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    conn.Open connStr
    Exit Sub
Try12:
    On Error GoTo Fail
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    Exit Sub
Fail:
    Err.Raise vbObjectError + 5100, "OpenSandboxConnection", "Microsoft Access Database Engine (ACE OLE DB) not installed. Install the 64-bit Access Database Engine matching your Office bitness."
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
    ProcessSandboxWorkflow
    
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
    ShowSandboxWorkflowStatus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error viewing records: " & Err.Description, vbCritical, "View Error"
End Sub

' Export dashboard data
Public Sub ExportDashboardData()
    On Error GoTo ErrorHandler
    
    ' Call the existing export function from BASH_Flow_Sandbox.vba
    ExportSandboxData
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vbCritical, "Export Error"
End Sub

' Show dashboard settings
Public Sub ShowDashboardSettings()
    On Error GoTo ErrorHandler
    
    Dim settingsMsg As String
    settingsMsg = "Dashboard Settings:" & vbCrLf & vbCrLf
    settingsMsg = settingsMsg & "� Refresh Interval: " & REFRESH_INTERVAL & " seconds" & vbCrLf
    settingsMsg = settingsMsg & "� Data Source: C:\BASHFlowSandbox\TestDatabase.accdb" & vbCrLf
    settingsMsg = settingsMsg & "� Auto-refresh: Disabled" & vbCrLf & vbCrLf
    settingsMsg = settingsMsg & "To modify settings, edit the constants in BASHWorkflowDashboard.bas"
    
    MsgBox settingsMsg, vbInformation, "Dashboard Settings"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing settings: " & Err.Description, vbCritical, "Settings Error"
End Sub

' Close dashboard
Public Sub CloseDashboard()
    On Error GoTo ErrorHandler
    
    If Not dashboardForm Is Nothing Then
        dashboardForm.Hide
        Set dashboardForm = Nothing
    End If
    
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
    
    IsSandboxAvailable = True
    
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
    If Not dashboardForm Is Nothing Then
        RefreshDashboardData
    End If
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
    InitializeSandbox
End Sub

Public Sub FixDatabaseTableRibbon(control As IRibbonControl)
    FixDatabaseTable
End Sub

Public Sub CheckDatabaseStructureRibbon(control As IRibbonControl)
    CheckDatabaseStructure
End Sub

Public Sub RunSandboxDiagnosticsRibbon(control As IRibbonControl)
    RunSandboxDiagnostics
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
            logContent = left(logContent, 3000) & vbCrLf & "... (truncated)"
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
    
    ' This function should call the BASHCheckDatabase.bas module
    ' For now, we'll create a simple implementation
    MsgBox "Database structure check would be performed here. Please run CheckDatabaseStructure() from BASHCheckDatabase.bas for detailed analysis.", vbInformation, "Database Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking database structure: " & Err.Description, vbCritical, "Database Check Error"
End Sub

' Run sandbox diagnostics
Private Sub RunSandboxDiagnostics()
    On Error GoTo ErrorHandler
    
    ' This function should call the BASHSandboxDiagnostic.bas module
    ' For now, we'll create a simple implementation
    MsgBox "System diagnostics would be performed here. Please run RunSandboxDiagnostics() from BASHSandboxDiagnostic.bas for detailed analysis.", vbInformation, "System Diagnostics"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error running diagnostics: " & Err.Description, vbCritical, "Diagnostics Error"
End Sub


