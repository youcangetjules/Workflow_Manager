'===============================================================================
' BASH Flow Dashboard Setup Script
' Run this script to set up the dashboard integration with Outlook
'===============================================================================

Option Explicit

'===============================================================================
' SETUP FUNCTIONS
'===============================================================================

' Main setup function - call this to set up the dashboard
Public Sub SetupDashboard()
    On Error GoTo ErrorHandler
    
    Dim setupMsg As String
    setupMsg = "BASH Flow Dashboard Setup" & vbCrLf & vbCrLf
    setupMsg = setupMsg & "This will set up the dashboard integration with Outlook." & vbCrLf & vbCrLf
    setupMsg = setupMsg & "Steps:" & vbCrLf
    setupMsg = setupMsg & "1. Load WorkflowDashboard.vba module" & vbCrLf
    setupMsg = setupMsg & "2. Set up ribbon customization" & vbCrLf
    setupMsg = setupMsg & "3. Test dashboard functionality" & vbCrLf & vbCrLf
    setupMsg = setupMsg & "Do you want to proceed?"
    
    If MsgBox(setupMsg, vbYesNo + vbQuestion, "Dashboard Setup") = vbYes Then
        LoadDashboardModule
        SetupRibbonCustomization
        TestDashboardFunctionality
        MsgBox "Dashboard setup completed successfully!", vbInformation, "Setup Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during setup: " & Err.Description, vbCritical, "Setup Error"
End Sub

' Load the dashboard module
Private Sub LoadDashboardModule()
    On Error GoTo ErrorHandler
    
    ' This would normally load the module, but in VBA it's done manually
    ' We'll just show instructions
    MsgBox "Please manually load the WorkflowDashboard.vba module into your Outlook VBA project." & vbCrLf & vbCrLf & _
           "Steps:" & vbCrLf & _
           "1. Open VBA Editor (Alt+F11)" & vbCrLf & _
           "2. Right-click on your project" & vbCrLf & _
           "3. Select 'Import File...'" & vbCrLf & _
           "4. Select WorkflowDashboard.vba", vbInformation, "Load Module"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading module: " & Err.Description, vbCritical, "Module Load Error"
End Sub

' Set up ribbon customization
Private Sub SetupRibbonCustomization()
    On Error GoTo ErrorHandler
    
    MsgBox "Ribbon customization setup:" & vbCrLf & vbCrLf & _
           "1. Copy DashboardRibbon.xml to your Outlook customization folder" & vbCrLf & _
           "2. Register the customization in Outlook" & vbCrLf & _
           "3. Restart Outlook to see the new 'BASH Flow' tab", vbInformation, "Ribbon Setup"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error setting up ribbon: " & Err.Description, vbCritical, "Ribbon Setup Error"
End Sub

' Test dashboard functionality
Private Sub TestDashboardFunctionality()
    On Error GoTo ErrorHandler
    
    ' Test if sandbox is available
    If IsSandboxAvailable() Then
        MsgBox "Sandbox environment is available. Dashboard should work correctly.", vbInformation, "Test Result"
    Else
        MsgBox "Sandbox environment not found. Please run InitializeSandbox() first.", vbExclamation, "Test Result"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error testing functionality: " & Err.Description, vbCritical, "Test Error"
End Sub

' Check if sandbox is available (copied from WorkflowDashboard.vba)
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
' QUICK ACCESS FUNCTIONS
'===============================================================================

' Quick access to dashboard
Public Sub QuickDashboard()
    On Error GoTo ErrorHandler
    
    ' Try to show dashboard directly
    ShowWorkflowDashboard
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Dashboard not available. Please run SetupDashboard() first.", vbExclamation, "Dashboard Not Available"
End Sub

' Quick access to sandbox initialization
Public Sub QuickSandboxInit()
    On Error GoTo ErrorHandler
    
    ' Try to initialize sandbox
    InitializeSandbox
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Sandbox initialization failed. Please check the BASH_Flow_Sandbox.vba module is loaded.", vbCritical, "Sandbox Error"
End Sub

'===============================================================================
' HELPER FUNCTIONS
'===============================================================================

' Show dashboard help
Public Sub ShowDashboardHelp()
    On Error GoTo ErrorHandler
    
    Dim helpMsg As String
    helpMsg = "BASH Flow Dashboard Help" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Main Functions:" & vbCrLf
    helpMsg = helpMsg & "• ShowWorkflowDashboard() - Opens the main dashboard" & vbCrLf
    helpMsg = helpMsg & "• RefreshDashboardData() - Refreshes dashboard statistics" & vbCrLf
    helpMsg = helpMsg & "• ProcessNewWorkflowItems() - Processes new workflow items" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Setup Functions:" & vbCrLf
    helpMsg = helpMsg & "• SetupDashboard() - Complete setup process" & vbCrLf
    helpMsg = helpMsg & "• QuickDashboard() - Quick access to dashboard" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Requirements:" & vbCrLf
    helpMsg = helpMsg & "• Sandbox environment must be initialized" & vbCrLf
    helpMsg = helpMsg & "• BASH_Flow_Sandbox.vba module must be loaded" & vbCrLf
    helpMsg = helpMsg & "• Microsoft Access Database Engine must be installed"
    
    MsgBox helpMsg, vbInformation, "Dashboard Help"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing help: " & Err.Description, vbCritical, "Help Error"
End Sub
