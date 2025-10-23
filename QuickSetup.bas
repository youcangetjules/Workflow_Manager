Attribute VB_Name = "QuickSetup"
'===============================================================================
' Quick Setup Script for BASH Flow VBA Modules
' Run this script in Outlook VBA Editor to automatically import all modules
'===============================================================================

Option Explicit

'===============================================================================
' MAIN SETUP FUNCTION
'===============================================================================

' Main setup function - imports all VBA modules
Public Sub SetupBASHFlowModules()
    On Error GoTo ErrorHandler
    
    Dim modulePath As String
    Dim successCount As Integer
    Dim totalCount As Integer
    
    ' Set the path to your VBA modules
    modulePath = "C:\Lumen\Workflow Manager\"
    successCount = 0
    totalCount = 8
    
    ' Show progress
    MsgBox "Starting BASH Flow module setup..." & vbCrLf & vbCrLf & _
           "This will import 8 VBA modules into your Outlook project.", vbInformation, "BASH Flow Setup"
    
    ' Import each module
    If ImportModule(modulePath & "BASH_Flow_Sandbox.bas", "BASH_Flow_Sandbox") Then successCount = successCount + 1
    If ImportModule(modulePath & "WorkflowManager.bas", "WorkflowManager") Then successCount = successCount + 1
    If ImportModule(modulePath & "WorkflowDashboard.bas", "WorkflowDashboard") Then successCount = successCount + 1
    If ImportModule(modulePath & "DashboardSetup.bas", "DashboardSetup") Then successCount = successCount + 1
    If ImportModule(modulePath & "SandboxDiagnostic.bas", "SandboxDiagnostic") Then successCount = successCount + 1
    If ImportModule(modulePath & "CheckDatabase.bas", "CheckDatabase") Then successCount = successCount + 1
    If ImportModule(modulePath & "RibbonSetup.bas", "RibbonSetup") Then successCount = successCount + 1
    If ImportModule(modulePath & "CreateRibbonCustomization.bas", "CreateRibbonCustomization") Then successCount = successCount + 1
    
    ' Show results
    If successCount = totalCount Then
        MsgBox "✅ Setup Complete!" & vbCrLf & vbCrLf & _
               "Successfully imported " & successCount & " out of " & totalCount & " modules." & vbCrLf & vbCrLf & _
               "You can now use:" & vbCrLf & _
               "• MSCANBASH_Flow_Sandbox.InitializeSandbox() - Set up sandbox environment" & vbCrLf & _
               "• MSCANWorkflowDashboard.ShowWorkflowDashboard() - Open dashboard" & vbCrLf & _
               "• MSCANDashboardSetup.SetupDashboard() - Complete dashboard setup", vbInformation, "Setup Complete"
    Else
        MsgBox "⚠️ Partial Setup Complete" & vbCrLf & vbCrLf & _
               "Imported " & successCount & " out of " & totalCount & " modules." & vbCrLf & _
               "Some modules may already exist or there may have been errors.", vbExclamation, "Setup Warning"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Setup Error: " & Err.Description, vbCritical, "Setup Failed"
End Sub

'===============================================================================
' HELPER FUNCTIONS
'===============================================================================

' Import a single module, checking if it already exists
Private Function ImportModule(modulePath As String, moduleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim comp As Object
    
    ' Check if module already exists
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Name = moduleName Then
            ' Module exists, remove it first
            Application.VBE.ActiveVBProject.VBComponents.Remove comp
            Exit For
        End If
    Next comp
    
    ' Import the module
    Application.VBE.ActiveVBProject.VBComponents.Import modulePath
    
    ImportModule = True
    Exit Function
    
ErrorHandler:
    ImportModule = False
End Function

'===============================================================================
' TESTING FUNCTIONS
'===============================================================================

' Test if all modules are loaded correctly
Public Sub TestBASHFlowModules()
    On Error GoTo ErrorHandler
    
    Dim moduleNames As String
    Dim comp As Object
    Dim foundCount As Integer
    Dim expectedModules As String
    
    expectedModules = "BASH_Flow_Sandbox,WorkflowManager,WorkflowDashboard,DashboardSetup,SandboxDiagnostic,CheckDatabase,RibbonSetup,CreateRibbonCustomization"
    foundCount = 0
    
    ' Check each expected module
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If InStr(expectedModules, comp.Name) > 0 Then
            foundCount = foundCount + 1
            moduleNames = moduleNames & "✅ " & comp.Name & vbCrLf
        End If
    Next comp
    
    ' Show results
    If foundCount = 8 Then
        MsgBox "🎯 All BASH Flow modules are loaded correctly!" & vbCrLf & vbCrLf & _
               "Found modules:" & vbCrLf & moduleNames, vbInformation, "Module Test - PASSED"
    Else
        MsgBox "⚠️ Some modules are missing!" & vbCrLf & vbCrLf & _
               "Found " & foundCount & " out of 8 expected modules:" & vbCrLf & vbCrLf & _
               moduleNames & vbCrLf & _
               "Please run SetupBASHFlowModules() to import missing modules.", vbExclamation, "Module Test - PARTIAL"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Test Error: " & Err.Description, vbCritical, "Test Failed"
End Sub

'===============================================================================
' CLEANUP FUNCTIONS
'===============================================================================

' Remove all BASH Flow modules (for clean reinstall)
Public Sub RemoveBASHFlowModules()
    On Error GoTo ErrorHandler
    
    Dim comp As Object
    Dim removedCount As Integer
    
    If MsgBox("This will remove all BASH Flow modules from your project." & vbCrLf & vbCrLf & _
              "Are you sure you want to continue?", vbYesNo + vbQuestion, "Remove Modules") = vbNo Then
        Exit Sub
    End If
    
    ' Remove each BASH Flow module
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Name Like "*BASH*" Or comp.Name Like "*Workflow*" Or comp.Name Like "*Dashboard*" Or comp.Name Like "*Sandbox*" Or comp.Name Like "*Check*" Then
            If comp.Name <> "QuickSetup" Then ' Don't remove this module
                Application.VBE.ActiveVBProject.VBComponents.Remove comp
                removedCount = removedCount + 1
            End If
        End If
    Next comp
    
    MsgBox "✅ Removed " & removedCount & " BASH Flow modules." & vbCrLf & vbCrLf & _
           "You can now run SetupBASHFlowModules() to reinstall them.", vbInformation, "Modules Removed"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ Removal Error: " & Err.Description, vbCritical, "Removal Failed"
End Sub

'===============================================================================
' QUICK ACCESS FUNCTIONS
'===============================================================================

' Quick access to dashboard setup
Public Sub QuickDashboardSetup()
    On Error GoTo ErrorHandler
    
    ' Test if modules are loaded
    TestBASHFlowModules
    
    ' Try to run dashboard setup
    SetupDashboard
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Dashboard setup failed. Please ensure all modules are loaded first.", vbExclamation, "Dashboard Setup Error"
End Sub

' Quick access to sandbox initialization
Public Sub QuickSandboxInit()
    On Error GoTo ErrorHandler
    
    ' Try to initialize sandbox
    MSCANBASH_Flow_Sandbox.InitializeSandbox
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Sandbox initialization failed. Please ensure BASH_Flow_Sandbox module is loaded.", vbExclamation, "Sandbox Error"
End Sub

' Quick access to ribbon setup
Public Sub QuickRibbonSetup()
    On Error GoTo ErrorHandler
    
    ' Try to run ribbon creation
    CreateBASHFlowRibbon
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ribbon setup failed. Please ensure CreateRibbonCustomization module is loaded.", vbExclamation, "Ribbon Error"
End Sub

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

' Show help information
Public Sub ShowBASHFlowHelp()
    On Error GoTo ErrorHandler
    
    Dim helpMsg As String
    helpMsg = "🚀 BASH Flow VBA Integration Help" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Setup Functions:" & vbCrLf
    helpMsg = helpMsg & "• SetupBASHFlowModules() - Import all modules" & vbCrLf
    helpMsg = helpMsg & "• TestBASHFlowModules() - Test module loading" & vbCrLf
    helpMsg = helpMsg & "• RemoveBASHFlowModules() - Clean removal" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Quick Access:" & vbCrLf
    helpMsg = helpMsg & "• QuickDashboardSetup() - Set up dashboard" & vbCrLf
    helpMsg = helpMsg & "• QuickSandboxInit() - Initialize sandbox" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "Main Functions (after setup):" & vbCrLf
    helpMsg = helpMsg & "• MSCANBASH_Flow_Sandbox.InitializeSandbox() - Set up sandbox environment" & vbCrLf
    helpMsg = helpMsg & "• MSCANWorkflowDashboard.ShowWorkflowDashboard() - Open main dashboard" & vbCrLf
    helpMsg = helpMsg & "• MSCANBASH_Flow_Sandbox.ProcessSandboxWorkflowEmails() - Process workflow items" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "File Location: C:\Lumen\Workflow Manager\" & vbCrLf
    helpMsg = helpMsg & "Documentation: VBA_Integration_Guide.md"
    
    MsgBox helpMsg, vbInformation, "BASH Flow Help"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing help: " & Err.Description, vbCritical, "Help Error"
End Sub
