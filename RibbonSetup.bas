Attribute VB_Name = "RibbonSetup"
'===============================================================================
' BASH Flow Ribbon Setup
' Alternative method to set up ribbon customization via VBA code
' Use this if the .exportedUI import doesn't work
'===============================================================================

Option Explicit

' Global ribbon object
Private ribbonUI As IRibbonUI

'===============================================================================
' RIBBON INITIALIZATION
'===============================================================================

' Initialize ribbon (called by Outlook)
Public Sub InitializeRibbon(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub

' Get ribbon XML (called by Outlook)
Public Function GetRibbonXML() As String
    GetRibbonXML = GetCustomRibbonXML()
End Function

'===============================================================================
' RIBBON XML DEFINITION
'===============================================================================

' Generate the ribbon XML
Private Function GetCustomRibbonXML() As String
    Dim xml As String
    
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xml = xml & "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    xml = xml & "  <ribbon>" & vbCrLf
    xml = xml & "    <tabs>" & vbCrLf
    xml = xml & "      <tab id=""BASHFlowTab"" label=""BASH Flow"">" & vbCrLf
    xml = xml & "        <group id=""DashboardGroup"" label=""Dashboard"">" & vbCrLf
    xml = xml & "          <button id=""ShowDashboard"" label=""Show Dashboard"" size=""large"" imageMso=""DataConnectionGallery"" onAction=""ShowDashboardRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""RefreshDashboard"" label=""Refresh"" size=""normal"" imageMso=""Refresh"" onAction=""RefreshDashboardRibbon"" />" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "        <group id=""WorkflowGroup"" label=""Workflow"">" & vbCrLf
    xml = xml & "          <button id=""ProcessNew"" label=""Process New"" size=""large"" imageMso=""MailMergeGoToFirstRecord"" onAction=""ProcessNewWorkflowItemsRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""ViewStatus"" label=""View Status"" size=""normal"" imageMso=""ViewWorkflowTasks"" onAction=""ViewAllRecordsRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""ExportData"" label=""Export"" size=""normal"" imageMso=""ExportExcel"" onAction=""ExportDashboardDataRibbon"" />" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "        <group id=""SandboxGroup"" label=""Sandbox"">" & vbCrLf
    xml = xml & "          <button id=""InitializeSandbox"" label=""Initialize Sandbox"" size=""large"" imageMso=""DatabaseOpen"" onAction=""InitializeSandboxRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""FixDatabase"" label=""Fix Database"" size=""normal"" imageMso=""DatabaseRepair"" onAction=""FixDatabaseTableRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""CheckDatabase"" label=""Check Database"" size=""normal"" imageMso=""DatabaseView"" onAction=""CheckDatabaseStructureRibbon"" />" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "        <group id=""DiagnosticGroup"" label=""Diagnostics"">" & vbCrLf
    xml = xml & "          <button id=""RunDiagnostics"" label=""Run Diagnostics"" size=""large"" imageMso=""Diagnostics"" onAction=""RunSandboxDiagnosticsRibbon"" />" & vbCrLf
    xml = xml & "          <button id=""ViewLogs"" label=""View Logs"" size=""normal"" imageMso=""ViewLogs"" onAction=""ViewSandboxLogsRibbon"" />" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "      </tab>" & vbCrLf
    xml = xml & "    </tabs>" & vbCrLf
    xml = xml & "  </ribbon>" & vbCrLf
    xml = xml & "</customUI>"
    
    GetCustomRibbonXML = xml
End Function

'===============================================================================
' RIBBON CALLBACKS
'===============================================================================

' Dashboard ribbon callbacks
Public Sub ShowDashboardRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.ShowWorkflowDashboard
    Exit Sub
ErrorHandler:
    MsgBox "Error opening dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

Public Sub RefreshDashboardRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.RefreshDashboardData
    Exit Sub
ErrorHandler:
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical, "Refresh Error"
End Sub

' Workflow ribbon callbacks
Public Sub ProcessNewWorkflowItemsRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.ProcessNewWorkflowItems
    Exit Sub
ErrorHandler:
    MsgBox "Error processing workflow items: " & Err.Description, vbCritical, "Processing Error"
End Sub

Public Sub ViewAllRecordsRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.ViewAllRecords
    Exit Sub
ErrorHandler:
    MsgBox "Error viewing records: " & Err.Description, vbCritical, "View Error"
End Sub

Public Sub ExportDashboardDataRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.ExportDashboardData
    Exit Sub
ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vbCritical, "Export Error"
End Sub

' Sandbox ribbon callbacks
Public Sub InitializeSandboxRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANBASH_Flow_Sandbox.InitializeSandbox
    Exit Sub
ErrorHandler:
    MsgBox "Error initializing sandbox: " & Err.Description, vbCritical, "Sandbox Error"
End Sub

Public Sub FixDatabaseTableRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANBASH_Flow_Sandbox.FixDatabaseTable
    Exit Sub
ErrorHandler:
    MsgBox "Error fixing database: " & Err.Description, vbCritical, "Database Error"
End Sub

Public Sub CheckDatabaseStructureRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANCheckDatabase.CheckDatabaseStructure
    Exit Sub
ErrorHandler:
    MsgBox "Error checking database: " & Err.Description, vbCritical, "Database Check Error"
End Sub

' Diagnostic ribbon callbacks
Public Sub RunSandboxDiagnosticsRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANSandboxDiagnostic.RunSandboxDiagnostics
    Exit Sub
ErrorHandler:
    MsgBox "Error running diagnostics: " & Err.Description, vbCritical, "Diagnostics Error"
End Sub

Public Sub ViewSandboxLogsRibbon(control As IRibbonControl)
    On Error GoTo ErrorHandler
    MSCANWorkflowDashboard.ViewSandboxLogs
    Exit Sub
ErrorHandler:
    MsgBox "Error viewing logs: " & Err.Description, vbCritical, "Log Error"
End Sub

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

' Test ribbon setup
Public Sub TestRibbonSetup()
    On Error GoTo ErrorHandler
    
    Dim xml As String
    xml = GetCustomRibbonXML()
    
    ' Save XML to file for inspection
    Dim fso As Object
    Dim file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile("C:\Lumen\Workflow Manager\RibbonTest.xml", True)
    file.Write xml
    file.Close
    
    MsgBox "Ribbon XML generated and saved to RibbonTest.xml" & vbCrLf & vbCrLf & _
           "XML Length: " & Len(xml) & " characters" & vbCrLf & _
           "You can inspect the file to verify the ribbon definition.", vbInformation, "Ribbon Test"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error testing ribbon: " & Err.Description, vbCritical, "Ribbon Test Error"
End Sub

' Manual ribbon setup instructions
Public Sub ShowRibbonSetupInstructions()
    On Error GoTo ErrorHandler
    
    Dim instructions As String
    instructions = "🎯 BASH Flow Ribbon Setup Instructions" & vbCrLf & vbCrLf
    instructions = instructions & "Method 1 - Import File:" & vbCrLf
    instructions = instructions & "1. Go to File → Options → Customize Ribbon" & vbCrLf
    instructions = instructions & "2. Click Import/Export → Import customization file" & vbCrLf
    instructions = instructions & "3. Select: BASH_Flow_Ribbon.exportedUI" & vbCrLf & vbCrLf
    instructions = instructions & "Method 2 - Manual Setup:" & vbCrLf
    instructions = instructions & "1. Go to File → Options → Customize Ribbon" & vbCrLf
    instructions = instructions & "2. Create new tab: 'BASH Flow'" & vbCrLf
    instructions = instructions & "3. Add groups: Dashboard, Workflow, Sandbox, Diagnostics" & vbCrLf
    instructions = instructions & "4. Add buttons and assign macros" & vbCrLf & vbCrLf
    instructions = instructions & "Method 3 - VBA Code:" & vbCrLf
    instructions = instructions & "1. Run TestRibbonSetup() to generate XML" & vbCrLf
    instructions = instructions & "2. Copy the generated XML" & vbCrLf
    instructions = instructions & "3. Use it in your ribbon customization"
    
    MsgBox instructions, vbInformation, "Ribbon Setup Guide"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing instructions: " & Err.Description, vbCritical, "Instructions Error"
End Sub
