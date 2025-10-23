Attribute VB_Name = "CreateRibbonCustomization"
'===============================================================================
' Create Ribbon Customization for BASH Flow
' This module helps create and export proper ribbon customizations
'===============================================================================

Option Explicit

'===============================================================================
' MAIN RIBBON CREATION FUNCTIONS
'===============================================================================

' Create and export ribbon customization
Public Sub CreateBASHFlowRibbon()
    On Error GoTo ErrorHandler
    
    Dim msg As String
    msg = "BASH Flow Ribbon Creation" & vbCrLf & vbCrLf
    msg = msg & "This will help you create a ribbon customization for BASH Flow." & vbCrLf & vbCrLf
    msg = msg & "Steps:" & vbCrLf
    msg = msg & "1. Go to File → Options → Customize Ribbon" & vbCrLf
    msg = msg & "2. Create a new tab called 'BASH Flow'" & vbCrLf
    msg = msg & "3. Add the groups and buttons as shown in the guide" & vbCrLf
    msg = msg & "4. Export the customization as .exportedUI file" & vbCrLf & vbCrLf
    msg = msg & "Would you like to see the detailed instructions?"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "Create Ribbon") = vbYes Then
        ShowDetailedRibbonInstructions
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating ribbon: " & Err.Description, vbCritical, "Ribbon Creation Error"
End Sub

' Show detailed ribbon creation instructions
Private Sub ShowDetailedRibbonInstructions()
    On Error GoTo ErrorHandler
    
    Dim instructions As String
    instructions = "🎯 BASH Flow Ribbon Creation Instructions" & vbCrLf & vbCrLf
    
    instructions = instructions & "STEP 1: Open Ribbon Customization" & vbCrLf
    instructions = instructions & "• Go to File → Options → Customize Ribbon" & vbCrLf & vbCrLf
    
    instructions = instructions & "STEP 2: Create New Tab" & vbCrLf
    instructions = instructions & "• Click 'New Tab' at the bottom" & vbCrLf
    instructions = instructions & "• Right-click the new tab → Rename" & vbCrLf
    instructions = instructions & "• Enter: 'BASH Flow'" & vbCrLf & vbCrLf
    
    instructions = instructions & "STEP 3: Create Groups" & vbCrLf
    instructions = instructions & "Create these 4 groups under BASH Flow tab:" & vbCrLf & vbCrLf
    
    instructions = instructions & "Group 1: Dashboard" & vbCrLf
    instructions = instructions & "• Add button: 'Show Dashboard'" & vbCrLf
    instructions = instructions & "• Assign macro: ShowDashboardRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'Refresh'" & vbCrLf
    instructions = instructions & "• Assign macro: RefreshDashboardRibbon" & vbCrLf & vbCrLf
    
    instructions = instructions & "Group 2: Workflow" & vbCrLf
    instructions = instructions & "• Add button: 'Process New'" & vbCrLf
    instructions = instructions & "• Assign macro: ProcessNewWorkflowItemsRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'View Status'" & vbCrLf
    instructions = instructions & "• Assign macro: ViewAllRecordsRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'Export'" & vbCrLf
    instructions = instructions & "• Assign macro: ExportDashboardDataRibbon" & vbCrLf & vbCrLf
    
    instructions = instructions & "Group 3: Sandbox" & vbCrLf
    instructions = instructions & "• Add button: 'Initialize Sandbox'" & vbCrLf
    instructions = instructions & "• Assign macro: InitializeSandboxRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'Fix Database'" & vbCrLf
    instructions = instructions & "• Assign macro: FixDatabaseTableRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'Check Database'" & vbCrLf
    instructions = instructions & "• Assign macro: CheckDatabaseStructureRibbon" & vbCrLf & vbCrLf
    
    instructions = instructions & "Group 4: Diagnostics" & vbCrLf
    instructions = instructions & "• Add button: 'Run Diagnostics'" & vbCrLf
    instructions = instructions & "• Assign macro: RunSandboxDiagnosticsRibbon" & vbCrLf
    instructions = instructions & "• Add button: 'View Logs'" & vbCrLf
    instructions = instructions & "• Assign macro: ViewSandboxLogsRibbon" & vbCrLf & vbCrLf
    
    instructions = instructions & "STEP 4: Export Customization" & vbCrLf
    instructions = instructions & "• Click 'Import/Export' button" & vbCrLf
    instructions = instructions & "• Select 'Export all customizations'" & vbCrLf
    instructions = instructions & "• Save as 'BASH_Flow_Ribbon.exportedUI'" & vbCrLf & vbCrLf
    
    instructions = instructions & "STEP 5: Test" & vbCrLf
    instructions = instructions & "• Restart Outlook" & vbCrLf
    instructions = instructions & "• Verify 'BASH Flow' tab appears" & vbCrLf
    instructions = instructions & "• Test each button" & vbCrLf & vbCrLf
    
    instructions = instructions & "Need help? Run ShowMacroList() to see available macros."
    
    MsgBox instructions, vbInformation, "Detailed Ribbon Instructions"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing instructions: " & Err.Description, vbCritical, "Instructions Error"
End Sub

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

' Show list of available macros for ribbon assignment
Public Sub ShowMacroList()
    On Error GoTo ErrorHandler
    
    Dim macroList As String
    macroList = "📋 Available BASH Flow Macros for Ribbon Assignment" & vbCrLf & vbCrLf
    
    macroList = macroList & "Dashboard Functions:" & vbCrLf
    macroList = macroList & "• ShowDashboardRibbon" & vbCrLf
    macroList = macroList & "• RefreshDashboardRibbon" & vbCrLf & vbCrLf
    
    macroList = macroList & "Workflow Functions:" & vbCrLf
    macroList = macroList & "• ProcessNewWorkflowItemsRibbon" & vbCrLf
    macroList = macroList & "• ViewAllRecordsRibbon" & vbCrLf
    macroList = macroList & "• ExportDashboardDataRibbon" & vbCrLf & vbCrLf
    
    macroList = macroList & "Sandbox Functions:" & vbCrLf
    macroList = macroList & "• InitializeSandboxRibbon" & vbCrLf
    macroList = macroList & "• FixDatabaseTableRibbon" & vbCrLf
    macroList = macroList & "• CheckDatabaseStructureRibbon" & vbCrLf & vbCrLf
    
    macroList = macroList & "Diagnostic Functions:" & vbCrLf
    macroList = macroList & "• RunSandboxDiagnosticsRibbon" & vbCrLf
    macroList = macroList & "• ViewSandboxLogsRibbon" & vbCrLf & vbCrLf
    
    macroList = macroList & "Setup Functions:" & vbCrLf
    macroList = macroList & "• SetupBASHFlowModules" & vbCrLf
    macroList = macroList & "• TestBASHFlowModules" & vbCrLf
    macroList = macroList & "• CreateBASHFlowRibbon" & vbCrLf & vbCrLf
    
    macroList = macroList & "Copy these macro names when assigning buttons in ribbon customization."
    
    MsgBox macroList, vbInformation, "Available Macros"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing macro list: " & Err.Description, vbCritical, "Macro List Error"
End Sub

' Test if all required modules are loaded
Public Sub TestRibbonRequirements()
    On Error GoTo ErrorHandler
    
    Dim moduleCheck As String
    Dim missingModules As String
    Dim foundCount As Integer
    Dim requiredModules As String
    Dim comp As Object
    
    requiredModules = "BASH_Flow_Sandbox,WorkflowManager,WorkflowDashboard,DashboardSetup,SandboxDiagnostic,CheckDatabase,RibbonSetup"
    foundCount = 0
    missingModules = ""
    
    ' Check each required module
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If InStr(requiredModules, comp.Name) > 0 Then
            foundCount = foundCount + 1
        End If
    Next comp
    
    ' Check for missing modules
    Dim moduleArray As Variant
    moduleArray = Split(requiredModules, ",")
    
    Dim i As Integer
    For i = 0 To UBound(moduleArray)
        Dim moduleName As String
        moduleName = Trim(moduleArray(i))
        
        Dim found As Boolean
        found = False
        
        For Each comp In Application.VBE.ActiveVBProject.VBComponents
            If comp.Name = moduleName Then
                found = True
                Exit For
            End If
        Next comp
        
        If Not found Then
            If missingModules <> "" Then missingModules = missingModules & ", "
            missingModules = missingModules & moduleName
        End If
    Next i
    
    ' Show results
    If foundCount = 7 And missingModules = "" Then
        moduleCheck = "✅ All required modules are loaded!" & vbCrLf & vbCrLf
        moduleCheck = moduleCheck & "Found modules:" & vbCrLf
        moduleCheck = moduleCheck & "• BASH_Flow_Sandbox" & vbCrLf
        moduleCheck = moduleCheck & "• WorkflowManager" & vbCrLf
        moduleCheck = moduleCheck & "• WorkflowDashboard" & vbCrLf
        moduleCheck = moduleCheck & "• DashboardSetup" & vbCrLf
        moduleCheck = moduleCheck & "• SandboxDiagnostic" & vbCrLf
        moduleCheck = moduleCheck & "• CheckDatabase" & vbCrLf
        moduleCheck = moduleCheck & "• RibbonSetup" & vbCrLf & vbCrLf
        moduleCheck = moduleCheck & "You can now create the ribbon customization."
    Else
        moduleCheck = "⚠️ Some modules are missing!" & vbCrLf & vbCrLf
        moduleCheck = moduleCheck & "Found: " & foundCount & " out of 7 required modules" & vbCrLf
        If missingModules <> "" Then
            moduleCheck = moduleCheck & "Missing: " & missingModules & vbCrLf
        End If
        moduleCheck = moduleCheck & vbCrLf & "Please run SetupBASHFlowModules() first."
    End If
    
    MsgBox moduleCheck, IIf(foundCount = 7, vbInformation, vbExclamation), "Ribbon Requirements Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking requirements: " & Err.Description, vbCritical, "Requirements Check Error"
End Sub

' Quick ribbon setup with macro assignment guide
Public Sub QuickRibbonSetup()
    On Error GoTo ErrorHandler
    
    ' First check if modules are loaded
    TestRibbonRequirements
    
    ' Then show creation instructions
    CreateBASHFlowRibbon
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in quick setup: " & Err.Description, vbCritical, "Quick Setup Error"
End Sub

'===============================================================================
' EXPORT FUNCTIONS
'===============================================================================

' Generate ribbon XML for reference
Public Sub GenerateRibbonXML()
    On Error GoTo ErrorHandler
    
    Dim xmlContent As String
    Dim fso As Object
    Dim file As Object
    
    ' Generate the XML content
    xmlContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xmlContent = xmlContent & "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    xmlContent = xmlContent & "  <ribbon>" & vbCrLf
    xmlContent = xmlContent & "    <tabs>" & vbCrLf
    xmlContent = xmlContent & "      <tab id=""BASHFlowTab"" label=""BASH Flow"">" & vbCrLf
    xmlContent = xmlContent & "        <group id=""DashboardGroup"" label=""Dashboard"">" & vbCrLf
    xmlContent = xmlContent & "          <button id=""ShowDashboard"" label=""Show Dashboard"" size=""large"" imageMso=""DataConnectionGallery"" onAction=""ShowDashboardRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""RefreshDashboard"" label=""Refresh"" size=""normal"" imageMso=""Refresh"" onAction=""RefreshDashboardRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "        </group>" & vbCrLf
    xmlContent = xmlContent & "        <group id=""WorkflowGroup"" label=""Workflow"">" & vbCrLf
    xmlContent = xmlContent & "          <button id=""ProcessNew"" label=""Process New"" size=""large"" imageMso=""MailMergeGoToFirstRecord"" onAction=""ProcessNewWorkflowItemsRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""ViewStatus"" label=""View Status"" size=""normal"" imageMso=""ViewWorkflowTasks"" onAction=""ViewAllRecordsRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""ExportData"" label=""Export"" size=""normal"" imageMso=""ExportExcel"" onAction=""ExportDashboardDataRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "        </group>" & vbCrLf
    xmlContent = xmlContent & "        <group id=""SandboxGroup"" label=""Sandbox"">" & vbCrLf
    xmlContent = xmlContent & "          <button id=""InitializeSandbox"" label=""Initialize Sandbox"" size=""large"" imageMso=""DatabaseOpen"" onAction=""MSCANRibbonSetup.InitializeSandboxRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""FixDatabase"" label=""Fix Database"" size=""normal"" imageMso=""DatabaseRepair"" onAction=""MSCANRibbonSetup.FixDatabaseTableRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""CheckDatabase"" label=""Check Database"" size=""normal"" imageMso=""DatabaseView"" onAction=""MSCANRibbonSetup.CheckDatabaseStructureRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "        </group>" & vbCrLf
    xmlContent = xmlContent & "        <group id=""DiagnosticGroup"" label=""Diagnostics"">" & vbCrLf
    xmlContent = xmlContent & "          <button id=""RunDiagnostics"" label=""Run Diagnostics"" size=""large"" imageMso=""Diagnostics"" onAction=""MSCANRibbonSetup.RunSandboxDiagnosticsRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "          <button id=""ViewLogs"" label=""View Logs"" size=""normal"" imageMso=""ViewLogs"" onAction=""MSCANRibbonSetup.ViewSandboxLogsRibbon"" />" & vbCrLf
    xmlContent = xmlContent & "        </group>" & vbCrLf
    xmlContent = xmlContent & "      </tab>" & vbCrLf
    xmlContent = xmlContent & "    </tabs>" & vbCrLf
    xmlContent = xmlContent & "  </ribbon>" & vbCrLf
    xmlContent = xmlContent & "</customUI>"
    
    ' Save to file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile("C:\Lumen\Workflow Manager\BASH_Flow_Ribbon_Reference.xml", True)
    file.Write xmlContent
    file.Close
    
    MsgBox "✅ Ribbon XML generated successfully!" & vbCrLf & vbCrLf & _
           "File saved as: BASH_Flow_Ribbon_Reference.xml" & vbCrLf & vbCrLf & _
           "This is for reference only. Use the manual setup method to create the actual ribbon.", vbInformation, "XML Generated"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating XML: " & Err.Description, vbCritical, "XML Generation Error"
End Sub
