'===============================================================================
' BASH Flow Management - Front End Question Sheet
' Provides a user-friendly interface for creating workflow entries
'===============================================================================

Option Explicit

'===============================================================================
' Data Structures
'===============================================================================

' State Information
Public Type StateInfo
    StateCode As String
    StateName As String
    DatabasePath As String
End Type

' Stage Information
Public Type StageInfo
    StageName As String
    MilestoneStartDate As Date
    Status As String
End Type

' Workflow Entry
Public Type WorkflowEntry
    State As String
    Stage As String
    MilestoneStartDate As Date
    Status As String
    Description As String
    CreatedDate As Date
End Type

'===============================================================================
' Configuration Constants
'===============================================================================

Private Const FORM_TITLE As String = "BASH Flow Management - Front End Question Sheet"
Private Const FORM_WIDTH As Integer = 600
Private Const FORM_HEIGHT As Integer = 500

'===============================================================================
' Global Variables
'===============================================================================

Private frontEndForm As Object

'===============================================================================
' Main Front End Functions
'===============================================================================

' Show the front end question sheet
Public Sub ShowFrontEndQuestionSheet()
    On Error GoTo ErrorHandler
    
    ' Create and show the front end form
    CreateFrontEndForm
    LoadStateData
    LoadStageData
    frontEndForm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing front end question sheet: " & Err.Description, vbCritical, "Front End Error"
End Sub

' Create the front end form
Private Sub CreateFrontEndForm()
    On Error GoTo ErrorHandler
    
    ' Create user form
    Set frontEndForm = CreateObject("Forms.Form")
    
    ' Set form properties
    With frontEndForm
        .caption = FORM_TITLE
        .width = FORM_WIDTH
        .height = FORM_HEIGHT
        .StartUpPosition = 2 ' Center on screen
        .BackColor = RGB(248, 248, 255)
        .BorderStyle = 3 ' Fixed Dialog
    End With
    
    ' Add controls to form
    AddFrontEndControls
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating front end form: " & Err.Description, vbCritical, "Form Creation Error"
End Sub

' Add controls to the front end form
Private Sub AddFrontEndControls()
    On Error GoTo ErrorHandler
    
    ' Title Label
    AddLabel "lblTitle", "BASH Flow Management - Front End Question Sheet", 20, 20, 560, 30, True, 14
    
    ' State Section
    AddFrame "fraState", "State Selection", 20, 70, 560, 80
    AddLabel "lblState", "Select State:", 40, 100, 100, 20
    AddComboBox "cboState", 150, 100, 200, 25
    
    ' Stage Section
    AddFrame "fraStage", "Stage Selection", 20, 170, 560, 80
    AddLabel "lblStage", "Select Stage:", 40, 200, 100, 20
    AddComboBox "cboStage", 150, 200, 200, 25
    AddLabel "lblMilestoneDate", "Milestone Start Date:", 40, 230, 120, 20
    AddTextBox "txtMilestoneDate", 170, 230, 120, 25
    
    ' Status Section
    AddFrame "fraStatus", "Status Selection", 20, 270, 560, 80
    AddLabel "lblStatus", "Select Status:", 40, 300, 100, 20
    AddComboBox "cboStatus", 150, 300, 200, 25
    
    ' Description Section
    AddFrame "fraDescription", "Description", 20, 370, 560, 60
    AddLabel "lblDescription", "Description:", 40, 390, 100, 20
    AddTextBox "txtDescription", 150, 390, 380, 25
    
    ' Control Buttons
    AddButton "btnSubmit", "Submit Entry", 200, 450, 100, 30
    AddButton "btnClear", "Clear Form", 320, 450, 100, 30
    AddButton "btnCancel", "Cancel", 440, 450, 80, 30
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding front end controls: " & Err.Description, vbCritical, "Control Creation Error"
End Sub

'===============================================================================
' Control Creation Helpers
'===============================================================================

' Add a label control
Private Sub AddLabel(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer, Optional bold As Boolean = False, Optional fontSize As Integer = 9)
    On Error GoTo ErrorHandler
    
    Dim lbl As Object
    Set lbl = frontEndForm.Controls.Add("Forms.Label.1", controlName)
    
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
    Set fra = frontEndForm.Controls.Add("Forms.Frame.1", controlName)
    
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

' Add a combobox control
Private Sub AddComboBox(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim cbo As Object
    Set cbo = frontEndForm.Controls.Add("Forms.ComboBox.1", controlName)
    
    With cbo
        .left = left
        .top = top
        .width = width
        .height = height
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding combobox " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

' Add a textbox control
Private Sub AddTextBox(controlName As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim txt As Object
    Set txt = frontEndForm.Controls.Add("Forms.TextBox.1", controlName)
    
    With txt
        .left = left
        .top = top
        .width = width
        .height = height
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding textbox " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

' Add a button control
Private Sub AddButton(controlName As String, caption As String, left As Integer, top As Integer, width As Integer, height As Integer)
    On Error GoTo ErrorHandler
    
    Dim btn As Object
    Set btn = frontEndForm.Controls.Add("Forms.CommandButton.1", controlName)
    
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
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding button " & controlName & ": " & Err.Description, vbCritical, "Control Error"
End Sub

'===============================================================================
' Data Loading Functions
'===============================================================================

' Load state data into the state dropdown
Private Sub LoadStateData()
    On Error GoTo ErrorHandler
    
    Dim cboState As Object
    Dim states As Variant
    Dim i As Integer
    
    Set cboState = frontEndForm.Controls("cboState")
    
    ' Define US States with their database paths
    states = Array( _
        "AL,Alabama,C:\BASHFlowSandbox\States\Alabama.accdb", _
        "AK,Alaska,C:\BASHFlowSandbox\States\Alaska.accdb", _
        "AZ,Arizona,C:\BASHFlowSandbox\States\Arizona.accdb", _
        "AR,Arkansas,C:\BASHFlowSandbox\States\Arkansas.accdb", _
        "CA,California,C:\BASHFlowSandbox\States\California.accdb", _
        "CO,Colorado,C:\BASHFlowSandbox\States\Colorado.accdb", _
        "CT,Connecticut,C:\BASHFlowSandbox\States\Connecticut.accdb", _
        "DE,Delaware,C:\BASHFlowSandbox\States\Delaware.accdb", _
        "FL,Florida,C:\BASHFlowSandbox\States\Florida.accdb", _
        "GA,Georgia,C:\BASHFlowSandbox\States\Georgia.accdb", _
        "HI,Hawaii,C:\BASHFlowSandbox\States\Hawaii.accdb", _
        "ID,Idaho,C:\BASHFlowSandbox\States\Idaho.accdb", _
        "IL,Illinois,C:\BASHFlowSandbox\States\Illinois.accdb", _
        "IN,Indiana,C:\BASHFlowSandbox\States\Indiana.accdb", _
        "IA,Iowa,C:\BASHFlowSandbox\States\Iowa.accdb", _
        "KS,Kansas,C:\BASHFlowSandbox\States\Kansas.accdb", _
        "KY,Kentucky,C:\BASHFlowSandbox\States\Kentucky.accdb", _
        "LA,Louisiana,C:\BASHFlowSandbox\States\Louisiana.accdb", _
        "ME,Maine,C:\BASHFlowSandbox\States\Maine.accdb", _
        "MD,Maryland,C:\BASHFlowSandbox\States\Maryland.accdb", _
        "MA,Massachusetts,C:\BASHFlowSandbox\States\Massachusetts.accdb", _
        "MI,Michigan,C:\BASHFlowSandbox\States\Michigan.accdb", _
        "MN,Minnesota,C:\BASHFlowSandbox\States\Minnesota.accdb", _
        "MS,Mississippi,C:\BASHFlowSandbox\States\Mississippi.accdb", _
        "MO,Missouri,C:\BASHFlowSandbox\States\Missouri.accdb", _
        "MT,Montana,C:\BASHFlowSandbox\States\Montana.accdb", _
        "NE,Nebraska,C:\BASHFlowSandbox\States\Nebraska.accdb", _
        "NV,Nevada,C:\BASHFlowSandbox\States\Nevada.accdb", _
        "NH,New Hampshire,C:\BASHFlowSandbox\States\New Hampshire.accdb", _
        "NJ,New Jersey,C:\BASHFlowSandbox\States\New Jersey.accdb", _
        "NM,New Mexico,C:\BASHFlowSandbox\States\New Mexico.accdb", _
        "NY,New York,C:\BASHFlowSandbox\States\New York.accdb", _
        "NC,North Carolina,C:\BASHFlowSandbox\States\North Carolina.accdb", _
        "ND,North Dakota,C:\BASHFlowSandbox\States\North Dakota.accdb", _
        "OH,Ohio,C:\BASHFlowSandbox\States\Ohio.accdb", _
        "OK,Oklahoma,C:\BASHFlowSandbox\States\Oklahoma.accdb", _
        "OR,Oregon,C:\BASHFlowSandbox\States\Oregon.accdb", _
        "PA,Pennsylvania,C:\BASHFlowSandbox\States\Pennsylvania.accdb", _
        "RI,Rhode Island,C:\BASHFlowSandbox\States\Rhode Island.accdb", _
        "SC,South Carolina,C:\BASHFlowSandbox\States\South Carolina.accdb", _
        "SD,South Dakota,C:\BASHFlowSandbox\States\South Dakota.accdb", _
        "TN,Tennessee,C:\BASHFlowSandbox\States\Tennessee.accdb", _
        "TX,Texas,C:\BASHFlowSandbox\States\Texas.accdb", _
        "UT,Utah,C:\BASHFlowSandbox\States\Utah.accdb", _
        "VT,Vermont,C:\BASHFlowSandbox\States\Vermont.accdb", _
        "VA,Virginia,C:\BASHFlowSandbox\States\Virginia.accdb", _
        "WA,Washington,C:\BASHFlowSandbox\States\Washington.accdb", _
        "WV,West Virginia,C:\BASHFlowSandbox\States\West Virginia.accdb", _
        "WI,Wisconsin,C:\BASHFlowSandbox\States\Wisconsin.accdb", _
        "WY,Wyoming,C:\BASHFlowSandbox\States\Wyoming.accdb" _
    )
    
    ' Add states to combobox
    For i = LBound(states) To UBound(states)
        cboState.AddItem Split(states(i), ",")(1) ' Add state name
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading state data: " & Err.Description, vbCritical, "Data Loading Error"
End Sub

' Load stage data into the stage dropdown
Private Sub LoadStageData()
    On Error GoTo ErrorHandler
    
    Dim cboStage As Object
    Dim stages As Variant
    Dim i As Integer
    
    Set cboStage = frontEndForm.Controls("cboStage")
    
    ' Define project stages with their milestone start dates
    stages = Array( _
        "Planning,2024-01-01", _
        "Design,2024-01-15", _
        "Development,2024-02-01", _
        "Testing,2024-03-01", _
        "Deployment,2024-04-01", _
        "Maintenance,2024-05-01", _
        "Review,2024-06-01", _
        "Completion,2024-07-01" _
    )
    
    ' Add stages to combobox
    For i = LBound(stages) To UBound(stages)
        cboStage.AddItem Split(stages(i), ",")(0) ' Add stage name
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading stage data: " & Err.Description, vbCritical, "Data Loading Error"
End Sub

'===============================================================================
' Event Handlers
'===============================================================================

' Handle stage selection change
Private Sub OnStageChange()
    On Error GoTo ErrorHandler
    
    Dim cboStage As Object
    Dim txtMilestoneDate As Object
    Dim selectedStage As String
    Dim milestoneDate As Date
    
    Set cboStage = frontEndForm.Controls("cboStage")
    Set txtMilestoneDate = frontEndForm.Controls("txtMilestoneDate")
    
    selectedStage = cboStage.Value
    
    ' Set milestone start date based on selected stage
    Select Case selectedStage
        Case "Planning"
            milestoneDate = DateValue("2024-01-01")
        Case "Design"
            milestoneDate = DateValue("2024-01-15")
        Case "Development"
            milestoneDate = DateValue("2024-02-01")
        Case "Testing"
            milestoneDate = DateValue("2024-03-01")
        Case "Deployment"
            milestoneDate = DateValue("2024-04-01")
        Case "Maintenance"
            milestoneDate = DateValue("2024-05-01")
        Case "Review"
            milestoneDate = DateValue("2024-06-01")
        Case "Completion"
            milestoneDate = DateValue("2024-07-01")
        Case Else
            milestoneDate = Date
    End Select
    
    txtMilestoneDate.Value = Format(milestoneDate, "mm/dd/yyyy")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error handling stage change: " & Err.Description, vbCritical, "Stage Change Error"
End Sub

' Handle form submission
Private Sub SubmitForm()
    On Error GoTo ErrorHandler
    
    Dim cboState As Object
    Dim cboStage As Object
    Dim cboStatus As Object
    Dim txtMilestoneDate As Object
    Dim txtDescription As Object
    Dim workflowEntry As WorkflowEntry
    
    ' Get form controls
    Set cboState = frontEndForm.Controls("cboState")
    Set cboStage = frontEndForm.Controls("cboStage")
    Set cboStatus = frontEndForm.Controls("cboStatus")
    Set txtMilestoneDate = frontEndForm.Controls("txtMilestoneDate")
    Set txtDescription = frontEndForm.Controls("txtDescription")
    
    ' Validate required fields
    If cboState.Value = "" Then
        MsgBox "Please select a state.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If cboStage.Value = "" Then
        MsgBox "Please select a stage.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    If cboStatus.Value = "" Then
        MsgBox "Please select a status.", vbExclamation, "Validation Error"
        Exit Sub
    End If
    
    ' Create workflow entry
    workflowEntry.State = cboState.Value
    workflowEntry.Stage = cboStage.Value
    workflowEntry.MilestoneStartDate = DateValue(txtMilestoneDate.Value)
    workflowEntry.Status = cboStatus.Value
    workflowEntry.Description = txtDescription.Value
    workflowEntry.CreatedDate = Now()
    
    ' Save to database
    If SaveWorkflowEntry(workflowEntry) Then
        MsgBox "Workflow entry saved successfully!", vbInformation, "Success"
        ClearForm
    Else
        MsgBox "Error saving workflow entry.", vbExclamation, "Save Error"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error submitting form: " & Err.Description, vbCritical, "Submission Error"
End Sub

' Clear the form
Private Sub ClearForm()
    On Error GoTo ErrorHandler
    
    Dim cboState As Object
    Dim cboStage As Object
    Dim cboStatus As Object
    Dim txtMilestoneDate As Object
    Dim txtDescription As Object
    
    Set cboState = frontEndForm.Controls("cboState")
    Set cboStage = frontEndForm.Controls("cboStage")
    Set cboStatus = frontEndForm.Controls("cboStatus")
    Set txtMilestoneDate = frontEndForm.Controls("txtMilestoneDate")
    Set txtDescription = frontEndForm.Controls("txtDescription")
    
    cboState.Value = ""
    cboStage.Value = ""
    cboStatus.Value = ""
    txtMilestoneDate.Value = ""
    txtDescription.Value = ""
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing form: " & Err.Description, vbCritical, "Clear Error"
End Sub

'===============================================================================
' Database Functions
'===============================================================================

' Save workflow entry to database
Private Function SaveWorkflowEntry(entry As WorkflowEntry) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    sql = "INSERT INTO BASHFlowSandbox ([State], [Type], [Date], [Description], [Status], [Date Created]) VALUES (" & _
          "'" & EscapeString(entry.State) & "', " & _
          "'" & EscapeString(entry.Stage) & "', " & _
          "'" & Format(entry.MilestoneStartDate, "yyyy-mm-dd") & "', " & _
          "'" & EscapeString(entry.Description) & "', " & _
          "'" & EscapeString(entry.Status) & "', " & _
          "'" & Format(entry.CreatedDate, "yyyy-mm-dd hh:nn:ss") & "')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    SaveWorkflowEntry = True
    
    Exit Function
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    SaveWorkflowEntry = False
End Function

' Escape SQL strings to prevent injection
Private Function EscapeString(inputStr As String) As String
    EscapeString = Replace(inputStr, "'", "''")
End Function

'===============================================================================
' Utility Functions
'===============================================================================

' Close the front end form
Public Sub CloseFrontEndForm()
    On Error GoTo ErrorHandler
    
    If Not frontEndForm Is Nothing Then
        frontEndForm.Hide
        Set frontEndForm = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error closing front end form: " & Err.Description, vbCritical, "Close Error"
End Sub

'===============================================================================
' Ribbon Integration
'===============================================================================

' Front end ribbon callback
Public Sub ShowFrontEndRibbon(control As IRibbonControl)
    ShowFrontEndQuestionSheet
End Sub
