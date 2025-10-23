'===============================================================================
' BASH Flow Management - Question Sheet
' Simple front-end interface with State dropdown and Stage dropdown
'===============================================================================

Option Explicit

'===============================================================================
' Data Structures
'===============================================================================

Public Type QuestionSheetEntry
    State As String
    Stage As String
    MilestoneStartDate As Date
    Status As String
    CreatedDate As Date
End Type

'===============================================================================
' Main Functions
'===============================================================================

' Show the question sheet
Public Sub ShowQuestionSheet()
    On Error GoTo ErrorHandler
    
    ShowQuestionInterface
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing question sheet: " & Err.Description, vbCritical, "Question Sheet Error"
End Sub

' Show question interface using input boxes
Private Sub ShowQuestionInterface()
    On Error GoTo ErrorHandler
    
    Dim state As String
    Dim stage As String
    Dim status As String
    Dim entry As QuestionSheetEntry
    Dim continueEntry As Integer
    
    Do
        ' Get State selection
        state = GetStateSelection()
        If state = "" Then Exit Sub
        
        ' Get Stage selection
        stage = GetStageSelection()
        If stage = "" Then Exit Sub
        
        ' Get Status selection
        status = GetStatusSelection()
        If status = "" Then Exit Sub
        
        ' Create entry
        entry.State = state
        entry.Stage = stage
        entry.MilestoneStartDate = GetMilestoneDate(stage)
        entry.Status = status
        entry.CreatedDate = Now()
        
        ' Save entry
        If SaveQuestionEntry(entry) Then
            Dim successMsg As String
            successMsg = "Question sheet entry saved successfully!"
            successMsg = successMsg & vbCrLf & vbCrLf
            successMsg = successMsg & "State: " & entry.State & vbCrLf
            successMsg = successMsg & "Stage: " & entry.Stage & vbCrLf
            successMsg = successMsg & "Milestone Start Date: " & Format(entry.MilestoneStartDate, "mm/dd/yyyy") & vbCrLf
            successMsg = successMsg & "Status: " & entry.Status
            MsgBox successMsg, vbInformation, "Success"
        Else
            MsgBox "Error saving question sheet entry.", vbExclamation, "Save Error"
        End If
        
        ' Ask if user wants to continue
        continueEntry = MsgBox("Would you like to create another entry?", vbYesNo + vbQuestion, "Continue?")
        
    Loop While continueEntry = vbYes
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in question interface: " & Err.Description, vbCritical, "Interface Error"
End Sub

'===============================================================================
' Input Helper Functions
'===============================================================================

' Get state selection from user
Private Function GetStateSelection() As String
    On Error GoTo ErrorHandler
    
    Dim stateList As String
    Dim selectedState As String
    
    ' Create state list for user
    stateList = "Please select a state:" & vbCrLf & vbCrLf
    stateList = stateList & "1. Alabama" & vbCrLf & "2. Alaska" & vbCrLf & "3. Arizona" & vbCrLf
    stateList = stateList & "4. Arkansas" & vbCrLf & "5. California" & vbCrLf & "6. Colorado" & vbCrLf
    stateList = stateList & "7. Connecticut" & vbCrLf & "8. Delaware" & vbCrLf & "9. Florida" & vbCrLf
    stateList = stateList & "10. Georgia" & vbCrLf & "11. Hawaii" & vbCrLf & "12. Idaho" & vbCrLf
    stateList = stateList & "13. Illinois" & vbCrLf & "14. Indiana" & vbCrLf & "15. Iowa" & vbCrLf
    stateList = stateList & "16. Kansas" & vbCrLf & "17. Kentucky" & vbCrLf & "18. Louisiana" & vbCrLf
    stateList = stateList & "19. Maine" & vbCrLf & "20. Maryland" & vbCrLf & "21. Massachusetts" & vbCrLf
    stateList = stateList & "22. Michigan" & vbCrLf & "23. Minnesota" & vbCrLf & "24. Mississippi" & vbCrLf
    stateList = stateList & "25. Missouri" & vbCrLf & "26. Montana" & vbCrLf & "27. Nebraska" & vbCrLf
    stateList = stateList & "28. Nevada" & vbCrLf & "29. New Hampshire" & vbCrLf & "30. New Jersey" & vbCrLf
    stateList = stateList & "31. New Mexico" & vbCrLf & "32. New York" & vbCrLf & "33. North Carolina" & vbCrLf
    stateList = stateList & "34. North Dakota" & vbCrLf & "35. Ohio" & vbCrLf & "36. Oklahoma" & vbCrLf
    stateList = stateList & "37. Oregon" & vbCrLf & "38. Pennsylvania" & vbCrLf & "39. Rhode Island" & vbCrLf
    stateList = stateList & "40. South Carolina" & vbCrLf & "41. South Dakota" & vbCrLf & "42. Tennessee" & vbCrLf
    stateList = stateList & "43. Texas" & vbCrLf & "44. Utah" & vbCrLf & "45. Vermont" & vbCrLf
    stateList = stateList & "46. Virginia" & vbCrLf & "47. Washington" & vbCrLf & "48. West Virginia" & vbCrLf
    stateList = stateList & "49. Wisconsin" & vbCrLf & "50. Wyoming" & vbCrLf & vbCrLf
    stateList = stateList & "Enter the number (1-50) or state name:"
    
    selectedState = InputBox(stateList, "State Selection", "")
    
    If selectedState = "" Then
        GetStateSelection = ""
        Exit Function
    End If
    
    ' Convert number to state name if needed
    GetStateSelection = ConvertToStateName(selectedState)
    
    Exit Function
    
ErrorHandler:
    GetStateSelection = ""
End Function

' Convert input to state name
Private Function ConvertToStateName(inputValue As String) As String
    On Error GoTo ErrorHandler
    
    Dim stateNames(49) As String
    Dim stateNumber As Integer
    Dim i As Integer
    
    ' Initialize state names array
    stateNames(0) = "Alabama"
    stateNames(1) = "Alaska"
    stateNames(2) = "Arizona"
    stateNames(3) = "Arkansas"
    stateNames(4) = "California"
    stateNames(5) = "Colorado"
    stateNames(6) = "Connecticut"
    stateNames(7) = "Delaware"
    stateNames(8) = "Florida"
    stateNames(9) = "Georgia"
    stateNames(10) = "Hawaii"
    stateNames(11) = "Idaho"
    stateNames(12) = "Illinois"
    stateNames(13) = "Indiana"
    stateNames(14) = "Iowa"
    stateNames(15) = "Kansas"
    stateNames(16) = "Kentucky"
    stateNames(17) = "Louisiana"
    stateNames(18) = "Maine"
    stateNames(19) = "Maryland"
    stateNames(20) = "Massachusetts"
    stateNames(21) = "Michigan"
    stateNames(22) = "Minnesota"
    stateNames(23) = "Mississippi"
    stateNames(24) = "Missouri"
    stateNames(25) = "Montana"
    stateNames(26) = "Nebraska"
    stateNames(27) = "Nevada"
    stateNames(28) = "New Hampshire"
    stateNames(29) = "New Jersey"
    stateNames(30) = "New Mexico"
    stateNames(31) = "New York"
    stateNames(32) = "North Carolina"
    stateNames(33) = "North Dakota"
    stateNames(34) = "Ohio"
    stateNames(35) = "Oklahoma"
    stateNames(36) = "Oregon"
    stateNames(37) = "Pennsylvania"
    stateNames(38) = "Rhode Island"
    stateNames(39) = "South Carolina"
    stateNames(40) = "South Dakota"
    stateNames(41) = "Tennessee"
    stateNames(42) = "Texas"
    stateNames(43) = "Utah"
    stateNames(44) = "Vermont"
    stateNames(45) = "Virginia"
    stateNames(46) = "Washington"
    stateNames(47) = "West Virginia"
    stateNames(48) = "Wisconsin"
    stateNames(49) = "Wyoming"
    
    ' Check if input is a number
    If IsNumeric(inputValue) Then
        stateNumber = CInt(inputValue)
        If stateNumber >= 1 And stateNumber <= 50 Then
            ConvertToStateName = stateNames(stateNumber - 1)
            Exit Function
        End If
    End If
    
    ' Check if input matches a state name
    For i = 0 To 49
        If UCase(Trim(inputValue)) = UCase(stateNames(i)) Then
            ConvertToStateName = stateNames(i)
            Exit Function
        End If
    Next i
    
    ' If no match found, return the input as-is
    ConvertToStateName = inputValue
    
    Exit Function
    
ErrorHandler:
    ConvertToStateName = inputValue
End Function

' Get stage selection from user
Private Function GetStageSelection() As String
    On Error GoTo ErrorHandler
    
    Dim stageList As String
    Dim selectedStage As String
    
    stageList = "Please select a stage:" & vbCrLf & vbCrLf
    stageList = stageList & "1. Planning" & vbCrLf & "2. Design" & vbCrLf
    stageList = stageList & "3. Development" & vbCrLf & "4. Testing" & vbCrLf
    stageList = stageList & "5. Deployment" & vbCrLf & "6. Maintenance" & vbCrLf
    stageList = stageList & "7. Review" & vbCrLf & "8. Completion" & vbCrLf & vbCrLf
    stageList = stageList & "Enter the number (1-8) or stage name:"
    
    selectedStage = InputBox(stageList, "Stage Selection", "")
    
    If selectedStage = "" Then
        GetStageSelection = ""
        Exit Function
    End If
    
    ' Convert number to stage name if needed
    GetStageSelection = ConvertToStageName(selectedStage)
    
    Exit Function
    
ErrorHandler:
    GetStageSelection = ""
End Function

' Convert input to stage name
Private Function ConvertToStageName(inputValue As String) As String
    On Error GoTo ErrorHandler
    
    Dim stageNames(7) As String
    Dim stageNumber As Integer
    Dim i As Integer
    
    ' Initialize stage names array
    stageNames(0) = "Planning"
    stageNames(1) = "Design"
    stageNames(2) = "Development"
    stageNames(3) = "Testing"
    stageNames(4) = "Deployment"
    stageNames(5) = "Maintenance"
    stageNames(6) = "Review"
    stageNames(7) = "Completion"
    
    ' Check if input is a number
    If IsNumeric(inputValue) Then
        stageNumber = CInt(inputValue)
        If stageNumber >= 1 And stageNumber <= 8 Then
            ConvertToStageName = stageNames(stageNumber - 1)
            Exit Function
        End If
    End If
    
    ' Check if input matches a stage name
    For i = 0 To 7
        If UCase(Trim(inputValue)) = UCase(stageNames(i)) Then
            ConvertToStageName = stageNames(i)
            Exit Function
        End If
    Next i
    
    ' If no match found, return the input as-is
    ConvertToStageName = inputValue
    
    Exit Function
    
ErrorHandler:
    ConvertToStageName = inputValue
End Function

' Get status selection from user
Private Function GetStatusSelection() As String
    On Error GoTo ErrorHandler
    
    Dim statusList As String
    Dim selectedStatus As String
    
    statusList = "Please select a status:" & vbCrLf & vbCrLf
    statusList = statusList & "1. To Do" & vbCrLf & "2. In Progress" & vbCrLf
    statusList = statusList & "3. Blocked" & vbCrLf & "4. Done" & vbCrLf & vbCrLf
    statusList = statusList & "Enter the number (1-4) or status name:"
    
    selectedStatus = InputBox(statusList, "Status Selection", "")
    
    If selectedStatus = "" Then
        GetStatusSelection = ""
        Exit Function
    End If
    
    ' Convert number to status name if needed
    GetStatusSelection = ConvertToStatusName(selectedStatus)
    
    Exit Function
    
ErrorHandler:
    GetStatusSelection = ""
End Function

' Convert input to status name
Private Function ConvertToStatusName(inputValue As String) As String
    On Error GoTo ErrorHandler
    
    Dim statusNames(3) As String
    Dim statusNumber As Integer
    Dim i As Integer
    
    ' Initialize status names array
    statusNames(0) = "To Do"
    statusNames(1) = "In Progress"
    statusNames(2) = "Blocked"
    statusNames(3) = "Done"
    
    ' Check if input is a number
    If IsNumeric(inputValue) Then
        statusNumber = CInt(inputValue)
        If statusNumber >= 1 And statusNumber <= 4 Then
            ConvertToStatusName = statusNames(statusNumber - 1)
            Exit Function
        End If
    End If
    
    ' Check if input matches a status name
    For i = 0 To 3
        If UCase(Trim(inputValue)) = UCase(statusNames(i)) Then
            ConvertToStatusName = statusNames(i)
            Exit Function
        End If
    Next i
    
    ' If no match found, return the input as-is
    ConvertToStatusName = inputValue
    
    Exit Function
    
ErrorHandler:
    ConvertToStatusName = inputValue
End Function

' Get milestone date based on stage
Private Function GetMilestoneDate(stage As String) As Date
    On Error GoTo ErrorHandler
    
    Select Case UCase(Trim(stage))
        Case "PLANNING"
            GetMilestoneDate = DateValue("2024-01-01")
        Case "DESIGN"
            GetMilestoneDate = DateValue("2024-01-15")
        Case "DEVELOPMENT"
            GetMilestoneDate = DateValue("2024-02-01")
        Case "TESTING"
            GetMilestoneDate = DateValue("2024-03-01")
        Case "DEPLOYMENT"
            GetMilestoneDate = DateValue("2024-04-01")
        Case "MAINTENANCE"
            GetMilestoneDate = DateValue("2024-05-01")
        Case "REVIEW"
            GetMilestoneDate = DateValue("2024-06-01")
        Case "COMPLETION"
            GetMilestoneDate = DateValue("2024-07-01")
        Case Else
            GetMilestoneDate = Date
    End Select
    
    Exit Function
    
ErrorHandler:
    GetMilestoneDate = Date
End Function

'===============================================================================
' Database Functions
'===============================================================================

' Save question entry to database
Private Function SaveQuestionEntry(entry As QuestionSheetEntry) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BASHFlowSandbox\TestDatabase.accdb;"
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    
    sql = "INSERT INTO BASHFlowSandbox ([State], [Type], [Date], [Description], [Status], [Date Created]) VALUES ("
    sql = sql & "'" & EscapeString(entry.State) & "', "
    sql = sql & "'" & EscapeString(entry.Stage) & "', "
    sql = sql & "'" & Format(entry.MilestoneStartDate, "yyyy-mm-dd") & "', "
    sql = sql & "'Question Sheet Entry', "
    sql = sql & "'" & EscapeString(entry.Status) & "', "
    sql = sql & "'" & Format(entry.CreatedDate, "yyyy-mm-dd hh:nn:ss") & "')"
    
    cmd.CommandText = sql
    cmd.Execute
    
    conn.Close
    
    SaveQuestionEntry = True
    
    Exit Function
    
ErrorHandler:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    SaveQuestionEntry = False
End Function

' Escape SQL strings to prevent injection
Private Function EscapeString(inputStr As String) As String
    EscapeString = Replace(inputStr, "'", "''")
End Function

'===============================================================================
' Ribbon Integration
'===============================================================================

' Question sheet ribbon callback
Public Sub ShowQuestionSheetRibbon(control As IRibbonControl)
    ShowQuestionSheet
End Sub