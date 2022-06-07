Attribute VB_Name = "copySchedules"
''''''''''''''''''''''''''''''''''''''''''''''
' Author: Ezra John Guia
' www.linkedin.com/in/ezrajohn-guia
'
' Pre-requisites:
'   - All files must be in the same folder
'   - All files must comply to the same format
''''''''''''''''''''''''''''''''''''''''''''''

' Global Variables to be Used
Dim myFile As String

Sub RunCopySchedule()
' Run this macro

    ' Dim myFile As String
    myFile = InputBox("What is your file name? (Please include the extension!)", _
        "Schedule Transfer", _
        "Enter your file name HERE.")
    
    ' Main Execution
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Call OpenFiles
    Call CopySched
    Call CloseFile

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox ("File: " & myFile & " processed.")

End Sub

Sub OpenFiles()
' Open files to be used

    'Schedule given will be opened
    Workbooks.Open Filename:=ThisWorkbook.Path & "\" & myFile
    
    'Cheat sheet given will be opened
    Workbooks.Open Filename:=ThisWorkbook.Path & "\" & "Scheduling Cheat Sheet.xlsm"
    
End Sub

Sub CloseFile()
' Files opened will be closed

    'Schedule given will be closed
    Windows(myFile).Activate
    ActiveWorkbook.Close SaveChanges := False

    'Cheat sheet will be closed
    Windows("Scheduling Cheat Sheet.xlsm").Activate
    ActiveWorkbook.Close SaveChanges := False

End Sub

Sub SelectEmpty(StoreNum, NorthOrSouth)
' Checks the store sum cell in the template if empty/duplicate and selects

    ' Stepping through North or South stores
    For row = NorthOrSouth(0) To NorthOrSouth(1) Step 8

        ' Stepping through columns
        For col = 2 To 22 Step 4

            ' Checks if it's empty
            If trim(Cells(row, col).Value & vbnullstring) = vbnullstring Then
            ' https://stackoverflow.com/questions/14108948/excel-vba-check-if-entry-is-empty-or-not-space

                ' Paste the Store Number
                ActiveSheet.Cells(row, col).Value = StoreNum

                ' Once empty box is found, select the cell exit the sub
                ActiveSheet.Cells(row + 1, col + 1).Select
                Exit Sub

            End If

        Next col

    Next row

End Sub

Function Extract_Number_from_Text(Phrase) As Double
' Grabs the number from the string
' Returns 0 if empty
' https://www.automateexcel.com/vba/extract-number-from-string/
    Dim Length_of_String As Integer
    Dim Current_Pos As Integer
    Dim Temp As String
    Length_of_String = Len(Phrase)
    Temp = ""
        
        For Current_Pos = 1 To Length_of_String
        
            If (Mid(Phrase, Current_Pos, 1) = "-") Then
                Temp = Temp & Mid(Phrase, Current_Pos, 1)
            End If
    
            If (Mid(Phrase, Current_Pos, 1) = ".") Then
                Temp = Temp & Mid(Phrase, Current_Pos, 1)
            End If
    
            If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
                Temp = Temp & Mid(Phrase, Current_Pos, 1)
            End If
        
        Next Current_Pos
    
        If Len(Temp) = 0 Then
            Extract_Number_from_Text = 0
        Else
            Extract_Number_from_Text = CDbl(Temp)
        End If

End Function

Function IsNorth(StoreNum)
' Checks in the cheat sheet if the store is located north
' Returns the Array of the North or South

    ' Activates the Cheat Sheet Window
    Windows("Scheduling Cheat Sheet.xlsm").Activate
    Worksheets("Corporate Store Listing").Activate

    ' Sets the Array location of North or South stores
    If Range("A1:A100").Find(StoreNum, lookat:=xlPart).Offset(, 10).Value = "N" Then
        IsNorth = Array(2, 34)
    Else
        IsNorth = Array(43, 67)
    End If
     
End Function

Sub CopySched()
' Copy and Pastes the Selected Schedule

    ' Activate given schedule
    Windows(myFile).Activate
    Worksheets(1).Activate

    ' Creates an array of Weeks Selected of 3 rows and 2 columns
    Dim weekLocations(2,1) as Variant
    weekLocations(0, 0) = "B1"
    weekLocations(0, 1) = "E2:E8"

    weekLocations(1, 0) = "B11"
    weekLocations(1, 1) = "E12:E18"

    weekLocations(2, 0) = "B21"
    weekLocations(2, 1) = "E22:E28"

    ' Creates a Variable to save the Week, Store number and Location
    Dim thisWeek As String
    Dim thisStoreNum As String
    thisStoreNum = Extract_Number_from_Text(Range("A1").Value)
    
    ' Returns range of North or South store
    Dim NorthOrSouth
    NorthOrSouth = IsNorth(thisStoreNum)

    ' 3 Loops for POTENTIAL given schedules
    For week = 0 to 2
    
        ' Activate given Schedule
        Windows(myFile).Activate
        Sheets(1).Select
        
        ' Checks if Week is N/A
        thisWeek = Extract_Number_from_Text(Range(weekLocations(week, 0)).Value)
        if thisWeek = 0 Then GoTo NextIteration
        thisWeek = "Week " & thisWeek

        ' Checks if Schedule is N/A
        if WorksheetFunction.CountA(Range(weekLocations(week, 1))) = 0 Then
            GoTo NextIteration

        ' Copying the schedule
        Range(weekLocations(week, 1)).Select
        Selection.Copy

        ' Checks that week's schedule and pastes
        Windows("F21 Safeway Schedules.xlsm").Activate
        Sheets(thisWeek).Select
        Call SelectEmpty(thisStoreNum, NorthOrSouth)
        ActiveSheet.Paste

        NextIteration:

    Next

End Sub