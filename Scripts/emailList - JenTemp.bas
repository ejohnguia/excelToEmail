Attribute VB_Name = "emailList"
''''''''''''''''''''''''''''''''''''''''''''''
' Author: Ezra John Guia
' www.linkedin.com/in/ezrajohn-guia
'
' Pre-requisites:
'   - All files must be in the same folder
'   - All files must comply to the same format
''''''''''''''''''''''''''''''''''''''''''''''
Dim schedStr, cheatStr As String

Sub RunEmailList()
' *** RUN THIS ***
' Sends Emails with HTML Tables

    ' Setting global variables
    schedStr = ActiveWorkbook.Name
    cheatStr = "Scheduling Cheat Sheet.xlsm"

    ' Main Execution
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Call OpenFiles
    Call EmailSend
    Call CloseFiles

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub OpenFiles()
' Open files to be used

    'Setting Global Variables
    Workbooks.Open Filename:=ThisWorkbook.Path & "\" & cheatStr
    
End Sub

Sub CloseFiles()
' Files opened will be closed

    'Cheat Sheet will be closed
    Windows(cheatStr).Activate
    ActiveWorkbook.Close savechanges:=False

End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
' https://www.rondebruin.nl/win/s1/outlook/bmail2.htm
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function findEmails(personName) As String
' Find emails that belong to the person given

    Windows("Scheduling Cheat Sheet.xlsm").Activate
    Worksheets("Floater Contact List").Activate

    ' Sets the range of employees
    Dim rng As Range: Set rng = Range(Range("B2"), Range("B2").End(xlDown))

    Dim row As Integer
    For row = 1 to rng.Rows.Count
        
        ' If First Name + Last Name = Person
        If (Trim(rng.Cells(RowIndex := row, ColumnIndex := 2).Value) & _
            " " & Trim(rng.Cells(RowIndex := row, ColumnIndex := 1).Value) = personName) Then

            ' findEmails = Personal Email + Sobeys
            ' Row offset for dif. Range
            findEmails = Trim(Cells(row + 1, 7).Value) & "; " & _
                Trim(Cells(row + 1, 8).Value)

            Exit Function

        End If

    Next row

End Function

Function ccEmails(floaterRng As Range) As String
' Creates the emails to be sent by store numbers

    ' Creates an empty string to be appended by
    ccEmails = ""
    Dim store As Range

    ' Creates a dictionary
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")

    For Each store in floaterRng

        v = store.Offset(, -1).Value

        ' Ensures that the row is not hidden by a filter and not a duplicate
        If (store.EntireRow.Hidden = False) and Not IsEmpty(v) _
            and (v <> "Store") and Not dict.Exists(v) Then

            dict(v) = ""
            ccEmails = ccEmails & "RX" & v & "@sobeys.com; "
        
        End If

    Next
    
End Function

Sub EmailSend()
' Creates the email with fields filled in

    ' Setup variables
    Dim schedBotRow As Integer
    Dim schedTable As Range
    Dim floaterRng As Range
    Dim floatTable As Range
    Dim preTableStr, weekName As String

    ' Opens Outlook
    Dim OutApp, OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")

    ' Setup dictionary of floaters
    Dim floatDict: Set floatDict = CreateObject("Scripting.Dictionary")
    Dim floatArr() As Variant
    Dim flt As Variant

    ' Activate scheduling sheet
    Windows(schedStr).Activate
    weekName = ActiveSheet.Name

    ' Clear filters and convert range of floaters to array
    Range("A3:K3").AutoFilter Field:=1
    Range("A3:K3").AutoFilter Field:=2
    Set floatTable = Range(Range("B4"), Range("B4").End(xlDown))
    floatArr = floatTable

    ' Loop through all the floaters
    For Each flt in floatArr

        If Not floatDict.Exists(flt) Then
            
            ' Adds floater to the dictionary to avoid duplicates
            floatDict(flt) = ""

            ' Create Outlook mail
            Set OutMail = OutApp.CreateItem(0)
            
            ' Filters the Current Floater
            Range("A3:K3").AutoFilter Field:=2, Criteria1:=flt

            ' Gets range of floaters and bottom row number
            Set floaterRng = Range(Range("B3"), Range("B3").End(xlDown))
            schedBotRow = floaterRng.row + floaterRng.Rows.Count - 1
            
            ' Set the size/location of the table to be sent
            Set schedTable = Worksheets(weekName).Range(Cells(1, 1), Cells(schedBotRow, 11))

            ' HTML Style
            preTableStr = "<BODY style = font-size:11pt;font-family:Calibri (Body)>" & _
                "Hello, <br><br>Below is your " & weekName & " schedule.<br>"

            ' Fill in email fields
            With OutMail
                .to = findEmails(flt)
                .CC = ccEmails(floaterRng)
                .Subject = weekName & " Schedule"
                .Display
                .HTMLBody = preTableStr & RangetoHTML(schedTable) & .HTMLBody
            End With
            
            ' Return to the Excel Sheet
            Windows(schedStr).Activate
            
            ' Resets memory
            Set OutMail = Nothing

        End if
    
    Next

    ' Clear filters
    Range("A3:K3").AutoFilter Field:=2

    Set OutApp = Nothing

End Sub