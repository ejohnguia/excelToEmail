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
    Call CloseFile

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub OpenFiles()
' Open files to be used

    'Setting Global Variables
    Workbooks.Open Filename:=ThisWorkbook.Path & "\" & cheatStr
    
End Sub

Sub CloseFile()
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

    ' Selects the range of employees
    Dim rng As Range
    Range(Range("B2"), Range("B2").End(xlDown)).Select
    Set rng = Selection

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

Function ccEmails(storeNs As Range) As String
' Creates the emails to be sent by store numbers

    ' Creates an array from the range 
    Dim Emails() As Variant
    Emails = storeNs
    ccEmails = ""

    'Create a dictionary only containing unique keys
    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    Dim val As Variant

    ' Gets unique keys from a list
    For Each val in Emails

        If Not dict.Exists(val) and Not IsEmpty(val) Then

            ' Add the store number to the dictionary and email string
            dict(val) = "cheese monkey"
            ccEmails = ccEmails & "pharmmgr" & val & "@sobeys.com; " _
                & "technician" & val & "@sobeys.com; "
        
        End If

    Next

    Set dict = Nothing
    
End Function

Sub EmailSend()
' Creates the email with fields filled in

    Dim OutApp, OutMail As Object

    Dim schedTable As Range
    Dim preTableStr, weekName As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Set the size/location of the table
    Windows(schedStr).Activate
    weekName = ActiveSheet.Name
    Set schedTable = Worksheets(weekName).Range(ActiveCell.Offset(-1), _
        ActiveCell.Offset(7, 1))

    ' HTML Style
    preTableStr = "<BODY style = font-size:11pt;font-family:Calibri (Body)>" & _
        "Hello, <br><br>Below is your " & weekName & " schedule.<br>"

    ' Fill in email fields
    With OutMail
        .to = findEmails(ActiveCell.Value)
        .CC = ccEmails(Range(schedTable.Cells(RowIndex := 3, ColumnIndex := 4), _
            schedTable.Cells(RowIndex := 9, ColumnIndex := 4)))
        .Subject = weekName & " Schedule"
        .Display
        .HTMLBody = preTableStr & RangetoHTML(schedTable) & .HTMLBody
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub