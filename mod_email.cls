Function sendEMailWithAtt(rcpt, sbjt, strBody, attachments)
   Dim iMsg
   Dim iConf
   Dim Flds
    Dim i As Long
    Dim errMsg
    Dim SDrive
    
    SDrive = "S:"

   On Error Resume Next
   Set iMsg = CreateObject("CDO.Message")
   Set iConf = CreateObject("CDO.Configuration")
   iConf.Load -1    ' CDO Source Defaults
   Set Flds = iConf.Fields
   user = getLine(SDrive & "\tools\VB Script\EmailAccount.txt", 1)
   pswd = getLine(SDrive & "\tools\VB Script\EmailAccount.txt", 2)
   With Flds
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = user
      .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pswd
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
      .Update
   End With
   With iMsg
      Set .Configuration = iConf
      .To = rcpt
      .CC = ""
      .BCC = ""
      .from = user
      .Subject = sbjt
      .HTMLBody = strBody
      For i = LBound(attachments) To UBound(attachments)
      
      .addattachment (attachments(i))
      Next i
      
      .Send
   End With
   If Err.Description <> "" Then
      sendEMailWithAtt = ecpt & Chr(10) & sbjt & Chr(10) & strBody & Chr(10) & Err.Description
   End If
End Function

Function sendEMail(rcpt, sbjt, strBody)
   Dim iMsg
   Dim iConf
   Dim Flds
 
    Dim errMsg
    Dim SDrive
    
    SDrive = "S:"

   On Error Resume Next
   Set iMsg = CreateObject("CDO.Message")
   Set iConf = CreateObject("CDO.Configuration")
   iConf.Load -1    ' CDO Source Defaults
   Set Flds = iConf.Fields
   user = getLine(SDrive & "\tools\VB Script\EmailAccount.txt", 1)
   pswd = getLine(SDrive & "\tools\VB Script\EmailAccount.txt", 2)
   With Flds
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = user
      .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pswd
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
      .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
      .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
      .Update
   End With
   With iMsg
      Set .Configuration = iConf
      .To = rcpt
      .CC = ""
      .BCC = ""
      .from = user
      .Subject = sbjt
      .TextBody = strBody
      .Send
   End With
   If Err.Description <> "" Then
      sendEMail = ecpt & Chr(10) & sbjt & Chr(10) & strBody & Chr(10) & Err.Description
   End If
End Function

Sub Mail_Selection_Range_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "ron@debruin.nl"
        .CC = ""
        .BCC = ""
        .Subject = "This is the Subject line"
        .HTMLBody = RangetoHTML(rng)
        .Send   'or use .Display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
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
         fileName:=TempFile, _
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




Function RangetoHTMLFile(rng As Range, path As String)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = path

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
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
   RangetoHTMLFile = ts.readall
    ts.Close
    RangetoHTMLFile = Replace(RangetoHTMLFile, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
