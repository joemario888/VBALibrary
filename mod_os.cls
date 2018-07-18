Function fileExist(fileName)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(fileName)) Then
       fileExist = True
    Else
       fileExist = False
    End If
End Function

Function getLine(fname, lineNum)
    Dim pbjFSO
    Dim objFile
    Dim i
    
    If fileExist(fname) Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.GetFile(fname)
        If objFile.Size = 0 Then
            getLine = ""
            Exit Function
        End If

        Set objFile = objFSO.OpenTextFile(fname, 1)
        For i = 1 To lineNum
            getLine = objFile.ReadLine()
        Next
    Else
    getLine = "No"
    End If
End Function


Function selectFileBrowser() As Workbook
    Dim txtFileName As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
      .AllowMultiSelect = False
      .Title = "Please select the file."
      .Filters.Clear
      .Filters.Add "Excel 2003", "*.xls"
      .Filters.Add "All Files", "*.*"
    
      If .Show = True Then
        txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
      End If
    End With
    
    Set selectFileBrowser = Workbooks.Open(txtFileName)
End Function


Dim PreviousCalculationSetting As Variant
Sub StartEnviroment()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .CalculationInterruptKey = xlEscKey
        PreviousCalculationSetting = .Calculation
        .Calculation = xlCalculationManual
        .Calculate
    End With
End Sub
Sub EndEnviroment()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = False
        .CalculationInterruptKey = xlAnyKey
        .Calculation = PreviousCalculationSetting
        .Calculate
    End With
End Sub

