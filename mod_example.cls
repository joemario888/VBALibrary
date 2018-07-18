
Function RunChecks() As Boolean
Dim wb As Workbook
Set wb = ActiveWorkbook

If CheckHeadingMapping(wb) And checkFileNotRunYet(wb) And checkDataMatchesMapping(wb) Then
RunChecks = True
Else
RunChecks = False
End If
End Function
Function checkDataMatchesMapping(wb As Workbook) As Boolean
Dim mapping As Variant, header As Variant
Dim i As Long

With ThisWorkbook
header = rangeBlock(.Sheets("Data").Range("DataStart"), 1, 0)
mapping = rangeBlock(.Sheets("Mapping").Range("MapStart").Offset(1, 0), 0, 1)

mapping = TransposeArray(mapping)

For i = LBound(header) To UBound(header)
If header(1, i) <> mapping(1, i) Then

checkDataMatchesMapping = False
MsgBox ("The data headers do not match the mapping file")
Exit Function
End If
Next i
End With
checkDataMatchesMapping = True
End Function


Sub runcheckDataMatchesMapping()
Dim wb As Workbook
Set wb = ActiveWorkbook
Debug.Print checkDataMatchesMapping(wb)

End Sub

Sub runcheckFileNotRunYet()
Dim wb As Workbook
Set wb = ActiveWorkbook
Debug.Print checkFileNotRunYet(wb)

End Sub
Function checkFileNotRunYet(wb As Workbook) As Boolean
Dim rowCheck As Long

On Error Resume Next
rowCheck = Application.Match(Trim(wb.Sheets(1).Name), ThisWorkbook.Sheets("Data").Columns(1), 0)
On Error GoTo 0

If rowCheck > 0 Then
MsgBox ("the file has already been run")
checkFileNotRunYet = False
Else
checkFileNotRunYet = True
End If

End Function
Function CheckHeadingMapping(wb As Workbook) As Boolean
Dim heading As Range, header As Range, data As Variant
Dim allmapping As Variant
Dim headingIndex As Long
Dim numCol As Long, numRow As Long

allmapping = getMapping(wb)

With ThisWorkbook.Sheets("Mapping")
Set header = rangeBlock(wb.Sheets(1).Range("A1"), 1, 0)
End With

For Each heading In header
On Error Resume Next
headingIndex = Application.VLookup(Trim(heading), allmapping, 2, 0)
On Error GoTo 0

If headingIndex > 0 Then
' no problems
Else
MsgBox ("Missing column:" & Trim(heading))
CheckHeadingMapping = False
Exit Function
End If
headingIndex = 0
Next heading

CheckHeadingMapping = True
End Function


Option Explicit

Sub LoopThroughFiles()
Dim StrFile As String
Dim wb As Workbook

Call StartEnviroment

StrFile = Dir("V:\PipelineCapacity\Files\")
Do While Len(StrFile) > 0
Set wb = Workbooks.Open("V:\PipelineCapacity\Files\" & StrFile)
Call main

wb.Close (False)
StrFile = Dir

Loop

Call EndEnviroment
End Sub

Sub main()
If ActiveWorkbook.Name <> ThisWorkbook.Name Then
If RunChecks Then
Call RunCopy
Else
Stop
End If
Else
Stop
End If
End Sub

Sub RunCopy()
Dim wb As Workbook
Dim data As Variant, allmapping As Variant
Dim colIndex As Long, numRows As Long, rowStart As Long, numCols As Long, j As Long
Dim pasteStart As Range, header As Range
Dim debugArray As Variant

Set wb = ActiveWorkbook
allmapping = getMapping(wb)
data = getData(wb)
Set header = rangeBlock(wb.Sheets(1).Range("A1"), 1, 0)

numRows = lengtharray(data, 1)
numCols = header.Count - 1
Set pasteStart = ThisWorkbook.Sheets("Data").Range("A60000").End(xlUp).Offset(1, 0)

rangeBlock(pasteStart, numRows, 1).Value = wb.Sheets(1).Name
For j = 1 To numCols
On Error Resume Next
colIndex = Application.VLookup(Trim(header(, j)), allmapping, 2, 0)
On Error GoTo 0

If colIndex > 0 Then
With pasteStart
debugArray = getSlice(data, 0, j)
rangeBlock(.Offset(0, colIndex - 1), numRows, 1).Value = getSlice(data, 0, j)

End With

Else
Stop 'need to alert
End If
colIndex = -1
Next j


End Sub


Function getData(wb As Workbook) As Variant
Dim data As Variant
Dim outData As Variant
Dim numCol As Long, numRow As Long

numCol = rangeBlock(wb.Sheets(1).Range("A1"), 1, 0).Count - 1
numRow = wb.Sheets(1).Range("A10000").End(xlUp).row

data = ArrayFromRangeWithoutBlankRows(wb.Sheets(1).Range("A2"), numRow - 1, numCol)

getData = data
End Function
Function getMapping(wb As Workbook)
Dim heading As Range, header As Range, data As Variant
Dim allmapping As Variant

With ThisWorkbook.Sheets("Mapping")
Set header = rangeBlock(wb.Sheets(1).Range("A1"), 1, 0)
allmapping = rangeBlock(.Range("MapStart").Offset(1, 0), 0, 2).Value
allmapping = matrixToVector(allmapping, True)
End With

getMapping = allmapping
End Function



