
Function getSlice(data As Variant, row As Long, col As Long) As Variant
Dim outData As Variant
Dim i As Long, j As Long

If row = 0 And col > 0 Then
ReDim outData(LBound(data, 1) To UBound(data, 1), 0) As Variant
For i = LBound(data, 1) To UBound(data, 1)
outData(i, 0) = data(i, col - 1)
Next i
ElseIf row > 0 And col = 0 Then
ReDim outData(0, LBound(data, 2) To UBound(data, 2)) As Variant
For i = LBound(data, 2) To UBound(data, 2)
outData(0, i) = data(row - 1, 0)
Next i
Else

End If
getSlice = outData
End Function

Function appendArray(data1 As Variant, data2 As Variant) As Variant
'must be a single column array
Dim mergedArray As Variant
Dim j As Integer
Dim i As Integer

ReDim Preserve data1(lengtharray(data1, 1) + lengtharray(data2) + 1)

For i = (lengtharray(data1) + 1) To UBound(data2, 1)
arr(j) = ar(i, 1)
j = j + 1
Next i

appendArray = data1

End Sub


Function lengtharray(data As Variant, ltype As Long) As Long
lengtharray = UBound(data, ltype) - LBound(data, ltype) + 1
End Function
Function TransposeArray(InputArr As Variant) As Variant

Dim RowNdx, ColNdx, LB1, LB2, UB1, UB2 As Long, tmpArray As Variant

LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)

ReDim tmpArray(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)

For RowNdx = LB2 To UB2
For ColNdx = LB1 To UB1
tmpArray(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
Next ColNdx
Next RowNdx

TransposeArray = tmpArray

End Function

Function ArrayFromRangeWithoutBlankRows(startRng As Range, numRow As Long, numCol As Long) As Variant
Dim data As Variant

data = rangeBlock(startRng, numRow, numCol)
data = arrayRemoveBlankRows(data)

ArrayFromRangeWithoutBlankRows = data
End Function

Function arrayRemoveBlankRows(data As Variant) As Variant
Dim outData As Variant
Dim i, j, x, y As Long
Dim blankFlag As Boolean
Dim blankCount As Long

ReDim outData(lengtharray(data, 1) - 1, lengtharray(data, 2) - 1)
blankCount = 0
x = 0
For i = LBound(data, 1) To UBound(data, 1)

'Check if row is blank
blankFlag = True
For j = LBound(data, 2) To UBound(data, 2)
If Not IsEmpty(data(i, j)) Then
blankFlag = False
Exit For
End If
Next j

If Not blankFlag Then
y = 0
For j = LBound(data, 2) To UBound(data, 2)
outData(x, y) = data(i, j)
y = y + 1
Next j

x = x + 1
Else
blankCount = blankCount + 1
End If


Next i

'Can only redim last dimension
outData = TransposeArray(outData)
ReDim Preserve outData(lengtharray(outData, 1) - 1, lengtharray(outData, 2) - 1 - blankCount)
arrayRemoveBlankRows = TransposeArray(outData)
End Function

Function matrixToVector(data As Variant, includeIndex As Boolean) As Variant
If includeIndex Then
matrixToVector = matrixToVectorWithIndex(data)
Else
matrixToVector = matrixToVectorWithNoIndex(data)
End If
End Function

Function matrixToVectorWithNoIndex(data As Variant) As Variant
Dim i, j As Long
Dim vector As Variant

ReDim vector(lengtharray(data, 1) * lengtharray(data, 2))

For j = LBound(data, 2) To UBound(data, 2)
For i = LBound(data, 1) To UBound(data, 1)
vector(i + (j - 1) * lengtharray(data, 1)) = data(i, j)
Next i
Next j

matrixToVectorWithNoIndex = vector
End Function

Function matrixToVectorWithIndex(data As Variant) As Variant
Dim i, j As Long
Dim vector As Variant

ReDim vector(lengtharray(data, 1) * lengtharray(data, 2), 0 To 1)

For j = LBound(data, 2) To UBound(data, 2)
For i = LBound(data, 1) To UBound(data, 1)
vector(i + (j - 1) * lengtharray(data, 1), 0) = data(i, j)
vector(i + (j - 1) * lengtharray(data, 1), 1) = i
Next i
Next j

matrixToVectorWithIndex = vector
End Function

