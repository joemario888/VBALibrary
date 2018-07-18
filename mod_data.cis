Option Explicit
Function rangeBlock(rng As Range, rowLength As Long, colLength As Long) As Range
    Dim nRow As Long
    Dim nCol As Long

    If rowLength = 0 Then
        If rng.Offset(1, 0).Formula = "" Then
            nRow = 0
        Else
            nRow = rng.End(xlDown).Row - rng.Row
        End If
    Else
        nRow = rowLength - 1
    End If

    If colLength = 0 Then
        If rng.Offset(0, 1).Formula = "" Then
            nCol = 0
        Else
            nCol = rng.End(xlToRight).Column - rng.Column
        End If
    Else
        nCol = colLength - 1
    End If
    Set rangeBlock = Range(rng, rng.Offset(nRow, nCol))
End Function

Sub fillMissing(sht As Worksheet, colOffset As Long, cellStart As Range)
    Dim rng As Range
    With sht
        colCount = rangeblock(cellStart, 1, 0).Count - colOffset
        rowCount = rangeblock(cellStart, 0, 1).Count
            
        With cellStart
            For i = 1 To colCount
                If .Offset(0, i) = "Total" Then
                    Exit For
                Else
                    With .Offset(rowCount, i)
                        If rowCount > (.End(xlUp).Row - cellStart.Row + 1) Then
                            Set rng = rangeblock(.End(xlUp).Offset(1, 0), rowCount - .End(xlUp).Row + cellStart.Row - 1, 1)
                            rng.Value = .End(xlUp).Value
                            rng.Font.Color = -16776961
                        End If
                    End With
                End If
            Next i
        End With
            
    End With


Function getConn(serverName As String, dbname As String) As ADODB.Connection
    Dim conn As ADODB.Connection, connString As String
    
    Set conn = New ADODB.Connection
    With conn
        .Provider = "SQLOLEDB"
        .Properties("Initial Catalog") = dbname
        .Properties("Data Source") = serverName
        .Properties("Integrated Security") = "SSPI"
        .CommandTimeout = 600
        .ConnectionTimeout = 30
        .Open ConnectionString:=connString
    End With
    Set getConn = conn
End Function
Function getData(conn As ADODB.Connection, sql As String) As Variant
    Dim rs As ADODB.Recordset, values As Variant

    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = conn
        .Open Source:=sql, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly
    End With
    
    If Not rs.EOF Then
        values = rs.GetRows
        getData = TransposeArray(values)
    End If
End Function