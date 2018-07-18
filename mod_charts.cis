Function SaveCharts()
    Sheets("EIAChart").Shapes("Chart 1").Height = 288
    Sheets("EIAChart").Shapes("Chart 1").Width = 504
    Sheets("EIAChart").ChartObjects("Chart 1").Chart.Export "NatEIALoadForecastEvolution.png"
End Function

Sub UpdateAreaChartAxis(SelectedChart As ChartObject, yMin As Long, Optional yMax As Long, Optional yUnit As Long)
    With SelectedChart.Chart
        With .Axes(xlValue)
            If yMax <> Null Then
                .MaximumScale = yMax
            End If
            .MinimumScale = yMin
            If yUnit <> Null Then
                .MajorUnit = yUnit
            End If
        End With
    End With
End Sub

Sub saveTablePNG()
    Const FName As String = "EIANatLoadTable.png"
    Dim pic_rng As Range
    Dim ShTemp As Worksheet
    Dim ChTemp As Chart
    Dim PicTemp As Picture
    Dim area As Range
    
    Application.ScreenUpdating = False
    Set pic_rng = Worksheets("EIAChart").Range("A1:I14")
    Set ShTemp = Worksheets.Add
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:=ShTemp.Name
    Set area = Sheets("EIAChart").Range(Sheets("EIAChart").PageSetup.PrintArea)
    area.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    Set ChTemp = ActiveChart
    
    
    ChTemp.Paste
    Set PicTemp = Selection
    With ChTemp.Parent
    .Width = area.Width
    .Height = area.Height
    End With
    ChTemp.Export Filename:="EIANatLoadTable.png", FilterName:="png"
    'UserForm1.Image1.Picture = LoadPicture(FName)
    ShTemp.Delete
End Sub

Sub ColorLines(SelectedChart As ChartObject)
    Dim myChartObject As ChartObject
    Dim seriesCount As Long, i As Long
    Dim mySrs As Series
    Dim myPts As Points
    
    seriesCount = SelectedChart.Chart.SeriesCollection.Count
    For i = seriesCount To 1 Step -1
        Set mySrs = SelectedChart.Chart.SeriesCollection(i)
        If mySrs.Name = "1/1/1900 - Norms" Then
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            mySrs.Format.Line.DashStyle = msoLineSysDash
            
        ElseIf mySrs.Name = "1/1/1900 - Actuals" Or mySrs.Name = "1/1/1900 - Actual" Then
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            mySrs.Format.Line.DashStyle = msoLineSolid
            mySrs.Format.Line.Weight = 3
         ElseIf Left(mySrs.Name, 18) = "1/1/1900 - StdDev1" Then
            mySrs.Format.Line.ForeColor.RGB = RGB(80, 80, 80)
            mySrs.Format.Line.DashStyle = msoLineSysDash
            mySrs.Format.Line.Weight = 1.5
            mySrs.MarkerStyle = -4142
        ElseIf Left(mySrs.Name, 18) = "1/1/1900 - StdDev2" Then
            mySrs.Format.Line.ForeColor.RGB = RGB(160, 160, 160)
            mySrs.Format.Line.DashStyle = msoLineSysDash
            mySrs.Format.Line.Weight = 1.5
            mySrs.MarkerStyle = -4142
        ElseIf Right(mySrs.Name, 18) = "1/1/1900 - Norm" Then
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            mySrs.Format.Line.Weight = 1.5
        ElseIf i = seriesCount Then
            mySrs.Format.Line.Weight = 3
            mySrs.Format.Line.ForeColor.RGB = RGB(66, 28, 94)
            mySrs.MarkerSize = 2
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 1 Then
            mySrs.Format.Line.Weight = 2.5
            mySrs.Format.Line.ForeColor.RGB = RGB(120, 100, 150)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 2 Then
            mySrs.Format.Line.Weight = 2
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 130, 192)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 3 Then
            mySrs.Format.Line.Weight = 1.75
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 176, 240)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 4 Then
            mySrs.Format.Line.Weight = 1.6
            mySrs.Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 5 Then
            mySrs.Format.Line.Weight = 1.5
            mySrs.Format.Line.ForeColor.RGB = RGB(146, 208, 80)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 6 Then
            mySrs.Format.Line.Weight = 1.4
            mySrs.Format.Line.ForeColor.RGB = RGB(200, 200, 0)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 7 Then
            mySrs.Format.Line.Weight = 1.25
            mySrs.Format.Line.ForeColor.RGB = RGB(220, 220, 0)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 8 Then
            mySrs.Format.Line.Weight = 1
            mySrs.Format.Line.ForeColor.RGB = RGB(255, 255, 0)
            Call FormatMarkers(mySrs)
        ElseIf i = seriesCount - 9 Then
            mySrs.Format.Line.Weight = 0.7
            Call FormatMarkers(mySrs)
        End If
    Next i
End Sub

Function FormatMarkers(mySrs As Series)
    mySrs.MarkerStyle = 8
    mySrs.MarkerSize = 2
    mySrs.Format.Line.Visible = msoTrue
'    With mySrs.MarkerStyle.Format.Fill
'        .Visible = msoTrue
'        .ForeColor.ObjectThemeColor = msoThemeColorText1
'        .ForeColor.TintAndShade = 0
'        .Solid
'    End With
End Function
