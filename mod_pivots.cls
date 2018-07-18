Sub RefreshPivotTables()
    Sheets("EIAChart").PivotTables("PivotTable1").PivotCache.Refresh
    Sheets("EIAChart").PivotTables("PivotTable1").PivotFields("AsOfDate").AutoSort _
        xlAscending, "AsOfDate"
    Sheets("EIAChart").PivotTables("PivotTable1").PivotFields("ForecastedDate").AutoSort _
        xlAscending, "ForecastedDate"
    Calculate
End Sub
