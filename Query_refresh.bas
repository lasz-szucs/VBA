'Refresh all queries in the workbook

For Each objConnection In ThisWorkbook.Connections
    bBackground = objConnection.OLEDBConnection.BackgroundQuery
    objConnection.OLEDBConnection.BackgroundQuery = False
    objConnection.Refresh
    objConnection.OLEDBConnection.BackgroundQuery = bBackground
Next