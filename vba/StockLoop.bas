Attribute VB_Name = "StockLoop"
Option Explicit

Sub StockLoop()
'   -----------------------------------
'   Instantiate a Collection object
'   -----------------------------------
    Dim theStocks As Collection
    Set theStocks = New Collection
'   -----------------------------------
'   Find the last row in the table
'   -----------------------------------
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
'   --------------------------------------------
'   Clear existing contents/formats, set headers
'   --------------------------------------------
    With Cells
        .ClearFormats
        .HorizontalAlignment = xlCenter
    End With
    Range("I1:Z" & lastRow).ClearContents
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Year"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("O1").Value = "Category"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Columns().ColumnWidth = 8
    Columns("H").ColumnWidth = 1
    Columns("N").ColumnWidth = 1
    Columns("M").ColumnWidth = 11
    Columns("G").ColumnWidth = 11
    Columns("Q").ColumnWidth = 11
    Columns("O").AutoFit
'   -----------------------------------
'   Add formatting to headers
'   -----------------------------------
    With Range("A1:G1,I1:M1,O1:Q1,O2:O4")
        .Interior.Color = RGB(200, 200, 200)
        .Borders.Color = 0
    End With
    Rows(1).WrapText = True
'   -----------------------------------
'   Dim variables for the loop
'   -----------------------------------
    Dim row As Long
    Dim theStock As CStock
    Dim theYear As Integer
    Dim id As String
    Dim exists As Boolean
'   -----------------------------------
'   Loop over all the rows
'   -----------------------------------
    For row = 2 To lastRow
        ' Call util function to get the year
        theYear = YearFromDate(Cells(row, 2).Value)
        ' Create a unique ID to use as key in collection
        id = Cells(row, 1).Value & "_" & theYear
        ' Call util function to check if a stock already exists
        exists = ExistsInCollection(id, theStocks)
        ' If it doesn't exist yet, instantiate and set initial values
        If exists = False Then
            Set theStock = New CStock
            With theStock
                .TickerID = Cells(row, 1).Value
                .StockYear = theYear
                .InitialDate = Cells(row, 2).Value
                .FinalDate = Cells(row, 2).Value
                .InitialValue = Cells(row, 3).Value
                .FinalValue = Cells(row, 6).Value
                .IncTotalVolume Cells(row, 7).Value
            End With
            ' Add to collection with ID as key
            theStocks.Add theStock, id
        ' If it already exists, then edit it (with some logic)
        ElseIf exists = True Then
            Set theStock = theStocks.Item(id)
            With theStock
                If .InitialDate > Cells(row, 2).Value Then
                    .InitialDate = Cells(row, 2).Value
                    .InitialValue = Cells(row, 3).Value
                End If
                If .FinalDate < Cells(row, 2).Value Then
                    .FinalDate = Cells(row, 2).Value
                    .FinalValue = Cells(row, 6).Value
                End If
                .IncTotalVolume Cells(row, 7).Value
            End With
        End If
    Next row
'   -----------------------------------
'   Loop variables
'   -----------------------------------
    Dim aStock As CStock
    Dim i As Integer
    i = 2
'   -----------------------------------
'   Greatest category variables
'   -----------------------------------
    Dim greatestPctIncrease As CStock
    Dim greatestPctDecrease As CStock
    Dim greatestTotalVolume As CStock
    Set greatestPctIncrease = New CStock
    Set greatestPctDecrease = New CStock
    Set greatestTotalVolume = New CStock
'   -----------------------------------
'   Populate main table and find greatests
'   -----------------------------------
    For Each aStock In theStocks
        ' Populate main table
        With aStock
            Cells(i, 9).Value = .TickerID
            Cells(i, 10).Value = .StockYear
            Cells(i, 11).Value = .YearlyChange
            Cells(i, 12).Value = .PercentChange
            Cells(i, 13).Value = .TotalVolume
            i = i + 1
        End With
        ' Check if this is the greatest in any categories
        If aStock.PercentChange > greatestPctIncrease.PercentChange Then
            Set greatestPctIncrease = aStock
        End If
        If aStock.PercentChange < greatestPctDecrease.PercentChange Then
            Set greatestPctDecrease = aStock
        End If
        If aStock.TotalVolume > greatestTotalVolume.TotalVolume Then
            Set greatestTotalVolume = aStock
        End If
    Next aStock
'   -----------------------------------
'   Basic table formatting
'   -----------------------------------
    With Range("A2:G" & lastRow & ",I2:M" & theStocks.Count + 1 & ",O2:Q4")
        .Borders.Color = 0
        .VerticalAlignment = xlCenter
    End With
    Range("L2:L" & theStocks.Count + 1 & ",Q2:Q3").NumberFormat = "0.00%"
    Range("M2:M" & theStocks.Count + 1 & ",O2:O4").ShrinkToFit = True
'   -----------------------------------
'   Populate greatest categories table
'   -----------------------------------
    Range("P2").Value = greatestPctIncrease.TickerID
    Range("Q2").Value = greatestPctIncrease.PercentChange
    Range("P3").Value = greatestPctDecrease.TickerID
    Range("Q3").Value = greatestPctDecrease.PercentChange
    Range("P4").Value = greatestTotalVolume.TickerID
    Range("Q4").Value = greatestTotalVolume.TotalVolume
'   ----------------------------------------
'   Conditional formatting for yearly change
'   ----------------------------------------
    With Columns("K")
        .FormatConditions.Delete
        .FormatConditions.Add(xlCellValue, xlGreater, "=0") _
            .Interior.Color = RGB(50, 255, 50)
        .FormatConditions.Add(xlCellValue, xlLess, "=0") _
            .Interior.Color = RGB(255, 50, 50)
    End With
    Range("K1").FormatConditions.Delete
'   ----------------------------------------
'   Lastly, select cell A1 (as a nicety)
'   ----------------------------------------
    Range("A1").Select

End Sub

' ------------------------------------
' Util to check if object exists in collection
' (weird VBA workaround using error handler)
' ------------------------------------
Function ExistsInCollection(ByVal key As String, _
                            ByRef coll As Collection) _
                            As Boolean
    Dim obj As Object
    On Error GoTo err
        ' If no obj with this key, it fires an error
        Set obj = coll.Item(key)
        ' If there is an object, it goes here and returns True
        ExistsInCollection = True
        Exit Function
err:
        ' Error handler just returns false
        ExistsInCollection = False
End Function

' ------------------------------------
' Util to extract year from date Long
' ------------------------------------
Function YearFromDate(fullDate As Long)
    YearFromDate = (fullDate - fullDate Mod 10000) / 10000
End Function
