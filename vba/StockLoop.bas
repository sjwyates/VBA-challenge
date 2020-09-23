Attribute VB_Name = "StockLoop"
Option Explicit

Sub StockLoop()
'   -----------------------------------
'   Start a timer to check performance
'   -----------------------------------
    Dim startTime, stopTime, totalTime As Single
    startTime = Timer
'   -----------------------------------
'   Instantiate a Collection object
'   -----------------------------------
    Dim theStocks As Collection
    Set theStocks = New Collection
'   ------------------------------------
'   Loop #1: Loop through all worksheets
'   ------------------------------------
    Dim CurrentWS As Worksheet
    For Each CurrentWS In Worksheets
    '   -----------------------------------
    '   Some variables for loops
    '   -----------------------------------
        Dim lastRow As Long
        lastRow = CurrentWS.Cells(Rows.Count, 1).End(xlUp).row
        Dim i As Long
    '   -----------------------------------
    '   Copy data from Range to 2D array
    '   -----------------------------------
        Dim sheetData As Variant
        sheetData = CurrentWS.Range("A2:G" & lastRow).Value
    '   -----------------------------------
    '   Loop #2: Get data from 2D array
    '   -----------------------------------
        Dim theStock As CStock
        Dim theYear As Integer
        Dim id As String
        Dim exists As Boolean
        For i = 1 To UBound(sheetData, 1)
            theYear = Round(sheetData(i, 2) / 10000, 0)
            ' Unique ID for each stock, for each year
            id = sheetData(i, 1) & "_" & theYear
            ' Custom util function (below end sub)
            exists = ExistsInCollection(id, theStocks)
            If exists = False Then
                Set theStock = New CStock
                With theStock
                    .TickerID = sheetData(i, 1)
                    .StockYear = theYear
                    .InitializeValues sheetData(i, 3), _
                                      sheetData(i, 6), _
                                      sheetData(i, 2), _
                                      sheetData(i, 7)
                    theStocks.Add theStock, id
                End With
            ElseIf exists = True Then
                Set theStock = theStocks.Item(id)
                With theStock
                    .UpdateValues sheetData(i, 3), _
                                  sheetData(i, 6), _
                                  sheetData(i, 2), _
                                  sheetData(i, 7)
                End With
            End If
        Next i
    Next CurrentWS
'   ------------------------------------
'   New worksheet to store combined data
'   ------------------------------------
    Set CurrentWS = ActiveWorkbook.Sheets.Add
    CurrentWS.Name = "Analysis"
'   -----------------------------------
'   Instantiate the record-holders
'   -----------------------------------
    Dim greatestPctIncrease, _
        greatestPctDecrease, _
        greatestTotalVolume _
        As CStock
    Set greatestPctIncrease = New CStock
    Set greatestPctDecrease = New CStock
    Set greatestTotalVolume = New CStock
'   ------------------------------------------
'   Loop #3: Find the record-holders
'   ------------------------------------------
    Dim recordHolders As Collection
    Set recordHolders = New Collection
    
    Dim aStock As CStock
    For Each aStock In theStocks
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
    
    recordHolders.Add greatestPctIncrease
    recordHolders.Add greatestPctDecrease
    recordHolders.Add greatestTotalVolume
'   ------------------------------------------
'   Loop #4: Populate stock table
'   ------------------------------------------
    i = 2
    For Each aStock In theStocks
        With CurrentWS
            .Cells(i, 1).Value = aStock.TickerID
            .Cells(i, 2).Value = aStock.StockYear
            .Cells(i, 3).Value = aStock.YearlyChange
            .Cells(i, 4).Value = aStock.PercentChange
            .Cells(i, 5).Value = aStock.TotalVolume
        End With
        i = i + 1
    Next aStock
'   -----------------------------------
'   Loop #5: Populate record-holders
'   -----------------------------------
    i = 2
    For Each aStock In recordHolders
        With CurrentWS
            .Cells(i, 8).Value = aStock.TickerID
            .Cells(i, 9).Value = aStock.StockYear
            If i < 4 Then
                .Cells(i, 10).Value = aStock.PercentChange
            Else
                .Cells(i, 10).Value = aStock.TotalVolume
            End If
        End With
        i = i + 1
    Next aStock
'   -----------------------------------
'   Loop #6: Add header titles
'   -----------------------------------
    Dim headerTitles() As String
    headerTitles = Split("Ticker|Year|Yearly Change|Percent Change|Total Stock Volume||Category|Ticker|Year|Value", "|")
    For i = 0 To UBound(headerTitles)
        CurrentWS.Cells(1, i + 1).Value = headerTitles(i)
    Next i
'   -----------------------------------
'   Loop #7: Add category names
'   -----------------------------------
    Dim categoryNames() As String
    categoryNames = Split("Greatest % Increase|Greatest % Decrease|Greatest Total Volume", "|")
    For i = 0 To UBound(categoryNames)
        CurrentWS.Cells(i + 2, 7).Value = categoryNames(i)
    Next i
'   -----------------------------------
'   Sort by ticker, then year
'   -----------------------------------
    With CurrentWS.Sort
        .SortFields.Add key:=Range("A1"), Order:=xlAscending
        .SortFields.Add key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A1:E" & theStocks.Count + 1)
        .Header = xlYes
        .Apply
    End With
'   -----------------------------------
'   Add basic formatting
'   -----------------------------------
    With CurrentWS
        .Cells.HorizontalAlignment = xlCenter
        .Columns("F").ColumnWidth = 1
        With .Range("A1:E1, G1:J1")
            .Interior.Color = RGB(200, 200, 200)
            .Borders.Color = 0
            .WrapText = True
        End With
        With .Range("G2:J4, A2:E" & theStocks.Count + 1)
            .Borders.Color = 0
            .EntireColumn.AutoFit
        End With
        .Range("D2:D" & theStocks.Count + 1 & ",J2:J3") _
            .NumberFormat = "0.00%"
    End With
'   ----------------------------------------
'   New conditional formatting (color scale)
'   ----------------------------------------
    Dim pctChangeColorScale As ColorScale
    Set pctChangeColorScale = CurrentWS.Range("D2:D" & theStocks.Count + 1) _
        .FormatConditions.AddColorScale(ColorScaleType:=3)
    With pctChangeColorScale
        With .ColorScaleCriteria(1)
            .FormatColor.Color = RGB(255, 50, 50)
            .Type = xlConditionValueNumber
            .Value = -1
        End With
        With .ColorScaleCriteria(2)
            .FormatColor.Color = RGB(255, 255, 255)
            .Type = xlConditionValueNumber
            .Value = 0
        End With
        With .ColorScaleCriteria(3)
            .FormatColor.Color = RGB(50, 255, 50)
            .Type = xlConditionValueNumber
            .Value = 1
        End With
    End With
'   ------------------------------------
'   Old conditional formatting (boolean)
'   ------------------------------------
'    With CurrentWS.Range("C2:C" & theStocks.Count + 1).FormatConditions
'        .Add(xlCellValue, xlGreater, "=0") _
'            .Interior.Color = RGB(50, 255, 50)
'        .Add(xlCellValue, xlLess, "=0") _
'            .Interior.Color = RGB(255, 50, 50)
'    End With

'   -----------------------------------
'   Get stop time & display duration
'   -----------------------------------
    stopTime = Timer
    totalTime = Round(stopTime - startTime, 2)
    With CurrentWS
        With .Range("G8")
            .Value = "Duration:"
            .Interior.Color = RGB(200, 200, 200)
            .Borders.Color = 0
        End With
        With .Range("G9")
            .Value = totalTime & " seconds"
            .Borders.Color = 0
        End With
    End With
    
End Sub

' --------------------------------------------
' Util to check if object exists in collection
' (weird VBA workaround using error handling)
' --------------------------------------------
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


