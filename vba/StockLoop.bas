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
    '   Loop #2: Get data from CurrentWS
    '   -----------------------------------
        Dim sheetData As Variant
        sheetData = CurrentWS.Range("A2:G" & lastRow).Value
        
        Dim theStock As CStock
        Dim theYear As Integer
        Dim id As String
        Dim exists As Boolean
        For i = 1 To UBound(sheetData, 1)
            theYear = Round(sheetData(i, 2) / 10000, 0)
            id = sheetData(i, 1) & "_" & theYear
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
'   Greatest category variables
'   -----------------------------------
    Dim greatestPctIncrease, _
        greatestPctDecrease, _
        greatestTotalVolume _
    As CStock
    Set greatestPctIncrease = New CStock
    Set greatestPctDecrease = New CStock
    Set greatestTotalVolume = New CStock
'   ------------------------------------------
'   Loop #3: Find the greatests
'   ------------------------------------------
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
'   ------------------------------------------
'   Loop #4: Populate stock table
'   ------------------------------------------
    i = 2
    For Each aStock In theStocks
        With aStock
            CurrentWS.Cells(i, 1).Value = .TickerID
            CurrentWS.Cells(i, 2).Value = .StockYear
            CurrentWS.Cells(i, 3).Value = .YearlyChange
            CurrentWS.Cells(i, 4).Value = .PercentChange
            CurrentWS.Cells(i, 5).Value = .TotalVolume
            i = i + 1
        End With
    Next aStock
'   -----------------------------------
'   Populate greatest categories table
'   -----------------------------------
    With CurrentWS
        .Range("H2").Value = greatestPctIncrease.TickerID
        .Range("I2").Value = greatestPctIncrease.PercentChange
        .Range("H3").Value = greatestPctDecrease.TickerID
        .Range("I3").Value = greatestPctDecrease.PercentChange
        .Range("H4").Value = greatestTotalVolume.TickerID
        .Range("I4").Value = greatestTotalVolume.TotalVolume
    End With
'   -----------------------------------
'   Add header titles
'   -----------------------------------
    Dim headerTitles() As String
    headerTitles = Split("Ticker|Year|Yearly Change|Percent Change|Total Stock Volume||Category|Ticker|Value", "|")
    For i = 0 To UBound(headerTitles)
        CurrentWS.Cells(1, i + 1).Value = headerTitles(i)
    Next i
'   -----------------------------------
'   Add category names
'   -----------------------------------
    Dim categoryNames() As String
    categoryNames = Split("Greatest % Increase|Greatest % Decrease|Greatest Total Volume", "|")
    For i = 0 To UBound(categoryNames)
        CurrentWS.Cells(i + 2, 7).Value = categoryNames(i)
    Next i
'   -----------------------------------
'   Add formatting
'   -----------------------------------
    With CurrentWS
        With .Cells
            .HorizontalAlignment = xlCenter
            .EntireColumn.AutoFit
        End With
        With .Range("A1:E1, G1:I1")
            .Interior.Color = RGB(200, 200, 200)
            .Borders.Color = 0
            .WrapText = True
        End With
        .Range("G2:I4, A2:E" & theStocks.Count + 1).Borders.Color = 0
        .Range("D2:D" & theStocks.Count + 1 & ",I2:I3").NumberFormat = "0.00%"
        With .Columns("C")
            .FormatConditions.Add(xlCellValue, xlGreater, "=0") _
                .Interior.Color = RGB(50, 255, 50)
            .FormatConditions.Add(xlCellValue, xlLess, "=0") _
                .Interior.Color = RGB(255, 50, 50)
        End With
        .Range("C1").FormatConditions.Delete
    End With
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
