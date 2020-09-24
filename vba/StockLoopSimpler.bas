Attribute VB_Name = "StockLoopSimpler"
Sub StockLoopSimpler()

    ' Start timer to check perfomance
    Dim startTime, stopTime, totalTime As Single
    startTime = Timer

    ' Dim variables to reference worksheets
    Dim DataWS, AnalysisWS As Worksheet
    
    ' Make a new worksheet for the results
    Set AnalysisWS = ActiveWorkbook.Sheets.Add
    AnalysisWS.Name = "Analysis"
    
    ' Dim some useful variables
    Dim dLastRow, aLastRow As Long
    Dim i, j As Long
    
    ' Select the first worksheet
    For Each DataWS In Worksheets
    
        ' Ignore analysis WS
        If DataWS.Name <> "Analysis" Then
        
            With DataWS
                ' Find the last row in the current worksheet
                dLastRow = .Cells(Rows.Count, 1).End(xlUp).row
                
                ' Dim loop variables
                Dim currentTicker As String
                Dim yearlyOpen, _
                    yearlyClose, _
                    yearlyChange, _
                    percentChange, _
                    totalVolume As Variant
                
                ' Set initial values for first stock
                currentTicker = .Cells(2, 1).Value
                yearlyOpen = .Cells(2, 3).Value
                totalVolume = 0
                
                ' The first row of the analysis table
                j = 1
                
                ' Loop to get all unique tickers & total volumes
                For i = 2 To dLastRow
                    ' Always increment totalVolume
                    totalVolume = totalVolume + .Cells(i, 7).Value
                    ' When you're on the last row of a stock...
                    If currentTicker <> .Cells(i + 1, 1).Value Then
                        ' Package up data for current stock...
                        yearlyClose = .Cells(i, 6).Value
                        yearlyChange = yearlyClose - yearlyOpen
                        If yearlyOpen = 0 Then
                            percentChange = 0
                        Else
                            percentChange = yearlyChange / yearlyOpen
                        End If
                        ' Add data to the analysis spreadsheet...
                        AnalysisWS.Cells(j + aLastRow, 1).Value = currentTicker
                        AnalysisWS.Cells(j + aLastRow, 2).Value = yearlyChange
                        AnalysisWS.Cells(j + aLastRow, 3).Value = percentChange
                        AnalysisWS.Cells(j + aLastRow, 4).Value = totalVolume
                        ' Increment the row on the analysis table
                        j = j + 1
                        ' Then reset everything for the next stock
                        currentTicker = .Cells(i + 1, 1).Value
                        yearlyOpen = .Cells(i + 1, 3).Value
                        totalVolume = 0
                    End If
                Next i
                
            End With
            
        End If
        
        ' Reset the last row of analysis worksheet
        aLastRow = AnalysisWS.Cells(Rows.Count, 1).End(xlUp).row
        
    Next DataWS
    
    ' Dim variables for record-holders
    Dim recordIncrease, recordDecrease As Double
    Dim recordVolume As Variant
    Dim recordIncreaseTic, _
        recordDecreaseTic, _
        recordVolumeTic As String
    
    ' Set initial values
    recordIncrease = 0
    recordDecrease = 0
    recordVolume = 0
    
    With AnalysisWS
        
        ' Loop to find the record-holders
        For i = 2 To aLastRow
            If .Cells(i, 3).Value > recordIncrease Then
                recordIncrease = .Cells(i, 3).Value
                recordIncreaseTic = .Cells(i, 1).Value
            End If
            If .Cells(i, 3).Value < recordDecrease Then
                recordDecrease = .Cells(i, 3).Value
                recordDecreaseTic = .Cells(i, 1).Value
            End If
        
            If .Cells(i, 4).Value > recordVolume Then
                recordVolume = .Cells(i, 4).Value
                recordVolumeTic = .Cells(i, 1).Value
            End If
        Next i
    
        ' Add record-holders table
        .Range("G2").Value = recordIncreaseTic
        .Range("H2").Value = recordIncrease
        .Range("G3").Value = recordDecreaseTic
        .Range("H3").Value = recordDecrease
        .Range("G4").Value = recordVolumeTic
        .Range("H4").Value = recordVolume
        
        ' Add row headers
        Dim headerTitles() As String
        headerTitles = Split("Ticker|Yearly Change|Percent Change|Total Stock Volume||Category|Ticker|Value", "|")
        For i = 0 To UBound(headerTitles)
            .Cells(1, i + 1).Value = headerTitles(i)
        Next i
        
        ' Add category titles
        Dim categoryNames() As String
        categoryNames = Split("Greatest % Increase|Greatest % Decrease|Greatest Total Volume", "|")
        For i = 0 To UBound(categoryNames)
            .Cells(i + 2, 6).Value = categoryNames(i)
        Next i
        
        ' Basic formatting
        .Cells.HorizontalAlignment = xlCenter
        .Columns("E").ColumnWidth = 1
        With .Range("A1:D1, F1:H1")
            .Interior.Color = RGB(200, 200, 200)
            .Borders.Color = 0
            .WrapText = True
        End With
        With .Range("A2:D" & aLastRow & ", F2:H4")
            .Borders.Color = 0
            .EntireColumn.AutoFit
        End With
        .Range("C2:C" & aLastRow & ", H2:H3") _
            .NumberFormat = "0.00%"

        'Conditonal formatting
        Dim pctChangeColorScale As ColorScale
        Set pctChangeColorScale = .Range("C2:C" & aLastRow) _
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
        
    End With
    
    ' Stop timer and display duration
    stopTime = Timer
    totalTime = Round(stopTime - startTime, 2)
    With AnalysisWS
        With .Range("F8")
            .Value = "Duration:"
            .Interior.Color = RGB(200, 200, 200)
            .Borders.Color = 0
        End With
        With .Range("F9")
            .Value = totalTime & " seconds"
            .Borders.Color = 0
        End With
    End With
    
End Sub

    
