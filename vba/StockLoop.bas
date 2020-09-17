Attribute VB_Name = "StockLoop"
Option Explicit

Sub StockLoop()

'   -----------------------------------
'   Create a collection
'   -----------------------------------
    Dim stocks As Collection
    Set stocks = New Collection
'   -----------------------------------
'   Dim reusable variables
'   -----------------------------------
    Dim stock As CStock
    Dim id As String
    Dim exists As Boolean
    Dim row As Long
    Dim i As Integer
    Dim ticker As String
    Dim yearlyChange As Long
    Dim percentChange As Double
'   -----------------------------------
'   Find the last row
'   -----------------------------------
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
'   -----------------------------------
'   Start looping
'   -----------------------------------
    For row = 2 To lastRow
        exists = False
        id = IDFromTickerAndDate(Cells(row, 1).Value, Cells(row, 2).Value)
        For Each stock In stocks
            If stock.TickerID = id Then
                exists = True
                If stock.InitialDate > Cells(row, 2).Value Then
                    stock.InitialDate = Cells(row, 2).Value
                    stock.InitialValue = Cells(row, 3).Value
                End If
                If stock.FinalDate < Cells(row, 2).Value Then
                    stock.FinalDate = Cells(row, 2).Value
                    stock.FinalValue = Cells(row, 6).Value
                End If
                stock.IncTotalVolume Cells(row, 7).Value
                Exit For
            End If
        Next stock
        If exists = False Then
            Set stock = New CStock
            stock.TickerID = id
            stock.InitialDate = Cells(row, 2).Value
            stock.FinalDate = Cells(row, 2).Value
            stock.InitialValue = Cells(row, 3).Value
            stock.FinalValue = Cells(row, 6).Value
            stock.IncTotalVolume Cells(row, 7).Value
            stocks.Add stock, id
        End If
    Next row

    For i = 1 To stocks.Count
        ticker = Split(stocks(i).TickerID, "_")(0)
        yearlyChange = stocks(i).FinalValue - stocks(i).InitialValue
        percentChange = yearlyChange / stocks(i).InitialValue
        
        Cells(i + 1, 9).Value = ticker
        Cells(i + 1, 10).Value = yearlyChange
        Cells(i + 1, 11).Value = percentChange
        Cells(i + 1, 12).Value = stocks(i).TotalVolume
    Next i

End Sub

' ------------------------------------
' Util to create ID from ticker/date
' ------------------------------------
Function IDFromTickerAndDate(ticker As String, fullDate As Long)
    Dim year As Integer
    year = (fullDate - fullDate Mod 10000) / 10000
    IDFromTickerAndDate = ticker & "_" & year
End Function


