VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------
' Declare fields
' ----------------------------------
Private mStockTickerID As String
Private mStockYear As Integer
Private mStockInitialValue As Double
Private mStockInitialDate As Long
Private mStockFinalValue As Double
Private mStockFinalDate As Long
Private mStockTotalVolume As Variant
Private mStockPercentChange As Double

' ----------------------------------
' Getters
' ----------------------------------
Property Get TickerID() As String
    TickerID = mStockTickerID
End Property

Property Get StockYear() As Integer
    StockYear = mStockYear
End Property

Property Get InitialValue() As Double
    InitialValue = mStockInitialValue
End Property

Property Get InitialDate() As Long
    InitialDate = mStockInitialDate
End Property

Property Get FinalValue() As Double
    FinalValue = mStockFinalValue
End Property

Property Get FinalDate() As Long
    FinalDate = mStockFinalDate
End Property

Property Get TotalVolume() As Variant
    TotalVolume = mStockTotalVolume
End Property

Property Get YearlyChange() As Double
    YearlyChange = mStockFinalValue - mStockInitialValue
End Property

Property Get PercentChange() As Double
    If mStockInitialValue = 0 Then
        PercentChange = 0
    Else
        PercentChange = (mStockFinalValue - mStockInitialValue) / mStockInitialValue
    End If
End Property

' ----------------------------------
' Setters
' ----------------------------------
Property Let TickerID(ByVal newTickerID As String)
    mStockTickerID = newTickerID
End Property

Property Let StockYear(ByVal year As Integer)
    mStockYear = year
End Property

Property Let InitialValue(ByVal stockOpeningValue As Double)
    mStockInitialValue = stockOpeningValue
End Property

Property Let InitialDate(ByVal stockDate As Long)
    mStockInitialDate = stockDate
End Property

Property Let FinalValue(ByVal stockClosingValue As Double)
    mStockFinalValue = stockClosingValue
End Property

Property Let FinalDate(ByVal stockDate As Long)
    mStockFinalDate = stockDate
End Property

' ----------------------------------
' Increment total volume
' ----------------------------------
Public Sub IncTotalVolume(ByVal stockVolume As Long)
    mStockTotalVolume = mStockTotalVolume + stockVolume
End Sub

' ----------------------------------
' Constructor
' ----------------------------------
Private Sub Class_Initialize()
    ' nothing to see here...
End Sub
