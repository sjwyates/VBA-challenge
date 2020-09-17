VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------
' Declare Fields
' ----------------------------------
Private mStockTickerID As String
Private mStockYear As Integer
Private mStockInitialValue As Double
Private mStockInitialDate As Integer
Private mStockFinalValue As Double
Private mStockFinalDate As Integer
Private mStockTotalVolume As Long

' ----------------------------------
' Getters
' ----------------------------------
Property Get TickerID() As String
    TickerID = mStockTickerID
End Property

Property Get Year() As Integer
    Year = mStockYear
End Property

Property Get InitialValue() As Double
    InitialValue = mStockInitialValue
End Property

Property Get InitialDate() As Integer
    InitialDate = mStockInitialDate
End Property

Property Get FinalValue() As Double
    FinalValue = mStockFinalValue
End Property

Property Get FinalDate() As Integer
    FinalDate = mStockFinalDate
End Property

Property Get TotalVolume() As Long
    TotalVolume = mStockTotalVolume
End Property

' ----------------------------------
' Setters
' ----------------------------------
Property Let Year(ByVal stockDate As Integer)
    ' Prevent year being overwritten if already set
    If mStockYear <> 0 Then
        mStockYear = YearFromDate(stockDate)
    End If
End Property

Property Let InitialValue(ByVal stockOpeningValue As Double)
    mStockInitialValue = stockOpeningValue
End Property

Property Let InitialDate(ByVal stockDate As Integer)
    mStockInitialDate = stockDate
End Property

Property Let FinalValue(ByVal stockClosingValue As Double)
    mStockFinalValue = stockClosingValue
End Property

Property Let FinalDate(ByVal stockDate As Integer)
    mStockFinalDate = stockDate
End Property

' ----------------------------------
' Total Volume Incrementer
' ----------------------------------
Public Sub IncTotalVolume(ByVal stockVolume As Long)
    mStockTotalVolume = mStockTotalVolume + stockVolume
End Sub

' ----------------------------------
' Get Year From Date Function
' ----------------------------------
Private Function YearFromDate(fullDate As Integer)
    YearFromDate = (fullDate - fullDate Mod 10000) / 10000
End Function


' ----------------------------------
' Constructor
' ----------------------------------
Private Sub Class_Initialize(ByVal ticker As String)
    stockTickerID = ticker
    mStockYear = 0
    mStockTotalVolume = 0
End Sub
