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
' ----------------------------------
' Getters
' ----------------------------------
Property Get TickerID() As String
    TickerID = mStockTickerID
End Property

Property Get StockYear() As Integer
    StockYear = mStockYear
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
Property Let TickerID(ByVal theTickerID As String)
    mStockTickerID = theTickerID
End Property

Property Let StockYear(ByVal theYear As Integer)
    mStockYear = theYear
End Property
' ----------------------------------------
' Set initial values
' ----------------------------------------
Public Sub InitializeValues(ByVal theOpeningValue As Double, _
                              ByVal theClosingValue As Double, _
                              ByVal theDate As Long, _
                              ByVal theTotalVolume As Long)
    mStockInitialValue = theOpeningValue
    mStockInitialDate = theDate
    mStockFinalValue = theClosingValue
    mStockFinalDate = theDate
    mStockTotalVolume = theTotalVolume
End Sub
' ----------------------------------------
' Check whether values need to update
' ----------------------------------------
Public Sub UpdateValues(ByVal theOpeningValue As Double, _
                              ByVal theClosingValue As Double, _
                              ByVal theDate As Long, _
                              ByVal theTotalVolume As Long)
    If theDate < mStockInitialDate Then
        mStockInitialValue = theOpeningValue
        mStockInitialDate = theDate
    ElseIf theDate > mStockFinalDate Then
        mStockFinalValue = theClosingValue
        mStockFinalDate = theDate
    End If
    mStockTotalVolume = mStockTotalVolume + theTotalVolume
End Sub
' ----------------------------------
' Constructor
' ----------------------------------
Private Sub Class_Initialize()
    ' nothing to see here...
End Sub
