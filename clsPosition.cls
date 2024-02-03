VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public symbol As String '1
Public costFormulaString As String '2
Public profitFormulaString As String '3
Public salesFormulaString As String '4
Public cogsFormulaString As String '5
Public quantity As Integer '6 --------------------
Public cost As Currency '7
Public profit As Currency '8
Public sales As Currency '9
Public cogs As Currency '10
Public symbolAddress As Integer '11

Function getPosition(counter As Integer)
symbol = Sheets("Inventory").Cells(counter, 1).Value2  ' SYMBOL column
costFormulaString = Sheets("Inventory").Cells(counter, 2).Formula
profitFormulaString = Sheets("Inventory").Cells(counter, 3).Formula
salesFormulaString = Sheets("Inventory").Cells(counter, 4).Formula
cogsFormulaString = Sheets("Inventory").Cells(counter, 5).Formula
quantity = Sheets("Inventory").Cells(counter, 6).Value2
cost = Sheets("Inventory").Cells(counter, 2).Value2
profit = Sheets("Inventory").Cells(counter, 3).Value2
sales = Sheets("Inventory").Cells(counter, 4).Value2
cogs = Sheets("Inventory").Cells(counter, 5).Value2
symbolAddress = counter

End Function

Function getProperty(prop As String) As Variant
Dim property As Variant

Select Case prop
Case "symbol"
    property = symbol
Case "costFormulaString"
    property = costFormulaString
Case "profitFormulaString"
    property = profitFormulaString
Case "salesFormulaString"
    property = salesFormulaString
Case "cogsFormulaString"
    property = cogsFormulaString
Case "costFormula"
    property = costFormulaString
Case "profitFormula"
    property = profitFormulaString
Case "salesFormula"
    property = salesFormulaString
Case "cogsFormula"
    property = cogsFormulaString
Case "quantity"
    property = quantity
Case "profit"
    property = profit
Case "sales"
    property = sales
Case "cogs"
    property = cogs
Case "symbolAddress"
    property = symbolAddress

End Select
getProperty = property
End Function

Sub printPosProperties(Optional comment As String)
Debug.Print vbCr & "-----------------------------------------------------"
Debug.Print comment & " Symbol: " & symbol
Debug.Print comment & " Cost formula: " & costFormulaString
Debug.Print comment & " Profit formula: " & profitFormulaString
Debug.Print comment & " Sales formula: " & salesFormulaString
Debug.Print comment & " COGS formula: " & cogsFormulaString
Debug.Print comment & " Quantity: " & quantity
Debug.Print comment & " Cost value: " & cost
Debug.Print comment & " Profit value: " & profit
Debug.Print comment & " Sales value: " & sales
Debug.Print comment & " COGS value: " & cogs
Debug.Print comment & " Symbol address: " & symbolAddress
End Sub
