Attribute VB_Name = "accounting"
Option Explicit

' Account related variables
Dim accountSymbols As New Dictionary
Dim accountInventoryStrings As New Dictionary
Dim accountProfitStrings As New Dictionary
Dim accountInventoryValues As New Dictionary
Dim accountProfitValues As New Dictionary
Dim accountCashBalance As Double
Dim accountInterestEarned As Double
Dim accountRegFeesPaid As Double
'---- End Account related variables

Dim trAmount As Double
Dim trPrice As Double
Dim trQuantity As Integer

Sub bookTransaction()
' Transaction related variables, continued from above
Dim trDate As Date
Dim trDescription As String
Dim trSymbol As String
Dim trSymbolAddress As Integer
Dim trCommission As Double
Dim trRegFee As Double

Dim i As Integer
Dim length As Integer

Dim item As Variant
Dim sellShort As Integer
Dim buyToCover As Integer

Dim iInventoryCost As String             ' numeric values stored as String
Dim iInventoryFormulaString As String    ' follows " = + 1 + 2.... " format
Dim iProfitOrLoss As String              ' numeric values stored as String
Dim iPurchaseArray() As String
Dim iSaleArray() As String
Dim iOldInventoryArray() As String
Dim iNewInventoryArray() As String

Dim transaction As Range
Dim trSymbolCell As Range
Dim invCell As Range
Dim profitCell As Range

Dim test As String
Dim testArray() As String
'Dim inventoryString As String
'Dim oldInventoryValue As Double

'STEP 0 - Potentially, here we can import trading data as new sheet

'STEP 1 - Retrieve existing information

' getInventoryProp returns inventory property based on Id that corresponds to column:
' 2 for inventory string, 3 for profit & loss string, 4 inventory value and 5 for P&L value. 0 gives row number
' Instead of using 5 different dictionaries that share the same key, explore using object with 3 different fields
Set accountSymbols = getInventoryProp(0)
Set accountInventoryStrings = getInventoryProp(2)
Set accountProfitStrings = getInventoryProp(3)
Set accountInventoryValues = getInventoryProp(4)
Set accountProfitValues = getInventoryProp(5)

accountCashBalance = Sheets("Inventory").Cells(1, 7).Value2
accountInterestEarned = Sheets("Inventory").Cells(1, 9).Value2
accountRegFeesPaid = Sheets("Inventory").Cells(1, 11).Value2

Debug.Print "Number of Symbols added to dictionary :" & accountSymbols.count

'STEP 2 - Get number of rows and transactions
ActiveWorkbook.Worksheets("Transactions").Activate
ActiveSheet.Range("A2").CurrentRegion.Select
length = Selection.Rows.count
'Debug.Print "How many rows will be processed? " & length
Debug.Print "Current cash balance " & accountCashBalance

' STEP 3 - Iterate through all transactions
' In this section accountInventoryStrings and accountInventoryValues will be written to worksheet at the end of the iterating loop
' Write a new Sub doing that

For i = 2 To 5 'length
Debug.Print vbCr & "----------------------------------------------------- Iteration Number " & (i - 1) & " -----------------------------------------------------"
'    transaction = Range(Cells(i, 1), Cells(i, 10))
    
    trSymbol = Cells(i, 5).Value2  ' SYMBOL column
    trAmount = Cells(i, 8).Value2
    trDescription = Cells(i, 3).Value2   ' DESCRIPTION column
    
    ' If no trSymbol, then it's either cash transaction or internal transfer
    If trSymbol = "" Then
        processCash trAmount, trDescription
        Debug.Print "Cash balance after '" & trDescription & "' transaction is " & accountCashBalance & ", and interest earned so far is " & accountInterestEarned
        GoTo NextIteration
    End If
    
    trDate = Cells(i, 1).Value ' DATE column
    trQuantity = Cells(i, 4).Value2   ' QUANTITY column
    trPrice = Cells(i, 6).Value2
    trRegFee = Cells(i, 9).Value2
    iInventoryCost = accountInventoryValues.item(trSymbol)
    iProfitOrLoss = accountProfitValues.item(trSymbol)
'    sellShort = InStr(trDescription, "Short")
'    buyToCover = InStr(trDescription, "Cover")
'    Debug.Print "Is it a Short Sale? - " & (sellShort > 0)
'    Debug.Print "Is it a Buy To Cover? - " & (buyToCover > 0)
'    Debug.Print "Short Sale? - " & sellShort
'    Debug.Print "Buy To Cover? - " & buyToCover
'    Debug.Print "i = " & i:      Debug.Print trDate:     Debug.Print trQuantity:     Debug.Print trPrice:     Debug.Print trRegFee
'    Debug.Print "Symbol = " & trSymbol
'    Debug.Print "The row " & trSymbol & " occupies is " & trSymbolAddress

    
    Debug.Print "Amount = " & trAmount & ". Existing cost of " & trSymbol & " inventory: " & iInventoryCost & " with accumulated profit / loss of " & iProfitOrLoss & "."
    
    accountCashBalance = accountCashBalance + trAmount
    trSymbolAddress = accountSymbols.item(trSymbol)
    iInventoryFormulaString = accountInventoryStrings.item(trSymbol)
    iOldInventoryArray = convertToArray(iInventoryFormulaString)
    Debug.Print "Old inventory from spreadsheet : " & iInventoryFormulaString
        
    ' If trAmount is negative (we spend cash) then book purchase
    ' The purchase should work well both with BUY and BUY TO COVER
    ' BUY TO COVER triggers profit calculation and can result in inventory changing sign
    If trAmount < 0 Then
    
    Debug.Print "Cash after purchase transaction " & accountCashBalance
'    Debug.Print "Inventory purchase string from getExpanded function : " & getExpanded(trPrice, trQuantity)
        ' If it's regular BUY
        If buyToCover = 0 Then
        Debug.Print "Processing regular BUY"
'            iProfitOrLoss = accountProfitStrings.item(trSymbol)
'            If iInventoryFormulaString = "" Then
'            iInventoryFormulaString = "= "
'            End If
            iPurchaseArray = getArray(trPrice, trQuantity)
            ReDim iNewInventoryArray(1 To (UBound(iOldInventoryArray) + UBound(iPurchaseArray)))
            iNewInventoryArray = mergeAndSortArray(iOldInventoryArray, iPurchaseArray)
        ' If it's BUY TO COVER
            Else
            Debug.Print "Processing BUY TO COVER"
            If iInventoryCost = "" Then
                Debug.Print "Inventory is empty. You cannot Buy To Cover, if there is no (negative) inventory"
                GoTo NextIteration
            End If
            ' Do buyToCover activities here
            ' Don't forget that like regular SELL, BTC generates profit or loss
        End If

'        buyPosition trSymbol, trPrice, trQuantity
'        iOldInventoryArray() = convertToArray(iInventoryFormulaString)

    ' If trAmount is positive (we get cash) then book sale
    ElseIf trAmount > 0 Then
    
    accountCashBalance = accountCashBalance - trRegFee
    accountRegFeesPaid = accountRegFeesPaid + trRegFee
    Debug.Print "Cash after sale transaction " & accountCashBalance & ". RegFee paid: " & trRegFee & ", and  RegFees paid so far: " & accountRegFeesPaid
    ' This should seemlessly process the sales in the following situations:
    ' 1) With regards to profit & loss:
    '   a) Sale when all inventory items result in profit
    '   b) When all inventory items result in a loss
    '   c) When some generate profit and some loss
    ' 2) With respect to inventory
    '   a) Positive inventory at the beginning, positive inventory at the end
    '       i) Also, how do record a sale at a loss? If there's no inventory, either positive or negative,
    '           the difference in sale and carrying cost goes to P&L
    '   b) Positive inventory ends up with negative inventory
    '   c) Negative inventory grows more negative
        ' If it's regular SELL
        If sellShort = 0 Then
            ' Do SELL activities here
            iSaleArray = getArray(trPrice, trQuantity)
            ReDim iNewInventoryArray(1 To (UBound(iOldInventoryArray) - UBound(iSaleArray)))
            iNewInventoryArray = shrinkArray(iOldInventoryArray, iSaleArray)
            If iInventoryCost = "" Then
                Debug.Print "Inventory is empty. You cannot regular Sell, if there is no (positive) inventory"
                GoTo NextIteration
            End If
            
        Else
            ' Do SHORT SELL activities here
            ' This involves adding to negative inventory
        End If
    End If
    
'    test = "= +115.19+115.20+115.31+115.42+115.53+115.64+115.75+115.86+115.91+115.98"
'    testArray = convertToArray(test)
'    For Each item In testArray
'    Debug.Print "Using array: " & item
'    Next item
'    Debug.Print "testArray size: " & UBound(testArray)
'    iOldInventoryArray(1) = 115.63
'    iOldInventoryArray(2) = 115.63

    
'    For Each item In iOldInventoryArray
'        Debug.Print "iOldInventoryArray item : " & item
'    Next item
'    For Each item In iPurchaseArray
'       Debug.Print "New purchase: " & item
'    Next item
'    For Each item In iNewInventoryArray
'       Debug.Print "Combined array: " & item
'    Next item
    
    Debug.Print "Processed inventory array's size is: " & UBound(iNewInventoryArray)
    Debug.Print "Processed array converted to string: " & convertArrayToString(iNewInventoryArray)
    
NextIteration:
Next


End Sub


Function mergeAndSortArray(oldInventory() As String, newPurchase() As String) As String()

If UBound(oldInventory) < 1 Then
mergeAndSortArray = newPurchase
Exit Function
End If

'If oldInventory(1) = "" Or oldInventory(1) = "=" Or oldInventory(1) = "= " Then

Dim combinedArray() As String
Dim numArray() As Double
Dim counter As Integer
Dim combinedArraySize As Integer
Dim oldArraySize As Integer
Dim newArraySize As Integer
Dim unitPrice As Double
Dim existingUnitPrice As Double         'Will change while iterating
Dim lastExistingUnitPrice As Double
Dim newPurchasePrice As Double          'This one is fixed
Dim item As Variant
Dim coll As New Collection
Set coll = New Collection

oldArraySize = UBound(oldInventory)
newArraySize = UBound(newPurchase)
Debug.Print "old :" & oldArraySize
Debug.Print "new :" & newArraySize

counter = 1
combinedArraySize = oldArraySize + newArraySize
ReDim combinedArray(1 To combinedArraySize + 1)
ReDim numArray(1 To combinedArraySize + 1)

newPurchasePrice = CDbl(newPurchase(1))
existingUnitPrice = CDbl(oldInventory(0))
lastExistingUnitPrice = CDbl(oldInventory(oldArraySize))

Debug.Print "existingUnitPrice with index 0: " & existingUnitPrice
Debug.Print "newPurchasePrice: " & newPurchasePrice

'Step 1: Splice 2 arrays
    'First array
For Each item In newPurchase
'    Debug.Print "new unitPrice " & item
    numArray(counter) = CDbl(item)
    coll.Add item
    counter = counter + 1
Next item
    'Second array
For Each item In oldInventory
'Debug.Print "old unitPrice " & item
    numArray(counter) = CDbl(item)
    coll.Add item
    counter = counter + 1
Next item

'Step 2: Sort the numeric array

'Debug.Print "unitPrice" & unitPrice
'For Each item In coll
'Debug.Print "Unsorted collection :" & item
'Next item

counter = 1
Set coll = sortCollection(coll)
For Each item In coll
'Debug.Print "Sorted collection :" & item
combinedArray(counter) = CStr(item)
counter = counter + 1
Next item



    mergeAndSortArray = combinedArray
End Function

Function shrinkArray(oldInventory() As String, saleArray() As String) As String()
Dim item As Variant
Dim counter As Integer
counter = 1

' We should also consider 2 situations when opposite is true: when we have either positive or negative inventory
If UBound(oldInventory) < 1 Then
' If the below condition is true, then basically our new combined inventory array will consist of negative values of sale
For Each item In saleArray
'    Debug.Print "new unitPrice " & item
    shrinkArray(counter) = "-" & item
    counter = counter + 1
Next item
Exit Function
End If

'If oldInventory(1) = "" Or oldInventory(1) = "=" Or oldInventory(1) = "= " Then

Dim smallerArray() As String
Dim numArray() As Double
Dim smallerArraySize As Integer
Dim oldArraySize As Integer
Dim newArraySize As Integer
Dim unitPrice As Double
Dim minUnitCost As Double
Dim maxUnitCost As Double
Dim saleArrayPrice As Double
'Dim coll As New Collection
'Set coll = New Collection

oldArraySize = UBound(oldInventory)
saleArraySize = UBound(saleArray)
Debug.Print "old inventory :" & oldArraySize
Debug.Print "sale inventory :" & saleArraySize

smallerArraySize = oldArraySize - saleArraySize
ReDim smallerArray(1 To smallerArraySize + 1)
ReDim numArray(1 To smallerArraySize + 1)

saleArrayPrice = CDbl(saleArray(1))
minUnitCost = CDbl(oldInventory(0))
maxUnitCost = CDbl(oldInventory(oldArraySize))

Debug.Print "minUnitCost with index 0: " & minUnitCost
Debug.Print "saleArrayPrice: " & saleArrayPrice

'Step 1: Determine where sale price falls

counter = 1
If saleArrayPrice >= maxUnitCost Then
    ' All units sold at profit. Select 'quanitity' from the end to minize total profit
    ' Two loops: one to create new array after sale, and another to calculate profit
        Do
            shrinkArray(counter) = oldInventory(counter + trQuantity)
            counter = counter + 1
        Loop Until counter = smallerArraySize
        
        counter = 1
        Do
            shrinkArray(counter) = oldInventory(counter + trQuantity)
            counter = counter + 1
        Loop Until counter = smallerArraySize
ElseIf saleArrayPrice < minUnitCost Then
    ' All units sold at loss. Select 'quanitity' from the start to minize total loss
    ' Two loops: one to create new array after sale, and another to calculate loss
Else
    ' Units sold as well as profit and loss will depend on each individual case

counter = 1
    'First array
For Each item In saleArray
'    Debug.Print "new unitPrice " & item
    numArray(counter) = CDbl(item)
    coll.Add item
    counter = counter + 1
Next item
    'Second array
For Each item In oldInventory
'Debug.Print "old unitPrice " & item
    numArray(counter) = CDbl(item)
    coll.Add item
    counter = counter + 1
Next item

'Step 2: Sort the numeric array

'Debug.Print "unitPrice" & unitPrice
'For Each item In coll
'Debug.Print "Unsorted collection :" & item
'Next item

counter = 1
Set coll = sortCollection(coll)
For Each item In coll
'Debug.Print "Sorted collection :" & item
smallerArray(counter) = CStr(item)
counter = counter + 1
Next item



    shrinkArray = smallerArray
End Function

Function sortCollection(coll As Collection) As Collection

    Dim newColl As Collection
    Dim vItm As Variant, item As Variant
    Dim i As Long, j As Long
    Dim vTemp As Variant

    Set newColl = New Collection

    'fill the collection
    For Each item In coll
    newColl.Add item
    Next

    'Two loops to bubble sort
    For i = 1 To newColl.count - 1
        For j = i + 1 To newColl.count
            If CDbl(newColl(i)) > CDbl(newColl(j)) Then
                'store the lesser item
                vTemp = newColl(j)
                'remove the lesser item
                newColl.Remove j
                're-add the lesser item before the
                'greater Item
                newColl.Add vTemp, vTemp, i
            End If
        Next j
    Next i

    'Test it
'    For Each vItm In newColl
'        Debug.Print vItm
'    Next vItm
Set sortCollection = newColl
End Function




Function getArray(trPrice As Double, trQuantity As Integer) As String()
Dim arr() As String
Dim i As Integer
Debug.Print "Quantity passed to getArray function: " & trQuantity
ReDim arr(1 To trQuantity)

For i = 1 To trQuantity
arr(i) = CStr(trPrice)
Next
getArray = arr
End Function

Function convertToArray(inventoryString As String) As String()
Dim arr() As String
inventoryString = Replace(inventoryString, "= +", "")
inventoryString = Replace(inventoryString, "= ", "")
inventoryString = Replace(inventoryString, "=", "")
arr() = Split(inventoryString, "+")

convertToArray = arr
End Function

Function getInventoryProp(column As Integer) As Dictionary
' Property Id corresponds to column: 2 for inventory and 3 for profit & loss. 0 gives row number
Dim dict As New Dictionary
Dim trSymbol As String
Dim property As String
Dim counter As Integer
counter = 2
'Debug.Print "Column " & column
If column = 2 Or column = 3 Then
    Do
        trSymbol = Sheets("Inventory").Cells(counter, 1).Value2
        If trSymbol = "" Then
            Exit Do 'We reached the end of the list
        End If
        property = Sheets("Inventory").Cells(counter, column).Formula
        dict.Add trSymbol, property
'        Debug.Print "Column " & column & ". Iteration " & (counter - 1) & " " & trSymbol
        counter = counter + 1
        Loop Until counter = 120
    
Else
    Do
        trSymbol = Sheets("Inventory").Cells(counter, 1).Value2
        If trSymbol = "" Then
            Exit Do 'We reached the end of the list
        End If
'        Debug.Print "Column " & column & ". Iteration " & (counter - 1) & " " & trSymbol
        If column = 0 Then
            property = CStr(counter)
            ElseIf column = 4 Or column = 5 Then
            property = Sheets("Inventory").Cells(counter, column - 2).Value2
        End If
        dict.Add trSymbol, property
        counter = counter + 1
        Loop Until counter = 120
End If
Set getInventoryProp = dict
End Function
Function findSymbol(target As String) As Integer
Dim trSymbol As String
Dim counter As Integer
counter = 2

    Do
        trSymbol = Sheets("Inventory").Cells(counter, 1).Value2
'        Debug.Print "Iteration " & (counter - 1) & " " & trSymbol
        If trSymbol = target Then
            findSymbol = counter
            Exit Do
        End If
        counter = counter + 1
    Loop Until counter = 120

End Function


Function convertArrayToString(arr() As String) As String
Dim returnString As String
Dim symbol As String
Dim counter As Integer
Dim item As Variant
convertArrayToString = "= "
symbol = "+"
If Left(arr(1), 1) = "-" Then
symbol = "-"
End If

For Each item In arr()
convertArrayToString = convertArrayToString & symbol & item
Next

End Function

Function buyPosition(trSymbol As String, trPrice As Double, trQuantity As Integer) As String

Dim transUnits As New Collection
Dim newPuchaseString As String
newPuchaseString = ""

Set transUnits = getCollection(trPrice, trQuantity)
    Debug.Print trans.count
    
    For Each item In trans
    Debug.Print "This is item: " & item
    Next item


    
End Function

Function getCollection(trPrice As Double, trQuantity As Integer) As Collection
Dim coll As New Collection
Dim i As Integer

For i = 1 To trQuantity
coll.Add (trPrice)
Next
Set getCollection = coll
End Function

Sub processCash(trAmount As Double, trDescription As String)
Select Case trDescription
Case "ELECTRONIC NEW ACCOUNT FUNDING"
    accountCashBalance = accountCashBalance + trAmount
Case "CLIENT REQUESTED ELECTRONIC FUNDING RECEIPT (FUNDS NOW)"
    accountCashBalance = accountCashBalance + trAmount
Case "PERSONAL CHECK RECEIPT"
    accountCashBalance = accountCashBalance + trAmount
Case "MARGIN INTEREST ADJUSTMENT"
    accountInterestEarned = accountInterestEarned + trAmount
    accountCashBalance = accountCashBalance + trAmount
Case "FREE BALANCE INTEREST ADJUSTMENT"
    accountInterestEarned = accountInterestEarned + trAmount
    accountCashBalance = accountCashBalance + trAmount
Case "CLIENT REQUESTED ELECTRONIC FUNDING DISBURSEMENT (FUNDS NOW)"
    accountCashBalance = accountCashBalance - trAmount
End Select

End Sub

Function getExpanded(trPrice As Double, trQuantity As Integer) As String
Dim expanded As String
Dim i As Integer
expanded = ""

For i = 1 To trQuantity
expanded = expanded & "+" & trPrice
Next

getExpanded = expanded
End Function

