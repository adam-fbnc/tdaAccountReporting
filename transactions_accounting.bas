Attribute VB_Name = "transactions_accounting"
Option Explicit

' Account related variables
Dim accountSymbols As New Dictionary
Dim accountInventoryFormulaStrings As New Dictionary 'All ending with "String" variables will be retrieved at the beginning and saved at the end of execution, updated just before from the Array
Dim accountProfitFormulaStrings As New Dictionary
Dim accountInventoryValues As New Dictionary  ' "Values" variables will be retrieved once at the beginning; there's no need to resave them
Dim accountProfitValues As New Dictionary
Dim accountInventoryArrays As New Dictionary  ' "Arrays" variables will be populated based on "String" variables above. Like "Values" they don't persist, only "Formula"s do
Dim accountProfitArrays As New Dictionary
Dim accountCashBalance As Double
Dim accountInterestEarned As Double
Dim accountRegFeesPaid As Double
Dim accountCommissionPaid As Double

' Transaction related variables
Dim trDate As Date
Dim trDescription As String
Dim trSymbol As String
Dim trSymbolAddress As Integer
Dim trCommission As Double
Dim trRegFee As Double
Dim trAmount As Double
Dim trPrice As Double
Dim trQuantity As Integer
Dim trType As String

' Variables for iterations
Dim iInventoryCost As String             ' numeric value stored as String
Dim iInventoryFormulaString As String    ' follows " = + 1 + 2.... " format
Dim iOldInventoryArray() As String
Dim iProfitOrLoss As String              ' numeric value stored as String
Dim iProfitOrLossFormulaString As String ' follows " = + 1 + 2.... " format
Dim iNumericProfitArray() As Double
Dim iPurchaseArray() As String
Dim iSaleArray() As String
Dim iNewInventoryArray() As String
Dim item As Variant

Sub bookTransaction()

Dim i As Integer
Dim j As Integer
Dim length As Integer

Dim transaction As Range
Dim trSymbolCell As Range
Dim invCell As Range
Dim profitCell As Range

'Dim inventoryString As String
'Dim oldInventoryValue As Double

'STEP 0 - Potentially, here we can import trading data as new sheet

' STEP 1 - Retrieve persistent data stored in spreadsheet
Call setPublicVariables

' STEP 2 - Get number of rows and transactions
ActiveWorkbook.Worksheets("Transactions").Activate
ActiveSheet.Range("A2").CurrentRegion.Select
length = Selection.Rows.count
'Debug.Print "How many rows will be processed? " & length
Debug.Print "Current cash balance " & accountCashBalance

' STEP 3 - Iterate through all transactions
' In this section accountInventoryFormulaStrings and accountProfitFormulaStrings will be written to worksheet at the end of the iterating loop
' Write a new Sub doing that

For i = 2 To 6 'length ' number of iterations
    Call setIterVariables(i)
    trType = getTransactionType(trDescription)
    Debug.Print vbCr & "----------------------------------------------------- Iteration Number " & (i - 1) & ": " & trType & " -----------------------------------------------------"

    ' If no trSymbol, then it's either cash transaction or internal transfer
'    Debug.Print "IGNORE THIS FOR NOW! " & trDescription
    If trSymbol = "" Then
        processCash trAmount, trDescription
        Debug.Print "Cash balance after '" & trDescription & "' transaction is " & accountCashBalance & ", and interest earned so far is " & accountInterestEarned
        GoTo NextIteration
    End If
    
    ReDim iPurchaseArray(1 To trQuantity)
    ReDim iSaleArray(1 To trQuantity)
    
    Debug.Print "Amount = " & trAmount & ". Existing cost of " & trSymbol & " inventory: " & iInventoryCost & " with accumulated profit / loss of " & iProfitOrLoss & "."
    Debug.Print "Old inventory from spreadsheet: " & iInventoryFormulaString
        
    ' If trAmount is negative (we spend cash) then book purchase
    ' The purchase should work well both with BUY and BUY TO COVER
    ' BUY TO COVER triggers profit calculation and can result in inventory changing sign
    If trAmount < 0 Then
    Debug.Print "Cash balance after purchase transaction " & accountCashBalance
        ' If it's regular BUY
        If trType = "BUY" Then
            Call regularBuy
        Else ' If it's BUY TO COVER
            Debug.Print "Processing BUY TO COVER"
            If iInventoryCost = "" Then
                Debug.Print "Inventory is empty. You cannot Buy To Cover, if there is no (negative) inventory"
                GoTo NextIteration
            End If
            ' Do buyToCover activities here
            ' Don't forget that like regular SELL, BTC generates profit or loss
        End If

        iOldInventoryArray() = convertStringToArray(iInventoryFormulaString)

    ' If trAmount is positive (we get cash) then book sale
    End If
    If trAmount > 0 Then
        
'        accountCashBalance = accountCashBalance - trRegFee
        accountRegFeesPaid = accountRegFeesPaid + trRegFee
        Debug.Print "Cash balance after sale transaction " & accountCashBalance & ". RegFee paid: " & trRegFee & ", and  RegFees paid so far: " & accountRegFeesPaid
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
        If trType = "SELL" Then
            ' Do SELL activities here
            If iInventoryCost = "" Then
                Debug.Print "Inventory is empty. You cannot perform regular SELL, if there is no (positive) inventory"
                GoTo NextIteration
            End If
            Call regularSell
            iProfitOrLoss = getNewProfitOrLoss
        Else
            ' Do SHORT SELL activities here
            ' This involves adding to negative inventory
        End If
    End If

    Debug.Print "Processed inventory array's size is: " & UBound(iNewInventoryArray)
    Debug.Print "Processed array converted to string: " & iInventoryFormulaString
    
' STEP N-1 - Save the inventory and P&L changes caused by transaction (1 iteration) to dictionary entry for that particular Symbol
    accountInventoryFormulaStrings(trSymbol) = iInventoryFormulaString
    accountInventoryValues(trSymbol) = iInventoryCost
    accountInventoryArrays(trSymbol) = convertStringToArray(iInventoryFormulaString)
    accountProfitFormulaStrings(trSymbol) = iProfitOrLossFormulaString
    accountProfitValues(trSymbol) = iProfitOrLoss
    Debug.Print "Processed inventory cost is: " & iInventoryCost
    
    Debug.Print "Profit realized for " & trSymbol & " so far: " & iProfitOrLoss
    Debug.Print "Profit / loss after sale as string: " & iProfitOrLossFormulaString
    
NextIteration:
Next

' STEP N - Save all changes in dictionaries to the "Inventory" sheet
' Write a Sub that iterates through Row 2 to 82 in that sheet and writes new values based on key
End Sub

Function shrinkArray(oldInventory() As String, saleArray() As String) As String()
Dim counter As Integer
counter = 1

' We should also consider 2 situations when opposite is true: when we have either positive or negative inventory
'If UBound(oldInventory) < 1 Then
'' If the the above condition is true, then basically our new combined inventory array will consist of negative values of sale array
'Call printArray(saleArray)
'Exit Function
'End If

Dim remainingInvArray() As String
Dim remainingInvArraySize As Integer
Dim soldInventoryArray() As String
Dim soldInventorySize As Integer
Dim oldArraySize As Integer
Dim saleArraySize As Integer
Dim sliceIndex As Integer
Dim minUnitCost As Double
Dim maxUnitCost As Double
Dim minProfitOrMinLossArray() As String

oldArraySize = UBound(oldInventory)
saleArraySize = UBound(saleArray)
soldInventorySize = saleArraySize
remainingInvArraySize = oldArraySize - saleArraySize + 1

Debug.Print "old inventory :" & oldArraySize + 1
Debug.Print "sold inventory :" & saleArraySize
Debug.Print "new inventory (remainingInvArraySize): " & remainingInvArraySize

ReDim remainingInvArray(1 To remainingInvArraySize)
ReDim iNumericProfitArray(1 To trQuantity)
ReDim soldInventoryArray(1 To soldInventorySize)

minUnitCost = CDbl(oldInventory(0))
maxUnitCost = CDbl(oldInventory(oldArraySize))

Debug.Print "minUnitCost with index 0: " & minUnitCost
Debug.Print "maxUnitCost with index " & oldArraySize & ": " & maxUnitCost
Debug.Print "salePrice: " & trPrice

'Step 1: Determine where sale price falls

If trPrice >= maxUnitCost Then
    ' All units sold at profit. Select 'quanitity' from the right to minimize total profit
        Debug.Print "Remaining inventory will iterate " & remainingInvArraySize & " times"
        remainingInvArray = sliceArray(oldInventory, 1, remainingInvArraySize)
        soldInventoryArray = sliceArray(oldInventory, remainingInvArraySize + 1, saleArraySize)
End If
If trPrice < minUnitCost Then
    ' All units sold at loss. Select 'quantity' from the left to minize total loss
        Debug.Print "Remaining inventory will iterate " & remainingInvArraySize & " times"
        remainingInvArray = sliceArray(oldInventory, saleArraySize + 1, remainingInvArraySize)
        soldInventoryArray = sliceArray(oldInventory, 1, saleArraySize)
End If
If trPrice < maxUnitCost And trPrice >= minUnitCost Then
    ' Units sold as well as profit and loss will depend on each individual case
    ' Calculate profit/loss and find minimal p/l from right to left and from left to right
    ' Write a function that will return an index if where inventory sale will start
    ' Underneath write another function that uses
    minProfitOrMinLossArray = minProfitMinLoss(iOldInventoryArray)
    ' 1- For least profit amount. 2- startFrom index for sliceArray function. 3- For least loss amount. 3- startFrom index
    If minProfitOrMinLossArray(1) > 0 Then
        sliceIndex = minProfitOrMinLossArray(2)
    Else
        sliceIndex = minProfitOrMinLossArray(4)
    End If
    Debug.Print "Slice index = " & sliceIndex
    soldInventoryArray = sliceArray(oldInventory, sliceIndex, saleArraySize)
    remainingInvArray = cutOutSoldArray(oldInventory, sliceIndex)
End If
    iNumericProfitArray = getArrayOfProfits(soldInventoryArray)
'   Debug.Print "Member count profit array: " & UBound(iNumericProfitArray)
    iInventoryFormulaString = convertArrayToString(remainingInvArray)
    Debug.Print "Remaining inventory as string :" & iInventoryFormulaString
    iProfitOrLossFormulaString = iProfitOrLoss & convertDoubleArrayToString(iNumericProfitArray)
    shrinkArray = remainingInvArray
    Debug.Print "Remaining inventory value " & sumCollection(getCollectionFromArray(remainingInvArray))
End Function

Function cutOutSoldArray(oldInvArray() As String, sliceIndex As Integer) As String()
Dim resultArray() As String
Dim r1Array() As String
Dim r2Array() As String
Dim counter As Integer
Dim size As Integer
Dim r1 As Integer
Dim r2 As Integer
Dim i1 As Integer
Dim i2 As Integer
r1 = sliceIndex - 1
r2 = r1 + trQuantity + 1
size = UBound(oldInvArray) - trQuantity + 1
ReDim resultArray(1 To size)
counter = 1

r1Array = sliceArray(oldInvArray, 1, r1)
r2Array = sliceArray(oldInvArray, r2, size - r1)
'Call printArray(r1Array, "Left array")
'Call printArray(r2Array, "Right array")

For Each item In r1Array
Debug.Print "Left array member with index " & counter & " is " & r1Array(counter) & " Which is the same as the item " & item
resultArray(counter) = CStr(item)
counter = counter + 1
Next

For Each item In r2Array
Debug.Print "Right array member with index " & counter - r1 & " is " & r2Array(counter - r1) & " Which is the same as the item " & item
resultArray(counter) = CStr(item)
counter = counter + 1
Next

cutOutSoldArray = resultArray

End Function

Function getArrayOfProfits(inventory() As String) As Double()
Dim profitsArray() As Double
Dim counter As Integer
Dim size As Integer
size = trQuantity
ReDim profitsArray(1 To size)
Debug.Print "Calculating profit for selling " & trQuantity & " units at sale price " & trPrice & " (amount: " & trAmount & ") and inventory cost of: " & convertArrayToString(inventory)
counter = 1
Do
    profitsArray(counter) = Round(trPrice - CDbl(inventory(counter)), 2)
    Debug.Print "Sold item with " & inventory(counter) & " cost. Profit from iteration " & counter & " is " & profitsArray(counter)
    counter = counter + 1
Loop Until counter = trQuantity + 1

getArrayOfProfits = profitsArray

End Function
Function mergeAndSortArray(oldInventory() As String, newPurchase() As String) As String()
Debug.Print "old Inv size " & UBound(oldInventory) + 1
If UBound(oldInventory) < 1 Then
    iInventoryFormulaString = convertArrayToString(newPurchase)
    Debug.Print "There was no inventory so we just added the purchase: " & iInventoryFormulaString
    mergeAndSortArray = newPurchase
Exit Function
End If

Dim combinedArray() As String
Dim numArray() As Double
Dim counter As Integer
Dim combinedArraySize As Integer
Dim oldArraySize As Integer
Dim newPurchaseArraySize As Integer
Dim firstExistingUnitPrice As Double         'Will change while iterating
Dim lastExistingUnitPrice As Double
Dim item As Variant
Dim coll As New Collection
Set coll = New Collection

oldArraySize = UBound(oldInventory)
newPurchaseArraySize = UBound(newPurchase)
Debug.Print "old: " & oldArraySize + 1
Debug.Print "new: " & newPurchaseArraySize

counter = 1
combinedArraySize = oldArraySize + newPurchaseArraySize
ReDim combinedArray(1 To combinedArraySize + 1)
ReDim numArray(1 To combinedArraySize + 1)

firstExistingUnitPrice = CDbl(oldInventory(0))
lastExistingUnitPrice = CDbl(oldInventory(oldArraySize))

Debug.Print "firstExistingUnitPrice with index 0: " & firstExistingUnitPrice
Debug.Print "newPurchasePrice: " & trPrice

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


'Set coll = sortCollection(coll)
'For Each item In coll
''Debug.Print "Sorted collection :" & item
'combinedArray(counter) = CStr(item)
'counter = counter + 1
'Next item

counter = 1
Call shellSortAscending(numArray, combinedArraySize + 1)
For Each item In numArray
'Debug.Print "Sorted collection :" & item
combinedArray(counter) = CStr(item)
counter = counter + 1
Next item
    iInventoryFormulaString = convertArrayToString(combinedArray)
    Debug.Print "Existing inventory after purchase as string :" & iInventoryFormulaString
    mergeAndSortArray = combinedArray
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


Function getArrayFromPriceAndQuantity(trPrice As Double, trQuantity As Integer) As String()
Dim arr() As String
Dim i As Integer
'Debug.Print "Quantity passed to getArrayFromPriceAndQuantity function: " & trQuantity
ReDim arr(1 To trQuantity)

For i = 1 To trQuantity
arr(i) = CStr(trPrice)
Next
getArrayFromPriceAndQuantity = arr
End Function

Function convertStringToArray(inventoryString As String) As String()
Dim arr() As String
inventoryString = Replace(inventoryString, "= +", "")
inventoryString = Replace(inventoryString, "= ", "")
inventoryString = Replace(inventoryString, "=", "")
arr() = Split(inventoryString, "+")

convertStringToArray = arr
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
If item <> "" Then
    convertArrayToString = convertArrayToString & symbol & item
'    Debug.Print "Convert array to string item: " & item
End If
Next

End Function

Function convertDoubleArrayToString(arr() As Double) As String
Dim returnString As String
Dim symbol As String
Dim counter As Integer
Dim item As Variant
convertDoubleArrayToString = ""
symbol = "+"
If Left(arr(1), 1) = "-" Then
symbol = ""
End If

For Each item In arr()
If item <> "" Then
    convertDoubleArrayToString = convertDoubleArrayToString & symbol & CStr(item)
'    Debug.Print "Convert array to string item: " & item
End If
Next

End Function

Sub regularBuy()
Dim existingInventoryValue As Double
Dim addedInventoryValue As Double
Debug.Print "Processing regular BUY"
iPurchaseArray = getArrayFromPriceAndQuantity(trPrice, trQuantity)
Debug.Print "New array size after purchase " & (UBound(iOldInventoryArray) + UBound(iPurchaseArray) + 1)
ReDim iNewInventoryArray(1 To (UBound(iOldInventoryArray) + UBound(iPurchaseArray)))
iNewInventoryArray = mergeAndSortArray(iOldInventoryArray, iPurchaseArray)
existingInventoryValue = CDbl(iInventoryCost)
addedInventoryValue = Round(trPrice * trQuantity, 2)
iInventoryCost = CStr(existingInventoryValue + addedInventoryValue)
End Sub

Sub regularSell()

iSaleArray = getArrayFromPriceAndQuantity(trPrice, trQuantity)
Debug.Print "New array size after sale " & (UBound(iOldInventoryArray) - UBound(iSaleArray) + 1)
ReDim iNewInventoryArray(1 To (UBound(iOldInventoryArray) - UBound(iSaleArray)) + 1)
iNewInventoryArray = shrinkArray(iOldInventoryArray, iSaleArray)
    
End Sub

Function getCollectionFromPriceAndQuantity(trPrice As Double, trQuantity As Integer) As Collection
Dim coll As New Collection
Dim i As Integer

For i = 1 To trQuantity
coll.Add (trPrice)
Next
Set getCollectionFromPriceAndQuantity = coll
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

Sub setPublicVariables()

' getInventoryProp returns inventory property based on Id that corresponds to column:
' 2 for inventory string, 3 for profit & loss string, 4 inventory value and 5 for P&L value. 0 gives row number
' Instead of using 5 different dictionaries that share the same key, explore using object with 3 different fields
Set accountSymbols = getInventoryProp(0)
Set accountInventoryFormulaStrings = getInventoryProp(2)
Set accountProfitFormulaStrings = getInventoryProp(3)
Set accountInventoryValues = getInventoryProp(4)
Set accountProfitValues = getInventoryProp(5)

accountCashBalance = Sheets("Inventory").Cells(1, 7).Value2
accountInterestEarned = Sheets("Inventory").Cells(1, 9).Value2
accountRegFeesPaid = Sheets("Inventory").Cells(1, 11).Value2
accountCommissionPaid = Sheets("Inventory").Cells(1, 13).Value2

Debug.Print "Number of Symbols added to dictionary: " & accountSymbols.count
End Sub
 
Sub setIterVariables(i As Integer)
    trSymbol = Cells(i, 5).Value2  ' SYMBOL column
    trAmount = Cells(i, 8).Value2
    trDescription = Cells(i, 3).Value2   ' DESCRIPTION column
    
    trDate = Cells(i, 1).Value ' DATE column
    trQuantity = Cells(i, 4).Value2   ' QUANTITY column
    trPrice = Cells(i, 6).Value2
    trRegFee = Cells(i, 9).Value2
    iInventoryCost = accountInventoryValues.item(trSymbol)
    iProfitOrLoss = accountProfitValues.item(trSymbol)
    iProfitOrLossFormulaString = accountProfitFormulaStrings.item(trSymbol)
    
    If iProfitOrLossFormulaString = "" Or iProfitOrLossFormulaString = "+" Or iProfitOrLossFormulaString = "-" Then
    iProfitOrLossFormulaString = "="
    End If
        accountCashBalance = accountCashBalance + trAmount
    trSymbolAddress = accountSymbols.item(trSymbol)
    iInventoryFormulaString = accountInventoryFormulaStrings.item(trSymbol)
    iOldInventoryArray = convertStringToArray(iInventoryFormulaString)
End Sub

Function getNewProfitOrLoss() As String
Dim sum As Double
sum = CDbl(iProfitOrLoss)
getNewProfitOrLoss = iProfitOrLoss
For Each item In iNumericProfitArray
'    Debug.Print "Items of iNumericProfitArray " & item
    sum = sum + Round(item, 2)
Next
getNewProfitOrLoss = CStr(sum)
End Function

Sub test()
trQuantity = 4
trAmount = 109 * trQuantity

Dim newColl As Collection
Dim test As String
Dim testArray() As String
Dim newArray() As String
Dim sumsArray() As Double
Dim numArray() As Double
'ReDim newArray(1 To trQuantity)
'Set newColl = New Collection

'    test = "= +105.19+106.20+107.31+108.42+109.53+110.64+111.75+112.86+113.91+115.08"
    test = "= +112.86+110.64+115.08+105.19+111.75+106.20+113.91+108.42+109.53+107.31+114.23"
    testArray = convertStringToArray(test)
    numArray = convertStringsToDoubles(testArray)
'    Call printArray(testArray)
'    For Each item In testArray
'    Debug.Print "Using array: " & item
'    Next item
'    Debug.Print "testArray size: " & UBound(testArray) + 1

'newArray = sliceArray(testArray, 3, trQuantity)
'Debug.Print "newArray size: " & UBound(newArray)
'Call printArray(newArray)

'Set newColl = getCollectionFromArray(newArray)
'Debug.Print "Size of collection " & newColl.count
'Call printCollection(newColl)
'Debug.Print "Sum of collection " & sumCollection(newColl)
Dim size As Integer
size = UBound(numArray)
Debug.Print "How many elements in the above array? - " & size

'sumsArray = getArrayOfSums(testArray)
'Call printArray(testArray, "Strings")
'Call printDoubleArray(numArray, "Unsorted")
Call shellSortAscending(numArray, size)
Call printDoubleArray(numArray, "Should be sorted")

End Sub


Private Sub shellSortAscending(ByRef a() As Double, N As Integer)
    ' Adapted from jdrago @ https://stackoverflow.com/questions/11504418/excel-vba-quickest-way-to-sort-an-array-of-numbers-in-descending-order
    ' requires a(1..N)
'    Debug.Print UBound(a)
    Debug.Assert LBound(a) = 1

    ' setup

    Dim i, j, inc As Integer
    Dim v As Double
    inc = 1

    ' determine the starting incriment

    Do
        inc = inc * 3
        inc = inc + 1
    Loop While inc <= N

    ' loop over the partial sorts

    Do
        inc = inc / 3

        ' Outer loop of straigh insertion

        For i = inc + 1 To N
            v = a(i)
            j = i

            ' Inner loop of straight insertion
            ' switch to a(j - inc) > v for ascending, replace sign to < for descending

            Do While a(j - inc) > v
                a(j) = a(j - inc)
                j = j - inc
                If j <= inc Then Exit Do
            Loop
            a(j) = v
        Next i
    Loop While inc > 1
End Sub


Function sliceArray(arr() As String, startFrom As Integer, count As Integer) As String()
'--SF-->>(SF+Q)-- Returns a slice of the Array that starts at startFrom and moves left -> trQuantity number of steps
' The trQuantity variable used for number of iterations is public
Dim arraySlice() As String
ReDim arraySlice(1 To count)
Dim counter As Integer
counter = 1

'Debug.Print "Quant " & trQuantity
'Debug.Print "Index " & startFrom

For counter = 1 To count
arraySlice(counter) = arr(counter + startFrom - 2)
'Debug.Print "sliceArray (" & counter & ") is " & arraySlice(counter)
Next


sliceArray = arraySlice

End Function

Function minProfitMinLoss(arr() As String) As String()
' Inputs are: 1- inventory array arr(), 2- transaction Amount trAmount and 3- transaction Quantity trQuantity
' Only inventory is specicified as parameter because, while inventory array is know, it may change
Dim result(1 To 4) As String
' 1- For least profit amount. 2- startFrom index for sliceArray function. 3- For least loss amount. 3- startFrom index
' If there's no profit / loss then amount should be set to 0
Debug.Print "Size of string array passed " & UBound(arr) + 1

Dim sumsArray() As Double
Dim profitsArray() As Double
Dim minLoss As Double
Dim minProfit As Double
Dim profitOrLoss As Double
Dim counter As Integer
'ReDim sumsArray(1 To UBound(sumsArray) + 1)
sumsArray = getArrayOfSums(arr)
Debug.Print "sumsArray size " & UBound(sumsArray)
counter = 1
minLoss = -sumsArray(UBound(sumsArray))
minProfit = trAmount

Dim tempColl As New Collection
Dim i As Integer


'Set coll = getCollectionFromArray(arr)

For Each item In sumsArray
    profitOrLoss = Round(trAmount - CDbl(item), 2)
    
'    Debug.Print "Sum " & counter & " of " & trQuantity & " items sold at " & trAmount & " is " & item & " and profit is " & profitOrLoss
    If profitOrLoss > 0 And profitOrLoss < minProfit Then
        minProfit = profitOrLoss
        result(2) = CStr(counter)
    End If
    
    If profitOrLoss < 0 And profitOrLoss > minLoss Then
        minLoss = profitOrLoss
        result(4) = CStr(counter)
    End If
    counter = counter + 1
Next

result(1) = CStr(minProfit)
result(3) = CStr(minLoss)
Call printArray(result, "Min/Max Profit & Loss")
minProfitMinLoss = result
End Function

Function getArrayOfSums(arr() As String) As Double()
Dim sumsArray() As Double
Dim size As Integer
Dim tempArray() As String
Dim tempColl As New Collection
Dim i As Integer
'Create array slices that trQuantity long

size = UBound(arr) - trQuantity + 2
ReDim sumsArray(1 To size)
Debug.Print "How many sums (from inside function)? - " & size

For i = 1 To size
    tempArray = sliceArray(arr, i, trQuantity)
    Set tempColl = getCollectionFromArray(tempArray)
    sumsArray(i) = sumCollection(tempColl)
'    Debug.Print "Sum of (" & i & ") of array Slice is " & sumsArray(i)
Next
getArrayOfSums = sumsArray
End Function
Function convertStringsToDoubles(arr() As String) As Double()
Dim numsArray() As Double
Dim size As Integer
Dim tempArray() As String
Dim tempColl As New Collection
Dim i As Integer
'Create array slices that trQuantity long

size = UBound(arr) + 1
ReDim numsArray(1 To size)
'Debug.Print "How many elements in array? - " & size
i = 1
For Each item In arr
numsArray(i) = CDbl(item)
i = i + 1
Next
convertStringsToDoubles = numsArray
End Function

Function getCollectionFromArray(arr() As String) As Collection
Dim coll As New Collection
Dim size As Integer
Dim i As Integer

size = UBound(arr)
'Debug.Print "Size of array passed " & size

For Each item In arr()
coll.Add (CStr(item))
Next

Set getCollectionFromArray = coll

End Function

Function sumCollection(coll As Collection) As Double
Dim sum As Double
Dim size As Integer
Dim i As Integer

sum = 0
size = coll.count
'Debug.Print "Size of collection passed " & size

For Each item In coll
sum = sum + Round(CDbl(item), 2)
'Debug.Print "Sum so far " & sum
Next

sumCollection = sum

End Function

Sub printArray(arr() As String, Optional comment As String)
Dim i As Integer
Dim note As String
If comment = "" Then
note = ""
Else
note = comment & " : "
End If
i = 1

    For Each item In arr
       Debug.Print note & "Member (" & i & ") of Array is " & item
       i = i + 1
    Next item
End Sub

Sub printDoubleArray(arr() As Double, Optional comment As String)
Dim i As Integer
Dim note As String
If comment = "" Then
note = ""
Else
note = comment & " : "
End If
i = 1

    For Each item In arr
       Debug.Print note & "Member (" & i & ") of Array is " & item
       i = i + 1
    Next item
End Sub


Sub printCollection(coll As Collection)
Dim i As Integer
i = 1

    For Each item In coll
       Debug.Print "Member (" & i & ") of Collection is " & item
       i = i + 1
    Next item
End Sub

Function getTransactionType(description As String) As String
Dim sellShort As Integer
Dim buyToCover As Integer
Dim normalSell As Integer
Dim normalBuy As Integer
sellShort = InStr(description, "Short")
buyToCover = InStr(description, "Cover")
normalBuy = InStr(description, "Bought")
normalSell = InStr(description, "Sold")

getTransactionType = "CASH & OTHER"

If normalBuy > 0 Then
    getTransactionType = "BUY"
    If buyToCover > 0 Then
        getTransactionType = "BUY TO COVER"
    End If
End If
If normalSell > 0 Then
    getTransactionType = "SELL"
    If sellShort > 0 Then
        getTransactionType = "SELL SHORT"
    End If
End If

End Function
