/* ---------------------------------------------------------------------------
 *  VBA program
 *  
 *  Author: Thanh Tran
 *          thanh.t.tran1301@gmail.com
 *  --------------------------------------------------------------------------- 
 */

Sub Main1()
UserForm1.Show
'this sub will calculate the price given_
'the spot price, the volatility and risk Free Rate
'since the project is not completed, I set those variable to be constant
'but it would not be a problem to pull out at the end


Dim spotPrice As Double, volatility As Double, riskFreeRate As Double
Dim i As Integer, j As Integer
'declare a array for the price
Dim oilPriceArray() As Double
'temporarily set the number of iteration to be 10000
ReDim oilPriceArray(1 To 10000)
Dim k As Integer
Dim m As Integer
Dim priceRange As Range
Dim graphTittle As String, xTitle As String, yTittle As String
Dim graphPos As Integer 'the position to place the graph

graphPos = 400


'm=2 indicates the second row of the oil price column
m = 2
'temporarily set these variables to be constant
spotPrice = Cells(4, 2).Value
volatility = Cells(5, 2).Value
riskFreeRate = Cells(9, 2).Value
Cells(11, 6).Value = Cells(4, 2).Value
j = 0
'repeat 10000times
'call the oil price generator sub
'put all the values in the K column
For i = 1 To 10000
    oilPriceArray(i) = oilPriceGen(spotPrice, volatility, riskFreeRate)
    Cells(m, 11).Value = oilPriceArray(i)
    m = m + 1
    spotPrice = 678
Next

'sort the array
Call sort(oilPriceArray())
'find all the breaks and the frequency of the array
Call hist(10000, 100, oilPriceArray(1), oilPriceArray(10000), oilPriceArray())

'graph the oil price distribution
Set priceRange = Range(Cells(1, 8), Cells(101, 9))
graphTittle = "Oil Price"
xTitle = "Price per Tonne ($)"
yTittle = "Frequency"


Call graph(priceRange, graphTittle, xTitle, yTittle, graphPos)

    
End Sub
Sub sort(oilPriceArray() As Double)
'source : http://www.mrexcel.com/forum/excel-questions/690718-vi_
'sual-basic-applications-sort-array-numbers.html

Dim i As Integer
Dim j As Integer
Dim SrtTemp As Variant
 
 'compare the ith item with the (i-1)th item in the array
 

 For i = LBound(oilPriceArray) To UBound(oilPriceArray)
   For j = i + 1 To UBound(oilPriceArray)
             If oilPriceArray(i) > oilPriceArray(j) Then
                 SrtTemp = oilPriceArray(j)
                 oilPriceArray(j) = oilPriceArray(i)
                 oilPriceArray(i) = SrtTemp
             End If
         Next j
     Next i

 
  End Sub

Function oilPriceGen(spotPrice, volatility, riskFreeRate)
'this is the function to calculate new oil price

Dim genOilPrice As Double, expectedReturn As Double, noise As Double
Dim change As Double 'percentage change in oil price
Dim i As Integer
Dim j As Integer
Dim sum As Double
Dim average As Double ' average of oil price over the 12 periods
Dim k As Integer


j = 12
expectedReturn = (riskFreeRate / 100 - 0.5 * (volatilty / 100) ^ 2) * (1 / 12) * 100
sum = 0



'there are 12 periods so it will loop 12 times
'also calcualtes the sum over 12 periods
For i = 1 To 12
    noise = ((volatility / 100) * Sqr(1 / 12) * normRandNum()) * 100
    change = Exp(expectedReturn / 100 + noise / 100) * 100
    genOilPrice = spotPrice / 100 * change
    sum = sum + genOilPrice
    spotPrice = genOilPrice
Next
   

'calculate the average

average = sum / 12
oilPriceGen = average

End Function


Sub hist(n As Variant, m As Long, Start As Double, Right As Double, arr() As Double)
    'this sub is used to find all the breaks and the frequency in that break
    'Lenth: the length of one break to the subsequent
    'm is the number of breaks
    
    Dim i As Long, j As Long
    Dim Length As Double
    ReDim breaks(m) As Single
    ReDim freq(m) As Single
    
    'set all the elements in the range to be 0
    For i = 1 To m
        freq(i) = 0
    Next i

    Length = (Right - Start) / m
    'find the value for all the breaks
    For i = 1 To m
        breaks(i) = Start + Length * i
    Next i
    'compare each element in array to the break
    ' it is less than the break, it will be put in that break
    For i = 1 To n
        If (arr(i) <= breaks(1)) Then freq(1) = freq(1) + 1
        If (arr(i) >= breaks(m - 1)) Then freq(m) = freq(m) + 1
        For j = 2 To m - 1
            If (arr(i) > breaks(j - 1) And arr(i) <= breaks(j)) Then freq(j) = freq(j) + 1
        Next j
    Next i
    'put the values in 2 columns
    For i = 1 To m
        Cells(i + 1, 8) = breaks(i)
        Cells(i + 1, 9) = freq(i)
    Next i
End Sub

Function normRandNum()
' Box-Muller Method
'genrate a random number with u=0 and SD=1
Dim fac As Double, r As Double, V1 As Double, V2 As Double
repeat: V1 = 2 * Rnd - 1
        V2 = 2 * Rnd - 1
        r = V1 ^ 2 + V2 ^ 2
        If (r >= 1) Then GoTo repeat
        fac = Sqr(-2 * Log(r) / r)
        normRandNum = V2 * fac
End Function
Sub Main2()
'this sub will calculate the net income
' the user will put in the price elasticity of demand
' again I set those endogenous variables to be constant
' they will be pulled out when the project is completed
UserForm2.Show

Dim newTicketDemand As Double, newFuelExpense As Double
Dim netIncome() As Double
Dim operRevenue As Double, expense1 As Double, expense2 As Double
Dim totalExpense As Double, tax As Double
Dim oilPrice() As Double
ReDim oilPrice(1 To 10000)
Dim demandChange As Double
Dim i As Integer, j As Integer, k As Integer
ReDim netIncome(1 To 10000)
Dim elasticity As Double
Dim r As Integer
Dim incomeRange As Range
Dim graphTittle As String, xTitle As String, yTittle As String
Dim graphPos As Integer

graphPos = 600


r = 2
'get the range of oil price in sheet 1
Sheets(1).Activate
k = 2
For j = 1 To 10000
oilPrice(j) = Cells(k, 11).Value
k = k + 1
Next
spotPrice = Cells(4, 2).Value
'reactivate sheet 2
Sheets(2).Activate
elasticity = Cells(1, 20).Value



'calculate the expense and revenue on the income statement
operRevenue = WorksheetFunction.sum(Range("B7:B8"))
expense1 = WorksheetFunction.sum(Range("B14:B20"))
expense2 = WorksheetFunction.sum(Range("B27:B30"))
tax = Cells(34, 2).Value


For i = 1 To 10000
'change in demand will affect tick demand and fuel purchased fee
    newFuelExpense = Cells(13, 2).Value / spotPrice * oilPrice(i)
    demandChange = (oilPrice(i) - spotPrice) / spotPrice * 100 * elasticity
    newTicketDemand = Cells(6, 2).Value / 100 * (demandChange + 100)
    netIncome(i) = operRevenue + newTicketDemand - expense1 - expense2 - tax - newFuelExpense
    Cells(r, 11).Value = netIncome(i)
    r = r + 1
Next
' call the sort and hist subs
Call sort(netIncome())
Call hist(10000, 200, netIncome(1), netIncome(10000), netIncome())

' graph the distribution of the net income
Set incomeRange = Range(Cells(1, 8), Cells(201, 9))
graphTittle = "Income"
xTittle = "Net income in millions dollars ($)"
yTittle = "Frequency"
Call graph(incomeRange, graphTittle, xTitle, yTittle, graphPos)

End Sub


Sub graph(priceRange As Range, graphTittle As String, xTittle As String, yTittle As String, graphPos As Integer)
   'graph the distribution
    On Error Resume Next
    ActiveSheet.ChartObjects.Delete
    ActiveWindow.SmallScroll Down:=5
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.SetSourceData Source:=priceRange
    With ActiveChart
    .HasTitle = True
    .ChartTitle.Characters.Text = graphTittle
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = xTittle
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = yTittle
    .Parent.Top = graphPos
    .Parent.Left = 0
End With
End Sub

Sub getData()
'get data from Yaho Finace through Access
Sheets(4).Cells.Clear
With Sheets(4).QueryTables.Add(Connection:= _
        "URL;http://finance.yahoo.com/q/fc?s=CLM15.NYM+Futures+Chain", Destination:= _
        Sheets(4).Range("$A$1"))
         .Name = "fc?s=CLM15.NYM+Futures+Chain"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "10"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub getFuturePrice(futurePriceArray() As Double)
'put teh future price in an array
Sheets(4).Activate

Dim i As Double
i = 1
ReDim futurePriceArray(1 To 16)
For i = 1 To 16
futurePriceArray(i) = Left(Cells(i + 2, 3).Value, 5)
Next

End Sub


Sub monthlyOilGen(averMonthlyOilPrice() As Double)

'generate monthly future prices

Dim spotPrice As Double, volatility As Double, riskFreeRate As Double
Dim genOilPrice As Double, expectedReturn As Double, noise As Double
Dim change As Double 'percentage change in oil price
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim monthlyOilrPrice() As Double
ReDim monthlyOilrPrice(1 To 16, 1 To 10000)

ReDim averMonthlyOilPrice(1 To 16)
Dim sum As Long

Sheets(1).Activate
sum = 0
spotPrice = Cells(4, 2).Value
volatility = Cells(5, 2).Activate
riskFreeRate = Cells(9, 2).Activate

Sheets(3).Activate
expectedReturn = (riskFreeRate / 100 - 0.5 * (volatilty / 100) ^ 2) * (1 / 12) * 100
sum = 0


'there are 12 periods so it will loop 16 times
'also calcualtes the sum over 12 periods

For j = 1 To 10000
For i = 1 To 16
    noise = ((volatility / 100) * Sqr(1 / 12) * normRandNum()) * 100
    change = Exp(expectedReturn / 100 + noise / 100) * 100
    genOilPrice = spotPrice / 100 * change
    monthlyOilrPrice(i, j) = genOilPrice
    spotPrice = genOilPrice
Next
spotPrice = 678
Next

For i = 1 To 16
    For j = 1 To 10000
        sum = sum + monthlyOilrPrice(i, j)
    Next
        averMonthlyOilPrice(i) = sum / 10000
        sum = 0
Next

End Sub

Sub crossHedge()
'calculate the cross-hedge ratio
'source http://financetrain.com/minimum-variance-hedge-ratio

Dim futurePriceArray() As Double
Dim averMonthlyOilPrice() As Double
Dim i As Integer
Dim fuelCost As Double
Dim nContracts As Double
Dim spotPrice

spotPrice = Sheets(1).Cells(4, 2).Value
fuelCost = Sheets(2).Cells(13, 2).Value

'get data, simulate monthly price, and get real future price for heating oils

Call getData
Call getFuturePrice(futurePriceArray)
Call monthlyOilGen(averMonthlyOilPrice)
'put the values in a table
For i = 1 To 16
    Cells(i + 1, 2).Value = futurePriceArray(i)
    Cells(i + 1, 3).Value = averMonthlyOilPrice(i)
Next
'calculate the percentage change
For i = 1 To 15
    Cells(i + 2, 5).Value = (Cells(i + 2, 2) - Cells(i + 1, 2)) / Cells(i + 1, 2)
    Cells(i + 2, 6).Value = (Cells(i + 2, 3) - Cells(i + 1, 3)) / Cells(i + 1, 3)
Next
'find standard deviation and corelation coeeficient

Cells(19, 2) = WorksheetFunction.StDev_S(Range(Cells(3, 5), Cells(17, 5)))
Cells(20, 2) = WorksheetFunction.StDev_S(Range(Cells(3, 6), Cells(17, 6)))
Cells(21, 2) = WorksheetFunction.Correl(Range(Cells(3, 5), Cells(17, 5)), Range(Cells(3, 6), Cells(17, 6)))
Cells(22, 2) = Cells(21, 2).Value * Cells(20, 2).Value / Cells(19, 2).Value
'find the number of future contracts

If Cells(22, 2).Value <= 0 Then
    Cells(23, 2).Value = 0
    MsgBox "You should not hedge"
Else
    nContracts = (fuelCost * 1000000 / (spotPrice * 140)) * (Cells(21, 2).Value) * (Cells(20, 2).Value) / (Cells(19, 2).Value)
    Cells(23, 2).Value = CInt(nContracts)
    MsgBox "You should buy " & CInt(nContracts) & " heating oil contracts"
End If

End Sub


Sub probability()
'this sub will calculate the probability
'for a given future price and projected net income

UserForm3.Show

Dim priceArray() As Double
Dim netIncomeArray() As Double
Dim i As Integer, j As Integer
Dim counter1 As Double, counter2 As Double, counter3 As Double
Dim percent1 As Double
ReDim priceArray(1 To 10000)
ReDim netIncomeArray(1 To 10000)
counter1 = 0
counter2 = 0
counter3 = 0
'get the array of simulated prices and net income

Sheets(1).Activate
For i = 1 To 10000
 priceArray(i) = Cells(i + 1, 11).Value
Next

Sheets(2).Activate

For j = 1 To 10000
    netIncomeArray(j) = Cells(j + 1, 11).Value
Next

Sheets(3).Activate
'calculate the probability when future spot price >= input

For i = 1 To 10000
    If priceArray(i) >= Cells(4, 12).Value Then
        counter1 = counter1 + 1
End If
Next
'this is the probability for projected net income
For j = 1 To 10000
    If netIncomeArray(j) >= Cells(7, 12).Value Then
        counter2 = counter2 + 1
End If
Next

Cells(2, 12).Value = counter1 / 10000
Cells(5, 12).Value = counter2 / 10000
'calculate the conditional probability
For i = 1 To 10000
    If priceArray(i) >= Cells(4, 12).Value Then
        If netIncomeArray(i) >= Cells(7, 12).Value Then
            counter3 = counter3 + 1
        End If
    End If
Next

Cells(10, 12).Value = counter3 / 10000


End Sub
