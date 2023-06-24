' Module 2 Challenge VBA Script

Sub VBAStocks()

' To loop it to work through each worksheet in the file
For Each ws In Worksheets

' Assigning dimensions as needed and assigning initial value
Dim tickername As String
Dim STOCKvolume As Double
STOCKvolume = 0
Dim newtable As Integer
newtable = 2

 ' Assigning all the variables used
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim totalPrice As Double
totalPrice = 0
Dim PercentChange As Double
Dim Volume As Double
Dim Row As Long
Dim LastRow As Long
Dim greatestinc As Double
Dim greatestdec As Double
Dim greatestvol As Double

'Column headers for New Table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("k1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'finding the opening value of the Stock
OpenPrice = ws.Cells(2, 3).Value
greatestinc = 0  ' assigning the lowest value
greatestvol = 0
greatestdec = 99999999  ' assigning the highest value

'getting the row number of the last row with data and establish a value for it
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
' assigning the range of Row
For Row = 2 To LastRow

'set the ticker name
tickername = ws.Cells(Row, 1).Value

'finding total number of the stock
totalPrice = totalPrice + ws.Cells(Row, 7).Value

'finding the total ticker value
If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1) Then
ws.Cells(newtable, 9) = tickername
STOCKvolume = STOCKvolume + ws.Cells(Row, 7).Value
ws.Cells(newtable, 12).Value = STOCKvolume
ClosePrice = ws.Cells(Row, 6).Value

ws.Cells(newtable, 10).Value = ClosePrice - OpenPrice
ws.Cells(newtable, 10).NumberFormat = "$#.##" 'formatting the colum with $ sign
ws.Cells(newtable, 11).Value = ((ClosePrice - OpenPrice) / OpenPrice)

' the loop runs and tries to find the greatest increase and dreccrease and greatest value while the ticker is being generated in the ccolumn.
' determining the greatest % increase
If ws.Cells(newtable, 11).Value > greatestinc Then
greatestinc = ws.Cells(newtable, 11).Value
ws.Range("O2") = ws.Cells(newtable, 9).Value
ws.Range("P2").Value = ws.Cells(newtable, 11).Value
End If

'determining the greatest % decrease
If ws.Cells(newtable, 11).Value < greatestdec Then
greatestdec = ws.Cells(newtable, 11).Value
ws.Range("O3") = ws.Cells(newtable, 9).Value
ws.Range("P3").Value = ws.Cells(newtable, 11).Value
End If

'determining the greatest total volume
If Cells(newtable, 12).Value > greatestvol Then
greatestvol = ws.Cells(newtable, 12).Value
ws.Range("O4") = ws.Cells(newtable, 9).Value
ws.Range("P4").Value = ws.Cells(newtable, 12).Value
End If

 
' to change the colour based on the conditions for yearly change
 
If ws.Cells(newtable, 10).Value >= 0 Then
ws.Cells(newtable, 10).Interior.ColorIndex = 4 'change to colour green

Else
ws.Cells(newtable, 10).Interior.ColorIndex = 3 ' change to colour red
End If

If ws.Cells(newtable, 11).Value >= 0 Then
ws.Cells(newtable, 11).Interior.ColorIndex = 4 'change to colour green

Else
ws.Cells(newtable, 11).Interior.ColorIndex = 3 ' change to colour red


End If


'finding the new opening value
OpenPrice = ws.Cells(Row + 1, 3)

STOCKvolume = 0

ws.Cells(newtable, 10).NumberFormat = "%$.$$" 'formatting the colum with $ sign

' reset variables for new stock ticker
newtable = newtable + 1

Else
STOCKvolume = STOCKvolume + ws.Cells(Row, 7).Value

 End If
 


 
>>Next Row


Next ws


End Sub

