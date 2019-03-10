Attribute VB_Name = "Module1"
Sub stock_sum()

Dim ws As Worksheet

For Each ws In Worksheets

   ws.Range("I1").Value = "Stock_Name"
   ws.Range("J1").Value = "Total Stock Volume"

   Dim Stock_Name As String
   Dim Stock_Total As Double
   Stock_Total = 0

   Dim row As Integer
   row = 2
   Dim i As Long

   Dim Lastrow As Long
   Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

       For i = 2 To Lastrow
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               Stock_Name = ws.Cells(i, 1).Value
               Stock_Total = Stock_Total + ws.Cells(i, 7).Value
               ws.Range("I" & row).Value = Stock_Name
               ws.Range("j" & row).Value = Stock_Total
               row = row + 1
               Stock_Total = 0
           Else
               Stock_Total = Stock_Total + ws.Cells(i, 7).Value
           End If
        Next i
Next

End Sub

