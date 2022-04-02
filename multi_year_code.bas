Attribute VB_Name = "Module1"
Sub multiyearstock()

Dim i As Long
Dim Summary_Row As Integer
Dim totalvolume As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim ws As Worksheet
Dim openingprice As Double
Dim closingprice As Double
Dim LastRow As Long

For Each ws In Worksheets
Summary_Row = 2
totalvolume = 0
yearlychange = 0
percentchange = 0

'Loop through all worksheets, we are going to be referring to these as ws.
openingprice = ws.Cells(2, 3).Value


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    'Loop throug all rows (use looping variable i)
For i = 2 To LastRow

        'totalvolume = totalvolume + "the value in column 7)"
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
        'if current cell ticker <> next cell ticker then
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

             'Put current ticker over in our summary table based on summary_row
             ws.Cells(Summary_Row, 10).Value = ws.Cells(i, 1).Value
            
            'Put totalvolume over in our summary table based on summary_row
            ws.Cells(Summary_Row, 13).Value = totalvolume
            
            'Put yearlychange over in our summary table based on summary_row
            
            closingprice = ws.Cells(i, 6).Value
            
            yearlychange = closingprice - openingprice
            
            ws.Cells(Summary_Row, 11).Value = yearlychange
            
            
           
            
            If openingprice > 0 Then
             percentchange = yearlychange / openingprice
            
          Else
                percentchange = 0
          
          End If
            
            openingprice = ws.Cells(i, 3).Value
            
            'Put percentchange over in our summary table based on summar_row
            ws.Cells(Summary_Row, 12).Value = percentchange
              
          
          'Based on the yearly change, is it positive or negative, do the coloring for red and green
          If yearlychange > 0 Then
          ws.Cells(Summary_Row, 11).Interior.ColorIndex = 4
          Else
          ws.Cells(Summary_Row, 11).Interior.ColorIndex = 3
          End If
          
          'increment summary_row by 1
            Summary_Row = Summary_Row + 1
          
          End If
          
Next i

'This is where the worsheet loop ends

Next ws

End Sub
