Attribute VB_Name = "Module1"


Sub test():

' declare variables '

' initial stock volume is set to zero '
Dim totalvolume As Double
    totalvolume = 0
' declare and calculate final row '
Dim finalrow As Long
    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    
Dim tickername As String
Dim tablerow As Double
    tablerow = 2
    
Dim yearlychange As Double
Dim percentchange As Double

Dim openprice As Double
Dim closeprice As Double

Dim red As String
Dim green As String

Dim ws As Worksheet

' -- only working in first two sheets* ?? -- '
For Each ws In Worksheets

' part I headers for new table columns '
' -- headers only showing up on first worksheet ?? -- '
ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "yearly change"
ws.Range("K1").Value = "percent change"
ws.Range("L1").Value = "total stock volume"

' part II headers/ titles for table '
ws.Range("O2").Value = "greatest % increase"
ws.Range("O3").Value = "greatest % decrease"
ws.Range("O4").Value = "greatest total volume"
ws.Range("P1").Value = "ticker"
ws.Range("Q1").Value = "value"

openprice = ws.Range("C2").Value

    For i = 2 To finalrow
    
        ' if the following cell is different from the previous... '
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' retreive the ticker name and print it in the table '
            tickername = ws.Cells(i, 1).Value
            ws.Cells(tablerow, 9).Value = tickername
            
            ' add the volume of the given row to the totalvolume and print it in the table '
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Cells(tablerow, 12).Value = totalvolume
            
            
            ' yearly change... '
            closeprice = ws.Cells(i, 6).Value
            yearlychange = closeprice - openprice
            ws.Cells(tablerow, 10).Value = yearlychange
            
            ' changes color according to yearlychange '
            red = yearlychange <= 0
            green = yearlychange > 0
            If red Then ws.Cells(tablerow, 10).Interior.ColorIndex = 3
            If green Then ws.Cells(tablerow, 10).Interior.ColorIndex = 4
                
            ' percent change... '
            percentchange = (yearlychange / openprice) * 100
            ws.Cells(tablerow, 11).Value = percentchange
            
            ' reset total stock volune to zero '
            totalvolume = 0
            
            tablerow = tablerow + 1
            
            ' resets open price... '
            openprice = ws.Cells(i + 1, 3).Value
               
    
        ' if the following cell is the same... '
        Else
            
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If

        Dim maxticker As String
        Dim minticker As String
        Dim maxvolume As String
         
    ' finds and prints greatest % increase and decrease and greatest total volume '
            ws.Range("Q2").Value = WorksheetFunction.max(ws.Range("K:K"))
            ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("Q4").Value = WorksheetFunction.max(ws.Range("L:L"))
    
        ' stores answers in variables '
        maxticker = ws.Range("Q2").Value
        minticker = ws.Range("Q3").Value
        maxvolume = ws.Range("Q4").Value
    
    ' if greatest % increase matches given row, pull ticker name from same row '
       If maxticker = ws.Cells(tablerow, 11).Value Then
            ws.Range("P2").Value = ws.Cells(tablerow, 11 - 2).Value
           End If
           
            If minticker = ws.Cells(tablerow, 11).Value Then
                ws.Range("P3").Value = ws.Cells(tablerow, 11 - 2).Value
                End If
                
            ' -- not printing ticker ?? -- '
                If maxvolume = ws.Cells(tablerow, 12).Value Then
                     ws.Range("P4").Value = ws.Cells(tablerow, 12 - 3).Value
                     
            
        End If

    Next i
    
    tablerow = 2
    
Next ws

End Sub

