Attribute VB_Name = "Module2"
Option Explicit

Sub stock()

'set worksheet variable

    Dim ws As Worksheet
    
'loop through the worksheets

    For Each ws In Worksheets
    
    
'set variables

    Dim ticker As String
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim lastrow As Long
    Dim i As Long
    Dim opening_value As Double
    Dim closing_value As Double
    Dim change As Double
    
    
    
    
'start the total stock volume counter at zero
    total_stock_volume = 0

    
'keep track of the location for each ticker name in the summary table
    Dim summary_table_row As Integer
        summary_table_row = 2
        
    
'lastrow shortcut
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'set opening value
    opening_value = ws.Range("C2").Value
    

'loop through all the rows
    For i = 2 To lastrow
    

'check to see if the ticker in the row directly below a ticker has the same name and if it does, then...
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

'add to the total stock volume
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
    
    

    
'if the ticker below does not have the same name
    Else

'set the ticker name
    ticker = ws.Cells(i, 1).Value
    
'set the close value
    closing_value = ws.Cells(i, 6).Value
    
'set the yearly change
    change = closing_value - opening_value
    
'set the percent change
    percent_change = (change / opening_value)
    
'set the new opening value
    opening_value = ws.Cells(i + 1, 3).Value
    

'add to the total stock volume
   total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
   

'print the ticker name in the summary table
    ws.Range("I" & summary_table_row).Value = ticker
    
'print the total stock volume amount to the summary table
    ws.Range("L" & summary_table_row).Value = total_stock_volume
    
'print the yearly change amount to the summary table
    ws.Range("J" & summary_table_row).Value = change
    
'print the percent change amount to the summary table****************************************************8
    ws.Range("K" & summary_table_row).Value = percent_change
    
    
'add one to the summary table row because we're done checking all of the same ticker and moving down a row on the summary table to the next ticker
    summary_table_row = summary_table_row + 1
    
'reset the total stock volume to 0 because we're done with that ticker and moving down a row to the next ticker's total stock volume
    total_stock_volume = 0
    
   

    End If
    

    Next i
    


'find the largest percent change increase and place that value in Q3
    ws.Range("Q3").Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
   
'find the largest percent change decrease and place that value in Q4
   ws.Range("Q4").Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
   
'find the largest total stock volume and place that value in Q5
    ws.Range("Q5").Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
   
    Dim j As Long


'loop through the rows to associate the ticker with the greatest % increase value
    
    For j = 2 To lastrow
    
    If ws.Range("Q3").Value = ws.Cells(j, 11).Value Then
    
        ws.Range("P3").Value = ws.Cells(j, 9).Value
        
    End If
    
    Next j
    

        
'and then to associate the ticker with the greatest % decrease

    Dim k As Long
    
    For k = 2 To lastrow

    
    If ws.Range("Q4").Value = ws.Cells(k, 11).Value Then
        
        ws.Range("P4").Value = ws.Cells(k, 9).Value
        
    End If
    
    Next k
    
'and then to associate the ticker with the greatest total stock volume. Dim the letter l as integer

    Dim l As Long
    
    For l = 2 To lastrow
    
    If ws.Range("Q5").Value = ws.Cells(l, 12).Value Then
    
        ws.Range("P5").Value = ws.Cells(l, 9).Value
        
    End If
    
    Next l
    

'declare for yearly change loop

    Dim m As Long
    
    For m = 2 To lastrow
    

'set positve yearly percent change to green
    If ws.Cells(m, 10).Value > 0 Then
        ws.Cells(m, 10).Interior.ColorIndex = 4

'set negative yearly percent change to red
        Else
        ws.Cells(m, 10).Interior.ColorIndex = 3
        
    End If
    
    Next m
    

'declare for percent change loop
    Dim n As Long
    
    For n = 2 To lastrow
    

'set postive raw yearly change to green
    If ws.Cells(n, 11).Value > 0 Then
        ws.Cells(n, 11).Interior.ColorIndex = 4

'set negative raw yearly change to red
    Else
        ws.Cells(n, 11).Interior.ColorIndex = 3
        
    End If
    
    Next n
    
'set percent change column as a percentage
    ws.Range("K:K").NumberFormat = "0.00%"
    
'set percent change column as a percentage
    ws.Range("Q:Q").NumberFormat = "0.00%"
    
'set greatest total volume cell back to an integer
    
    ws.Range("Q5").NumberFormat = "0"

    
    
'go to the next worksheet

    Next ws


    
    

End Sub

    
        
        
    
    
        
    
    
'
    
        
    
        
    

   








