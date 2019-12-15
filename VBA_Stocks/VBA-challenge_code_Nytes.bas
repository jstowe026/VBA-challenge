Attribute VB_Name = "Module1"
Sub stockdata()

Dim ws As Worksheet


'Loop through worksheets
For Each ws In ActiveWorkbook.Worksheets


'Naming Headers for Summary Data

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

'Define variables

Dim ticker As String

Dim year_begin As Double

Dim year_end As Double

Dim volume As Double
volume = 0

Dim summary_table_row As Integer
summary_table_row = 2

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

'last rows

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim summary_lastrow As Long
summary_lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    

'Loop
For I = 2 To lastrow

    'find year begining value
    If ws.Cells(I - 1, 1) <> ws.Cells(I, 1) Then
        
        'find year beginning value
        year_begin = ws.Cells(I, 3)
        
    
    
    ' Check if we are still within the same ticker, if it is not...
    ElseIf ws.Cells(I + 1, 1) <> ws.Cells(I, 1) Then

        'Set ticker
        ticker = ws.Cells(I, 1)
        
        'Add to volume
        volume = volume + ws.Cells(I, 7)
        
        'Find year end value
        year_end = ws.Cells(I, 6)
        
        'Calculate yearly change
        yearly_change = year_end - year_begin
        
        If year_begin = 0 Then
            percent_change = ws.Cells(I, 6)
        
            'Calculate percent change
            Else: percent_change = (ws.Cells(I, 6) - year_begin) / (year_begin)
            
            End If
        
        ' Print the ticker in the Summary Table
        ws.Range("I" & summary_table_row) = ticker
        
        ' Print the Volume Amount to the Summary Table
        ws.Range("L" & summary_table_row) = volume
        
        'Print the Yearly Change in the Summary Table
        ws.Range("J" & summary_table_row) = yearly_change
        
        'Print the Percent Change in the Summary Tabe
        ws.Range("K" & summary_table_row) = percent_change
        
        ' Add one to the summary table row
        summary_table_row = summary_table_row + 1
        
        'Reset yearly change
        yearly_change = 0
        
        'Reset percent change
        percent_change = 0
        
        ' Reset the volume
        volume = 0
        
        
    ' If the cell immediately following a row is the same ticker
    Else
    
        'Add to volume
        volume = volume + ws.Cells(I, 7)

        
End If
    
Next I



'loop through yearly change for final formatting
For x = 2 To summary_lastrow

    'Format percent change as a percentage
    ws.Cells(x, 11).NumberFormat = "0.00%"

    'Check value of yearly change
    If ws.Cells(x, 10) < 0 Then
    
        'format negative yearly change as red
        ws.Cells(x, 10).Interior.ColorIndex = 3
        
    Else
    
        'format positive yearly change as green
        ws.Cells(x, 10).Interior.ColorIndex = 4
        
    End If
    
Next x

'populate high level summary header and row labels

ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"

'define variables and values of those variables

Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Double

max_percent = Application.WorksheetFunction.max(Columns("K"))
min_percent = Application.WorksheetFunction.Min(Columns("K"))
max_volume = Application.WorksheetFunction.max(Columns("L"))

'loop through summary table to find the values and tickers associated with each
For y = 2 To summary_lastrow

    'Look for the maximum percent change, populate table, and format to percent
    If ws.Cells(y, 11) = max_percent Then
        
        'populate max percent and ticker
        ws.Cells(2, 17) = max_percent
        ws.Cells(2, 16) = ws.Cells(y, 9)
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
    'Look for the minimum percent change, populate table, and format to percent
    ElseIf ws.Cells(y, 11) = min_percent Then
    
        'populate min percent and ticker
        ws.Cells(3, 17) = min_percent
        ws.Cells(3, 16) = ws.Cells(y, 9)
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
    'Look for the maximum volume and populate table
    ElseIf ws.Cells(y, 12) = max_volume Then
    
        'populate max volume and ticker
        ws.Cells(4, 17) = max_volume
        ws.Cells(4, 16) = ws.Cells(y, 9)

 
 End If
 Next y
    
Next ws
    
End Sub

