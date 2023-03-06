# VBA-challenge
Challenge #2 - writing VBA macros to analyze generated stock market data

I created a VBA script/macro to analyze generated stock market data for three year's worth of data.  I created a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

In adddition, I made sure to use conditional formatting to highlight positive change in green and negative change in red.  

Below is the VBA code from my Excel workbook listed in plain text:

Sub StockDataMacro()
    
  Dim ws As Worksheet
  Dim ticker_type As String
  Dim yearly_chng As Double
  Dim openvalue As Double
  Dim closevalue As Double
  Dim volume_total As Double
  Dim lastrow As Double
  Dim Summary_Table_Row As Double

  WkSheets = Array("2018", "2019", "2020")
  
For Each ws In Sheets(Array("2018", "2019", "2020"))
    ws.Select
    ws.Activate
 
  Summary_Table_Row = 2
       
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For I = 2 To lastrow
  
  If openvalue = 0 Then
  openvalue = Cells(I, 3).Value

  End If
  
    ' Check if we are still within the same ticker type, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the ticker type
      ticker_type = Cells(I, 1).Value

      ' Add to the volume total
      volume_total = volume_total + Cells(I, 7).Value

      ' Print the ticker type in the Summary Table
      Range("i" & Summary_Table_Row).Value = ticker_type
      
    closevalue = Cells(I, 6).Value
    yearvalue = closevalue - openvalue
    
          
      If Cells(I, 10).Value >= 0 Then
            Cells(I, 10).Interior.ColorIndex = 4
    
        Else
            Cells(I, 10).Interior.ColorIndex = 3
    
        End If
      
    
    If openvalue = 0 Then
  percentagechange = 0
  Else
    
    percentagechange = (closevalue - openvalue) / openvalue
    End If
    

      ' Print the volume total to the Summary Table
      Range("l" & Summary_Table_Row).Value = volume_total
      
       ' Print the yearly change to the Summary Table
      Range("j" & Summary_Table_Row).Value = yearvalue
      
            
        If Range("j" & Summary_Table_Row).Value >= 0 Then
            Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
    
        Else
            Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
    
        End If
      
      
      ' Print the percentage change to the Summary Table
      Range("k" & Summary_Table_Row).Value = percentagechange

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the volume total
      volume_total = 0
      
      ' Reset the open value
      openvalue = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the volume total
      volume_total = volume_total + Cells(I, 7).Value

    End If
    
  Next I
  
Next ws
  
End Sub
