Sub Homework2()

  ' Define variables
  Dim Ticker As String
  Dim Open_Price As Double
  Dim Close_Price As Double
  Dim Yearly_Change As Double
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim Percent_Change As Double
  Dim Tot_Stock_Vol As Double
  Dim Start_Date As Double
  Dim Column_End As Double
  
  ' Loop through all stocks
  'Column_End = Cells(Rows.Count, 1)
  For i = 2 To 80000

    ' Check if we are still within the same stock, if it is not . . .
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker = Cells(i, 1).Value

      ' Generate the Open Price
      Start_Date = WorksheetFunction.Min(Range("B2:B80000"))
      Open_Price = WorksheetFunction.VLookup(Start_Date, Range("b2:c80000"), 2)
      
      ' Generate the Close Price
      Close_Price = Cells(i, 6).Value

      ' Generate the Yearly Change Total
      Yearly_Change = Close_Price - Open_Price

      ' Generate the Yearly Change Percentage
      Year_Change_Percent = (Yearly_Change / Open_Price)
      
      ' Generate the Total Stock Volume
      Tot_Stock_Vol = Cells(i, 7).Value

      ' Print the Ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Yearly Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_Change

      ' Print the Yearly Change Percentage to the Summary Table
      Range("K" & Summary_Table_Row) = Year_Change_Percent
      '??Range("K").NumberFormat = "0.00%"
      
      ' Print the Total Stock Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Tot_Stock_Vol
     
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Open_Price = 0
      Start_Date = 0
      Close_Price = 0
      Yearly_Change = 0
      Year_Change_Percent = 0
      Tot_Stock_Vol = 0

    ' If the cell immediately following a row is the same ticker . . .
    Else

      ' Add to the Ticker Total
      Open_Price = Open_Price + Cells(i, 3).Value

    End If

  Next i

End Sub