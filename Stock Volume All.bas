Attribute VB_Name = "Module1"
Sub StockVolume()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

  Dim Ticker As String
  Dim StockVolume As LongLong
  StockVolume = 0
  
  Range("I1") = "Ticker"
  Range("J1") = "Yearly Change"
  Range("K1") = "Percent Change"
  Range("L1") = "Total Stock Volume"

  Dim Summary_Table_Row As LongLong
  
  Summary_Table_Row = 2
    
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
  For I = 2 To lastrow

    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      Ticker = Cells(I, 1).Value

      StockVolume = StockVolume + Cells(I, 7).Value

      Range("I" & Summary_Table_Row).Value = Ticker

      Range("L" & Summary_Table_Row).Value = StockVolume

      Summary_Table_Row = Summary_Table_Row + 1
      
      StockVolume = 0
      
    Else

      StockVolume = StockVolume + Cells(I, 7).Value

    End If

  Next I
  
  ws.Cells(1, 1) = 1
Next

starting_ws.Activate

  
End Sub
