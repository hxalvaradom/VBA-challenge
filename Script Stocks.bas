Attribute VB_Name = "Módulo2"
Sub Stocks():

    For Each ws In Worksheets
    
        ' Variable to hold

        'Dim LastRow As Integer
    
        'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ticker As String

        Dim tickertotal As Double

        tickertotal = 0

        Dim Summary_table_row As Integer

        Summary_table_row = 2

    Dim openprice As Integer

    Dim closeprice As Integer

    Dim Yearlychange As Integer

    Dim Percentagechange As Integer

    Yearlychange = closeprice / openprice

    For I = 2 To 263

        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        ' set the ticker name
        
        ticker = Cells(I, 1).Value
    
        ' Add to the tickertotal
    
        tickertotal = tickertotal + Cells(I, 6).Value
    
        ' print the ticker in the summary table
        
        Range("I" & Summary_table_row).Value = ticker
        
        'print the ticker amount to the summary table
    
        Range("L" & Summary_table_row).Value = tickertotal
        
        ' print the yearly Change
        
        Yearlychange = closeprice - openprice
        
        ' add one to the summary table row
    
        Summary_table_row = Summary_table_row + 1
    
        ' Reset the ticker total
    
        tickertotal = 0
              
        ' print the close price
        
        closeprice = Cells(I, 6).Value
        
        Else
    
        'Add to the tickertotal
    
        tickertotal = tickertotal + Cells(I, 6).Value
        
        openprice = Cells(I, 3).Value
        
        
        End If
    
    Next I

End Sub


  'Dim WS_Count As Integer

'Dim I As Integer

'WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.

'For I = 1 To WS_Count

 '   MsgBox ActiveWorkbook.Worksheets(I).Name
    
  '  Next I
    



