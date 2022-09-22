Attribute VB_Name = "Module2"
Sub HW2_VBA():

'variables
    Dim Ticker As String
    Dim SummaryTableRow As Integer
    Dim volume As Double
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim WsPage As Long: WsPage = ThisWorkbook.Worksheets.Count
    
'loop
    For Z = 1 To WsPage
    Worksheets(Z).Cells(1, 9).Value = "Ticker"
    Worksheets(Z).Cells(1, 10).Value = "Yearly Change"
    Worksheets(Z).Cells(1, 11).Value = "Percent Change"
    Worksheets(Z).Cells(1, 12).Value = "Total Volume"
    
    'Start Point
        year_open = Worksheets(Z).Cells(2, 3).Value
    'last row
        lrow = Worksheets(Z).Cells(Worksheets(Z).Rows.Count, 1).End(xlUp).Row
        
    SummaryTableRow = 2
    volume = 0
    For i = 2 To lrow
    
'if Statement
    If Worksheets(Z).Cells(i + 1, 1).Value <> Worksheets(Z).Cells(i, 1).Value Then
    'Ticker
        Ticker = Worksheets(Z).Cells(i, 1).Value
    'Volume
        volume = volume + Worksheets(Z).Cells(i, 7).Value
    'Yearly Change
        year_close = Worksheets(Z).Cells(i, 6).Value
        yearly_change = year_close - year_open
    'Percent Change
        percent_change = (yearly_change / year_open)
        year_open = Worksheets(Z).Cells(i + 1, 3).Value
    'insert values
        Worksheets(Z).Range("I" & SummaryTableRow).Value = Ticker
        Worksheets(Z).Range("L" & SummaryTableRow).Value = volume
        Worksheets(Z).Range("J" & SummaryTableRow).Value = yearly_change
        Worksheets(Z).Range("K" & SummaryTableRow).Value = percent_change
        Worksheets(Z).Range("K" & SummaryTableRow).NumberFormat = "0.00%"
        
    'Color Format
        If (Worksheets(Z).Range("J" & SummaryTableRow) > 0) Then
            Worksheets(Z).Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        ElseIf (Worksheets(Z).Range("K" & SummaryTableRow) <= 0) Then
            Worksheets(Z).Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        End If
        
    'Add one to SummaryTableRow
    SummaryTableRow = SummaryTableRow + 1
    'Reset Volume
    volume = 0
    
    Else
    
    'Add to Volume total
    volume = volume + Worksheets(Z).Cells(i, 7).Value
    
    End If
    
    Next i
    Next Z
    

End Sub

