Attribute VB_Name = "Module1"
Sub test()
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate
    
    Dim column, list As Integer
    column = 1
    list = 2
    Dim open_p, close_p, change, percent As Double
    open_p = 0
    close_p = 0
    change = 0
    percent = 0
    
    Dim totvol, vol As Long
    
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yr Chg"
    Cells(1, 11).Value = "% Chg"
    Cells(1, 12).Value = "Tot Vol"
    
    open_p = Cells(2, 3).Value

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
    
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            
            close_p = Cells(i, 6).Value
            
            change = close_p - open_p
            percent = (change / open_p) * 100
            percent = Round(percent, 2)
            
            Cells(list, 11).Value = (CStr(percent) & "%")
            
            
            Cells(list, 10).Value = change
                If (change > 0) Then
                    Cells(list, 10).Interior.ColorIndex = 4
                    
                    ElseIf (change <= 0) Then
                    Cells(list, 10).Interior.ColorIndex = 3
                End If
        
            ticker = Cells(i, column).Value
            Cells(list, 9).Value = ticker
            
            Cells(list, 12).Value = totvol
        
            list = list + 1
            totvol = 0
            change = 0
            close_p = 0
            percent = 0
            open_p = Cells(i + 1, 3).Value
            
            Columns(12).AutoFit
            
            
            
            Else
            
            vol = Cells(i + 1, 7).Value
            totvol = vol + totvol
    
             
            End If
        
        Next i

    Next ws
End Sub
