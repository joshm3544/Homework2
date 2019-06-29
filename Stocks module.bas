Attribute VB_Name = "Module1"
Sub stocks():

    Dim ticker As String
    
    Dim volume As Double
    volume = 0
    
    Dim Summary_table_row As Integer
    Summary_table_row = 2
        
        For i = 2 To 70926
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ticker = Cells(i, 1).Value
                
                volume = volume + Cells(i, 7).Value
                
                Range("J" & Summary_table_row).Value = ticker
                
                Range("K" & Summary_table_row).Value = volume
                
                
                Summary_table_row = Summary_table_row + 1
                
                
                volume = 0
                
                
            Else
            
                volume = volume + Cells(i, 7).Value
                
            
            
            End If
            
            
        Next i
            
    
End Sub
