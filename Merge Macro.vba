
Sub MergeData()
    Dim nRowsOrg As Long
    Dim nRowsIn As Long
    Dim nColsIn As Long
    Dim nColsOrg As Long
    
    
    Dim r1 As Range
    Dim r2 As Range
    
    nRowsOrg = ThisWorkbook.Worksheets(1).Range("A1").CurrentRegion.Rows.Count
    nRowsIn = ActiveWorkbook.Worksheets(1).Range("A1").CurrentRegion.Rows.Count - 1
    nColsIn = ActiveWorkbook.Worksheets(1).Range("A1").CurrentRegion.Columns.Count
    nColsOrg = ThisWorkbook.Worksheets(1).Range("A1").CurrentRegion.Columns.Count
    ThisWorkbook.Worksheets(1).Rows(nRowsOrg + 1).Resize(nRowsIn).Insert
    
    Dim i As Integer

    For i = 1 To nColsIn Step 1 'itterating over columns to inport
        
        Dim j As Integer
        For j = 1 To nColsOrg Step 1 'itterating over columns in desitation sheet to find match
            'MsgBox ActiveWorkbook.Worksheets(1).Cells(1, i).Value & " vs " & ThisWorkbook.Worksheets(1).Cells(1, j).Value
            If ActiveWorkbook.Worksheets(1).Cells(1, i).Value = ThisWorkbook.Worksheets(1).Cells(2, j).Value Then ' if cloumn lables match

                    Set r1 = ActiveWorkbook.Worksheets(1).Range(Cells(2, i), Cells(nRowsIn + 1, i))
                    Set r2 = ThisWorkbook.Worksheets(1).Range(ThisWorkbook.Worksheets(1).Cells(nRowsOrg + 1, j), ThisWorkbook.Worksheets(1).Cells(nRowsOrg + nRowsIn, j))
            
                    r2.Value = r1.Value 'actaully copying the data 
                Exit For
            End If
        Next
    Next
    
    'MsgBox nRowsOrg
    'MsgBox nColsIn
    
    

    
    
End Sub
