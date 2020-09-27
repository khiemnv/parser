VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function Clear(sh, firstRow, firstCol, lastCol)
    max_count = 1000000
    iRow = firstRow
    For i = 0 To max_count
        For iCol = firstCol To lastCol
            zTxt = sh.Cells(iRow, iCol).Value
            If zTxt <> "" Then Exit For
        Next
        If zTxt = "" Then Exit For
        iRow = iRow + 1
    Next
    
    assert (iRow < max_count)
    Set tRng = sh.Range(sh.Cells(firstRow, firstCol), sh.Cells(iRow, lastCol))
    tRng.Clear
End Function



