VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myShReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NOTE: exit when meet empty rec
Public Function getShData(sh, baseRow, cols, outLst, _
Optional max_count = 1000000, _
Optional max_empty = 5)
    Set outLst = New Collection
    
    nCol = UBound(cols) + 1
    iRow = baseRow
    nEmpty = 0
    For i = 0 To max_count
        ReDim rec(nCol)
        rec(0) = iRow
        j = 1
        For Each iCol In cols
            rec(j) = sh.Cells(iRow, iCol).Value
            
            j = j + 1
        Next
        
        If IsAllEmpty(rec) Then
            nEmpty = nEmpty + 1
            If nEmpty > max_empty Then Exit For
        Else
            nEmpty = 0
            outLst.Add rec
        End If
        
        iRow = iRow + 1
    Next
    rc = 0
    If (i > max_count) Then rc = -1 '//out of range
    
    getShData = rc
End Function

Function IsAllEmpty(rec)
    IsAllEmpty = True
    For i = 1 To UBound(rec)
        If rec(i) <> "" Then
            IsAllEmpty = False
            Exit For
        End If
    Next
End Function

Public Function getSignals(sh, _
Optional baseRow = 1, Optional baseCol = 1)
    zTxt = sh.Cells(baseRow, baseCol).Value
    assert (InStr(zTxt, "Time") <> 0)
    
    max_count = 10000
    iRow = baseRow
    iCol = baseCol
    For i = 0 To max_count
        zTxt = sh.Cells(iRow + i, iCol).Value
        If zTxt = "" Then Exit For
    Next
    assert (i < max_count)
    nRow = i
    
    For i = 0 To max_count
        zTxt = sh.Cells(iRow, iCol + i).Value
        If zTxt = "" Then Exit For
    Next
    assert (i < max_count)
    nCol = i
    
    lastCol = baseCol + nCol - 1
    lastRow = baseRow + nRow - 1
    Set signalLst = New Collection
    For iCol = baseCol To lastCol
        Dim signal As mySignal
        Set signal = New mySignal
        signal.ord = iCol
        zName = sh.Cells(baseRow, iCol).Value
        signal.zName = zName
        Set signal.valList = New Collection
        For iRow = baseRow + 1 To lastRow
            signal.valList.Add sh.Cells(iRow, iCol).Value
        Next
        signalLst.Add signal
    Next
    Set getSignals = signalLst
End Function

