Attribute VB_Name = "Module2"

Sub preprocess()
    Set sh = ActiveSheet
    zFile = sh.Range("A1").Value
    Dim fso As New FileSystemObject
    zDir = fso.GetParentFolderName(zFile)
    zOut = zDir & "\out.c"
    
    Dim pp As New myPreProc
    pp.RemoveSw zFile
    
    Dim tLine As myLineCode
    
    Set ts = fso.CreateTextFile(zOut)
    i = 0
    For Each tLine In pp.mLines
        For i = i + 1 To tLine.mNo - 1
            ts.WriteLine ""
        Next
        ts.WriteLine tLine.mTxt
    Next
    ts.Close
End Sub
