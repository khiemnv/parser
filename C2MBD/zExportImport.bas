Attribute VB_Name = "zExportImport"
'//2020-07-13;KhiemNV1;compCode
Sub ExportModules(Optional msg)
    Dim wb As Workbook
    Dim fso As New FileSystemObject
    
    Set wb = ThisWorkbook
    szExportPath = wb.Path
    
    For Each cmpComponent In wb.VBProject.VBComponents
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case 2 '//vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case 1  '//vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case Else
                GoTo lSkip
        End Select
        
        szFileName = szExportPath & "\" & szFileName
        
        If fso.FileExists(szFileName) Then
            Set ts = fso.OpenTextFile(szFileName, ForReading)
            zOld = ts.ReadAll
            ts.Close
            
            l = cmpComponent.CodeModule.CountOfLines
            zCur = cmpComponent.CodeModule.Lines(1, cmpComponent.CodeModule.CountOfLines)
            If compCode(zOld, zCur, cmpComponent.Type) = 0 Then GoTo lSkip
        End If
        
        cmpComponent.Export szFileName
        
lSkip:
    Next
End Sub
Private Sub ImportModules()
    Dim wb As Workbook
    Dim fso As New Scripting.FileSystemObject
    Dim fin As Scripting.File
    
    Set wb = ThisWorkbook
    Set cmpComponents = wb.VBProject.VBComponents
    
    Dim tDict As New Dictionary
    Set tDict = New Dictionary
    For Each cmpComponent In cmpComponents
        Select Case cmpComponent.Type
        Case 1  '//vbext_ct_StdModule
            zKey = cmpComponent.Name & ".bas"
        Case 2 '//vbext_ct_ClassModule
            zKey = cmpComponent.Name & ".cls"
        Case Else
            GoTo lSkip
        End Select
        
        tDict.Add zKey, cmpComponent
lSkip:
    Next
    
    Set impLst = New Collection
    Dim reg As New RegExp
    reg.Pattern = "(.*)\.(cls|bas)"
    reg.IgnoreCase = False
    For Each fin In fso.GetFolder(wb.Path).Files
        Set m = reg.Execute(fin.Name)
        If m.Count = 0 Then GoTo lSkip2
        
        zMod = m(0).SubMatches(0)
        zKey = fin.Name
        If Not tDict.Exists(zKey) Then GoTo lSkip2
        
        Set ts = fin.OpenAsTextStream(ForReading)
        zTxt = ts.ReadAll
        ts.Close
        
        Set oldComp = tDict(zKey)
        oldTxt = oldComp.CodeModule.Lines(1, oldComp.CodeModule.CountOfLines)
        
        nDiff = compCode(zTxt, oldTxt)
        If nDiff = 0 Then GoTo lSkip2
        
        '//import
        If zKey = "zExportImport.bas" Then GoTo lSkip2
        cmpComponents.Remove oldComp
        impLst.Add fin.Path
lSkip2:
    Next
    
    For Each zFile In impLst
        cmpComponents.Import zFile
    Next
End Sub
Private Function getAllcmpnt(cmpComponents)

End Function
'if equal then return 0
'nType = 1: vbext_ct_StdModule
'nType = 2: vbext_ct_ClassModule
Private Function compCode(zOld, zCur, Optional nType = 1)
    aOld = Split(zOld, vbCrLf)
    aCur = Split(zCur, vbCrLf)
    
    Select Case nType
    Case 1  'vbext_ct_StdModule
        iOld = 1
    Case 2  'vbext_ct_ClassModule
        iOld = 9
    Case Else
        assert (False)
    End Select
    
    'ignore case
    'ignore empty line
    iCur = 0
    max_count = 1000000
    For i = 1 To max_count
        For iOld = iOld To UBound(aOld)
            zOld = aOld(iOld)
            If zOld <> "" Then Exit For
        Next
        
        For iCur = iCur To UBound(aCur)
            zCur = aCur(iCur)
            If zCur <> "" Then Exit For
        Next
        
        If zOld = "" Then Exit For
        If zCur = "" Then Exit For
        If StrComp(zOld, zCur, vbTextCompare) <> 0 Then Exit For
        
        iOld = iOld + 1
        iCur = iCur + 1
    Next
    
    compCode = StrComp(zOld, zCur, vbTextCompare)
End Function

Private Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub
