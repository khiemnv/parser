Attribute VB_Name = "Module1"
Sub mainxxx()
    Set sh = Sheets(1)
    zFile = sh.Range("A1").Value
    Set tLst = readData(zFile)
    Dim tDict As Dictionary
    Set tDict = crtNodeDict(tLst)
    Set tLnks = crtGraph(tDict)
    
    Set inLst = New Collection
    Set outLst = New Collection
    Dim tNode As myNode
    For i = 0 To tDict.Count - 1
        Set tNode = tDict.Items(i)
        Select Case tNode.mType
        Case eIn
            inLst.Add tNode.mName
        Case eOut
            outLst.Add tNode.mName
        End Select
    Next
    
    'export
    Set outSh = Sheets(2)
    iRow = 2
    iCol = 1
    For Each zIn In inLst
        outSh.Cells(iRow, iCol).Value = zIn
        iRow = iRow + 1
    Next
    
    iRow = 2
    iCol = 2
    For Each zOut In outLst
        outSh.Cells(iRow, iCol).Value = zOut
        iRow = iRow + 1
    Next
    
    iRow = 2
    iCol = 3
    For Each rec In tLnks
        outSh.Cells(iRow, iCol).Value = rec(0).mName
        outSh.Cells(iRow, iCol + 1).Value = rec(1)
        outSh.Cells(iRow, iCol + 2).Value = rec(2).mName
        outSh.Cells(iRow, iCol + 3).Value = rec(3)
        iRow = iRow + 1
    Next
    
End Sub
Private Function crtGraph(tDict As Dictionary)
    Dim tInDict As New Dictionary
    Dim tOutDict As New Dictionary
    Dim tLnks As New Collection
    Dim tNode As myNode
    Dim inNode As myNode
    Dim outNode As myNode
    For i = 0 To tDict.Count - 1
        Set tNode = tDict.Items(i)
        
        For Each rec In tNode.mInLst
            zIn = rec(0)
            nPort = rec(1)
            If tOutDict.Exists(zIn) Then
                'sys/out ->b/in
                srcRec = tOutDict(zIn)
                tLnks.Add Array(srcRec(0), srcRec(1), tNode, nPort)
            ElseIf tInDict.Exists(zIn) Then
                'sys/in ->b/in
                srcRec = tInDict(zIn)
                tLnks.Add Array(srcRec(0), srcRec(1), tNode, nPort)
            Else
                'add sys/in
                Set inNode = New myNode
                inNode.mName = zIn
                inNode.mType = eIn
                tDict.Add zIn, inNode
                tInDict.Add zIn, Array(inNode, 1)
                'sys/in ->b/in
                tLnks.Add Array(inNode, 1, tNode, nPort)
            End If
        Next
        
        For Each rec In tNode.mOutLst
            zOut = rec(0)
            nPort = rec(1)
            If tOutDict.Exists(zOut) Then
                'update out dict
                tOutDict(zOut) = Array(tNode, nPort)
            Else
                'add to out dict
                tOutDict.Add zOut, Array(tNode, nPort)
            End If
            If tInDict.Exists(zOut) Then
                'update sys/in: xxx ->xxx_prev
                Set inNode = tInDict(zOut)(0)
                assert (inNode.mType = eIn)
                inNode.mName = inNode.mName & "_prev"
                
                'update tDict
                tDict.Remove zOut
                tDict.Add inNode.mName, inNode
                
                'update inDict
                tInDict.Remove zOut
                tInDict.Add inNode.mName, Array(inNode, 1)
            End If
        Next
    Next
    For i = 0 To tOutDict.Count - 1
        zOut = tOutDict.Keys(i)
        srcRec = tOutDict.Items(i)
        
        Set outNode = New myNode
        outNode.mName = zOut
        outNode.mType = eOut
        tDict.Add zOut, outNode
        
        'b/out ->sys/out
        tLnks.Add Array(srcRec(0), srcRec(1), outNode, 1)
    Next
    
    Set crtGraph = tLnks
End Function
Private Function crtNodeDict(tLst)
    Dim tDict As New Dictionary 'node dict
    Dim tNode As myNode
    
    Dim reg As New RegExp
    reg.Pattern = "_r\d$"
    
    For Each rec In tLst
        zFunc = rec(1)
        zSignal = rec(2)
        zInout = rec(3)
        If tDict.Exists(zFunc) Then
            Set tNode = tDict(zFunc)
        Else
            Set tNode = New myNode
            tNode.mType = eFunc
            tNode.mName = zFunc
            Set tNode.mInLst = New Collection
            Set tNode.mOutLst = New Collection
            
            tDict.Add zFunc, tNode
        End If
        If zInout = "in" Then
            tNode.mInLst.Add Array(zSignal, tNode.mInLst.Count + 1)
        Else
            If reg.Test(zSignal) Then
                zOrg = reg.Replace(zSignal, "")
                tNode.mOutLst.Add Array(zOrg, tNode.mOutLst.Count + 1)
            Else
                tNode.mOutLst.Add Array(zSignal, tNode.mOutLst.Count + 1)
            End If
        End If
    Next
    
    Set crtNodeDict = tDict
End Function

Private Function readData(zFile)
    Set wb = Workbooks.Add(zFile)
    Set sh = wb.Sheets(1)
    Dim tLst As Collection
    Dim rd As myShReader
    Set rd = New myShReader
    rd.getShData sh, 2, Array(1, 2, 3), tLst
    wb.Close
    Set readData = tLst
End Function

Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub
