Attribute VB_Name = "Module1"
Enum eSts
    stsInit
    stsCont
    
    stsIf
    stsIfEnd
    stsElse
    stsElseEnd
End Enum

Sub mainxxx()
    Set sh = Workbooks("parser.xlsm").Sheets(1)
    zFile = sh.Range("A1").Value
    Dim par As New myParser
    Set funcObj = par.getFunc(zFile)
    
    Dim blockObj As myBlockCode
    Set blockObj = par.parseFun(funcObj)
    
    'evaluation
    Dim stmtObj As myStmt
    Dim evalz As New myEvaluator
    Set tLst = New Collection
    tLst.Add blockObj
    i = 1
    Do
        For Each stmtObj In blockObj.mLst
            If stmtObj.mType = stmtIf Then
                Dim cndStmt As myStmt
                For Each cndStmt In stmtObj.mCndLst
                    Set cndStmt.mVal = evalz.evalStmt(cndStmt.mLst, cndStmt.mId)
                Next
                For Each childB In stmtObj.mBlockLst
                    tLst.Add childB
                Next
            ElseIf stmtObj.mType = stmtSw Then
                Dim swStmt As myStmt
                Set swStmt = stmtObj
                Set cndStmt = swStmt.mCndLst(1)
                Set cndStmt.mVal = evalz.evalStmt(cndStmt.mLst, cndStmt.mId)
                Set caseBlock = swStmt.mBlockLst(1)
                For Each caseStmt In caseBlock.mLst
                    tLst.Add caseStmt.mBlockLst(1)
                Next
            ElseIf stmtObj.mLst(1).mType = tkStatic Then
                stmtObj.mType = stmtDeclare
            ElseIf stmtObj.mLst(1).mType = tkType Then
                stmtObj.mType = stmtDeclare
            ElseIf stmtObj.mLst(1).mType = tkReturn Then
                stmtObj.mType = stmtReturn
            Else
                Set stmtObj.mVal = evalz.evalStmt(stmtObj.mLst, stmtObj.mId)
                If stmtObj.mVal.mOrder = ooAssign Then
                    stmtObj.mType = stmtAssign
                End If
            End If
        Next
        
        If tLst.Count = i Then Exit Do
        i = i + 1
        Set blockObj = tLst(i)
    Loop
    
    'graph
    
    Dim gph As New myGraphBuilder
    Dim gphObj As myGraph
    Set gphObj = gph.crtGraph(tLst(1))
    gph.refineGraph gphObj
    Set tLst = gph.printGraph(gphObj)
    'print graph
    iRow = 11
    iCol = 2
    Dim cln As New myCleaner
    cln.Clear sh, iRow, iCol, iCol + 2
    For Each rec In tLst
        For i = 0 To 2
            If Not rec(i) Is Nothing Then
                sh.Cells(iRow, iCol + i).Value = rec(i).mId & "," & rec(i).mName
            End If
        Next
        iRow = iRow + 1
    Next
    
    Dim grd As New myGrid
    Dim d2 As Dictionary
    Set d2 = grd.arange(gphObj.mDict)
    
    Dim dr As New myDraw
    Set cmdLst = dr.genConvSub(d2, gphObj.mDict)
    zOut = sh.Range("A2").Value
    Dim fs As New FileSystemObject
    Set ts = fs.CreateTextFile(zOut, True)
    ts.WriteLine "function gen(sys)"
    For Each zCmd In cmdLst
        ts.WriteLine zCmd
    Next
    ts.Close
    
End Sub


Sub testEval()
    Dim tkz As New myTokenizer
    zTxt = "a = b + c - d;"
    zTxt = "a = b + c * d / e;"
    zTxt = "a = (b + c) % d;"
    zTxt = "a = (b + c) * (d + e);"
    zTxt = "(( tu8_var1 == OFF )&&( tu8_var3 != tu8_var2 ))"
    zTxt = "x++;"
    zTxt = "tu1_ret = (uint8)CONST_V1;"
    zTxt = "ts2_var6 = fs2_func2( );"
    zTxt = "ts2_var7 = fs2_func3( ts2_var2, ts2_var3 );"
    zTxt = "tu8_Arr[ IDX_3 ] = ZERO;"
    zTxt = "a = !(b&c|d^~e);"
    
    Dim tLst As Collection
    Set tLst = New Collection
    tkz.stepAll zTxt, tLst
    
    Dim evalz As New myEvaluator
    Set rc = evalz.evalStmt(tLst, 0)
End Sub

Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub

