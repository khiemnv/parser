VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myGraphBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//2020-09-09: type cast, inc: x++
'//2020-09-14: switch case
'//2020-09-15: fix elseif
Private mCurG As myGraph
Private mCtx(20)
Private mLvl
Private mCount

Private Function init()
    Set mCurG = New myGraph
    mLvl = 0
    mCount = 0
End Function

Public Function crtGraph(blockObj)
    If mCurG Is Nothing Then init
    
    Dim stmtObj As myStmt
    For Each stmtObj In blockObj.mLst
        Select Case stmtObj.mType
        Case stmtIf
            convStmtIf stmtObj
        Case stmtSw
            convStmtSw stmtObj
        Case stmtAssign
            convStmtAssign stmtObj
        Case Else
            convStmt stmtObj
        End Select
    Next
    
    Set crtGraph = mCurG
End Function
Public Function refineGraph(graphObj)
    Set mCurG = graphObj
    Dim tDict As Dictionary
    Dim lNode As myNode
    Dim rNode As myNode
    
    Set tDict = mCurG.mDict
    
    Dim reg As New RegExp
    Set reg = New RegExp
    reg.Pattern = "^(bind|tmp)\d"
    
    'remove temp node
    'set in/out
    For i = tDict.Count - 1 To 0 Step -1
        Dim tNode As myNode
        Set tNode = tDict.Items(i)
        If tNode.mInLst Is Nothing Then
            If tNode.mType = eConst Then
            ElseIf tNode.mType = eFunc Then
            Else
                tNode.mType = eIn
            End If
        ElseIf tNode.mOutLst Is Nothing Then
            tNode.mType = eOut
        ElseIf tNode.mType <> eNone Then
        ElseIf tNode.mInLst.Count + tNode.mOutLst.Count = 2 Then
            Set lNode = tNode.mInLst(1)
            Set rNode = tNode.mOutLst(1)
            rplNode rNode.mInLst, tNode, lNode
            rplNode lNode.mOutLst, tNode, rNode
            tDict.Remove tNode.mId
        Else
            ' lnode -> tnode ->rnode[i]
            Set lNode = tNode.mInLst(1)
            Set rNode = tNode.mOutLst(1)
            If lNode.mOutLst.Count = 1 Then
                Set lNode.mOutLst = tNode.mOutLst
                For Each rNode In tNode.mOutLst
                    rplNode rNode.mInLst, tNode, lNode
                Next
                tDict.Remove tNode.mId
            ElseIf rNode.mInLst.Count = 1 Then
                Set rNode.mInLst = tNode.mInLst
                For Each lNode In tNode.mInLst
                    rplNode lNode.mOutLst, tNode, rNode
                Next
                tDict.Remove tNode.mId
            Else
                assert (reg.Test(tNode.mName))
                assert (tNode.mInLst.Count = 1)
                rplNodeLst lNode.mOutLst, tNode, tNode.mOutLst
                For Each rNode In tNode.mOutLst
                    rplNode rNode.mInLst, tNode, lNode
                Next
                tDict.Remove tNode.mId
            End If
        End If
    Next
    
End Function
Private Function rplNodeLst(tLst, oldNode, newNodeLst)
    For i = 1 To tLst.Count
        If tLst(i) Is oldNode Then Exit For
    Next
    For Each newNode In newNodeLst
        tLst.Add newNode, after:=i
    Next
    tLst.Remove i
End Function
Private Function rplNode(tLst, oldNode, newNode)
    For i = 1 To tLst.Count
        If tLst(i) Is oldNode Then Exit For
    Next
    tLst.Add newNode, after:=i
    tLst.Remove i
End Function
Public Function printGraph(graphObj)
    Dim tDict As Dictionary
    Set tDict = graphObj.mDict
    Set tLst = New Collection
    For i = 0 To tDict.Count - 1
        Dim tNode As myNode
        Set tNode = tDict.Items(i)
        
        Dim desNode As myNode
        Dim srcNode As myNode
        n = 0
        If Not tNode.mInLst Is Nothing Then n = tNode.mInLst.Count
        m = 0
        If Not tNode.mOutLst Is Nothing Then m = tNode.mOutLst.Count
        s = n
        If n < m Then s = m
        ReDim arr(s - 1, 2) As myNode
        Set arr(0, 1) = tNode
        For j = 0 To n - 1
            Set arr(j, 0) = tNode.mInLst(j + 1)
        Next
        For j = 0 To m - 1
            Set arr(j, 2) = tNode.mOutLst(j + 1)
        Next
        
        For j = 0 To s - 1
            tLst.Add Array(arr(j, 0), arr(j, 1), arr(j, 2))
        Next
    Next
    Set printGraph = tLst
End Function
Private Function convStmtIf(ByVal stmtObj As myStmt)
    Set graphLst = New Collection
    
    'push
    Set parentG = mCurG
    Set mCtx(mLvl) = mCurG
    mLvl = mLvl + 1
    
    For Each childB In stmtObj.mBlockLst
        Set mCurG = New myGraph
        mCurG.mPath = childB.mPath
        mCurG.mId = childB.mId
        Set childG = crtGraph(childB)
        graphLst.Add childG
    Next
    
    'pop
    mLvl = mLvl - 1
    Set mCurG = mCtx(mLvl)
    
    'cnd
    Set cndNodes = New Collection
    For i = 1 To stmtObj.mCndLst.Count
        Set cndNode = convNode(stmtObj.mCndLst(i).mVal)
        cndNodes.Add cndNode
    Next
    
    'input, output list
    Dim inNodes As Dictionary
    Dim outNodes As Dictionary
    Set recLst = New Collection
    For Each childG In graphLst
        Set tmpLst = getInOut(inNodes, outNodes, childG)
        recLst.Add Array(inNodes, outNodes, tmpLst, childG)
    Next
    
    'missing else block
    If (stmtObj.mBlockLst.Count = stmtObj.mCndLst.Count) Then
        assert (False)
    End If
    
    'add out
    Dim outNames As Dictionary
    Set outNames = addOutToCtx(recLst)
    
    'add in
    addInToCtx recLst
    
    'bind input to ctx
    bindToCtx recLst
    
    'multi output
    For iOut = 0 To outNames.Count - 1
        zOut = outNames.Keys(iOut)
        Set tLst = New Collection
        
        For iRec = 1 To recLst.Count
            rec = recLst(iRec)
            Set outNodes = rec(1)
            tLst.Add outNodes(zOut)
        Next
        
        Dim swNode As myNode
        Set swNode = crtSwNode("sw" & stmtObj.mId & "_" & iOut + 1)
        Set tNode = tLst(1)
        Set cndNode = cndNodes(1)
        'true->sw
        link tNode, swNode
        link cndNode, swNode
        'sw->out
        Set outNode = crtOutNode(zOut)
        link swNode, outNode
        'false->sw
        For i = 2 To stmtObj.mCndLst.Count
            Set fNode = crtSwNode("sw" & stmtObj.mId & "_" & iOut + 1 & "_" & i)
            link fNode, swNode
            Set swNode = fNode
            Set tNode = tLst(i)
            link tNode, swNode
            Set cndNode = cndNodes(i)
            link cndNode, swNode
        Next
        Set fNode = tLst(i)
        link fNode, swNode
        
    Next
    
    getInOut inNodes, outNodes, mCurG
End Function
Private Function convStmtSw(ByVal stmtObj As myStmt)
    Set graphLst = New Collection
    
    'push
    Set parentG = mCurG
    Set mCtx(mLvl) = mCurG
    mLvl = mLvl + 1
    
    Set caseBlock = stmtObj.mBlockLst(1)
    For Each caseStmt In caseBlock.mLst
        Set mCurG = New myGraph
        Set childB = caseStmt.mBlockLst(1)
        mCurG.mPath = childB.mPath
        mCurG.mId = childB.mId
        Set childG = crtGraph(childB)
        graphLst.Add childG
    Next
    
    'pop
    mLvl = mLvl - 1
    Set mCurG = mCtx(mLvl)
    
    'sw(expr)
    zDataPortForDefault = "Last data port"
    zDataPortOrder = "Specify indices"
    zDataPortIndices = ""
    Set exprNode = convNode(stmtObj.mCndLst(1).mVal)
    'case expr:
    For Each caseStmt In caseBlock.mLst
        Set exprLst = caseStmt.mCndLst
        zExpr = ""
        For Each tk In exprLst
            If tk.mTxt = "default" Then
                If exprLst.Count = 1 Then
                    zDataPortForDefault = "Additional data port"
                End If
            Else
                zExpr = zExpr & tk.mTxt & " "
            End If
        Next
        If exprLst.Count > 1 Then
            zExpr = "[" & zExpr & "]"
        End If
        zDataPortIndices = zDataPortIndices & " " & zExpr
    Next
    zDataPortIndices = "{" & zDataPortIndices & "}"
    
    'input, output list
    Dim inNodes As Dictionary
    Dim outNodes As Dictionary
    Set recLst = New Collection
    For Each childG In graphLst
        Set tmpLst = getInOut(inNodes, outNodes, childG)
        recLst.Add Array(inNodes, outNodes, tmpLst, childG)
    Next
    
    'add out
    Dim outNames As Dictionary
    Set outNames = addOutToCtx(recLst)
    
    'add in
    addInToCtx recLst
    
    'bind input to ctx
    bindToCtx recLst
    
    'multi output
    For iOut = 0 To outNames.Count - 1
        zOut = outNames.Keys(iOut)
        Set tLst = New Collection
        
        For iRec = 1 To recLst.Count
            rec = recLst(iRec)
            Set outNodes = rec(1)
            tLst.Add outNodes(zOut)
        Next
        
        Dim msNode As myNode    'multiport switch
        Set msNode = crtMulSwNode("ms" & stmtObj.mId & "_" & iOut + 1)
        msNode.mDataPortIndices = zDataPortIndices
        msNode.mDataPortForDefault = zDataPortForDefault
        link exprNode, msNode
        For i = 1 To tLst.Count
            Set tNode = tLst(i)
            link tNode, msNode
        Next
        
        'multi sw->out
        Set outNode = crtOutNode(zOut)
        link msNode, outNode
    Next
    
    getInOut inNodes, outNodes, mCurG
End Function
Private Function addInToCtx(recLst)
    'crt in dict
    Dim inNames As New Dictionary
    For Each rec In recLst
        Dim inNodes As Dictionary
        Set inNodes = rec(0)
        For i = 0 To inNodes.Count - 1
            zIn = inNodes.Keys(i)
            If Not inNames.Exists(zIn) Then
                inNames.Add zIn, 0
            End If
        Next
    Next
    'crt new in (parent ctx)
    For i = 0 To inNames.Count - 1
        zIn = inNames.Keys(i)
        Set nodeObj = grpAddIn(zIn, mCurG)
    Next
End Function
Private Function addOutToCtx(recLst)
    'crt out dict
    Dim childG As myGraph
    Dim zOut As String
    Dim outNode As myNode
    Dim inNode As myNode
    Dim outNodes As Dictionary
    Dim outNames As Dictionary
    
    Set outNames = New Dictionary
    For Each rec In recLst
        Set outNodes = rec(1)
        For i = 0 To outNodes.Count - 1
            zOut = outNodes.Keys(i)
            If Not outNames.Exists(zOut) Then
                outNames.Add zOut, 0
            End If
        Next
    Next
    
    'add new out (child ctx)
    For i = 0 To outNames.Count - 1
        zOut = outNames.Keys(i)
        For Each rec In recLst
            'in,out,other,graph
            Set childG = rec(3)
            If Not rec(1).Exists(zOut) Then
                'crt new in->out (child ctx)
                If Not childG.mVarDict.Exists(zOut) Then
                    Set inNode = grpAddIn(zOut, childG)
                    rec(0).Add zOut, inNode 'update input dict
                Else
                    Set inNode = childG.mVarDict(zOut)
                End If
                Set outNode = grpAddOut(zOut, childG)
                rec(1).Add zOut, outNode    'update output dict
            End If
        Next
    Next
    
    Set addOutToCtx = outNames
End Function
'if not exist crt new
'else return existing
Private Function grpAddIn(zName, ByVal graphObj As myGraph)
    If Not graphObj.mVarDict.Exists(zName) Then
        Dim newNode As New myNode
        newNode.mName = zName
        newNode.mId = "in" & graphObj.mId & "_" & mCount
        mCount = mCount + 1
        graphObj.mDict.Add newNode.mId, newNode
        graphObj.mVarDict.Add zName, newNode
        Set grpAddIn = newNode
    Else
        Set grpAddIn = graphObj.mVarDict(zName)
    End If
End Function
Private Function grpAddOut(zName, ByVal graphObj As myGraph)
    Dim inNode As myNode
    Dim outNode As myNode
    assert (graphObj.mVarDict.Exists(zName))
    Set inNode = graphObj.mVarDict(zName)
    Set outNode = New myNode
    outNode.mId = "out" & graphObj.mId & "_" & mCount
    mCount = mCount + 1
    graphObj.mDict.Add outNode.mId, outNode
    Set graphObj.mVarDict(zName) = outNode
    outNode.mVersion = inNode.mVersion + 1
    link inNode, outNode
    
    Set grpAddOut = outNode
End Function
Private Function findNode(zName, nodeLst)
    
End Function
Private Function convStmtAssign(ByVal stmtObj As myStmt)
    Set inNode = convNode(stmtObj.mVal.mChilds(2))
    
    'x=x+1
    Dim enObj As myEvalNode
    Set enObj = stmtObj.mVal.mChilds(1)
    Set outNode = convNodeL(enObj)
    
    link inNode, outNode
End Function
Private Function convStmt(stmtObj As myStmt)
    'stmt declare, ...
End Function
'(in,out,tmpLst,graph)
Private Function bindToCtx(recLst)
    Dim tNode As myNode
    Dim zName As String
    For Each rec In recLst
        'in
        Dim inNodes As Dictionary
        Set inNodes = rec(0)
        For i = 0 To inNodes.Count - 1
            Set tNode = inNodes.Items(i)
            zName = tNode.mName
            If mCurG.mVarDict.Exists(zName) Then
                'link to existing input
                tNode.mName = "bind" & mCurG.mId & "_" & mCount & zName
                mCount = mCount + 1
                link mCurG.mVarDict(zName), tNode
                'add to ctx
                mCurG.mDict.Add tNode.mId, tNode
            Else
                'add new input
                mCurG.mVarDict.Add zName, tNode
                'add to ctx
                mCurG.mDict.Add tNode.mId, tNode
            End If
        Next
        
        'other
        For Each tNode In rec(2)
            zId = tNode.mId
            'add to ctx
            mCurG.mDict.Add zId, tNode
        Next
        
        'out
        Dim outNodes As Dictionary
        Set outNodes = rec(1)
        For i = 0 To outNodes.Count - 1
            Set tNode = outNodes.Items(i)
            zName = tNode.mName
            tNode.mName = "tmp" & mCurG.mId & "_" & mCount & zName
            mCount = mCount + 1
            'add to ctx
            mCurG.mDict.Add tNode.mId, tNode
        Next
    Next
End Function

Private Function getInOut(inNodes As Dictionary, outNodes As Dictionary, ByVal childG As myGraph)
    Set inNodes = New Dictionary
    Set outNodes = New Dictionary
    Set tLst = New Collection
    Dim tNode As myNode
    For i = 0 To childG.mDict.Count - 1
        Set tNode = childG.mDict.Items(i)
        If tNode.mType = eConst Then
            'const
            tLst.Add tNode
        ElseIf tNode.mType = eFunc Then
            'func Node
            tLst.Add tNode
        ElseIf tNode.mInLst Is Nothing Then
            inNodes.Add tNode.mName, tNode
        ElseIf tNode.mOutLst Is Nothing Then
            outNodes.Add tNode.mName, tNode
        Else
            'oper
            tLst.Add tNode
        End If
    Next
    Set getInOut = tLst
End Function

'conv evalNode to graphNode
Private Function convNode(ByVal enObj As myEvalNode)
    'Dim dictNode As Dictionary
    'Dim dictName As Dictionary
    
    'Set dictNode = mCurG.mDict
    'Set dictName = mCurG.mVarDict
    Dim newNode As myNode
    If enObj.mType = enVar Then
        'var
        zName = enObj.mName
        If mCurG.mVarDict.Exists(zName) Then
            Set newNode = mCurG.mVarDict(zName)
        Else
            Set newNode = New myNode
            newNode.mName = zName
            newNode.mId = enObj.mId
            mCurG.mVarDict.Add zName, newNode
            mCurG.mDict.Add newNode.mId, newNode
        End If
    ElseIf enObj.mType = enConst Then
        Set newNode = New myNode
        newNode.mId = enObj.mId
        newNode.mType = eConst
        newNode.mName = enObj.mName
        mCurG.mDict.Add newNode.mId, newNode
    ElseIf enObj.mType = enFunc Then
        Set newNode = New myNode
        newNode.mId = enObj.mId
        newNode.mType = eFunc
        newNode.nInport = enObj.mChilds.Count
        newNode.nOutport = 1
        newNode.mName = enObj.mName
        mCurG.mDict.Add newNode.mId, newNode
    ElseIf enObj.mOrder = ooTypeCast Then
        Set newNode = New myNode
        newNode.mId = enObj.mId
        newNode.mType = eTypeConv
        newNode.mDataType = enObj.mOper
        mCurG.mDict.Add newNode.mId, newNode
    Else
        'oper/const
        Set newNode = New myNode
        newNode.mId = enObj.mId
        mCurG.mDict.Add newNode.mId, newNode
        Select Case enObj.mOper
            Case "!"
                newNode.mType = eLogical
                newNode.mOperator = "NOT"
            Case "~"
                newNode.mType = eBitwise
                newNode.mOperator = "NOT"
            Case "*"
                newNode.mType = eProduct
                newNode.mOperator = "**"
            Case "/"
                newNode.mType = eProduct
                newNode.mOperator = "*/"
            Case "%"
                newNode.mType = eMath
                newNode.mOperator = "mod"  'Operator: mod
            Case "+"
                newNode.mType = eSum
                newNode.mOperator = "++"
            Case "-"
                newNode.mType = eSum
                newNode.mOperator = "+-"
            Case ">>"
                newNode.mType = eShift
                newNode.mOperator = "Right" 'BitShiftDirection: Right/Left
            Case "<<"
                newNode.mType = eShift
                newNode.mOperator = "Left"
            Case "==", ">", "<", ">= ", "<="
                newNode.mType = eRelation
                newNode.mOperator = enObj.mOper
            Case "!="
                newNode.mType = eRelation
                newNode.mOperator = "~="
            Case "&"
                newNode.mType = eBitwise
                newNode.mOperator = "AND"
            Case "|"
                newNode.mType = eBitwise
                newNode.mOperator = "OR"
            Case "^"
                newNode.mType = eBitwise
                newNode.mOperator = "XOR"
            Case "&&"
                newNode.mType = eLogical
                newNode.mOperator = "AND"
            Case "||"
                newNode.mType = eLogical
                newNode.mOperator = "OR"
            Case Else
                assert (False)
        End Select
    End If
    If Not enObj.mChilds Is Nothing Then
        For Each childEn In enObj.mChilds
            Set childNode = convNode(childEn)
            link childNode, newNode
        Next
    End If
    
    Set convNode = newNode
End Function
Private Function convNodeL(ByVal enObj As myEvalNode)
    assert (enObj.mType = enVar)
    zName = enObj.mName
    
    Dim newNode As myNode
    Set newNode = New myNode
    newNode.mId = enObj.mId
    newNode.mName = zName
    mCurG.mDict.Add newNode.mId, newNode
    
    If mCurG.mVarDict.Exists(zName) Then
        Set oldNode = mCurG.mVarDict(zName)
        newNode.mVersion = oldNode.mVersion + 1
        Set mCurG.mVarDict(zName) = newNode
    Else
        mCurG.mVarDict.Add zName, newNode
    End If
    
    Set convNodeL = newNode
End Function
Public Function link(ByVal src As myNode, ByVal des As myNode)
    If des.mInLst Is Nothing Then Set des.mInLst = New Collection
    If src.mOutLst Is Nothing Then Set src.mOutLst = New Collection
    des.mInLst.Add src, src.mId
    src.mOutLst.Add des, des.mId
End Function
Private Function crtSwNode(zId)
    Set crtSwNode = New myNode
    crtSwNode.mType = eSw
    crtSwNode.mId = zId
    mCurG.mDict.Add crtSwNode.mId, crtSwNode
End Function
Private Function crtMulSwNode(zId)
    Set crtMulSwNode = New myNode
    crtMulSwNode.mType = eMulSw
    crtMulSwNode.mId = zId
    mCurG.mDict.Add crtMulSwNode.mId, crtMulSwNode
End Function
Private Function crtMuxNode(zId, n)
    Set crtMuxNode = New myNode
    crtMuxNode.mType = eMux
    crtMuxNode.mId = zId
    crtMuxNode.nInport = n
    mDict.Add crtMuxNode.mId, crtMuxNode
End Function
Private Function crtInNode(zName) As myNode
    Dim newNode As myNode
    Set newNode = New myNode
    mCount = mCount + 1
    newNode.mId = "in" & mCount
    newNode.mName = zName
    mCurG.mDict.Add newNode.mId, newNode
    mCurG.mVarDict.Add zName, newNode
    Set crtInNode = newNode
End Function
Private Function crtOutNode(zName) As myNode
    Dim newNode As myNode
    Set newNode = New myNode
    mCount = mCount + 1
    newNode.mId = "out" & mCount
    newNode.mName = zName
    mCurG.mDict.Add newNode.mId, newNode
    If mCurG.mVarDict.Exists(zName) Then
        Set oldNode = mCurG.mVarDict(zName)
        newNode.mVersion = oldNode.mVersion + 1
        Set mCurG.mVarDict(zName) = newNode
    Else
        mCurG.mVarDict.Add newNode.mName, newNode
    End If
    Set crtOutNode = newNode
End Function
Private Function crtGotoFrom(gotoNode, fromNode)

End Function

Private Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub

