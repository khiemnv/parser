VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//2020-09-09: oper: ++, type cast
'//2020-09-09: func()
'//2020-09-15: bitwise, not
Enum eOperOrder
    
    ooNone = 0
    
    ooInc
    ooTypeCast
    ooNot
    
    ooProduct
    ooSum
    ooShift
    ooComp
    ooBitwise
    ooLogic
    
    ooAssign
    ooComma
End Enum

Private mValLst As Collection  'value
Private mOperLst As Collection  'oper
Private mLvl
Private mStack(20)
Private mLineNo
Private mCount

Public Function reset()
    Set mValLst = New Collection
    Set mOperLst = New Collection
    mLvl = 0
End Function

'lineNo ->unique id
Public Function evalStmt(tkLst, ByVal lineNo)
    reset
    mLineNo = lineNo
    mCount = 0
    
    Dim tk As myToken
    Dim castNode As myEvalNode
    For i = 1 To tkLst.Count
        Set tk = tkLst(i)
        If tk.mType = tkSemi Then Exit For
        
        If chkTypeCast(tkLst, i, castNode, iNext) Then
            'type cast
            pushNode mOperLst, castNode
            i = iNext
        ElseIf tk.mType = tkLBR Then
            sts = 0
            'function call
            If i > 1 Then
                Dim prevTk As myToken
                Set prevTk = tkLst(i - 1)
                If prevTk.mType = tkWord Then
                    Set funcNode = mValLst(mValLst.Count)
                    funcNode.mType = enFunc
                    sts = enFunc
                End If
            End If
            'push stack
            mStack(mLvl) = Array(mValLst, mOperLst, mLvl, sts)
            mLvl = mLvl + 1
            Set mValLst = New Collection
            Set mOperLst = New Collection
        ElseIf tk.mType = tkRBR Then
            For j = mOperLst.Count To 1 Step -1
                reduce
            Next
            Set childValLst = mValLst
            mStack(mLvl) = Array(mValLst, mOperLst, mLvl)
            
            'pop stack
            mLvl = mLvl - 1
            rec = mStack(mLvl)
            Set mValLst = rec(0)
            Set mOperLst = rec(1)
            sts = rec(3)
            
            If sts = enFunc Then
                'function call
                Set funcNode = mValLst(mValLst.Count)
                assert (funcNode.mType = enFunc)
                Set funcNode.mChilds = childValLst
            Else
                'other
                assert (childValLst.Count <= 1)
                If childValLst.Count = 1 Then
                    mValLst.Add childValLst(1)
                    childValLst(1).mName = "(" & childValLst(1).mName & ")"
                End If
            End If
        ElseIf tk.mType = tkLSBR Then
            'push stack
            mStack(mLvl) = Array(mValLst, mOperLst, mLvl)
            mLvl = mLvl + 1
            Set mValLst = New Collection
            Set mOperLst = New Collection
        ElseIf tk.mType = tkRSBR Then
            For j = mOperLst.Count To 1 Step -1
                reduce
            Next
            assert (mValLst.Count = 1)
            Set childValLst = mValLst
            mStack(mLvl) = Array(mValLst, mOperLst, mLvl)
            
            'pop stack
            mLvl = mLvl - 1
            rec = mStack(mLvl)
            Set mValLst = rec(0)
            Set mOperLst = rec(1)
            
            Dim arrNode As myEvalNode
            Set arrNode = crtOper(tk, ooNone)
            Dim tNode As myEvalNode
            Set tNode = popNode(mValLst)
            Set arrNode.mChilds = New Collection
            arrNode.mChilds.Add tNode
            arrNode.mChilds.Add childValLst(1)
            pushNode mValLst, arrNode
        Else
            'step
            rc = step(tk)
        End If
    Next
    For i = mOperLst.Count To 1 Step -1
        reduce
    Next
    assert (mValLst.Count = 1)
    Set evalStmt = mValLst(1)
End Function
Private Function step(ByVal tk As myToken)
    Dim curOper As myEvalNode
    Set curOper = Nothing
    
    Select Case tk.mType
    Case tkWord
        mValLst.Add crtVal(tk)
    Case tkNOT, tkBitNot
        Set curOper = crtOper(tk, ooNot)
    Case tkDec, tkInc
        Set curOper = crtOper(tk, ooInc)
    Case tkMul, tkDiv, tkMod
        Set curOper = crtOper(tk, ooProduct)
    Case tkAdd, tkSub
        Set curOper = crtOper(tk, ooSum)
    Case tkRShift, tkLShift
        Set curOper = crtOper(tk, ooShift)
    Case tkGT, tkLT, tkGTE, tkLTE, tkEQU, tkNEQ
        Set curOper = crtOper(tk, ooComp)
    Case tkBitAnd, tkBitOr, tkBitXor
        Set curOper = crtOper(tk, ooBitwise)
    Case tkAnd, tkOr
        Set curOper = crtOper(tk, ooLogic)
    Case tkAssign
        Set curOper = crtOper(tk, ooAssign)
    Case tkComma
        Set curOper = crtOper(tk, ooComma)
    End Select
    
    If Not curOper Is Nothing Then
        For i = mOperLst.Count To 1 Step -1
            Dim topOper As myEvalNode
            Set topOper = mOperLst(i)
            If curOper.mOrder >= topOper.mOrder Then
                rc = reduce()
            Else
                Exit For
            End If
        Next
        pushNode mOperLst, curOper
    End If
End Function

Private Function reduce()
    Dim topOper As myEvalNode
    Set topOper = popNode(mOperLst)
    
    Select Case topOper.mOrder
    Case ooComma
        'do nothing
    Case ooProduct, ooSum, ooShift, ooComp, ooBitwise, ooLogic, ooAssign
        Set rNode = popNode(mValLst)
        Set lNode = popNode(mValLst)
        Set topOper.mChilds = New Collection
        topOper.mChilds.Add lNode
        topOper.mChilds.Add rNode
        topOper.mName = lNode.mName & topOper.mTk.mTxt & rNode.mName
        pushNode mValLst, topOper
    Case ooInc
        Dim varNode As myEvalNode
        Dim var2Node As myEvalNode
        Dim sumNode As myEvalNode
        Dim oneNode As myEvalNode
        Dim assNode As myEvalNode
        
        Set varNode = popNode(mValLst)
        Set var2Node = crtVal(varNode.mTk)
        Set oneNode = crtConst("1")
        Set sumNode = crtOperZ("+", ooSum)
        Set assNode = crtOperZ("=", ooAssign)
        
        Set sumNode.mChilds = New Collection
        sumNode.mChilds.Add varNode
        sumNode.mChilds.Add oneNode
        
        Set assNode.mChilds = New Collection
        assNode.mChilds.Add var2Node
        assNode.mChilds.Add sumNode
        
        pushNode mValLst, assNode
        
        'Set topOper.mChilds = New Collection
        'topOper.mChilds.Add varNode
        'topOper.mName = varNode.mName & topOper.mTk.mTxt
        'pushNode mValLst, topOper
   Case ooTypeCast
        Set rNode = popNode(mValLst)
        Set topOper.mChilds = New Collection
        topOper.mChilds.Add rNode
        topOper.mName = "(" & topOper.mTk.mTxt & ")" & rNode.mName
        pushNode mValLst, topOper
    Case ooNot
        Set rNode = popNode(mValLst)
        Set topOper.mChilds = New Collection
        topOper.mChilds.Add rNode
        topOper.mName = topOper.mTk.mTxt & rNode.mName
        pushNode mValLst, topOper
     Case Else
        If topOper.mOrder > ooRBR Then
            '()
        Else
            assert (False)
        End If
    End Select
End Function
Private Function pushNode(tLst, tNode)
    tLst.Add tNode
End Function
Private Function popNode(tLst)
    Set popNode = tLst(tLst.Count)
    tLst.Remove tLst.Count
End Function
Private Function crtOper(tk, order)
    Dim newOper As New myEvalNode
    Set newOper.mTk = tk
    newOper.mOper = tk.mTxt
    newOper.mOrder = order
    newOper.mId = "op" & mLineNo & "_" & mCount
    mCount = mCount + 1
    Set crtOper = newOper
End Function
Private Function crtOperZ(zOper, order)
    Dim tk As New myToken
    tk.mTxt = zOper
    Set crtOperZ = crtOper(tk, order)
End Function
Private Function crtConst(zTxt)
    Dim tk As New myToken
    tk.mTxt = zTxt
    tk.mType = tkWord
    Set crtConst = crtVal(tk)
End Function
Private Function crtVal(ByVal tk As myToken)
    Dim newNode As New myEvalNode
    newNode.mName = tk.mTxt
    newNode.mType = enVar
    Set newNode.mTk = tk
    newNode.mId = "vl" & mLineNo & "_" & mCount
    mCount = mCount + 1
    If chkConst(tk.mTxt) Then
        newNode.mType = enConst
    End If
    Set crtVal = newNode
End Function
Private Function chkConst(zTxt) As Boolean
    Dim reg As New RegExp
    reg.Pattern = "^[A-Z_0-9]+$"
    chkConst = reg.Test(zTxt)
End Function
Private Function chkTypeCast(ByVal tkLst, ByVal iBase, ByRef castNode As myEvalNode, ByRef iNext)
    Dim tk As myToken
    s = 0
    s_type = 1
    s_succ = 9
    s_err = -1
    Dim typeTk As myToken
    For i = iBase + 1 To tkLst.Count
        Set tk = tkLst(i)
        Select Case s
        Case 0
            If tk.mType = tkType Then
                s = s_type
                Set typeTk = tk
            Else
                s = s_error
                Exit For
            End If
        Case s_type
            If tk.mType = tkRBR Then
                iNext = i
                Set castNode = crtOper(typeTk, ooTypeCast)
                s = s_succ
                Exit For
            Else
            End If
        End Select
        
    Next
    chkTypeCast = (s = s_succ)
End Function
'//debug
Private Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub

