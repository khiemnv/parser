VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myFuncParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2020-09-10:if/else parse
Private Enum parseSts
    stsInit
    stsIf
    stsIfEnd
    stsElse
    
    stsSw
    stsCase
    stsCaseEnd
    stsCaseBreak
End Enum

Dim mCurLvl
Dim mCtxStack(20)
Dim mCurLine As myLineCode
Dim mCurB As myBlockCode
Dim mCurStmt As myStmt
Dim mCurSts As parseSts

Public Function parseFun(funcObj As myFuncCode)
    Dim tkz As New myTokenizer
    Dim tLst As Collection
    Dim tk As myToken

    'init
    mCurLvl = 0
    mCurSts = stsInit
    Set mCurB = New myBlockCode
    mCurB.mPath = funcObj.mName
    Set mCurStmt = New myStmt
    mCtxStack(0) = Array(mCurB)
    
    For i = 1 To funcObj.mLines.Count
        Set mCurLine = funcObj.mLines(i)
        
        Set tLst = New Collection
        tkz.stepAll mCurLine.mTxt, tLst
        
        For Each tk In tLst
            step tk
        Next
    Next
    
    Set parseFun = mCtxStack(0)(0)
End Function
Private Function pushCtx()
    Set parentB = mCurB
    mCtxStack(mCurLvl) = Array(mCurB, mCurLvl, mCurSts, mCurStmt)
    mCurLvl = mCurLvl + 1
    mCurSts = stsInit
    Set mCurB = New myBlockCode
    mCurB.mId = mCurLine.mNo
    mCurB.mPath = parentB.mPath & "/" & mCurLine.mNo
    Set mCurStmt = New myStmt
    mCurStmt.mId = mCurLine.mNo
End Function
Private Function popCtx()
    Set childB = mCurB
    mCtxStack(mCurLvl) = Array(mCurB, mCurLvl, mCurSts, mCurStmt)
    'pop
    rec = mCtxStack(mCurLvl - 1)
    Set mCurB = rec(0)
    mCurLvl = mCurLvl - 1
    mCurSts = rec(2)
    Set mCurStmt = rec(3)
    
    If mCurSts = stsIf Then
        Set ifStmt = mCurStmt
        ifStmt.mBlockLst.Add childB
    ElseIf mCurSts = stsElse Then
        Set ifStmt = mCurStmt
        ifStmt.mBlockLst.Add childB
    ElseIf mCurSts = stsSw Then
        Set swStmt = mCurStmt
        swStmt.mBlockLst.Add childB
    ElseIf mCurSts = stsCaseEnd Then
        Set caseStmt = mCurStmt
        caseStmt.mBlockLst.Add childB
    End If
End Function
Private Function step(ByVal tk As myToken)
    Dim eTk As eTokenType
    eTk = tk.mType
        
    '{}
    If eTk = tkLCBR Then
        'push ctx
        pushCtx
    ElseIf eTk = tkRCBR Then
        'pop
        popCtx
    End If
    
    'if/else
    Select Case mCurSts
        Case stsInit
            If eTk = tkIf Then
                Set ifStmt = mCurStmt
                ifStmt.mId = mCurLine.mNo
                ifStmt.mType = stmtIf
                Set ifStmt.mCndLst = New Collection
                Set ifStmt.mBlockLst = New Collection
                mCurSts = stsIf

                Set cndStmt = New myStmt
                cndStmt.mId = mCurLine.mNo
                ifStmt.mCndLst.Add cndStmt
            ElseIf eTk = tkSw Then
                Set swStmt = mCurStmt
                swStmt.mId = mCurLine.mNo
                swStmt.mType = stmtSw
                Set swStmt.mCndLst = New Collection
                Set swStmt.mBlockLst = New Collection
                mCurSts = stsSw
                
                Set cndStmt = New myStmt
                cndStmt.mId = mCurLine.mNo
                swStmt.mCndLst.Add cndStmt
            ElseIf eTk = tkCase Then
                Set caseStmt = mCurStmt
                caseStmt.mId = mCurLine.mNo
                Set caseStmt.mCndLst = New Collection
                Set caseStmt.mBlockLst = New Collection
                mCurSts = stsCase
            ElseIf eTk = tkDefault Then
                Set caseStmt = mCurStmt
                caseStmt.mId = mCurLine.mNo
                Set caseStmt.mCndLst = New Collection
                Set caseStmt.mBlockLst = New Collection
                caseStmt.mCndLst.Add tk
                mCurSts = stsCase
            ElseIf eTk = tkBreak Then
                mCurSts = stsCaseBreak
            ElseIf eTk = tkSemi Then
                mCurB.mLst.Add mCurStmt
                mCurStmt.mId = mCurLine.mNo
                Set mCurStmt = New myStmt
            ElseIf eTk = tkLCBR Then
                'begin block
            Else
                mCurStmt.mLst.Add tk
            End If
        Case stsIf, stsIfEnd, stsElse
            stepIf tk
        'sw
        Case stsSw, stsCase, stsCaseBreak, stsCaseEnd
            stepSw tk
    End Select
End Function
Private Function stepSw(tk)
    Dim caseStmt As myStmt
    Dim swStmt As myStmt
    
    Dim eTk As eTokenType
    eTk = tk.mType
    Select Case mCurSts
    Case stsSw
        If eTk = tkRCBR Then
            'end sw
            Set swStmt = mCurStmt
            mCurB.mLst.Add swStmt
            
            'init state
            mCurSts = stsInit
            Set mCurStmt = New myStmt
            mCurStmt.mId = mCurLine.mNo
        Else
            '(...)
            Set swStmt = mCurStmt
            Set cndStmt = swStmt.mCndLst(swStmt.mCndLst.Count)
            cndStmt.mLst.Add tk
        End If
    Case stsCase
        If eTk = tkColon Then
            mCurSts = stsCaseEnd
        Else
            Set caseStmt = mCurStmt
            caseStmt.mCndLst.Add tk
        End If
    Case stsCaseEnd
        If eTk = tkCase Then
            mCurSts = stsCase
        Else
            'push ctx
            pushCtx
            mCurStmt.mLst.Add tk
        End If
    Case stsCaseBreak
        'break;
        popCtx
        Set caseStmt = mCurStmt
        mCurB.mLst.Add caseStmt
        
        mCurSts = stsInit
        Set mCurStmt = New myStmt
        mCurStmt.mId = mCurLine.mNo
        
    End Select
End Function
Private Function stepIf(tk)
    Dim eTk As eTokenType
    eTk = tk.mType
    Select Case mCurSts
    Case stsIf
        If eTk = tkRCBR Then
            mCurSts = stsIfEnd
        Else
            '(...)
            Set ifStmt = mCurStmt
            Set cndStmt = ifStmt.mCndLst(ifStmt.mCndLst.Count)
            cndStmt.mLst.Add tk
        End If
    Case stsIfEnd
        If eTk = tkElse Then
            mCurSts = stsElse
        Else
            'end if
            Set ifStmt = mCurStmt
            mCurB.mLst.Add ifStmt
            
            'init state
            mCurSts = stsInit
            Set mCurStmt = New myStmt
            mCurStmt.mId = mCurLine.mNo
            mCurStmt.mLst.Add tk
        End If
    Case stsElse
        If eTk = tkIf Then
            mCurSts = stsIf
            
            Set ifStmt = mCurStmt
            Set cndStmt = New myStmt
            cndStmt.mId = mCurLine.mNo
            ifStmt.mCndLst.Add cndStmt
        ElseIf eTk = tkRCBR Then
            'end if/else
            Set ifStmt = mCurStmt
            mCurB.mLst.Add ifStmt
            
            'init state
            mCurSts = stsInit
            Set mCurStmt = New myStmt
            mCurStmt.mId = mCurLine.mNo
        End If
    End Select
End Function
