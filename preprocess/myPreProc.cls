VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myPreProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'fix parent SW

Dim mtkz As myTokenizer
Dim mDefined As Dictionary
Dim mPathLst As Collection

'include
Dim mIncLvl As Integer

'sw
Dim mSwStack(20)
Dim mSwStackLvl As Integer
Dim mCurSw As Boolean
Dim mCurCnd As Boolean

'output
Public mLines  As Collection

Public Function RemoveSw(zFile)
    Dim tLine As myLineCode
    Dim tkLst As Collection
    Dim tk As myToken
    
    'init path
    If mPathLst Is Nothing Then Set mPathLst = New Collection
    Dim fso As New FileSystemObject
    zDir = fso.GetParentFolderName(zFile)
    mPathLst.Add zDir
    
    Set mtkz = New myTokenizer
    mSwStackLvl = 0
    mIncLvl = 0
    mCurSw = True
    Set mDefined = New Dictionary
    'init output
    Set mLines = New Collection
    
    'parse file
    Set tLst = readFile2(zFile)
    filestep tLst
End Function
Private Function filestep(tLst)
    Dim tkLst As Collection
    mIncLvl = mIncLvl + 1
    Dim tLine As myLineCode
    For Each tLine In tLst
        bAdd = True
        If tLine.isComment Then
            '//not parse
        Else
            Set tkLst = New Collection
            mtkz.stepAll tLine.mTxt, tkLst
            Set tk = tkLst(1)
            If tk.mType = tkSharp Then
                bAdd = False
                linestep tkLst
            End If
        End If
        
        If mIncLvl > 1 Then
            'ignore include
            bAdd = False
        ElseIf mCurSw Then
            bAdd = False
        End If
        
        If bAdd Then
            mLines.Add tLine
        End If
    Next
    mIncLvl = mIncLvl - 1
End Function
Private Function linestep(tkLst)
    Dim tk As myToken
    Set tk = tkLst(2)
    
    '#if else
    Select Case tk.mType
    Case tkIf
        'push ctx
        mSwStack(mSwStackLvl) = Array(mCurSw, mCurCnd)
        mSwStackLvl = mSwStackLvl + 1
        'eval cnd ->mCurCnd
        mCurCnd = evalCnd(tkLst, 3)
        parentSw = mCurCnd
        mCurSw = mCurSw And parentSw
    Case tkIfdef
        'push ctx
        mSwStack(mSwStackLvl) = Array(mCurSw, mCurCnd)
        mSwStackLvl = mSwStackLvl + 1
        'eval cnd ->mCurCnd
        mCurCnd = mDefined.Exists(tkLst(3).mTxt)
        parentSw = mCurCnd
        mCurSw = mCurSw And parentSw
    Case tkIfndef
        'push ctx
        mSwStack(mSwStackLvl) = Array(mCurSw, mCurCnd)
        mSwStackLvl = mSwStackLvl + 1
        'eval cnd ->mCurCnd
        mCurCnd = Not mDefined.Exists(tkLst(3).mTxt)
        parentSw = mCurCnd
        mCurSw = mCurSw And parentSw
    Case tkElif
        If Not mCurCnd Then
            'eval cnd ->mCurCnd
            mCurCnd = evalCnd(tkLst, 3)
            parentSw = mSwStack(mSwStackLvl - 1)(0)
            mCurSw = mCurSw And parentSw
        Else
            mCurSw = False
        End If
    Case tkElse
        If Not mCurCnd Then
            parentSw = mSwStack(mSwStackLvl - 1)(0)
            mCurSw = parentSw
        Else
            mCurSw = False
        End If
    Case tkEndif
        'pop ctx
        mSwStackLvl = mSwStackLvl - 1
        rec = mSwStack(mSwStackLvl)
        mCurSw = rec(0)
        mCurCnd = rec(1)
    End Select
    
    'sw on
    If mCurSw Then
        Select Case tk.mType
        Case tkDefine
            zKey = tkLst(3).mTxt
            zVal = ""
            If tkLst.Count > 3 Then
                zVal = tkLst(4).mTxt
            End If
            mDefined.Add zKey, zVal
        Case tkInclude
            'parse recursive
            zVal = tkLst(3).mTxt
            Set lineLst = readInclude(zVal)
            filestep lineLst
        End Select
    End If
End Function
Private Function evalCnd(tkLst, iStart)
    Dim tk As myToken
    Set valLst = New Collection
    Set opLst = New Collection
    
    For i = iStart To tkLst.Count
        Set tk = tkLst(i)
        Select Case tk.mType
            Case tkDefined   'defined(xxx)
                zDef = tkLst(i + 2).mTxt
                tVal = mDefined.Exists(zDef)
                valLst.Add tVal
                i = i + 3
            Case tkLBR
                opLst.Add tkLBR
            Case tkRBR
                'reduce to (
                reduce valLst, opLst, tkRBR
            Case tkNot, tkAnd, tkOr
                reduce valLst, opLst, tk.mType
                opLst.Add tk.mType
            Case tkWord
                '//#if 0
                If tk.mTxt = "0" Then
                    valLst.Add False
                ElseIf tk.mTxt = "1" Then
                    valLst.Add True
                End If
        End Select
    Next
    reduce valLst, opLst, tkRBR
    assert (valLst.Count = 1)
    evalCnd = valLst(1)
End Function
Private Function reduce(ByRef valLst, ByRef opLst, _
Optional ByVal topOp = tkInvalid)
    stopTk = tkInvalid

    If topOp = tkInvalid Then
        'pop
        topOp = opLst(opLst.Count)
        opLst.Remove (opLst.Count)
    End If
    
    If topOp = tkRBR Then
        stopTk = tkLBR
    End If
    
    For i = opLst.Count To 1 Step -1
        curOp = opLst(opLst.Count)
        If stopTk = tkInvalid Then
            If curOp = tkLBR Then
                Exit For
            ElseIf topOp < curOp Then
                Exit For
            End If
        Else
            '()
            If stopTk = curOp Then
                opLst.Remove opLst.Count
                Exit For
            End If
        End If
        
        'pop top oper
        opLst.Remove opLst.Count
        'calc
        Select Case curOp
        Case tkNot
            'pop
            tVal = valLst(valLst.Count)
            valLst.Remove (valLst.Count)
            'push
            valLst.Add (Not tVal)
        Case tkAnd
            'pop
            tVal = valLst(valLst.Count)
            valLst.Remove (valLst.Count)
            tVal2 = valLst(valLst.Count)
            valLst.Remove (valLst.Count)
            'push
            valLst.Add (tVal And tVal2)
        Case tkOr
            'pop
            tVal = valLst(valLst.Count)
            valLst.Remove (valLst.Count)
            tVal2 = valLst(valLst.Count)
            valLst.Remove (valLst.Count)
            'push
            valLst.Add (tVal Or tVal2)
        End Select
    Next
End Function

Private Function readInclude(zInc)
    Dim fso As New FileSystemObject
    Set tLst = New Collection
    For Each zDir In mPathLst
        fullpath = zDir & "\" & zInc & ".h"
        If fso.FileExists(fullpath) Then
            Set tLst = readFile(fullpath)
            Exit For
        End If
    Next
    Set readInclude = tLst
End Function

Private Function readFile(zFile)
    Dim fso As New FileSystemObject
    Set ts = fso.OpenTextFile(zFile)
    zData = ts.ReadAll
    n = ts.Line
    ts.Close
    
    
    Dim reg2 As New RegExp
    reg2.Pattern = "/\*.*\*/|//[^\r\n]*"
    reg2.Global = True
    reg2.MultiLine = True
    
    zTxt = reg2.Replace(zData, "")
    arr = Split(zData, vbCrLf)
    assert (UBound(arr) = n - 1)
    
    'remove empty/comment
    Dim tLine As myLineCode
    
    Dim reg As New RegExp
    reg.Pattern = "^[\s\t]*$"
    
    lineNo = 1
    Set tLst = New Collection
    For Each zLine In arr
        If reg.Test(zLine) Then GoTo lSkip
        
        Set tLine = New myLineCode
        tLine.mNo = lineNo
        tLine.mTxt = zLine
        tLst.Add tLine
lSkip:
        lineNo = lineNo + 1
    Next
    Set readFile = tLst
End Function

Private Function readFile2(zFile)
    Dim fso As New FileSystemObject
    Set ts = fso.OpenTextFile(zFile)
    zData = ts.ReadAll
    n = ts.Line
    ts.Close
    
    arr = Split(zData, vbCrLf)
    assert (UBound(arr) = n - 1)
    
    'skip empty line
    Dim tLine As myLineCode
    
    'remove comment
    Dim reg2 As New RegExp
    reg2.Pattern = "^([\s\t\n\r]|/\*.*\*/|//.*)+$"
    
    Dim reg As New RegExp
    reg.Pattern = "^[\s\n\r\t]*$"
    
    lineNo = 1
    Set tLst = New Collection
    For Each zLine In arr
        '//skip empty
        If reg.Test(zLine) Then GoTo lSkip
        
        Set tLine = New myLineCode
        tLine.mNo = lineNo
        tLine.mTxt = zLine
        tLine.isComment = reg2.Test(zLine)
        tLst.Add tLine
lSkip:
        lineNo = lineNo + 1
    Next
    Set readFile2 = tLst
End Function
