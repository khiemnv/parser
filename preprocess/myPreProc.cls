VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myPreProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Set tLst = readFile(zFile)
    filestep tLst
End Function
Private Function filestep(tLst)
    Dim tkLst As Collection
    mIncLvl = mIncLvl + 1
    For Each tLine In tLst
        Set tkLst = New Collection
        mtkz.stepAll tLine.mTxt, tkLst
        Set tk = tkLst(1)
        If tk.mType = tkSharp Then
            linestep tkLst
        ElseIf mIncLvl > 1 Then
            'ignore include
        ElseIf mCurSw Then
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
        mCurSw = mCurSw And mCurCnd
    Case tkIfdef
        'push ctx
        mSwStack(mSwStackLvl) = Array(mCurSw, mCurCnd)
        mSwStackLvl = mSwStackLvl + 1
        'eval cnd ->mCurCnd
        mCurCnd = mDefined.Exists(tkLst(3).mTxt)
        mCurSw = mCurSw And mCurCnd
    Case tkIfndef
        'push ctx
        mSwStack(mSwStackLvl) = Array(mCurSw, mCurCnd)
        mSwStackLvl = mSwStackLvl + 1
        'eval cnd ->mCurCnd
        mCurCnd = Not mDefined.Exists(tkLst(3).mTxt)
        mCurSw = mCurSw And mCurCnd
    Case tkElif
        If Not mCurCnd Then
            'eval cnd ->mCurCnd
            mCurCnd = evalCnd(tkLst, 3)
            mCurSw = mCurSw And mCurCnd
        Else
            mCurSw = False
        End If
    Case tkElse
        mCurSw = mCurSw And (Not mCurCnd)
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
        End Select
    Next
    evalCnd = valLst(1)
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
    
    arr = Split(zData, vbCrLf)
    assert (UBound(arr) = n - 1)
    
    'remove empty/comment
    Dim tLine As myLineCode
    Dim reg As New RegExp
    reg.Pattern = "^(\s+|/\*.*\*/)*$"
    lineNo = 1
    Set tLst = New Collection
    For Each zLine In arr
        If zLine = "" Then GoTo lSkip
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
