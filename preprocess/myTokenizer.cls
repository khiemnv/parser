VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mDict As Dictionary
Private mReg As RegExp

Private Sub Class_Initialize()
    Set mDict = New Dictionary
    Set mReg = New RegExp
    mReg.Pattern = ">>|<<|\+\+|\-\-|>=|<=|\|\||&&|==|!=|[=\(\);,\{\}\+\-\*\/%<>:\[\]&\^!\|~#]|\w+"
    
    mDict.Add "=", eTokenType.tkAssign
    mDict.Add ")", eTokenType.tkRBR
    mDict.Add "(", eTokenType.tkLBR
    mDict.Add ":", eTokenType.tkColon
    mDict.Add ";", eTokenType.tkSemi
    mDict.Add ",", eTokenType.tkComma
    mDict.Add "{", eTokenType.tkLCBR
    mDict.Add "}", eTokenType.tkRCBR
    mDict.Add "+", eTokenType.tkAdd
    mDict.Add "-", eTokenType.tkSub
    mDict.Add "*", eTokenType.tkMul
    mDict.Add "%", eTokenType.tkMod
    mDict.Add "/", eTokenType.tkDiv
    mDict.Add ">", eTokenType.tkGT
    mDict.Add "<", eTokenType.tkLT
    mDict.Add "[", eTokenType.tkLSBR
    mDict.Add "]", eTokenType.tkRSBR
    mDict.Add "|", eTokenType.tkBitOr
    mDict.Add "&", eTokenType.tkBitAnd
    mDict.Add "^", eTokenType.tkBitXor
    mDict.Add "~", eTokenType.tkBitNot
    mDict.Add "!", eTokenType.tkNot
    mDict.Add "#", eTokenType.tkSharp

    mDict.Add ">>", eTokenType.tkRShift
    mDict.Add "<<", eTokenType.tkLShift
    mDict.Add "++", eTokenType.tkInc
    mDict.Add "--", eTokenType.tkDec
    mDict.Add ">=", eTokenType.tkGTE
    mDict.Add "<=", eTokenType.tkLTE
    mDict.Add "||", eTokenType.tkOr
    mDict.Add "&&", eTokenType.tkAnd
    mDict.Add "==", eTokenType.tkEQU
    mDict.Add "!=", eTokenType.tkNEQ
    mDict.Add "uint32", eTokenType.tkType
    mDict.Add "uint16", eTokenType.tkType
    mDict.Add "uint8", eTokenType.tkType
    mDict.Add "int32", eTokenType.tkType
    mDict.Add "int16", eTokenType.tkType
    mDict.Add "int8", eTokenType.tkType

    mDict.Add "static", eTokenType.tkStatic
    mDict.Add "return", eTokenType.tkReturn
    mDict.Add "else", eTokenType.tkElse
    mDict.Add "if", eTokenType.tkIf
    mDict.Add "switch", eTokenType.tkSw
    mDict.Add "case", eTokenType.tkCase
    mDict.Add "default", eTokenType.tkDefault
    mDict.Add "break", eTokenType.tkBreak
    mDict.Add "elif", eTokenType.tkElif
    mDict.Add "endif", eTokenType.tkEndif
    mDict.Add "define", eTokenType.tkDefine
    mDict.Add "defined", eTokenType.tkDefined
    mDict.Add "ifdef", eTokenType.tkIfdef
    mDict.Add "ifndef", eTokenType.tkIfndef
    mDict.Add "include", eTokenType.tkInclude

End Sub

Public Function step1(ByVal arr, ByVal i, ByVal zTxt, _
ByRef iNext, ByRef zRemain, ByRef tLst As Collection)
    If tLst Is Nothing Then Set tLst = New Collection
    For iNext = i To UBound(arr)
        Set m = mReg.Execute(zTxt)
        If m.Count > 0 Then
            zTk = m(0).Value
            eTk = eTokenType.tkWord
            If mDict.Exists(zTk) Then eTk = mDict(zTk)
            zRemain = mId(zTxt, m(0).FirstIndex + m(0).Length + 1)
            Dim tk As myToken
            Set tk = New myToken
            tk.mTxt = zTk
            tk.mType = eTk
            tLst.Add tk
            Exit For
        End If
        
        zTxt = arr(iNext + 1).mTxt
    Next
    
    step1 = Array(eTk, zTk)
End Function
Public Function stepAll(ByVal zTxt, ByRef tLst As Collection)
    If tLst Is Nothing Then Set tLst = New Collection
    mReg.Global = True
    Set m = mReg.Execute(zTxt)
    mReg.Global = False
    For Each mi In m
        zTk = mi.Value
        eTk = eTokenType.tkWord
        If mDict.Exists(zTk) Then eTk = mDict(zTk)
        Dim tk As myToken
        Set tk = New myToken
        tk.mTxt = zTk
        tk.mType = eTk
        tLst.Add tk
    Next
End Function
Public Function steps(ByVal arr, ByVal i, ByVal zTxt, _
ByRef iNext, ByRef zRemain, ByRef tLst As Collection, _
n)
    For i = 1 To n
        tk = step1(arr, i, zTxt, iNext, zRemain, tLst)
        i = iNext
        zTxt = zRemain
    Next
    steps = tk
End Function

Public Function findTks(arr, iBase, tks)
    Dim tLst As Collection
    i = iBase
    zTxt = arr(i)
    For Each eTk In tks
        tk = step1(arr, i, zTxt, iNext, zRemain, tLst)
        If eTk <> tk(0) Then
            Exit For
        End If
        zTxt = zRemain
        i = iNext
    Next
    Set findTks = tLst
End Function

