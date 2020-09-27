VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function getFunc(zFile)
    Dim fso As New FileSystemObject
    Set ts = fso.OpenTextFile(zFile)
    zData = ts.ReadAll
    ts.Close
    
    arr = Split(zData, vbCrLf)
    s = 0
    Dim funcObj As myFuncCode
    Set funcObj = New myFuncCode
    Set funcObj.mLines = New Collection
    funcObj.mName = getFuncName(arr, 0)
    For i = 0 To UBound(arr)
        If arr(i) = "" Then
            GoTo lSkipEmpty
        End If
        
        If arr(i) = "}" Then
            s = 0
        End If
        
        If s = 1 Then
            Set lineObj = New myLineCode
            lineObj.mNo = i + 1
            lineObj.mTxt = arr(i)
            funcObj.mLines.Add lineObj
        End If
        
        If arr(i) = "{" Then
            s = 1
        End If
        
lSkipEmpty:
    Next
    
    Set getFunc = funcObj
End Function
Private Function getFuncName(arr, i)
    Dim tkz As New myTokenizer
    Set tLst = tkz.findTks(arr, i, Array(tkStatic, tkType, tkWord, tkLBR))
    If tLst.Count = 4 Then
        getFuncName = tLst(3).mTxt
    Else
        Set tLst = tkz.findTks(arr, i, Array(tkType, tkWord, tkLBR))
        If tLst.Count = 3 Then
            getFuncName = tLst(2).mTxt
        End If
    End If
End Function

Public Function parseFun(ByVal funcObj As myFuncCode)
    Dim fpar As New myFuncParser
    Set parseFun = fpar.parseFun(funcObj)
End Function
