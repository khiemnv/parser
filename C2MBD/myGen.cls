VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'gen sub harness
'2020-07-30;Khiemnv1;align bus horizontal
Public mX0
Public mY0
Public mEw
Public mEh
Public mInIdx
Public mMuxId
Public mCastId
Public mBusId
Public mOutIdx

Private par As myParser
Private rd As myDraw
'//__in inDict: <string, myInput>
Public Function genInConv2(ByVal inDict As Dictionary, ByVal cmdLst)
    Dim gd As New myGrid
    Dim dr As New myDraw
    Set dr.mCmdLst = cmdLst
    
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Dim d3 As Dictionary
    Set d1 = inDict
    
    Set subLst = New Collection
    subLst.Add d1
    
    '//split
    Set subLst = splitByIn(d1, 100)
    
    '//gen
    Dim fso As New FileSystemObject
    Dim tDict As Dictionary
    
    baseX = 0
    baseY = 0

    For Each tDict In subLst
        Set d2 = gd.arange(tDict)
        Set cmdLst = dr.genConvSub(d2, tDict, baseX, baseY)
        baseX = baseX + 80 * 5
    Next

End Function

'debug
Private Sub assert(expr, Optional msg)
    If Not expr Then
        ret = 1 / 0
    End If
End Sub
