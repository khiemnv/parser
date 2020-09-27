VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myEvalNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum eEvalNodeType
    enNone
    
    enVar
    enConst
    enNumber
    enOper
    enFunc
End Enum

Public mTk As myToken   'for debug
Public mId As String
Public mName As String
Public mType As eEvalNodeType
Public mOrder As eOperOrder
Public mChilds As Collection
Public mOper As String
