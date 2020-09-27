VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum eNodeType
    eNone
    eBusCrt
    eBusSel
    eIn
    eOut
    eTypeConv
    eDemux
    eMux
    eSubSys
    eGoto
    eFrom
    eVecConCat
    eSw
    eMulSw
    eSwCase
    eIf
    eIfAction
    eMath   '%
    eProduct
    eSum
    eShift
    eRelation
    eBitwise
    eLogical
    eConst
    
    eFunc
End Enum

Public mType As eNodeType
Public mName As String          'in/out var
Public mVersion As Integer      'update version x++

Public mCmnt As String
Public mId As String            'unique id
Public mInLst As Collection
Public mOutLst As Collection
Public mDataType As String      'bus, TypeConv
Public nOutport                 'demux, in
Public nInport
Public mOperator                'sum, product, ...
'//[in/out][n] - [src/des][n]
Public mCnnInfo As Collection   'demux, mux
Public mExt As Dictionary '<param, value>
'//multiport switch
Public mDataPortIndices
Public mDataPortForDefault

Public Function addCnnInfo(rec)
    If mCnnInfo Is Nothing Then Set mCnnInfo = New Collection
    mCnnInfo.Add rec
End Function

