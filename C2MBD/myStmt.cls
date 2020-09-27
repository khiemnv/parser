VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum eStmtType
    stmtNone
    stmtDeclare
    stmtReturn
    
    stmtIf
    stmtSw
    stmtAssign
    stmtFor
End Enum

Public mCndLst As Collection   'if, elseif condition
Public mBlockLst As Collection    'if, elseif, else block

Public mLine As myLineCode  'debug
Public mId As Long          '=lineNo
Public mLst As Collection   'tk list
Public mType As eStmtType
Public mVal As myEvalNode

Private Sub Class_Initialize()
    Set mLst = New Collection
End Sub
