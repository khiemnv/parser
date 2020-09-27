VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myBlockCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum eBlockType
    blkNone
    blkIf
End Enum

Public mPath As String
Public mId As Long
Public mLst As Collection
Public mType As eBlockType

Private Sub Class_Initialize()
    Set mLst = New Collection
End Sub
