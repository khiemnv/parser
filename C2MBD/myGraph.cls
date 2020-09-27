VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mDict  As Dictionary         'node dict
Public mVarDict As Dictionary
Public mPath As String
Public mId As Long

Private Sub Class_Initialize()
    Set mDict = New Dictionary
    Set mVarDict = New Dictionary
End Sub
