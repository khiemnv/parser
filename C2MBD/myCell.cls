VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mId As String
Public x1
Public x2
Public y1
Public y2
Public mNode As myNode
Public parentLst As Collection
Public childLst As Collection

Private Sub Class_Initialize()
    Set parentLst = New Collection
    Set childLst = New Collection
End Sub
