VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x
Public y
Public w
Public h
'//x1 y1 x2 y2
Public Property Let val(var As Variant)
    If IsArray(var) Then
        arr = var
    Else
        arr = Split(var, " ")
    End If
    x = CLng(arr(0))
    y = CLng(arr(1))
    w = CLng(arr(2)) - x
    h = CLng(arr(3)) - y
End Property
'//return x1 y1 x2 y2
Public Property Get val()
    val = x & " " & y & " " & x + w & " " & y + h
End Property
Public Property Let size(var As Variant)
    If IsArray(var) Then
        arr = var
    Else
        arr = Split(var, " ")
    End If
    w = arr(0)
    h = arr(1)
End Property
