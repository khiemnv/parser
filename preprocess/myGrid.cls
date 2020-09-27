VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mDict As Dictionary  '<key, cell>
Private mTraceLst As Collection 'cell collection
Private mDepth

Public Function arange(nodeD As Dictionary, _
Optional ByRef nCol, _
Optional ByRef nRow)
    Set mDict = New Dictionary
    Set mTraceLst = New Collection
    
    Dim tNode As myNode
    Dim tCell As myCell
    
    '//crt cell dict
    Set tLst = New Collection
    For i = 0 To nodeD.Count - 1
        Set tNode = nodeD.Items(i)
        Set tCell = New myCell
        tCell.mId = tNode.mId
        Set tCell.mNode = tNode
        mDict.Add tCell.mId, tCell
        If Not (tNode.mInLst Is Nothing) Then tLst.Add tNode
    Next
    For Each tNode In tLst
        For Each e In tNode.mInLst
            link mDict(tNode.mId), mDict(e.mId)
        Next
    Next
    
    '//y calculate
    '//dfs trace
    Set hLst = New Collection   'head
    Set tLst = New Collection   'tail
    For i = 0 To mDict.Count - 1
        Set tCell = mDict.Items(i)
        If tCell.parentLst.Count = 0 Then hLst.Add tCell
    Next
    nRow = 1
    nCol = 1
    mDepth = 0
    For Each tCell In hLst
        nRow = traceDFS(tCell, nRow, nCol)
        nRow = nRow + 1
    Next
    
    '//x calculate
    max_count = mTraceLst.Count
    For i = 1 To max_count
        Set tLst = New Collection
        nCol = 1
        For Each tCell In mTraceLst
            If tCell.childLst.Count = 0 Then
                tLst.Add tCell
                If tCell.x2 > nCol Then nCol = tCell.x2
            End If
        Next
        For Each tCell In tLst
            tCell.x2 = mDepth - nCol + 1
            mTraceLst.Remove tCell.mId
            For Each tparent In tCell.parentLst
                ulink tparent, tCell
            Next
        Next
        If mTraceLst.Count = 0 Then Exit For
    Next
    
    Set arange = mDict
    nCol = mDepth
End Function

Private Function traceDFS(parent As myCell, _
Optional ByVal nRow = 1, Optional ByVal nCol = 1)
    If nCol > mDepth Then mDepth = nCol
    
    On Error GoTo lCrash
    mTraceLst.Add parent, parent.mId
    '//skip if traced
    parent.y1 = nRow
    parent.x1 = nCol
lCrash:
    If parent.x2 > nCol Then nCol = parent.x2
    parent.y2 = nRow
    parent.x2 = nCol
    
    Dim tChild As myCell
    For Each tChild In parent.childLst
        nRow = traceDFS(tChild, nRow, nCol + 1)
        nRow = nRow + 1
    Next

    If parent.childLst.Count > 0 Then
        Set LastChild = parent.childLst(parent.childLst.Count)
        traceDFS = LastChild.y2
        parent.y2 = traceDFS
    Else
        traceDFS = parent.y2
    End If
End Function

Private Function link(ByVal parent As myCell, ByVal child As myCell)
    parent.childLst.Add child, child.mId
    child.parentLst.Add parent, parent.mId
End Function
Private Function ulink(ByVal parent As myCell, ByVal child As myCell)
    parent.childLst.Remove child.mId
    child.parentLst.Remove parent.mId
End Function
