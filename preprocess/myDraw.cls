VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2020-09-10: out revision
'2020-09-10: func
'2020-09-15: oper: mod, shift, bitwise
Public mCmdLst As Collection
Public Function genConvSub(cellD As Dictionary, _
nodeDict As Dictionary, _
Optional ByVal baseX = 0, _
Optional ByVal baseY = 0)
    If mCmdLst Is Nothing Then Set mCmdLst = New Collection
    Dim tCell As myCell
    Dim tNode As myNode
    ew = 80
    eh = 40
    dy = 5
    dx = 5
    
    Set tLst = New Collection   'in/out list
    Set fLst = New Collection   'func list
    Set cstLst = New Collection 'const list
    For i = 0 To nodeDict.Count - 1
        Set tNode = nodeDict.Items(i)
        Set tCell = cellD(tNode.mId)
                
        Select Case tCell.mNode.mType
            Case eIn
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zIn = crtIn(tCell.mNode.mId, tPos)
                tLst.Add tCell
            Case eOut
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zOut = crtOut(tCell.mNode.mId, tPos)
                tLst.Add tCell
            'bus
            Case eBusCrt
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, eh * (tCell.mNode.mInLst.Count - 0.5))
                zId = crtBusCrt(tCell.mNode.mId, tPos, tCell.mNode.mInLst.Count, tCell.mNode.mDataType)
            Case eMux
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, eh * (tCell.mNode.nInport - 0.5))
                zId = crtMux(tCell.mNode.mId, tPos, tCell.mNode.nInport)
            Case eDemux
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, eh * (tCell.mNode.nOutport - 0.5))
                zId = crtDemux(tCell.mNode.mId, tPos, tCell.mNode.nOutport)
            Case eTypeConv
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtTypecast(tCell.mNode.mId, tPos, tCell.mNode.mDataType)
            Case eVecConCat
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, eh * (tCell.y2 - tCell.y1 - 0.5))
                zId = crtVecConCat(tCell.mNode.mId, tPos, tCell.mNode.nInport, 2)
            'oper
            Case eProduct
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtProduct(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
             Case eSum
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtSum(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
            Case eMath
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtMath(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
            Case eShift
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtShift(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
            Case eRelation
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtRelation(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
            Case eBitwise
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtBitwise(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
            Case eLogical
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtLogical(tCell.mNode.mId, tPos, tCell.mNode.mOperator)
           'sw
            Case eSw
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtSw(tCell.mNode.mId, tPos)
            Case eMulSw
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, eh * tCell.mNode.mInLst.Count)
                zId = crtMulSw(tCell.mNode.mId, tPos, tCell.mNode.mDataPortIndices, tCell.mNode.mDataPortForDefault)
            Case eConst
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtConst(tCell.mNode.mId, tPos, tCell.mNode.mName)
                'cstLst.Add tCell
            Case eFunc
                Set tPos = New myPos
                tPos.val = Array(baseX + tCell.x2 * ew, baseY + tCell.y1 * eh, 0, 0)
                tPos.size = Array(40, 20)
                zId = crtFunc(tCell.mNode.mId, tPos, tCell.mNode.nInport, tCell.mNode.nOutport)
                fLst.Add tCell
            Case Else
                assert (False)
        End Select
    Next
    
    For i = 0 To nodeDict.Count - 1
        Set tNode = nodeDict.Items(i)
        n = 1
        If tNode.mInLst Is Nothing Then GoTo skipIn
        If tNode.mCnnInfo Is Nothing Then
            For Each child In tNode.mInLst
                crtLine child.mId, 1, tNode.mId, n
                n = n + 1
            Next
        Else
            For Each rec In tNode.mCnnInfo
                '//[in/out][n] - [src/des][n]
                crtLine rec(2), rec(3), tNode.mId, rec(1)
            Next
        End If
skipIn:
    Next
    
    'rename in/out
    For Each tCell In tLst
        newName = tCell.mNode.mName
        If tCell.mNode.mVersion > 0 Then
            newName = newName & "_" & tCell.mNode.mVersion
        End If
        renameInOut tCell.mNode.mId, newName
    Next
    
    'rename func
    Dim tDict As New Dictionary
    For Each tCell In fLst
        zName = tCell.mNode.mName
        If tDict.Exists(zName) Then
            tDict(zName) = tDict(zName) + 1
            newName = zName & "_" & tDict(zName)
            renameInOut tCell.mNode.mId, newName
        Else
            tDict.Add zName, 0
            renameInOut tCell.mNode.mId, zName
        End If
    Next
    
    'declare const
    i = 1
    For Each tCell In cstLst
        declareConst tCell.mNode.mName, i
        i = i + 1
    Next
    
    Set genConvSub = mCmdLst
End Function
'simulink/Signal Routing
Public Function crtSw(zId, ByVal pos As myPos)
    zTmp = "b = add_block('simulink/Signal Routing/Switch',[sys '/<id>']);"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtSw = zId
End Function
Public Function crtMulSw(zId, ByVal pos As myPos, zDataPortIndices, zDataPortForDefault)
    zTmp = "b = add_block('simulink/Signal Routing/Multiport Switch',[sys '/<id>']);"
    zTmp1 = "set_param(b, 'DataPortIndices', '<idx>');"
    zTmp2 = "set_param(b, 'DataPortForDefault', '<def>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add "set_param(b, 'DataPortOrder', 'Specify indices');"
    mCmdLst.Add Replace(zTmp1, "<idx>", zDataPortIndices)
    mCmdLst.Add Replace(zTmp2, "<def>", zDataPortForDefault)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtMulSw = zId
End Function
Public Function crtGoTo(zId, ByVal pos As myPos, tag)
    zTmp = "b = add_block('simulink/Signal Routing/Goto',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'GotoTag','<tag>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<tag>", tag)
    mCmdLst.Add "set_param(b,'ShowName','off');"
    crtGoTo = zId
End Function
Private Function joinLst(lst, Optional spl = " ")
    zTxt = ""
    For Each i In lst
        zTxt = zTxt & spl & i
    Next
    joinLst = mId(zTxt, 2)
End Function
Public Function crtBusSel(zId, ByVal pos As myPos, lst)
    zTmp = "b = add_block('simulink/Signal Routing/Bus Selector',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'OutputSignals','<lst>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<lst>", joinLst(lst, ","))
    crtBusSel = zId
End Function
Public Function crtBusCrt(zId, ByVal pos As myPos, n, _
zType)
    If InStr(zType, "Bus:") < 1 Then zType = "Bus: " & zType
    
    zTmp = "b = add_block('simulink/Signal Routing/Bus Creator',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'Inputs','<n>');"
    zTmp4 = "set_param(b,'OutDataTypeStr','<type>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<n>", n)
    mCmdLst.Add Replace(zTmp4, "<type>", zType)
    mCmdLst.Add "set_param(b,'InheritFromInputs','off');"
    crtBusCrt = zId
End Function
Public Function crtVecConCat(zId, ByVal pos As myPos, n, d)
    zTmp = "b = add_block('simulink/Signal Routing/Vector Concatenate',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'NumInputs','<n>');"
    zTmp4 = "set_param(b,'ConcatenateDimension','<d>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<n>", n)
    mCmdLst.Add Replace(zTmp4, "<d>", d)
    mCmdLst.Add "set_param(b,'Mode','Multidimensional array');"
    crtVecConCat = zId
End Function
'simulink/Sources
Public Function crtFromWS(zId, ByVal pos As myPos, zVar)
    zTmp = "b = add_block('simulink/Sources/From Workspace',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'VariableName','<var>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<var>", zVar)
    mCmdLst.Add "set_param(b,'SampleTime','-1');"
    crtFromWS = zId
End Function
Public Function crtConst(zId, ByVal pos As myPos, zVal)
    zTmp = "b = add_block('simulink/Sources/Constant',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'Value','<val>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<val>", zVal)
    crtConst = zId
End Function
Public Function crtTypecast(zId, ByVal pos As myPos, zType)
    zTmp = "b = add_block('simulink/Signal Attributes/Data Type Conversion',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b,'OutDataTypeStr','<type>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<type>", zType)
    crtTypecast = zId
End Function
Public Function crtSubsystem(zId, ByVal pos As myPos)
    zTmp = "b = add_block('simulink/Ports & Subsystems/Subsystem',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    crtSubsystem = zId
End Function
Public Function crtIn(zId, ByVal pos As myPos)
    zTmp = "b = add_block('simulink/Ports & Subsystems/In1',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    crtIn = zId
End Function
Public Function crtOut(zId, ByVal pos As myPos)
    zTmp = "b = add_block('simulink/Ports & Subsystems/Out1',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    crtOut = zId
End Function
Public Function crtDemux(zId, ByVal pos As myPos, n)
    zTmp = "b = add_block('simulink/Signal Routing/Demux',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b, 'Outputs', '<n>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<n>", n)
    crtDemux = zId
End Function
Public Function crtMux(zId, ByVal pos As myPos, n)
    zTmp = "b = add_block('simulink/Signal Routing/Mux',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Position', '[<pos>]');"
    zTmp3 = "set_param(b, 'Inputs', '<n>');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<pos>", pos.val)
    mCmdLst.Add Replace(zTmp3, "<n>", n)
    crtMux = zId
End Function
Public Function crtModel(zId, ByVal pos As myPos, zModelName)
    zTmp = "b = add_block('simulink/Ports & Subsystems/Model',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'ModelName', '<name>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<name>", zModelName)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtModel = zId
End Function
'simulink/Logic and Bit Operations
Public Function crtShift(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Logic and Bit Operations/Shift Arithmetic',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'BitShiftDirection', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtShift = zId
End Function
Public Function crtRelation(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Logic and Bit Operations/Relational Operator',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Operator', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtRelation = zId
End Function
Public Function crtBitwise(zId, ByVal pos As myPos, zOper)
    nPort = 2
    If zOper = "NOT" Then nPort = 1
    zTmp = "b = add_block('simulink/Logic and Bit Operations/Bitwise Operator',[sys '/<id>']);"
    zTmp1 = "set_param(b, 'NumInputPorts', '<n>');"
    zTmp2 = "set_param(b, 'logicop', '<op>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add "set_param(b, 'UseBitMask', 'off');"
    mCmdLst.Add Replace(zTmp1, "<n>", nPort)
    mCmdLst.Add Replace(zTmp2, "<op>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtBitwise = zId
End Function
Public Function crtLogical(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Logic and Bit Operations/Logical Operator',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Operator', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtLogical = zId
End Function
'sum/product
'simulink/Math Operations
Public Function crtProduct(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Math Operations/Product',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Inputs', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtProduct = zId
End Function
Public Function crtMath(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Math Operations/Math Function',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Operator', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtMath = zId
End Function
Public Function crtSum(zId, ByVal pos As myPos, zOper)
    zTmp = "b = add_block('simulink/Math Operations/Add',[sys '/<id>']);"
    zTmp2 = "set_param(b, 'Inputs', '<oper>');"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    mCmdLst.Add Replace(zTmp, "<id>", zId)
    mCmdLst.Add Replace(zTmp2, "<oper>", zOper)
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    crtSum = zId
End Function
Public Function crtFunc(zId, ByVal pos As myPos, nIn, nOut)
    'simulink/Ports & Subsystems/Subsystem
    zTmp = "b = add_func([sys '/<id>'],<in>,<out>);"
    zTmp3 = "set_param(b, 'Position', '[<pos>]');"
    zTmp = Replace(zTmp, "<id>", zId)
    zTmp = Replace(zTmp, "<in>", nIn)
    zTmp = Replace(zTmp, "<out>", nOut)
    mCmdLst.Add zTmp
    mCmdLst.Add Replace(zTmp3, "<pos>", pos.val)
    zTmp = zId
End Function
Public Function crtLine(src, iOut, des, iIn)
    zTmp = "add_line(sys ,'<src>','<des>','autorouting','on');"
    zTmp = Replace(zTmp, "<src>", src & "/" & iOut)
    zTmp = Replace(zTmp, "<des>", des & "/" & iIn)
    mCmdLst.Add zTmp
End Function

Public Function renameInOut(zId, zName)
    zTmp = "set_param([sys '/<id>'],'Name', '<name>');"
    zTmp = Replace(zTmp, "<name>", zName)
    mCmdLst.Add Replace(zTmp, "<id>", zId)
End Function

Public Function declareConst(zName, val)
    zTmp = "declare_const('<name>',<val>);"
    zTmp = Replace(zTmp, "<name>", zName)
    mCmdLst.Add Replace(zTmp, "<val>", val), before:=1
End Function
