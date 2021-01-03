Attribute VB_Name = "global_subs"
Public Sub AllPoolDataLoad() '全图图池数据录入（用于地图名的录入）
Dim i As Integer: Dim j As Integer: Dim strx As String
i = 0: j = 0

Open App.Path & "\map pool data\all racing maps_number.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, strx
    AllMapN_R(i) = Int(strx)
    i = i + 1
Loop
i = 0: j = 0
Close #1

Open App.Path & "\map pool data\all racing maps_order.txt" For Input As #1
Do While Not EOF(1)
   If AllMapN_R(j) <> 0 Then
      For i = 0 To AllMapN_R(j) - 1
         Line Input #1, strx
         AllMaporder_R(j, i) = Int(strx)
      Next
   End If
   j = j + 1
Loop
Close #1

i = 0
Open App.Path & "\item pool data\all item maps_number.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, strx
    AllMapN_I(i) = Int(strx)
    i = i + 1
Loop
i = 0: j = 0
Close #1

Open App.Path & "\item pool data\all item maps_order.txt" For Input As #1
Do While Not EOF(1)
   If AllMapN_I(j) <> 0 Then
      For i = 0 To AllMapN_I(j) - 1
         Line Input #1, strx
         AllMaporder_I(j, i) = Int(strx)
      Next
   End If
   j = j + 1
Loop
Close #1


End Sub


Public Sub PoolDataLoad_R(ByVal name As String) '竞速图池数据录入

Dim strx As String
Dim i As Integer: Dim j As Integer: Dim n As Integer
i = 0: j = 0: n = 0

Open App.Path & "\map pool data\" & name & "_number.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, strx
   MapN_R(i) = Int(strx)
   i = i + 1
Loop
Close #1

i = 0: j = 0
Open App.Path & "\map pool data\" & name & "_order.txt" For Input As #2
Do While Not EOF(2)
   If MapN_R(j) <> 0 Then
      For i = 0 To MapN_R(j) - 1
         Line Input #2, strx
         Maporder_R(j, i) = Int(strx)
      Next
   End If
   j = j + 1
Loop
Close #2

i = 0: j = 0: n = 0
Open App.Path & "\map pool data\mapname.txt" For Input As #3
Do While Not EOF(3)
   If AllMaporder_R(j, i) <> 0 Then
      Line Input #3, strx
      If Maporder_R(j, n) = AllMaporder_R(j, i) Then
         Mapname_R(j, n) = strx
         n = n + 1
      End If
      i = i + 1
   Else
      j = j + 1
      i = 0
      n = 0
   End If
Loop
Close #3

all_map_R = 0
For i = 0 To theme
    all_map_R = all_map_R + MapN_R(i)
Next

End Sub


Public Sub PoolDataLoad_I(ByVal name As String) '道具图池数据录入

Dim strx As String
Dim i As Integer: Dim j As Integer: Dim n As Integer
i = 0: j = 0: n = 0

Open App.Path & "\item pool data\" & name & "_number.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, strx
   MapN_I(i) = Int(strx)
   i = i + 1
Loop
Close #1

i = 0: j = 0
Open App.Path & "\item pool data\" & name & "_order.txt" For Input As #2
Do While Not EOF(2)
   If MapN_I(j) <> 0 Then
      For i = 0 To MapN_I(j) - 1
         Line Input #2, strx
         Maporder_I(j, i) = Int(strx)
      Next
   End If
   j = j + 1
Loop
Close #2

i = 0: j = 0: n = 0
Open App.Path & "\item pool data\mapname.txt" For Input As #3
Do While Not EOF(3)
   If AllMaporder_I(j, i) <> 0 Then
      Line Input #3, strx
      If Maporder_I(j, n) = AllMaporder_I(j, i) Then
         Mapname_I(j, n) = strx
         n = n + 1
      End If
      i = i + 1
   Else
      j = j + 1
      i = 0
      n = 0
   End If
Loop
Close #3

all_map_I = 0
For i = 0 To theme
    all_map_I = all_map_I + MapN_I(i)
Next

End Sub


Public Sub BPSubLoad(ByVal i As Integer, ByVal team As String, ByRef j As Integer)

Dim k As Integer: Dim l As Integer: Dim m As Integer: Dim k0 As Integer
k = 0: l = 0
Dim strx As String
Dim num2 As Integer: Dim bp2 As String
Dim num4(1) As Integer: Dim bp4(1) As String

If team = "B" Then
    strx = BlueBPdata(i)
Else
    strx = RedBPdata(i)
End If

If Len(strx) = 2 Then
    num2 = Int(Mid(strx, 1, 1))
    bp2 = Mid(strx, 2, 1)
    m = j + num2 - 1
    ReDim Preserve PointAttriArray(m)
    ReDim Preserve PointAttriArray(m).BeforeStop(num2 - 1)
    For k = 0 To num2 - 1
        l = j + k
        PointAttriArray(l).Turns = i
        PointAttriArray(l).TeamBPTypes = team & bp2
        PointAttriArray(l).StopOrNot = False
        PointAttriArray(m).BeforeStop(k) = l
    Next
    PointAttriArray(l).StopOrNot = True
    j = j + num2
Else
    num4(0) = Int(Mid(strx, 1, 1))
    num4(1) = Int(Mid(strx, 3, 1))
    bp4(0) = Mid(strx, 2, 1)
    bp4(1) = Mid(strx, 4, 1)
    m = j + num4(0) + num4(1) - 1
    ReDim Preserve PointAttriArray(m)
    ReDim Preserve PointAttriArray(m).BeforeStop(num4(0) + num4(1) - 1)
    For k = 0 To num4(0) - 1
        l = j + k
        PointAttriArray(l).Turns = i
        PointAttriArray(l).TeamBPTypes = team & bp4(0)
        PointAttriArray(l).StopOrNot = False
        PointAttriArray(m).BeforeStop(k) = l
    Next
    j = j + num4(0): k0 = k
    For k = 0 To num4(1) - 1
        l = j + k
        PointAttriArray(l).Turns = i
        PointAttriArray(l).TeamBPTypes = team & bp4(1)
        PointAttriArray(l).StopOrNot = False
        PointAttriArray(m).BeforeStop(k0 + k) = l
    Next
    PointAttriArray(l).StopOrNot = True
    j = j + num4(1)
End If


End Sub


Public Sub BPDataLoad()

Dim i As Integer: Dim j As Integer: Dim k As Integer
i = 0: j = 0
Dim s As String

For i = 0 To UBound(BlueBPdata)
    Call BPSubLoad(i, "B", j)
    Call BPSubLoad(i, "R", j)
Next
k = UBound(PointAttriArray)
For i = 0 To k
    If PointAttriArray(i).StopOrNot = True And i <> k Then
        j = 1
        s = Mid(PointAttriArray(i).TeamBPTypes, 1, 1)
        Do Until s = Mid(PointAttriArray(i + j).TeamBPTypes, 1, 1)
            ReDim Preserve PointAttriArray(i).AfterStop(j - 1)
            PointAttriArray(i).AfterStop(j - 1) = i + j
            j = j + 1
            If i + j > k Then
                Exit Do
            End If
        Loop
    End If
Next

ReDim blueconfirm(UBound(BlueBPdata)): ReDim redconfirm(UBound(RedBPdata))
For i = 0 To UBound(blueconfirm)
    blueconfirm(i) = 0: redconfirm(i) = 0
Next

s = Mid(PointAttriArray(0).TeamBPTypes, 1, 1) '初始预选提示
i = 0
Do Until s <> Mid(PointAttriArray(i).TeamBPTypes, 1, 1)
    Form01.Image6(i).Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg")
    i = i + 1
Loop
FirstRoundNumber = i

AllBPTurns = k

End Sub


Public Sub ImageOrderMove()

Dim i As Integer: Dim s As Integer: Dim strx As String
i = 0: s = UBound(PointAttriArray)
Dim bp As Integer: Dim bb As Integer: Dim rp As Integer: Dim rb As Integer
bp = 0: bb = 0: rp = 0: rb = 0


For i = 0 To 19
    If i > s Then
        Form01.Image6(i).Visible = False
        Form01.Image6(i).Enabled = False
    Else
        Form01.Image6(i).Visible = True
        Form01.Image6(i).Enabled = True
        strx = PointAttriArray(i).TeamBPTypes
        If strx = "BP" Then
            Set Form01.Image6(i).Container = Form01.Frame1
            Form01.Image6(i).ZOrder
            Form01.Image6(i).Top = 1200 + 1080 * bp
            Form01.Image6(i).Left = 240
            bp = bp + 1
        ElseIf strx = "BB" Then
            Set Form01.Image6(i).Container = Form01.Frame2
            Form01.Image6(i).ZOrder
            If bb < 3 Then
                Form01.Image6(i).Top = 480
                Form01.Image6(i).Left = 240 + 1560 * bb
            Else
                Form01.Image6(i).Top = 1560
                Form01.Image6(i).Left = 240 + 1560 * (bb - 3)
            End If
            bb = bb + 1
        ElseIf strx = "RP" Then
            Set Form01.Image6(i).Container = Form01.Frame3
            Form01.Image6(i).ZOrder
            Form01.Image6(i).Top = 1200 + 1080 * rp
            Form01.Image6(i).Left = 240
            rp = rp + 1
        Else
            Set Form01.Image6(i).Container = Form01.Frame4
            Form01.Image6(i).ZOrder
            If rb < 3 Then
                Form01.Image6(i).Top = 480
                Form01.Image6(i).Left = 3600 - 1560 * rb
            Else
                Form01.Image6(i).Top = 1560
                Form01.Image6(i).Left = 3600 - 1560 * (rb - 3)
            End If
            rb = rb + 1
        End If
    End If
Next

End Sub


Public Sub PoolLoadInPCmodule(ByVal Index As Integer) '图池创建模块的图池载入事件

Dim i As Integer: Dim j As Integer
Dim mapn_tem() As Integer
ReDim mapn_tem(theme)

pool_load_choice = Index

For i = 0 To theme
    For j = 0 To 12
        thememapsave(i, j) = False
    Next
Next

If Index = 0 Then '空图池
   For i = 0 To theme
      For j = 0 To 12
         thememapsave(i, j) = False
      Next
   Next
ElseIf Index = 1 Then '全部竞速图
    i = 0: j = 0
    Do Until i > theme
        If AllMaporder_R(i, j) <> 0 Then
            thememapsave(i, j + 1) = True
            j = j + 1
        Else
            j = 0
            i = i + 1
        End If
    Loop
Else '其他图池
    i = 0
    Open App.Path & "\map pool data\" & Form03.Label2(Index).Caption & "_number.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strx
        mapn_tem(i) = Int(strx)
        i = i + 1
    Loop
    Close #1
    i = 0: j = 0
    Open App.Path & "\map pool data\" & Form03.Label2(Index).Caption & "_order.txt" For Input As #2
    Do While Not EOF(2)
        If mapn_tem(j) <> 0 Then
            For i = 0 To mapn_tem(j) - 1
                Line Input #2, strx
                thememapsave(j, Int(strx)) = True
            Next
        End If
        j = j + 1
    Loop
    Close #2
End If

End Sub


Public Sub PoolLoadInPCmoduleI(ByVal Index As Integer) '图池创建模块的图池载入事件

Dim i As Integer: Dim j As Integer
Dim mapn_tem() As Integer
ReDim mapn_tem(theme)

pool_load_choiceI = Index

For i = 0 To theme
    For j = 0 To 12
        thememapsaveI(i, j) = False
    Next
Next

If Index = 0 Then '空图池
   For i = 0 To theme
      For j = 0 To 12
         thememapsaveI(i, j) = False
      Next
   Next
ElseIf Index = 1 Then '全部竞速图
    i = 0: j = 0
    Do Until i > theme
        If AllMaporder_I(i, j) <> 0 Then
            thememapsaveI(i, j + 1) = True
            j = j + 1
        Else
            j = 0
            i = i + 1
        End If
    Loop
Else '其他图池
    i = 0
    Open App.Path & "\item pool data\" & Form03.Label6(Index).Caption & "_number.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strx
        mapn_tem(i) = Int(strx)
        i = i + 1
    Loop
    Close #1
    i = 0: j = 0
    Open App.Path & "\item pool data\" & Form03.Label6(Index).Caption & "_order.txt" For Input As #2
    Do While Not EOF(2)
        If mapn_tem(j) <> 0 Then
            For i = 0 To mapn_tem(j) - 1
                Line Input #2, strx
                thememapsaveI(j, Int(strx)) = True
            Next
        End If
        j = j + 1
    Loop
    Close #2
End If

End Sub


Public Sub BPDelete(ByVal Index As Integer)
Dim i As Integer: Dim cpn As Integer
Dim name As String

name = Form11.Label8(Index).Caption

Kill App.Path & "\BP data\" & name & ".txt" '数据文件删除
   
For i = 0 To 17
    temname(i) = ""
Next

i = 0 '重新载入BP names文件
Open App.Path & "\BP data\BP names.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, temname(i)
    i = i + 1
Loop
Close #1
i = 0
Open App.Path & "\BP data\BP names.txt" For Output As #2
Do Until temname(i) = ""
    If temname(i) <> name Then
        Print #2, temname(i)
    End If
    i = i + 1
Loop
Close #2

Unload Form11.Label8(Index) '控件删除

Dim TemTop As Integer: Dim TemLeft As Integer
Dim TemCaption As String

If Index + 1 < BPnumber Then '下方控件上移
    For i = Index + 1 To BPnumber - 1
        '先删除控件，再重新载入控件
        TemTop = Form11.Label8(i).Top: TemLeft = Form11.Label8(i).Left
        TemCaption = Form11.Label8(i).Caption
        
        Unload Form11.Label8(i)
        
        Load Form11.Label8(i - 1)
        
        Form11.Label8(i - 1).Top = TemTop - 480
        Form11.Label8(i - 1).Left = TemLeft
        Form11.Label8(i - 1).Caption = TemCaption
        Form11.Label8(i - 1).Visible = True: Form11.Label8(i - 1).ZOrder 0
    Next
End If

BPnumber = BPnumber - 1

End Sub


Public Sub PoolDelete(ByVal Index As Integer)
Dim i As Integer: Dim cpn As Integer
Dim name As String

name = Form03.Label2(Index).Caption

'删除当前图池
Kill App.Path & "\map pool data\" & name & "_number.txt" '数据文件删除
Kill App.Path & "\map pool data\" & name & "_order.txt"
   
For i = 0 To 17
    temname(i) = ""
Next

i = 0 '重新载入pool name文件
Open App.Path & "\map pool data\pool name.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, temname(i)
    i = i + 1
Loop
Close #1
i = 0
Open App.Path & "\map pool data\pool name.txt" For Output As #2
Do Until temname(i) = ""
    If temname(i) <> name Then
        Print #2, temname(i)
    End If
    i = i + 1
Loop
Close #2

Unload Form03.Label2(Index) '控件删除

Dim TemTop As Integer: Dim TemLeft As Integer
Dim TemCaption As String

If Index + 1 < PoolNumber Then '下方控件上移
    For i = Index + 1 To PoolNumber - 1
        '先删除控件，再重新载入控件
        TemTop = Form03.Label2(i).Top: TemLeft = Form03.Label2(i).Left
        TemCaption = Form03.Label2(i).Caption
        
        Unload Form03.Label2(i)
        
        Load Form03.Label2(i - 1)
        
        Form03.Label2(i - 1).Top = TemTop - 360
        Form03.Label2(i - 1).Left = TemLeft
        Form03.Label2(i - 1).Caption = TemCaption
        Form03.Label2(i - 1).Visible = True: Form03.Label2(i - 1).ZOrder 0
        
    Next
End If

PoolNumber = PoolNumber - 1

End Sub


Public Sub PoolDeleteI(ByVal Index As Integer)
Dim i As Integer: Dim cpn As Integer
Dim name As String

name = Form03.Label6(Index).Caption

'删除当前图池
Kill App.Path & "\item pool data\" & name & "_number.txt" '数据文件删除
Kill App.Path & "\item pool data\" & name & "_order.txt"
   
For i = 0 To 17
    temname(i) = ""
Next

i = 0 '重新载入pool name文件
Open App.Path & "\item pool data\pool name.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, temname(i)
    i = i + 1
Loop
Close #1
i = 0
Open App.Path & "\item pool data\pool name.txt" For Output As #2
Do Until temname(i) = ""
    If temname(i) <> name Then
        Print #2, temname(i)
    End If
    i = i + 1
Loop
Close #2

Unload Form03.Label6(Index) '控件删除

Dim TemTop As Integer: Dim TemLeft As Integer
Dim TemCaption As String

If Index + 1 < PoolNumberI Then '下方控件上移
    For i = Index + 1 To PoolNumberI - 1
        '先删除控件，再重新载入控件
        TemTop = Form03.Label6(i).Top: TemLeft = Form03.Label6(i).Left
        TemCaption = Form03.Label6(i).Caption
        
        Unload Form03.Label6(i)
        
        Load Form03.Label6(i - 1)
        
        Form03.Label6(i - 1).Top = TemTop - 360
        Form03.Label6(i - 1).Left = TemLeft
        Form03.Label6(i - 1).Caption = TemCaption
        Form03.Label6(i - 1).Visible = True: Form03.Label6(i - 1).ZOrder 0
        
    Next
End If

PoolNumberI = PoolNumberI - 1

End Sub


Public Sub PLThemeShow(ByVal PicPath As String, ByVal Index As Integer, ByRef AM() As Integer, ByRef TMS() As Boolean) '图池创建模块的主题显示事件

Dim i As Integer: Dim j As Integer: Dim k As Integer
Dim index1 As Integer: Dim Index0 As Integer
Dim c As Integer
Index0 = theme - Index: index1 = Index0 + 1

'非显示全图的情况
'map pool select框中的主题显示
For i = 0 To AM(Index0) - 1
    Form03.Image1(i).Visible = True
    c = i + 1
    If TMS(Index0, c) = False Then
        Form03.Image1(i).Picture = LoadPicture(App.Path & PicPath & "no" & index1 & "_" & c & ".jpg")
        Form03.Image1(i).Enabled = True
    Else
        Form03.Image1(i).Picture = LoadPicture(App.Path & PicPath & "noo" & index1 & "_" & c & ".jpg")
        Form03.Image1(i).Enabled = False
    End If
Next
Do While i <= 11
    Form03.Image1(i).Visible = False
    i = i + 1
Loop
'map pool save框中的主题显示
For i = 0 To AM(Index0) - 1
    c = i + 1
    If TMS(Index0, c) = False Then
        Form03.Image4(i).Visible = False
    Else
        Form03.Image4(i).Picture = LoadPicture(App.Path & PicPath & "no" & index1 & "_" & c & ".jpg")
        Form03.Image4(i).Visible = True
    End If
Next
Do While i <= 11
    Form03.Image4(i).Visible = False
    i = i + 1
Loop

End Sub


Public Sub ShowAllMaps(ByVal p As String, ByVal A As Integer, ByVal MN As Integer, ByRef SaveMatrix() As Boolean)

Dim k As Integer: Dim B As Integer
Form05.VScroll1.Value = 0 '滚动条置顶
k = 0

For i = 0 To theme
    For j = 0 To 12
        If SaveMatrix(i, j) = True Then
            B = i + 1
            Form05.Image4(k).Picture = LoadPicture(App.Path & p & "no" & B & "_" & j & ".jpg")
            Form05.Image4(k).Visible = True
            k = k + 1
        End If
    Next
Next

Do While k <= A - 1
    Form05.Image4(k).Visible = False
    k = k + 1
Loop

If MN <= 30 Then
    Form05.VScroll1.Visible = False
Else
    Form05.VScroll1.Max = IIf((MN Mod 6) = 0, (MN - 30) \ 6, (MN - 30) \ 6 + 1)
    Form05.VScroll1.Visible = True
End If

Form05.Label1.Caption = MN & " maps total"

End Sub


Public Sub Hover()

If Not m_oCtlCancelMode Is Nothing Then
    m_oCtlCancelMode.CancelMode
    Set m_oCtlCancelMode = Nothing
End If

End Sub


Public Sub MappingArrayInitialize() '映射数组赋值

Dim n As Integer: Dim l As Integer: Dim i As Integer: Dim j As Integer
n = 0: l = 0: i = 0: j = 0

ReDim thememap2index(theme, 12): ReDim thememap2index_1(theme, 12) '变量清空
ReDim allindex2mapindex(PositionAmount): ReDim index2theme(PositionAmount): ReDim index2map(PositionAmount)

Do Until i > all_map - 1
    If Maporder(n, l) <> 0 Then
       index2theme(i) = n
       index2map(i) = Maporder(n, l)
       thememap2index(n, Maporder(n, l)) = i
       allindex2mapindex(i) = l
       l = l + 1
       i = i + 1
    Else
       n = n + 1
       l = 0
    End If
Loop
For i = 0 To theme
    For j = 0 To 12
        If Maporder(i, j) = 0 Then
           Exit For
        Else
           thememap2index_1(i, Maporder(i, j)) = j
        End If
    Next
Next

End Sub
