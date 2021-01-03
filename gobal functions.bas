Attribute VB_Name = "global_functions"
Public Function bpsearch(B As Integer, c As Integer) As Boolean '���������ͼ�Ƿ�BP
Dim i As Integer
Dim note As Boolean
note = False
For i = 0 To 19
    If B = mapbp(i, 0) And c = mapbp(i, 1) Then
       note = True
       Exit For
    End If
Next
bpsearch = note

End Function


Public Function keywordsearch(keyword As String, originstr As String) As Boolean '�ؼ�������
Dim n As Integer
Dim i As Integer
Dim record As Integer
Dim tem As Integer
record = 1
n = Len(keyword)

For i = 1 To n
     tem = InStr(record, originstr, Mid(keyword, i, 1))
     If tem = 0 Then
        Exit For
     Else
        record = tem + 1
     End If
Next
If tem = 0 Then
   keywordsearch = False
Else
   keywordsearch = True
End If

End Function


Public Function namesearch(name As String, p As String) As Boolean 'ͼ��������
Dim strx As String: Dim i As Integer
Dim note As Boolean
note = False

Open App.Path & p For Input As #1
Do While Not EOF(1)
   Line Input #1, strx
   If strx = name Then
      note = True
      Exit Do
   End If
Loop
Close #1

namesearch = note

End Function


Public Function PointAttriJudge(n As Integer) As points_judge

Dim precoun As Integer: Dim temteam As String
precoun = n - 1
If n <= AllBPTurns Then
    PointAttriJudge.NowTeam = Mid(PointAttriArray(n).TeamBPTypes, 1, 1)
Else
    PointAttriJudge.NowTeam = ""
End If

If precoun >= 0 Then
    PointAttriJudge.PreTeam = Mid(PointAttriArray(precoun).TeamBPTypes, 1, 1)
    temteam = PointAttriJudge.PreTeam
    PointAttriJudge.BlueStopOrNot = (temteam = "B" And PointAttriArray(precoun).StopOrNot) '�Ƿ�Ϊ����BPͣ�ٽڵ�
    PointAttriJudge.RedStopOrNot = (temteam = "R" And PointAttriArray(precoun).StopOrNot) '�Ƿ�Ϊ�췽BPͣ�ٽڵ�
    PointAttriJudge.BlueConfirmOrNot = (blueconfirm(PointAttriArray(precoun).Turns) = 0) 'ѡ���������BP�Ƿ�ȷ��
    PointAttriJudge.RedConfirmOrNot = (redconfirm(PointAttriArray(precoun).Turns) = 0)
Else
    PointAttriJudge.PreTeam = ""
    PointAttriJudge.BlueStopOrNot = False
    PointAttriJudge.RedStopOrNot = False
    PointAttriJudge.BlueConfirmOrNot = False
    PointAttriJudge.RedConfirmOrNot = False
End If

End Function


Public Function AnimationXofT(x As Double) As Double

Const pi = 3.14159265358

AnimationXofT = 1 - Cos(2 * pi * x)

End Function


Public Function TraverseAllNames(ByVal p As String, ByRef Contro As Object, ByVal Interval As Integer, Optional InitialIndex As Integer = 1, Optional Jump As Integer = 1) As Integer
'���������ļ����������ؼ�����
'�ؼ���ʼ�±�Ĭ��Ϊ1
'Ĭ��������һ��

Dim strx As String: Dim i As Integer: Dim j As Integer
i = InitialIndex: j = 1
Open App.Path & p For Input As #1
For j = 1 To Jump
    Line Input #1, strx '����Jump��
Next
Do While Not EOF(1)
   Line Input #1, strx
   Load Contro(i)
   Contro(i).Left = Contro(0).Left '�ؼ������
   Contro(i).Top = Contro(i - 1).Top + Interval '���ÿؼ���ļ��
   Contro(i).Visible = True: Contro(i).ZOrder 0
   Contro(i).Caption = strx
   i = i + 1
Loop
Close #1

TraverseAllNames = i '�������ؿؼ�������

End Function

'------------------------------------------

Public Function BPcheck(ByVal Bstrx As String, ByVal Rstrx As String) As Boolean 'BP������

Dim B1 As Integer: Dim B2 As Integer: Dim B3 As Integer
B1 = 0: B2 = 0: B3 = 0
Dim R1 As Integer: Dim R2 As Integer: Dim R3 As Integer
R1 = 0: R2 = 0: R3 = 0

If BPinputCheck(Bstrx, B1, B2, B3) And BPinputCheck(Rstrx, R1, R2, R3) Then
    If B1 = R1 And B1 >= 0 And B2 = R2 And B3 = R3 And B3 >= 0 Then
        BPcheck = True
        Exit Function
    Else
        BPcheck = False
        Exit Function
    End If
Else
    BPcheck = False
    Exit Function
End If

End Function

Public Function BPinputCheck(ByVal strx As String, ByRef NumOfTurns As Integer, ByRef NumB As Integer, ByRef NumP As Integer) As Boolean '��������BP�Ƿ���ϸ�ʽҪ��

Dim Inp() As String
Inp = Split(strx, "-")
Dim i As Integer: Dim temstrx As String

NumOfTurns = UBound(Inp) + 1
For i = 0 To UBound(Inp)
    temstrx = Inp(i)
    If Len(temstrx) = 2 Then ' ����ַ�������
        If NumOrNot(Mid(temstrx, 1, 1)) And BPOrNot(Mid(temstrx, 2, 1)) Then '��������ʽ
            '���ͨ����ͳ��B��P������
            If Mid(temstrx, 2, 1) = "B" Then
                NumB = NumB + Int(Mid(temstrx, 1, 1))
            Else
                NumP = NumP + Int(Mid(temstrx, 1, 1))
            End If
        Else
            BPinputCheck = False
            Exit Function
        End If
    ElseIf Len(temstrx) = 4 Then
        If NumOrNot(Mid(temstrx, 1, 1)) And BPOrNot(Mid(temstrx, 2, 1)) And NumOrNot(Mid(temstrx, 3, 1)) And BPOrNot(Mid(temstrx, 4, 1)) And Mid(temstrx, 4, 1) <> Mid(temstrx, 2, 1) Then
            If Mid(temstrx, 2, 1) = "B" Then
                NumB = NumB + Int(Mid(temstrx, 1, 1))
                NumP = NumP + Int(Mid(temstrx, 3, 1))
            Else
                NumP = NumP + Int(Mid(temstrx, 1, 1))
                NumB = NumB + Int(Mid(temstrx, 3, 1))
            End If
        Else
            BPinputCheck = False
            Exit Function
        End If
    Else
        BPinputCheck = False
        Exit Function
    End If
Next

BPinputCheck = True

End Function

Private Function NumOrNot(x As String) As Boolean
If Int(x) >= 1 And Int(x) <= 9 Then
    NumOrNot = True
Else
    NumOrNot = False
End If
End Function

Private Function BPOrNot(x As String) As Boolean
If x = "B" Or x = "P" Then
    BPOrNot = True
Else
    BPOrNot = False
End If
End Function

'------------------------------------------

