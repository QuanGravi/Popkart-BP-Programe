VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BP Management"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8160
   Icon            =   "BP_manage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DADADA&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   2520
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tips: Right click can delete a BP choice"
         BeginProperty Font 
            Name            =   "Nexa-Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "5P5B"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1020
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   0
         Picture         =   "BP_manage.frx":048A
         Stretch         =   -1  'True
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   0
      Left            =   2520
      ScaleHeight     =   6375
      ScaleWidth      =   5655
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   900
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00261EBF&
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   9
         Top             =   4260
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   2580
         Width           =   4455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   735
         Index           =   2
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   5280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Style           =   3
         AnimationDuration=   0.2
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   2
         Height          =   735
         Index           =   1
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   2
         Height          =   735
         Index           =   0
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00261EBF&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()

If Text1(2).Text = "" Or Text1(1).Text = "" Or Text1(0).Text = "" Then
    Beep
ElseIf namesearch(Trim(Text1(2).Text), "\BP data\BP names.txt") Then
    Beep
    Form11.Enabled = False
    Form12.Label1.Caption = "The name has been used. Please change your BP name."
    Form12.Show
ElseIf BPcheck(Text1(0).Text, Text1(1).Text) = False Then
    Beep
    Form11.Enabled = False
    Form12.Label1.Caption = "Your input does not satisfied the required format."
    Form12.Show
Else
    Open App.Path & "\BP data\" & Text1(2).Text & ".txt" For Output As #1  '建立文件
    Print #1, Trim(Text1(0).Text)
    Print #1, Trim(Text1(1).Text)
    Close #1

    Open App.Path & "\BP data\BP names.txt" For Append As #1 '追加BP名
    Print #1, Text1(2).Text
    Close #1

    BPnumber = TraverseAllNames("\BP data\BP names.txt", Label8, 480, BPnumber, BPnumber)
End If

End Sub

Private Sub Form_Load()

BPnumber = TraverseAllNames("\BP data\BP names.txt", Label8, 480)

Text1(0).Text = "1B-2B-1P-2P-2B-2P"
Text1(1).Text = "2B-1B-2P-1P1B-1B1P-1P"
Text1(2).Text = "5P5B"

BP_load_choice = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

Form01.Enabled = True

End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim strx As String: Dim i As Integer
Dim tembluebp As String: Dim temredbp As String
i = 0

If Button = 1 Then
    
    Image1.Top = 960 + Index * 480
    Text1(2).Text = Label8(Index).Caption
    BP_load_choice = Index
    
    Open App.Path & "\BP data\" & Label8(Index).Caption & ".txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strx
        Text1(i).Text = strx
        i = i + 1
    Loop
    Close #1
    
ElseIf Button = 2 Then

    If Index = 0 Then
        Beep
    Else
        Call BPDelete(Index)
        If BP_load_choice < Index Then
        ElseIf BP_load_choice = Index Then
            Image1.Top = 960
            Call Label8_MouseDown(0, 1, 0, 1, 1)
        Else
            BP_load_choice = BP_load_choice - 1
            Image1.Top = Image1.Top - 480
        End If
        
        If Int(BP_choice(0)) < Index Then
        ElseIf Int(BP_choice(0)) = Index Then
            
            BP_choice(0) = 0: BP_choice(1) = "5P5B"
            
            ReDim PointAttriArray(0) 'bp数组清空
            
            Open App.Path & "\BP data\" & BP_choice(1) & ".txt" For Input As #1
            Line Input #1, tembluebp
            Line Input #1, temredbp
            Close #1
            
            BlueBPdata() = Split(tembluebp, "-")
            RedBPdata() = Split(temredbp, "-")

            Call BPDataLoad
        
        Else
            BP_choice(0) = BP_choice(0) - 1
        End If
        
    End If

End If

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 0 Or Index = 1 Then '限制输入
    If KeyAscii = 45 Or (KeyAscii >= 49 And KeyAscii <= 57) Or KeyAscii = 66 Or KeyAscii = 80 Or KeyAscii = 8 Then
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        ctxNineButton1.SetFocus
        ctxNineButton1_Click
    Else
        KeyAscii = 0
        Beep
    End If
Else
    If KeyAscii = 13 Then
        KeyAscii = 0
        ctxNineButton1.SetFocus
        ctxNineButton1_Click
    End If
End If

End Sub
