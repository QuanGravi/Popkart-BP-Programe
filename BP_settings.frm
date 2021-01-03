VERSION 5.00
Begin VB.Form Form02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BP Settings"
   ClientHeight    =   6360
   ClientLeft      =   3240
   ClientTop       =   3195
   ClientWidth     =   8415
   Icon            =   "BP_settings.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8415
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   0
      Left            =   3000
      ScaleHeight     =   6375
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Racing Pool Choice"
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
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "all racing maps"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   870
         Width           =   3855
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "BP_settings.frx":048A
         Stretch         =   -1  'True
         Top             =   840
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   3000
      ScaleHeight     =   6375
      ScaleWidth      =   5415
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "all item maps"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   870
         Width           =   3855
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   2
         Left            =   0
         Picture         =   "BP_settings.frx":0C87
         Stretch         =   -1  'True
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Pool Choice"
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
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DADADA&
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   3000
      TabIndex        =   10
      Top             =   0
      Width           =   3000
      Begin VB.Image Image1 
         Height          =   495
         Left            =   0
         Picture         =   "BP_settings.frx":1484
         Stretch         =   -1  'True
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Choice"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Waiting Time"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Pool CHOICE"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Racing Pool CHOICE"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   10.5
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
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Settings"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   60
         TabIndex        =   11
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   3
      Left            =   3000
      ScaleHeight     =   6375
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Choice"
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
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "5P5B"
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   870
         Width           =   3255
      End
      Begin VB.Image Image5 
         Height          =   375
         Index           =   1
         Left            =   0
         Picture         =   "BP_settings.frx":1C82
         Stretch         =   -1  'True
         Top             =   840
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   2
      Left            =   3000
      ScaleHeight     =   6375
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2040
         TabIndex        =   5
         Text            =   "90"
         Top             =   870
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   495
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   1575
      End
      Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
         Height          =   615
         Left            =   2100
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Style           =   3
         AnimationDuration=   0.2
         Caption         =   "set"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Nexa-Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Waiting Time"
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
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()

If Text1.Text = "" Then
   waitingtime = 30
   Text1.Text = 30
Else
   If Text1.Text >= 30 Then
      waitingtime = Text1.Text
   Else
      waitingtime = 30
      Text1.Text = 30
   End If
End If

End Sub

Private Sub Form_Load()
Text1.Text = waitingtime
Label3.Caption = "1. The default BP waiting time is 30. To change it, type your setting time into the textbox." & _
vbCrLf & vbCrLf & "2. Please remember to click 'set' button to load your BP waiting time into the system." & _
vbCrLf & vbCrLf & "3. BP waiting time must be greater than 30, otherwise the system will turn it to 30 automatically."

Image5(0).Top = 840 + map_pool_choice * 360
Image5(1).Top = 840 + Int(BP_choice(0)) * 360
Image5(2).Top = 840 + map_pool_choice_I * 360

Call TraverseAllNames("\map pool data\pool name.txt", Label1, 360)
Call TraverseAllNames("\item pool data\pool name.txt", Label10, 360)
Call TraverseAllNames("\BP data\BP names.txt", Label5, 360)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Form01.Enabled = True
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer: Dim j As Integer

If Button = 1 Then '***************鼠标左击事件，选择图池*********************

    map_pool_choice = Index

    Image5(0).Top = 840 + Index * 360

    For i = 0 To theme '地图预置数组初始化
        MapN_R(i) = 0
        For j = 0 To 12
            Maporder_R(i, j) = 0
            Mapname_R(i, j) = ""
        Next
    Next

    Call PoolDataLoad_R(Label1(Index).Caption)

End If

End Sub


Private Sub Label10_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer: Dim j As Integer

If Button = 1 Then '***************鼠标左击事件，选择图池*********************

    map_pool_choice_I = Index

    Image5(2).Top = 840 + Index * 360

    For i = 0 To theme '地图预置数组初始化
        MapN_I(i) = 0
        For j = 0 To 12
            Maporder_I(i, j) = 0
            Mapname_I(i, j) = ""
        Next
    Next

    Call PoolDataLoad_I(Label10(Index).Caption)

End If

End Sub


Private Sub Label5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tembluebp As String: Dim temredbp As String

If Button = 1 Then '***************鼠标左击事件，选择BP机制*********************

    BP_choice(0) = Index
    BP_choice(1) = Label5(Index).Caption

    Image5(1).Top = 840 + Index * 360
    
    ReDim PointAttriArray(0) 'bp数组清空
    
    Open App.Path & "\BP data\" & BP_choice(1) & ".txt" For Input As #1
        Line Input #1, tembluebp
        Line Input #1, temredbp
    Close #1

    BlueBPdata() = Split(tembluebp, "-")
    RedBPdata() = Split(temredbp, "-")

    Call BPDataLoad

End If


End Sub

Private Sub Label8_Click(Index As Integer)

Dim i As Integer

Image1.Top = 480 * Index + 600

For i = 0 To 3
    If i = Index Then
        Picture1(i).Visible = True
    Else
        Picture1(i).Visible = False
    End If
Next



End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf KeyAscii = 13 Then
   KeyAscii = 0
   ctxNineButton1.SetFocus
   ctxNineButton1_Click
Else
   KeyAscii = 0
End If
End Sub



