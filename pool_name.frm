VERSION 5.00
Begin VB.Form Form04 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pool Name"
   ClientHeight    =   1800
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "pool_name.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   480
      TabIndex        =   0
      Top             =   630
      Width           =   7815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   8775
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   615
         Left            =   300
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Width           =   8175
      End
      Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
         Height          =   615
         Left            =   3840
         TabIndex        =   3
         Top             =   1185
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Style           =   3
         AnimationDuration=   0.2
         Caption         =   "OK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Manteka"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type in the pool name"
         BeginProperty Font 
            Name            =   "Nexa-Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   120
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()

Dim i As Integer: Dim j As Integer: Dim p As String
p = IIf(PoolRorNot, "\map pool data\", "\item pool data\")

If Text1.Text = "" Then
    Beep
ElseIf namesearch(Trim(Text1.Text), p & "pool name.txt") = True Then
    Beep
    Form04.Enabled = False
    Form10.Show
Else
    Open App.Path & p & Text1.Text & "_number.txt" For Output As #1  '建立文件
    For i = 0 To theme
        Print #1, Trim(mapnumber_save(i))
    Next i
    Close #1
   
    j = 0
    Open App.Path & p & Text1.Text & "_order.txt" For Output As #2
    For i = 0 To theme
        Do Until maporder_save(i, j) = 0
            Print #2, Trim(maporder_save(i, j))
            j = j + 1
        Loop
        j = 0
    Next
    Close #2
   
    Open App.Path & p & "pool name.txt" For Append As #3 '追加图池名
    Print #3, Text1.Text
    Close #3
    
    If PoolRorNot Then
        PoolNumber = TraverseAllNames(p & "pool name.txt", Form03.Label2, 360, PoolNumber, PoolNumber - 1) ' PoolNumber计入空图池, 故需要减一
    Else
        PoolNumberI = TraverseAllNames(p & "pool name.txt", Form03.Label6, 360, PoolNumberI, PoolNumberI - 1)
    End If
    
    Unload Form04
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form03.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   ctxNineButton1.SetFocus
   ctxNineButton1_Click
End If
End Sub


