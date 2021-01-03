VERSION 5.00
Begin VB.Form Form05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Maps"
   ClientHeight    =   8790
   ClientLeft      =   3825
   ClientTop       =   1425
   ClientWidth     =   14460
   Icon            =   "all_map_of_pool.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14460
   Begin VB.VScrollBar VScroll1 
      Height          =   8745
      Left            =   14200
      Max             =   15
      TabIndex        =   2
      Top             =   30
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   8820
      Left            =   0
      ScaleHeight     =   8760
      ScaleWidth      =   15075
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00403A35&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   40000
         Left            =   0
         ScaleHeight     =   40005
         ScaleWidth      =   15105
         TabIndex        =   1
         Top             =   -120
         Width           =   15100
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
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
            Index           =   1
            Left            =   1680
            TabIndex        =   5
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Racing"
            BeginProperty Font 
               Name            =   "Nexa-Bold"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00261EBF&
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Height          =   495
            Left            =   6240
            TabIndex        =   3
            Top             =   120
            Width           =   3135
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   0
            Left            =   360
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   1
            Left            =   2640
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   2
            Left            =   4920
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   3
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   4
            Left            =   9480
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   5
            Left            =   11760
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   6
            Left            =   360
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   7
            Left            =   2640
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   8
            Left            =   4920
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   9
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   10
            Left            =   9480
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   11
            Left            =   11760
            Stretch         =   -1  'True
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "Form05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

hwndVS_x2 = VScroll1.hwnd
OldWindowProc_x2 = GetWindowLong(VScroll1.hwnd, GWL_WNDPROC)
Call SetWindowLong(VScroll1.hwnd, GWL_WNDPROC, AddressOf NewWindowProc_x2)

Dim B As Integer: Dim A As Integer
VScroll1.Value = 0 '滚动条置顶
A = IIf(ALLMAP_R > ALLMAP_I, ALLMAP_R, ALLMAP_I)
sum_for_scroll = 0: sum_for_scrollI = 0

Dim m As Integer: m = 0
For i = 12 To A - 1
    Load Image4(i)
    If m <> 0 Then
       Image4(i).Left = Image4(i - 1).Left + 2280
       Image4(i).Top = Image4(i - 1).Top
       Image4(i).Visible = False
       Image4(i).ZOrder 0
       m = m + 1
       If m = 6 Then
          m = 0
       End If
    Else
       Image4(i).Left = Image4(i - 6).Left
       Image4(i).Top = Image4(i - 6).Top + 1695
       Image4(i).Visible = False
       Image4(i).ZOrder 0
       m = 1
    End If
Next

For i = 0 To theme
    For j = 0 To 12
        If thememapsave(i, j) = True Then
            sum_for_scroll = sum_for_scroll + 1 '统计竞速图数量
        End If
        If thememapsaveI(i, j) = True Then
            sum_for_scrollI = sum_for_scrollI + 1 '统计道具图数量
        End If
    Next
Next

Call ShowAllMaps("\picture\", A, sum_for_scroll, thememapsave) '进入界面时默认显示竞速图

End Sub

Private Sub Label2_Click(Index As Integer)

Dim A As Integer
A = IIf(ALLMAP_R > ALLMAP_I, ALLMAP_R, ALLMAP_I)

If Index = 0 Then
    Label2(0).ForeColor = &H261EBF
    Label2(1).ForeColor = &HFFFFFF
    Call ShowAllMaps("\picture\", A, sum_for_scroll, thememapsave)
Else
    Label2(0).ForeColor = &HFFFFFF
    Label2(1).ForeColor = &H8000000D
    Call ShowAllMaps("\item picture\", A, sum_for_scrollI, thememapsaveI)
End If

End Sub

Private Sub VScroll1_Change()
If VScroll1.Value = 0 Then '不均匀切分
    Picture2.Top = -120
Else
    Picture2.Top = -CLng(1695) * VScroll1.Value - 240
End If
If VScroll1.Visible = True Then
   VScroll1.SetFocus
End If
End Sub

Private Sub VScroll1_Scroll()
If VScroll1.Value = 0 Then
    Picture2.Top = -120
Else
    Picture2.Top = -CLng(1695) * VScroll1.Value - 240
End If
If VScroll1.Visible = True Then
   VScroll1.SetFocus
End If
End Sub

