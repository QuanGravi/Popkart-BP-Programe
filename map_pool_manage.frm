VERSION 5.00
Begin VB.Form Form03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Pool Management"
   ClientHeight    =   9435
   ClientLeft      =   690
   ClientTop       =   1425
   ClientWidth     =   20730
   Icon            =   "map_pool_manage.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   20730
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   720
      Width           =   1455
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00DADADA&
         BorderStyle     =   0  'None
         Height          =   41000
         Left            =   240
         ScaleHeight     =   40995
         ScaleWidth      =   1335
         TabIndex        =   6
         Top             =   0
         Width           =   1335
         Begin AoxLeague_Ver.AlphaBlendImage AlphaBlendImage1 
            Height          =   750
            Index           =   0
            Left            =   210
            Top             =   270
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1323
            Stretch         =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "GRAYSTROKE"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -15
            TabIndex        =   7
            Top             =   0
            Width           =   1215
         End
         Begin AoxLeague_Ver.AlphaBlendImage AlphaBlendImage2 
            Height          =   1095
            Left            =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   1931
            Opacity         =   0.5
            Stretch         =   -1  'True
            Picture         =   "map_pool_manage.frx":048A
         End
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   8415
         Left            =   0
         Max             =   19
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3696
      Left            =   1920
      ScaleHeight     =   3690
      ScaleWidth      =   14895
      TabIndex        =   2
      Top             =   5280
      Width           =   14895
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00403A35&
         BorderStyle     =   0  'None
         Height          =   45000
         Left            =   0
         ScaleHeight     =   45000
         ScaleMode       =   0  'User
         ScaleWidth      =   15105
         TabIndex        =   3
         Top             =   0
         Width           =   15100
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "click the image to add the map into the pool"
            BeginProperty Font 
               Name            =   "Nexa-Bold"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   4440
            TabIndex        =   8
            Top             =   0
            Width           =   6255
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   11
            Left            =   12360
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   10
            Left            =   9960
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   9
            Left            =   7560
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   8
            Left            =   5160
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   7
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   6
            Left            =   360
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   5
            Left            =   12360
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   4
            Left            =   9960
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   3
            Left            =   7560
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   2
            Left            =   5160
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   1
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   1335
            Index           =   0
            Left            =   360
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3695
      Left            =   1920
      ScaleHeight     =   3690
      ScaleWidth      =   14895
      TabIndex        =   0
      Top             =   885
      Width           =   14895
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00403A35&
         BorderStyle     =   0  'None
         Height          =   35000
         Left            =   0
         ScaleHeight     =   34995
         ScaleWidth      =   15105
         TabIndex        =   1
         Top             =   0
         Width           =   15100
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "click the image to delete the map from the pool"
            BeginProperty Font 
               Name            =   "Nexa-Bold"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   4320
            TabIndex        =   11
            Top             =   0
            Width           =   6615
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   11
            Left            =   12360
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   10
            Left            =   9960
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   9
            Left            =   7560
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   8
            Left            =   5160
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   7
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   6
            Left            =   360
            Stretch         =   -1  'True
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   5
            Left            =   12360
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   4
            Left            =   9960
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   3
            Left            =   7560
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   2
            Left            =   5160
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   1
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image4 
            Height          =   1335
            Index           =   0
            Left            =   360
            Stretch         =   -1  'True
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tips: Right click can delete a map pool"
      BeginProperty Font 
         Name            =   "Nexa-Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   17400
      TabIndex        =   18
      Top             =   120
      Width           =   2835
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Nexa-Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   19080
      TabIndex        =   15
      Top             =   900
      Width           =   735
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      Style           =   3
      AnimationDuration=   0.2
      Caption         =   "save pool"
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
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton2 
      Height          =   615
      Left            =   10200
      TabIndex        =   14
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      Style           =   3
      AnimationDuration=   0.2
      Caption         =   "Show All MAPS"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Map Pool"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   17280
      TabIndex        =   9
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Racing"
      BeginProperty Font 
         Name            =   "Nexa-Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00261EBF&
      Height          =   435
      Index           =   0
      Left            =   17640
      TabIndex        =   12
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "all racing maps"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   17280
      TabIndex        =   10
      Top             =   1920
      Width           =   3000
   End
   Begin VB.Image Image5 
      Height          =   45
      Index           =   12
      Left            =   20105
      Picture         =   "map_pool_manage.frx":0C97
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   305
   End
   Begin VB.Image Image5 
      Height          =   45
      Index           =   11
      Left            =   17160
      Picture         =   "map_pool_manage.frx":1495
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   300
   End
   Begin VB.Image Image5 
      Height          =   7915
      Index           =   10
      Left            =   17160
      Picture         =   "map_pool_manage.frx":1C93
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   45
   End
   Begin VB.Image Image5 
      Height          =   45
      Index           =   9
      Left            =   17160
      Picture         =   "map_pool_manage.frx":2491
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   3285
   End
   Begin VB.Image Image5 
      Height          =   7920
      Index           =   8
      Left            =   20400
      Picture         =   "map_pool_manage.frx":2C8F
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   45
   End
   Begin VB.Image Image5 
      Height          =   3780
      Index           =   7
      Left            =   16740
      Picture         =   "map_pool_manage.frx":348D
      Stretch         =   -1  'True
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   3780
      Index           =   6
      Left            =   1860
      Picture         =   "map_pool_manage.frx":3C8B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   5
      Left            =   1860
      Picture         =   "map_pool_manage.frx":4489
      Stretch         =   -1  'True
      Top             =   4485
      Width           =   15015
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   4
      Left            =   1920
      Picture         =   "map_pool_manage.frx":4C87
      Stretch         =   -1  'True
      Top             =   840
      Width           =   14895
   End
   Begin VB.Image Image5 
      Height          =   3780
      Index           =   3
      Left            =   16740
      Picture         =   "map_pool_manage.frx":5485
      Stretch         =   -1  'True
      Top             =   5235
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   3780
      Index           =   2
      Left            =   1860
      Picture         =   "map_pool_manage.frx":5C82
      Stretch         =   -1  'True
      Top             =   5235
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   1
      Left            =   1920
      Picture         =   "map_pool_manage.frx":647F
      Stretch         =   -1  'True
      Top             =   8880
      Width           =   14895
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   0
      Left            =   1920
      Picture         =   "map_pool_manage.frx":6C7C
      Stretch         =   -1  'True
      Top             =   5235
      Width           =   14895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Empty Map Pool"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   17280
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "All Item Maps"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   17280
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image Image5 
      Height          =   375
      Index           =   13
      Left            =   17160
      Picture         =   "map_pool_manage.frx":7479
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   3240
   End
   Begin VB.Image Image6 
      Height          =   9450
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "Form03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AlphaBlendImage1_Click(Index As Integer)

VScroll3.SetFocus

If PoolRorNot Then
    Call PLThemeShow("\picture\", Index, AllMapN_R, thememapsave)
    theme_poolcreation = theme - Index '记录当前主题号
Else
    Call PLThemeShow("\item picture\", Index, AllMapN_I, thememapsaveI)
    theme_poolcreationI = theme - Index '记录当前主题号
End If

AlphaBlendImage2.Top = Label1(Index).Top
AlphaBlendImage2.Visible = True

End Sub


Private Sub ctxNineButton1_Click()

Dim i As Integer: Dim j As Integer: Dim n As Integer
Dim MAPSUM As Integer

For i = 0 To theme '数据初始化
    mapnumber_save(i) = 0
    For j = 0 To 12
        maporder_save(i, j) = 0
    Next
Next

n = 0 '记录图池数据
For i = 0 To theme
   For j = 0 To 12
      If IIf(PoolRorNot, thememapsave(i, j), thememapsaveI(i, j)) Then
         maporder_save(i, n) = j
         n = n + 1
      End If
   Next
   n = 0
Next

j = 0
For i = 0 To theme
    Do Until maporder_save(i, j) = 0
       j = j + 1
    Loop
    mapnumber_save(i) = j
    j = 0
Next

MAPSUM = 0
For i = 0 To theme
   MAPSUM = MAPSUM + mapnumber_save(i)
Next
If MAPSUM >= 6 Then
   Form04.Show
   Form03.Enabled = False
Else
   Beep
   Form08.Show
   Form03.Enabled = False
End If

End Sub

Private Sub ctxNineButton2_Click()

Form05.Show

End Sub

Private Sub Form_Load()

hwndVS_y1 = VScroll3.hwnd
OldWindowProc_y1 = GetWindowLong(VScroll3.hwnd, GWL_WNDPROC)
Call SetWindowLong(VScroll3.hwnd, GWL_WNDPROC, AddressOf NewWindowProc_y1)

Image6.Picture = LoadPicture(App.Path & "\picture\Pool_Background.bmp")

Dim i As Integer

Set AlphaBlendImage1(0).Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\map" & theme & ".png")
Label1(0).Caption = ThemeName(theme)
For i = 1 To theme '生成主题选择框
    Load AlphaBlendImage1(i): Load Label1(i)
    AlphaBlendImage1(i).Left = AlphaBlendImage1(0).Left
    Label1(i).Left = Label1(0).Left
    AlphaBlendImage1(i).Top = AlphaBlendImage1(i - 1).Top + 1050
    Label1(i).Top = Label1(i - 1).Top + 1050
    AlphaBlendImage1(i).Visible = True: Label1(i).Visible = True
    AlphaBlendImage1(i).ZOrder 0: Label1(i).ZOrder 0
    Set AlphaBlendImage1(i).Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\map" & theme - i & ".png")
    Label1(i).Caption = ThemeName(theme - i)
Next
For i = 0 To 11
    Image1(i).Visible = False
Next
For i = 0 To 11
    Image4(i).Visible = False
Next

VScroll3.Max = theme - 7

'初始化图池加载模块
Image5(13).Top = 1560 + pool_load_choice * 360

PoolNumber = TraverseAllNames("\map pool data\pool name.txt", Label2, 360, 2) '载入所有竞速图池
PoolNumberI = TraverseAllNames("\item pool data\pool name.txt", Label6, 360, 2) '载入所有竞速图池

For i = 0 To PoolNumberI - 1 '进入界面时默认为竞速图池，隐藏道具选项
    Label6(i).Visible = False
Next
PoolRorNot = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form01.Enabled = True

End Sub

Private Sub Image1_Click(Index As Integer)

Dim B As Integer: Dim c As Integer: Dim p As String
c = Index + 1
B = IIf(PoolRorNot, theme_poolcreation + 1, theme_poolcreationI + 1)
p = IIf(PoolRorNot, "\picture\", "\item picture\")

Image1(Index).Picture = LoadPicture(App.Path & p & "noo" & B & "_" & c & ".jpg")
Image4(Index).Picture = LoadPicture(App.Path & p & "no" & B & "_" & c & ".jpg")
Image4(Index).Visible = True
Image1(Index).Enabled = False

If PoolRorNot Then
    thememapsave(theme_poolcreation, c) = True
Else
    thememapsaveI(theme_poolcreationI, c) = True
End If

End Sub


Private Sub Image4_Click(Index As Integer)

Dim B As Integer: Dim c As Integer: Dim p As String
c = Index + 1
B = IIf(PoolRorNot, theme_poolcreation + 1, theme_poolcreationI + 1)
p = IIf(PoolRorNot, "\picture\", "\item picture\")

Image1(Index).Picture = LoadPicture(App.Path & p & "no" & B & "_" & c & ".jpg")
Image4(Index).Visible = False
Image1(Index).Enabled = True

If PoolRorNot Then
    thememapsave(theme_poolcreation, c) = False
Else
    thememapsaveI(theme_poolcreationI, c) = False
End If

End Sub




Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then '鼠标左击图池载入事件
    Image5(13).Top = 1560 + 360 * Index
    Call PoolLoadInPCmodule(Index)
    Call AlphaBlendImage1_Click(0)
ElseIf Button = 2 Then '鼠标右击图池删除事件
    If Index <= 1 Then
        Beep
    Else
        Call PoolDelete(Index)
        If pool_load_choice < Index Then
        ElseIf pool_load_choice = Index Then '删除当前选中图池，则返回空图池
            Image5(13).Top = 1560
            Call PoolLoadInPCmodule(0)
            Call AlphaBlendImage1_Click(0)
        Else
            pool_load_choice = pool_load_choice - 1
            Image5(13).Top = Image5(13).Top - 360
        End If
        
        If map_pool_choice < Index - 1 Then
        ElseIf map_pool_choice = Index - 1 Then '若删除BP setting中的选中图池，则返回全地图图池
        
            map_pool_choice = 0 '载入全地图图池
            For i = 0 To theme
                MapN_R(i) = 0
                For j = 0 To 12
                    Maporder_R(i, j) = 0
                    Mapname_R(i, j) = ""
                Next
            Next
            Call PoolDataLoad_R("all racing maps")
        
        Else
            map_pool_choice = map_pool_choice - 1
        End If
        
    End If
Else
    
End If

End Sub

Private Sub Label5_Click(Index As Integer)

Dim i As Integer

If Index = 0 Then
    Label5(0).ForeColor = &H261EBF
    Label5(1).ForeColor = &HFFFFFF
    For i = 8 To 13
        Image5(i).Picture = LoadPicture(App.Path & "\picture\forchosen2.jpg")
    Next
    For i = 0 To PoolNumber - 1
        Label2(i).Visible = True
    Next
    For i = 0 To PoolNumberI - 1
        Label6(i).Visible = False
    Next
    PoolRorNot = True
    Image5(13).Top = 1560 + 360 * pool_load_choice
    Call AlphaBlendImage1_Click(0)
    'Call Label2_MouseDown(pool_load_choice, 1, 0, 1, 1) '切换模式时刷新页面
Else
    Label5(0).ForeColor = &HFFFFFF
    Label5(1).ForeColor = &H8000000D
    For i = 8 To 13
        Image5(i).Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg")
    Next
    For i = 0 To PoolNumber - 1
        Label2(i).Visible = False
    Next
    For i = 0 To PoolNumberI - 1
        Label6(i).Visible = True
    Next
    PoolRorNot = False
    Image5(13).Top = 1560 + 360 * pool_load_choiceI
    Call AlphaBlendImage1_Click(0)
    'Call Label6_MouseDown(pool_load_choiceI, 1, 0, 1, 1) '切换模式时刷新页面
End If

End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then '鼠标左击图池载入事件
    Image5(13).Top = 1560 + 360 * Index
    Call PoolLoadInPCmoduleI(Index)
    Call AlphaBlendImage1_Click(0)
ElseIf Button = 2 Then '鼠标右击图池删除事件
    If Index <= 1 Then
        Beep
    Else
        Call PoolDeleteI(Index)
        If pool_load_choiceI < Index Then
        ElseIf pool_load_choiceI = Index Then '删除当前选中图池，则返回空图池
            Image5(13).Top = 1560
            Call PoolLoadInPCmoduleI(0)
            Call AlphaBlendImage1_Click(0)
        Else
            pool_load_choiceI = pool_load_choiceI - 1
            Image5(13).Top = Image5(13).Top - 360
        End If
        
        If map_pool_choice_I < Index - 1 Then
        ElseIf map_pool_choice_I = Index - 1 Then '若删除BP setting中的选中图池，则返回全地图图池
        
            map_pool_choice_I = 0 '载入全地图图池
            For i = 0 To theme
                MapN_I(i) = 0
                For j = 0 To 12
                    Maporder_I(i, j) = 0
                    Mapname_I(i, j) = ""
                Next
            Next
            Call PoolDataLoad_I("all item maps")
        
        Else
            map_pool_choice_I = map_pool_choice_I - 1
        End If
        
    End If
Else
    
End If

End Sub

Private Sub VScroll3_Change()
Picture8.Top = -CLng(1050) * VScroll3.Value
If VScroll3.Visible = True Then
   VScroll3.SetFocus
End If
End Sub

Private Sub VScroll3_Scroll()
Picture8.Top = -CLng(1050) * VScroll3.Value
If VScroll3.Visible = True Then
   VScroll3.SetFocus
End If
End Sub




