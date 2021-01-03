VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form01 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popkart BP Ver4.0"
   ClientHeight    =   10320
   ClientLeft      =   1020
   ClientTop       =   1170
   ClientWidth     =   20055
   FillColor       =   &H00FF0000&
   ForeColor       =   &H00000000&
   Icon            =   "BP main interface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "BP main interface.frx":048A
   ScaleHeight     =   10320
   ScaleMode       =   0  'User
   ScaleWidth      =   20055
   Begin VB.TextBox Text5 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   12840
      TabIndex        =   28
      Top             =   590
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   7800
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   5300
      Left            =   17045
      Max             =   16
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UNDO"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Picture         =   "BP main interface.frx":7D44
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   2760
      ScaleHeight     =   5280
      ScaleWidth      =   1455
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
      Begin VB.VScrollBar VScroll1 
         Height          =   5280
         Left            =   0
         Max             =   22
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00DADADA&
         BorderStyle     =   0  'None
         Height          =   41000
         Left            =   240
         ScaleHeight     =   40995
         ScaleWidth      =   1335
         TabIndex        =   23
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "GRAYSTROKE"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   -15
            TabIndex        =   24
            Top             =   0
            Width           =   1215
         End
         Begin AoxLeague_Ver.AlphaBlendImage AlphaBlendImage4 
            Height          =   1095
            Left            =   60
            Top             =   0
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   1931
            Opacity         =   0.5
            Stretch         =   -1  'True
            Picture         =   "BP main interface.frx":8674
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   2760
      ScaleHeight     =   5295
      ScaleWidth      =   14535
      TabIndex        =   14
      Top             =   1200
      Width           =   14535
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00403A35&
         BorderStyle     =   0  'None
         Height          =   50000
         Left            =   0
         ScaleHeight     =   49995
         ScaleWidth      =   14535
         TabIndex        =   15
         Top             =   -120
         Width           =   14535
         Begin AoxLeague_Ver.ctxNineButton ctxNineButton5 
            Height          =   615
            Index           =   1
            Left            =   7920
            TabIndex        =   41
            Top             =   2400
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1085
            Style           =   3
            AnimationDuration=   0.2
            Caption         =   "Item BP Start"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Manteka"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin AoxLeague_Ver.ctxNineButton ctxNineButton5 
            Height          =   615
            Index           =   0
            Left            =   3720
            TabIndex        =   35
            Top             =   2400
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1085
            Style           =   3
            AnimationDuration=   0.2
            Caption         =   "Speed BP Start"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Manteka"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   0
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   1
            Left            =   4320
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   2
            Left            =   6600
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   3
            Left            =   8880
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   4
            Left            =   11160
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   5
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   6
            Left            =   4320
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   7
            Left            =   6600
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   8
            Left            =   8880
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   9
            Left            =   11160
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   10
            Left            =   2040
            Stretch         =   -1  'True
            Top             =   3720
            Width           =   2055
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   6360
            X2              =   6360
            Y1              =   1080
            Y2              =   3000
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   4320
            X2              =   6360
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   4320
            X2              =   4320
            Y1              =   1080
            Y2              =   3000
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   4320
            X2              =   6360
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   9600
            X2              =   9600
            Y1              =   1080
            Y2              =   3000
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   11640
            X2              =   11640
            Y1              =   1080
            Y2              =   3000
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   9600
            X2              =   11640
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   9600
            X2              =   11640
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   11
            Left            =   4320
            Stretch         =   -1  'True
            Top             =   3720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   12
            Left            =   6600
            Stretch         =   -1  'True
            Top             =   3720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   13
            Left            =   8880
            Stretch         =   -1  'True
            Top             =   3720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image21 
            Height          =   1335
            Index           =   14
            Left            =   11160
            Stretch         =   -1  'True
            Top             =   3720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   6240
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Select the winner"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   6720
            TabIndex        =   19
            Top             =   1200
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1575
            Left            =   4440
            TabIndex        =   18
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Blue"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1335
            Left            =   4440
            TabIndex        =   21
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1575
            Left            =   9720
            TabIndex        =   17
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Red"
            BeginProperty Font 
               Name            =   "Stencil"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1335
            Left            =   9720
            TabIndex        =   20
            Top             =   1200
            Width           =   2895
         End
      End
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6960
      TabIndex        =   13
      Text            =   "0"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   600
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   18000
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00353A40&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00261EBF&
         Height          =   600
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Red"
         Top             =   0
         Width           =   1695
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   16
         X1              =   120
         X2              =   1815
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   15
         X1              =   1560
         X2              =   1800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   14
         X1              =   120
         X2              =   915
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   13
         X1              =   120
         X2              =   120
         Y1              =   960
         Y2              =   6840
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   12
         X1              =   1800
         X2              =   1800
         Y1              =   960
         Y2              =   7680
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   11
         X1              =   120
         X2              =   1800
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   19
         Left            =   240
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   16
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   11
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   8
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   7
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick"
         BeginProperty Font 
            Name            =   "GRAYSTROKE"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00261EBF&
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00353A40&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Manteka"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Blue"
         Top             =   0
         Width           =   1575
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   5
         X1              =   120
         X2              =   1800
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   4
         X1              =   1800
         X2              =   1800
         Y1              =   960
         Y2              =   7695
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   3
         X1              =   120
         X2              =   120
         Y1              =   960
         Y2              =   6855
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   2
         X1              =   120
         X2              =   315
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   1
         X1              =   960
         X2              =   1800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   1815
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Image Image7 
         Height          =   0
         Left            =   120
         Picture         =   "BP main interface.frx":8E81
         Stretch         =   -1  'True
         Top             =   840
         Width           =   75
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   18
         Left            =   240
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   17
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   10
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   9
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   6
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick"
         BeginProperty Font 
            Name            =   "GRAYSTROKE"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   10
         X1              =   120
         X2              =   315
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   9
         X1              =   960
         X2              =   4920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   8
         X1              =   120
         X2              =   120
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   7
         X1              =   120
         X2              =   4920
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Index           =   6
         X1              =   4920
         X2              =   4920
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   14
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   13
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   4
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   3
         Left            =   240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "GRAYSTROKE"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00403A35&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   14640
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   21
         X1              =   5160
         X2              =   5160
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   20
         X1              =   360
         X2              =   5160
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   19
         X1              =   360
         X2              =   360
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   18
         X1              =   4905
         X2              =   5160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00261EBF&
         BorderWidth     =   3
         Index           =   17
         X1              =   360
         X2              =   4275
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   15
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   12
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   5
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   2
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   1
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "GRAYSTROKE"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00261EBF&
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Game Start"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      Picture         =   "BP main interface.frx":986E
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton10 
      Height          =   615
      Left            =   8400
      TabIndex        =   42
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "Manage BP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.AlphaBlendImage AlphaBlendImage3 
      Height          =   375
      Left            =   2160
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Stretch         =   -1  'True
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   6240
      TabIndex        =   40
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin AoxLeague_Ver.AlphaBlendImage AlphaBlendImage2 
      Height          =   975
      Left            =   8640
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      Stretch         =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00403A35&
      FillColor       =   &H00403A35&
      FillStyle       =   0  'Solid
      Height          =   30
      Index           =   3
      Left            =   12840
      Shape           =   4  'Rounded Rectangle
      Top             =   555
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00403A35&
      FillColor       =   &H00403A35&
      FillStyle       =   0  'Solid
      Height          =   30
      Index           =   2
      Left            =   12840
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00403A35&
      FillColor       =   &H00403A35&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   15600
      Shape           =   4  'Rounded Rectangle
      Top             =   560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00403A35&
      FillColor       =   &H00403A35&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   12720
      Shape           =   4  'Rounded Rectangle
      Top             =   555
      Visible         =   0   'False
      Width           =   135
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton9 
      Height          =   615
      Left            =   9260
      TabIndex        =   39
      Top             =   510
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Style           =   3
      AnimationDuration=   0.2
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton8 
      Height          =   615
      Left            =   15840
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      Style           =   5
      AnimationDuration=   0.2
      Caption         =   "Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton3 
      Height          =   615
      Left            =   8640
      TabIndex        =   33
      Top             =   6720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "Show All Maps"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
      Height          =   615
      Left            =   4440
      TabIndex        =   31
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "BP Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
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
      Left            =   12360
      TabIndex        =   32
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "Manage Map Pool"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton7 
      Height          =   615
      Left            =   16200
      TabIndex        =   37
      Top             =   9555
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      Style           =   9
      AnimationDuration=   0.2
      Caption         =   "BP Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton6 
      Height          =   615
      Left            =   1440
      TabIndex        =   36
      Top             =   9555
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      Style           =   6
      AnimationDuration=   0.2
      Caption         =   "BP Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton4 
      Height          =   615
      Left            =   8400
      TabIndex        =   34
      Top             =   7440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      Style           =   13
      AnimationDuration=   0.2
      Caption         =   "Instruction"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Produced by Aox Holographic Entanglement"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6240
      TabIndex        =   30
      Top             =   9600
      Width           =   8535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12960
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line11 
      X1              =   2880
      X2              =   2895
      Y1              =   360
      Y2              =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MAP POOL"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   9150
      TabIndex        =   9
      Top             =   75
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   10335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20085
   End
End
Attribute VB_Name = "Form01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UI的hover显示
Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub


Private Sub AlphaBlendImage1_Click(Index As Integer)

Dim i As Integer: Dim j As Integer: Dim k As Integer
Dim index1 As Integer: Dim c As Integer: Dim Index0 As Integer

Index0 = theme - Index: index1 = Index0 + 1: searchmode = 0

'非显示全图的情况
VScroll2.Visible = False
If MapN(Index0) <> 0 Then
    For i = 0 To MapN(Index0) - 1
        Image21(i).Visible = True
        c = Maporder(Index0, i)
        If thememapbp(Index0, c) = False Then
            Image21(i).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & index1 & "_" & c & ".jpg")
            Image21(i).Enabled = True
        Else
            Image21(i).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & index1 & "_" & c & ".jpg")
            Image21(i).Enabled = False
        End If
    Next
End If
Do While i <= all_map - 1
    Image21(i).Visible = False
    i = i + 1
Loop
mapindex = Index0

VScroll2.Value = 0 '滚动条置顶
VScroll1.SetFocus '保证滚动条焦点

AlphaBlendImage4.Top = Label5(Index).Top

AlphaBlendImage4.Visible = True

End Sub


Private Sub AlphaBlendImage3_Click()

If WindowsMediaPlayer1.URL = "" Then
    Set AlphaBlendImage3.Picture = AlphaBlendImage3.GdipLoadPicture(App.Path & "\picture\stop.png")
    WindowsMediaPlayer1.URL = App.Path & "\Music\BGM.mp3"
Else
    WindowsMediaPlayer1.URL = ""
    Set AlphaBlendImage3.Picture = AlphaBlendImage3.GdipLoadPicture(App.Path & "\picture\play.png")
End If

End Sub



Private Sub Command3_Click()
Dim i As Integer
Picture2.Top = 0
VScroll2.Visible = False
VScroll2.Value = 0
If coun = 20 And gamestart = 0 And SV = 1 Then
   gamestart = 1
   labelclick = 0
   Command3.Picture = LoadPicture(App.Path & "\picture\notchosen.jpg")
   Command3.Visible = False
   Command4.Visible = True
   Text5.Visible = False '隐藏搜索栏
   ctxNineButton8.Visible = False
   Label14.Visible = False
   ctxNineButton3.Visible = False
   SV = 0
   Label6.Visible = True
   Label8.Visible = False
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   Label12.Visible = True
   Label10.Enabled = True
   Label12.Enabled = True
   For i = 0 To all_map - 1
       If i <> 7 Then
          Image21(i).Visible = False
       End If
   Next
   Image21(7).Visible = True
   Image21(7).Enabled = False
   Image21(7).Picture = LoadPicture(App.Path & "\picture\firstmap.jpg")
   Picture3.Visible = False
   ctxNineButton6.Visible = False
   ctxNineButton7.Visible = False
Else
   Beep
End If

End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim j As Integer

If PCoun <> 0 Then

For i = 0 To 9
    IE(i, PCoun) = False
Next
LV(0, PCoun) = False: LV(1, PCoun) = False
L10E(PCoun) = False
L12E(PCoun) = False
L10C(PCoun) = False
L12C(PCoun) = False
L6C(PCoun) = ""
rc(PCoun) = 0
LC(PCoun) = 0
For i = 0 To 2
    For j = 0 To 9
        MR(j, i, PCoun) = 0
    Next
Next
PCoun = PCoun - 1

Else
   Beep
End If
   

If PCoun <> 0 Then
      For i = 0 To 9
          Image6(MPI(i)).Enabled = IE(i, PCoun)
      Next
      Line1.Visible = LV(0, PCoun): Line2.Visible = LV(0, PCoun): Line3.Visible = LV(0, PCoun): Line4.Visible = LV(0, PCoun)
      Line5.Visible = LV(1, PCoun): Line6.Visible = LV(1, PCoun): Line7.Visible = LV(1, PCoun): Line8.Visible = LV(1, PCoun)
      Label10.Enabled = L10E(PCoun)
      Label12.Enabled = L12E(PCoun)
      Label10.Caption = L10C(PCoun)
      Label12.Caption = L12C(PCoun)
      Label6.Caption = L6C(PCoun)
      racecoun = rc(PCoun)
      labelclick = LC(PCoun)
      For i = 0 To 2
          For j = 0 To 9
             maprace(j, i) = MR(j, i, PCoun)
          Next
      Next
     For i = 0 To 9
         B = mapbp(MPI(i), 0) + 1
         c = mapbp(MPI(i), 1)
         Image6(MPI(i)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
     Next
     For i = 0 To racecoun - 1
         B = maprace(i, 0) + 1
         c = maprace(i, 1)
         Image6(maprace(i, 2)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
     Next
     If racecoun <> 0 Then
        B = maprace(racecoun - 1, 0) + 1
        c = maprace(racecoun - 1, 1)
        Image21(7).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
     Else
        Image21(7).Picture = LoadPicture(App.Path & "\" & PicturePath & "\firstmap.jpg")
     End If
Else
      For i = 0 To 9
          Image6(MPI(i)).Enabled = False
      Next
      Line1.Visible = False: Line2.Visible = False: Line3.Visible = False: Line4.Visible = False
      Line5.Visible = False: Line6.Visible = False: Line7.Visible = False: Line8.Visible = False
      Label10.Enabled = True
      Label12.Enabled = True
      Label10.Caption = 0
      Label12.Caption = 0
      Label6.Caption = "Select the winner"
      racecoun = 0
      labelclick = 0
      Image21(7).Picture = LoadPicture(App.Path & "\picture\firstmap.jpg")
      For i = 0 To 2
          For j = 0 To 9
             maprace(j, i) = 100
          Next
      Next
      For i = 0 To 9
         B = mapbp(MPI(i), 0) + 1
         c = mapbp(MPI(i), 1)
         Image6(MPI(i)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
      Next
End If

End Sub


Private Sub ctxNineButton1_Click()

Form02.Show
Form01.Enabled = False

End Sub

Private Sub ctxNineButton10_Click()

Form11.Show
Form01.Enabled = False

End Sub

Private Sub ctxNineButton2_Click()

Form01.Enabled = False
Form03.Show

End Sub

Private Sub ctxNineButton3_Click()

Dim i As Integer: Dim j As Integer: Dim k As Integer
searchmode = 0

'显示全图的情况
i = 0: j = 0: k = 0
VScroll2.Visible = True
VScroll2.SetFocus
Do Until i > all_map - 1
    If Maporder(j, k) <> 0 Then
        Image21(i).Visible = True
        If indexbp(i) = False Then
            Image21(i).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & j + 1 & "_" & Maporder(j, k) & ".jpg")
            Image21(i).Enabled = True
        Else
            Image21(i).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & j + 1 & "_" & Maporder(j, k) & ".jpg")
            Image21(i).Enabled = False
        End If
        i = i + 1
        k = k + 1
    Else
        j = j + 1
        k = 0
    End If
Loop
mapindex = 100

VScroll2.Value = 0 '滚动条置顶

AlphaBlendImage4.Visible = False

End Sub

Private Sub ctxNineButton4_Click()

Form06.Show

End Sub

Private Sub ctxNineButton5_Click(Index As Integer)

anticlock = waitingtime * FirstRoundNumber '倒计时器获取等待时间


'--------------图池数据传递--------------

If Index = 0 Then
    ALLMAP = ALLMAP_R
    AllMapN = AllMapN_R
    AllMaporder = AllMaporder_R
    MapN = MapN_R
    Maporder = Maporder_R
    Mapname = Mapname_R
    all_map = all_map_R
    PicturePath = "picture"
Else
    ALLMAP = ALLMAP_I
    AllMapN = AllMapN_I
    AllMaporder = AllMaporder_I
    MapN = MapN_I
    Maporder = Maporder_I
    Mapname = Mapname_I
    all_map = all_map_I
    PicturePath = "item picture"
End If

'--------------传递结束--------------


'--------------控件预处理--------------

Label8.Visible = True: Label14.Visible = True
ctxNineButton5(0).Visible = False: ctxNineButton5(1).Visible = False: ctxNineButton6.Visible = True
ctxNineButton7.Visible = True: ctxNineButton8.Visible = True: ctxNineButton4.Visible = False
ctxNineButton3.Visible = True: ctxNineButton1.Visible = False: ctxNineButton2.Visible = False
ctxNineButton10.Visible = False
Picture3.Visible = True
Text5.Visible = True '显示搜索栏
For i = 0 To 3
    Shape1(i).Visible = True
Next
VScroll1.SetFocus

Dim m As Integer: m = 0
For i = 15 To 200
    Load Image21(i)
    If m <> 0 Then
        Image21(i).Left = Image21(i - 1).Left + 2280
        Image21(i).Top = Image21(i - 1).Top
        Image21(i).Visible = False: Image21(i).ZOrder 0
        m = m + 1
        If m = 5 Then
            m = 0
        End If
    Else
        Image21(i).Left = Image21(i - 5).Left
        Image21(i).Top = Image21(i - 5).Top + 1695
        Image21(i).Visible = False: Image21(i).ZOrder 0
        m = 1
    End If
Next

'--------------预处理结束--------------



'--------------映射数组赋值--------------

Call MappingArrayInitialize

'--------------赋值结束--------------


If (all_map Mod 5) = 0 Then
    VScroll2.Max = (all_map - 15) \ 5
Else
    VScroll2.Max = (all_map - 15) \ 5 + 1
End If
    
If Dir(App.Path & "\Music\BGM.mp3") <> "" Then
    WindowsMediaPlayer1.URL = App.Path & "\Music\BGM.mp3"
    AlphaBlendImage3.Visible = True
    Set AlphaBlendImage3.Picture = AlphaBlendImage3.GdipLoadPicture(App.Path & "\picture\stop.png")
End If

Call ImageOrderMove

'--------------动画相关--------------

Frame2.Visible = True: Frame4.Visible = True
Timer5.Enabled = True

End Sub

Private Sub ctxNineButton6_Click()

'保持搜索栏焦点
Text5.SetFocus

Dim i As Integer: Dim precoun As Integer

precoun = coun - 1: i = 0

If precoun >= 0 Then
    If PointAttriJudge(coun).BlueStopOrNot And blueconfirm(PointAttriArray(precoun).Turns) = 0 Then
        For i = 0 To UBound(PointAttriArray(precoun).AfterStop)
            Image6(PointAttriArray(precoun).AfterStop(i)).Picture = LoadPicture(App.Path & "\picture\forchosen2.jpg") '后方节点置于预选状态
        Next
        anticlock = waitingtime * (UBound(PointAttriArray(precoun).AfterStop) + 1)
        For i = 0 To UBound(PointAttriArray(precoun).BeforeStop)
            Image6(PointAttriArray(precoun).BeforeStop(i)).Enabled = False '前方节点无效化
        Next
        blueconfirm(PointAttriArray(precoun).Turns) = 1 '记录confirm数组
    Else
        Beep
    End If
Else
    Beep
End If


End Sub


Private Sub ctxNineButton7_Click()

'保持搜索栏焦点
Text5.SetFocus

Dim i As Integer: Dim precoun As Integer

precoun = coun - 1: i = 0

If precoun >= 0 Then
    If PointAttriJudge(coun).RedStopOrNot And redconfirm(PointAttriArray(precoun).Turns) = 0 Then
        If coun <= AllBPTurns Then
            For i = 0 To UBound(PointAttriArray(precoun).AfterStop)
                Image6(PointAttriArray(precoun).AfterStop(i)).Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg") '后方节点置于预选状态
            Next
            anticlock = waitingtime * (UBound(PointAttriArray(precoun).AfterStop) + 1)
        End If
        For i = 0 To UBound(PointAttriArray(precoun).BeforeStop)
            Image6(PointAttriArray(precoun).BeforeStop(i)).Enabled = False '前方节点无效化
        Next
        redconfirm(PointAttriArray(precoun).Turns) = 1 '记录confirm数组
    Else
        Beep
    End If
Else
    Beep
End If

If coun = AllBPTurns + 1 Then
   Command3.Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg")
   SV = 1
   Timer1.Enabled = False
   ctxNineButton9.Visible = True
End If

End Sub

Private Sub ctxNineButton8_Click()

Dim i As Integer: Dim j As Integer
Dim B As Integer: Dim c As Integer
Dim searchcoun As Integer

searchcoun = 0
Picture2.Top = 0 '初始化全图显示
VScroll2.Visible = False
VScroll2.Value = 0
searchmode = 1
For i = 0 To 4 '映射数组初始化
    searchindex2theme(i) = 0
    searchindex2map(i) = 0
Next
For i = 0 To theme
   For j = 0 To 12
      thememap2searchindex(i, j) = 0
   Next
Next


If searchsum > 5 Or searchsum = 0 Then
   Form07.Show
   Beep
Else
   i = 0
   j = 0
   For j = 0 To theme
     Do Until Maporder(j, i) = 0
        If keywordsearch(Text5.Text, Mapname(j, i)) = True Then
           B = j + 1
           c = Maporder(j, i)
           searchindex2theme(searchcoun) = j '映射数组赋值
           searchindex2map(searchcoun) = c
           thememap2searchindex(j, c) = searchcoun
           If bpsearch(B - 1, c) = True Then
              Image21(searchcoun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
              Image21(searchcoun).Enabled = False
           Else
              Image21(searchcoun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
              Image21(searchcoun).Enabled = True
           End If
           Image21(searchcoun).Visible = True
           searchcoun = searchcoun + 1
        End If
        i = i + 1
     Loop
     i = 0
   Next
   Do Until searchcoun > all_map - 1
      Image21(searchcoun).Visible = False
      searchcoun = searchcoun + 1
   Loop
End If

If searchsum = 1 And Image21(0).Enabled = True Then
    Image21_Click (0)
    If PointAttriJudge(coun).BlueStopOrNot Then
       ctxNineButton6.SetFocus
    ElseIf PointAttriJudge(coun).RedStopOrNot Then
       ctxNineButton7.SetFocus
    End If
End If

AlphaBlendImage4.Visible = False

End Sub

Private Sub ctxNineButton9_Click()

coun = 0
Dim i As Integer: Dim j As Integer: Dim k As Integer
Dim s As String

i = 0
Do While i <= 19
   Image6(i).Picture = LoadPicture(App.Path & "\picture\notchosen.jpg")
   Image6(i).Enabled = True
   i = i + 1
Loop

s = Mid(PointAttriArray(0).TeamBPTypes, 1, 1) '初始预选提示
i = 0
Do Until s <> Mid(PointAttriArray(i).TeamBPTypes, 1, 1)
    Image6(i).Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg")
    i = i + 1
Loop

For i = 0 To 19
    mapbp(i, 0) = 100
    mapbp(i, 1) = 100
Next

VScroll2.Visible = False
VScroll2.Value = 0
Picture2.Top = 0

i = 0: j = 0: k = 0
Do Until i > all_map - 1
      If Maporder(j, k) <> 0 Then
         Image21(i).Visible = False
         Image21(i).Enabled = True
         i = i + 1
         k = k + 1
      Else
         j = j + 1
         k = 0
      End If
Loop

For i = 0 To 100
   indexbp(i) = False
Next
For i = 0 To theme
   For j = 0 To 12
      thememapbp(i, j) = False
   Next
Next

For i = 0 To UBound(blueconfirm)
    blueconfirm(i) = 0: redconfirm(i) = 0
Next


anticlock = waitingtime * FirstRoundNumber
Timer1.Enabled = True

gamestart = 0
Label6.Visible = False: Label8.Visible = True: Label9.Visible = False: Label10.Visible = False
Label11.Visible = False: Label12.Visible = False: Label13.Visible = False
Label10.Caption = 0: Label12.Caption = 0


For i = 0 To 9
    maprace(i, 0) = 100
    maprace(i, 1) = 100
    maprace(i, 2) = 100
Next
racecoun = 0

Line1.Visible = False: Line2.Visible = False: Line3.Visible = False: Line4.Visible = False
Line5.Visible = False: Line6.Visible = False: Line7.Visible = False: Line8.Visible = False

Picture3.Visible = True
ctxNineButton6.Visible = True: ctxNineButton7.Visible = True: Command4.Visible = False
ctxNineButton3.Visible = True
PCoun = 0

ctxNineButton9.Visible = False

Text5.Visible = True '显示搜索栏
ctxNineButton8.Visible = True
Label14.Visible = True

'BP交换
'无畏杯联赛专用
'If PicturePath = "picture" Then
'    PicturePath = "item picture"
'    ALLMAP = ALLMAP_I
'    AllMapN = AllMapN_I
'    AllMaporder = AllMaporder_I
'    MapN = MapN_I
'    Maporder = Maporder_I
'    Mapname = Mapname_I
'    all_map = all_map_I
'Else
'    PicturePath = "picture"
'    ALLMAP = ALLMAP_R
'    AllMapN = AllMapN_R
'    AllMaporder = AllMaporder_R
'    MapN = MapN_R
'    Maporder = Maporder_R
'    Mapname = Mapname_R
'    all_map = all_map_R
'    If Dir(App.Path & "\Music\BGM.mp3") <> "" Then '重新播放音乐
'        WindowsMediaPlayer1.URL = ""
'        WindowsMediaPlayer1.URL = App.Path & "\Music\BGM.mp3"
'        AlphaBlendImage3.Visible = True
'        Set AlphaBlendImage3.Picture = AlphaBlendImage3.GdipLoadPicture(App.Path & "\picture\stop.png")
'    End If
'End If
'
'Call MappingArrayInitialize

End Sub

Private Sub Form_Load()

Dim i As Integer: Dim j As Integer: Dim n As Integer
Dim strx As String

'---------------变量初始化---------------'

'主题更新时更新这里
theme = 27: ALLMAP_R = 112: ALLMAP_I = 121: PositionAmount = 200

coun = 0: imagenumber = 15: gamestart = 0
co = 0: co1 = 0
racecoun = 0: recorder = 0: PCoun = 0

ReDim ThemeName(theme)
ThemeName(0) = "Ice": ThemeName(1) = "Desert": ThemeName(2) = "Forest": ThemeName(3) = "Village": ThemeName(4) = "Tomb"
ThemeName(5) = "Mine": ThemeName(6) = "Northeu": ThemeName(7) = "Factory": ThemeName(8) = "Gold": ThemeName(9) = "China"
ThemeName(10) = "Moonhill": ThemeName(11) = "Pirate": ThemeName(12) = "Fairy": ThemeName(13) = "Nymph": ThemeName(14) = "Castle"
ThemeName(15) = "Mechanic": ThemeName(16) = "WKC": ThemeName(17) = "Brodi": ThemeName(18) = "Beach": ThemeName(19) = "Jurassic"
ThemeName(20) = "World": ThemeName(21) = "Steam": ThemeName(22) = "Nemo": ThemeName(23) = "Sword": ThemeName(24) = "God"
ThemeName(25) = "Abyss": ThemeName(26) = "Camelot": ThemeName(27) = "Olympus"

ReDim MapN_I(theme): ReDim Maporder_I(theme, 12)
ReDim Mapname_I(theme, 12): ReDim AllMapN_I(theme): ReDim AllMaporder_I(theme, 12)
ReDim MapN_R(theme): ReDim Maporder_R(theme, 12)
ReDim Mapname_R(theme, 12): ReDim AllMapN_R(theme): ReDim AllMaporder_R(theme, 12)
ReDim thememapbp(theme, 12): ReDim thememap2searchindex(theme, 12)
ReDim thememapsave(theme, 12): ReDim thememapsaveI(theme, 12): ReDim th_map2index_pc(theme, 12)
ReDim maporder_save(theme, 12): ReDim mapnumber_save(theme)

'---------------初始化结束---------------'


'---------------滚动条---------------'

'取得控件的句柄
hwndVS = VScroll1.hwnd
'保存smMap控件的默认窗口消息处理函数地址
OldWindowProc = GetWindowLong(VScroll1.hwnd, GWL_WNDPROC)
'将smMap控件的消息处理函数指定为自定义函数NewWindowProc
Call SetWindowLong(VScroll1.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)

hwndVS2 = VScroll2.hwnd
OldWindowProc2 = GetWindowLong(VScroll2.hwnd, GWL_WNDPROC)
Call SetWindowLong(VScroll2.hwnd, GWL_WNDPROC, AddressOf NewWindowProc2)

'---------------滚动条结束---------------'


'---------------控件预处理---------------'

If Dir(App.Path & "\picture\background.jpg") <> "" Then '支持四种图片格式
    Image1.Picture = LoadPicture(App.Path & "\picture\background.jpg")
ElseIf Dir(App.Path & "\picture\background.gif") <> "" Then
    Image1.Picture = LoadPicture(App.Path & "\picture\background.gif")
ElseIf Dir(App.Path & "\picture\background.png") <> "" Then
    Set Image1.Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\background.png")
ElseIf Dir(App.Path & "\picture\background.bmp") <> "" Then
    Image1.Picture = LoadPicture(App.Path & "\picture\background.bmp")
End If

Label9.Visible = False: Label10.Visible = False: Label11.Visible = False: Label12.Visible = False

Set AlphaBlendImage1(0).Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\map" & theme & ".png")
Label5(0).Caption = ThemeName(theme)
For i = 1 To theme
    Load AlphaBlendImage1(i): Load Label5(i)
    AlphaBlendImage1(i).Left = AlphaBlendImage1(0).Left
    Label5(i).Left = Label5(0).Left
    AlphaBlendImage1(i).Top = AlphaBlendImage1(i - 1).Top + 1050
    Label5(i).Top = Label5(i - 1).Top + 1050
    AlphaBlendImage1(i).Visible = True: Label5(i).Visible = True
    AlphaBlendImage1(i).ZOrder 0: Label5(i).ZOrder 0
    Set AlphaBlendImage1(i).Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\map" & theme - i & ".png")
    Label5(i).Caption = ThemeName(theme - i)
Next
For i = 0 To imagenumber - 1
    Image21(i).Visible = False
Next

i = 0
Do While i <= 19
    Image6(i).Picture = LoadPicture(App.Path & "\picture\notchosen.jpg")
    i = i + 1
Loop

For i = 0 To 19
    Image6(i).Visible = False
Next

VScroll1.Max = theme - 4

Set AlphaBlendImage2.Picture = AlphaBlendImage1(0).GdipLoadPicture(App.Path & "\picture\Team_Icon.png")

'---------------预处理结束---------------'


MPI(0) = 7: MPI(1) = 8: MPI(2) = 11: MPI(3) = 16: MPI(4) = 19:
MPI(5) = 6: MPI(6) = 9: MPI(7) = 10: MPI(8) = 17: MPI(9) = 18:

RRRA(0) = 7: RRRA(1) = 8: RRRA(2) = 11: RRRA(3) = 16: RRRA(4) = 19:
RRRA(5) = 6: RRRA(6) = 9: RRRA(7) = 10: RRRA(8) = 17: RRRA(9) = 18:

For i = 0 To 19
    mapbp(i, 0) = 100
    mapbp(i, 1) = 100
    mapbp(i, 2) = 100
Next

For i = 0 To 9
    maprace(i, 0) = 100
    maprace(i, 1) = 100
    maprace(i, 2) = 100
Next


'---------------BP时间读取---------------'

Open App.Path & "\map pool data\current waiting time.txt" For Input As #1
Line Input #1, strx
waitingtime = Int(strx)
Close #1
'---------------时间读取结束---------------'


'---------------图池数据载入---------------'

Call AllPoolDataLoad

Open App.Path & "\map pool data\current pool index.txt" For Input As #1 '读取上一次关闭前的图池选择
Line Input #1, strx
    map_pool_choice = Int(strx)
Close #1
Open App.Path & "\item pool data\current pool index.txt" For Input As #1
Line Input #1, strx
    map_pool_choice_I = Int(strx)
Close #1

i = 0 '读取上一次关闭程序前选择的图池
Dim name As String
Open App.Path & "\map pool data\pool name.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, strx
   If map_pool_choice = i Then
      name = strx
      Exit Do
   End If
   i = i + 1
Loop
Close #1
Call PoolDataLoad_R(name)

i = 0
Open App.Path & "\item pool data\pool name.txt" For Input As #1
Do While Not EOF(1)
   Line Input #1, strx
   If map_pool_choice_I = i Then
      name = strx
      Exit Do
   End If
   i = i + 1
Loop
Close #1
Call PoolDataLoad_I(name)

'---------------载入结束---------------'


'---------------BP数据载入---------------'

Dim tembluebp As String: Dim temredbp As String '预置BP测试

Open App.Path & "\BP data\Current BP Choice.txt" For Input As #1
Line Input #1, strx
    BP_choice() = Split(strx, " ")
Close #1

Open App.Path & "\BP data\" & BP_choice(1) & ".txt" For Input As #1
Line Input #1, tembluebp
Line Input #1, temredbp
Close #1

BlueBPdata() = Split(tembluebp, "-")
RedBPdata() = Split(temredbp, "-")

Call BPDataLoad

'---------------载入结束---------------'


'---------------动画参数初始化---------------'

ReDim Amplitude(3): ReDim InitialPosition(3)
time = 0: Period = 1: StopTime = Period / 2: TimeInterval = Timer5.Interval / 1000
Amplitude(0) = (6630 - 750) / 2: Amplitude(1) = 5295 / 2
Amplitude(2) = (6630 - 750) / 2: Amplitude(3) = 5295 / 2
InitialPosition(0) = 750: InitialPosition(1) = 15
InitialPosition(2) = 750: InitialPosition(3) = 15

'---------------初始化结束---------------'

End Sub


Private Sub Form_Unload(Cancel As Integer)

Open App.Path & "\map pool data\current pool index.txt" For Output As #1 '保存关闭程序前的操作记录
Print #1, map_pool_choice
Close #1

Open App.Path & "\item pool data\current pool index.txt" For Output As #1
Print #1, map_pool_choice_I
Close #1

Open App.Path & "\map pool data\current waiting time.txt" For Output As #1
Print #1, waitingtime
Close #1

Open App.Path & "\BP data\Current BP Choice.txt " For Output As #1
Print #1, Trim(BP_choice(0)) & " " & Trim(BP_choice(1))
Close #1

End Sub



Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Hover

End Sub


Private Sub Image21_Click(Index As Integer)
Dim A As Integer: Dim B As Integer: Dim c As Integer
A = coun + 1
Dim temjudge As points_judge: Dim tembp As String

If coun <= AllBPTurns Then
    tembp = Mid(PointAttriArray(coun).TeamBPTypes, 2, 1)
Else
    tembp = ""
End If
temjudge = PointAttriJudge(coun)

If temjudge.BlueStopOrNot And temjudge.BlueConfirmOrNot Then
    Beep
ElseIf temjudge.RedStopOrNot And temjudge.RedConfirmOrNot Then
    Beep
ElseIf coun > AllBPTurns Then
    Beep
Else
  If searchmode <> 1 Then '非搜索模式下的bp
     If mapindex < theme + 1 Then '非显示全图的情况
        B = mapindex + 1
        c = Maporder(mapindex, Index)
        Image21(Index).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
        Image21(Index).Enabled = False
        If tembp = "B" Then 'Ban节点
           Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
        Else
           Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
        End If
        mapbp(coun, 0) = mapindex 'bp数组记录
        mapbp(coun, 1) = c
        mapbp(coun, 2) = Index
        indexbp(thememap2index(B - 1, c)) = True
        thememapbp(mapindex, c) = True
        coun = coun + 1
        If VScroll1.Visible = True Then
           VScroll1.SetFocus
        End If
     Else '显示全图的情况
        B = index2theme(Index) + 1
        c = index2map(Index)
        If tembp = "B" Then
           Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
        Else
           Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
        End If
        Image21(Index).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
        Image21(Index).Enabled = False
        mapbp(coun, 0) = B - 1 'bp数组记录
        mapbp(coun, 1) = c
        mapbp(coun, 2) = allindex2mapindex(Index)
        indexbp(Index) = True
        thememapbp(B - 1, c) = True
        coun = coun + 1
        If VScroll2.Visible = True Then
           VScroll2.SetFocus
        End If
     End If
   Else '搜索模式下的bp
      B = searchindex2theme(Index) + 1
      c = searchindex2map(Index)
      If tembp = "B" Then
         Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
      Else
         Image6(coun).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B & "_" & c & ".jpg")
      End If
      Image21(Index).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
      Image21(Index).Enabled = False
      mapbp(coun, 0) = B - 1 'bp数组记录
      mapbp(coun, 1) = c
      mapbp(coun, 2) = thememap2index_1(B - 1, c)
      indexbp(thememap2index(B - 1, c)) = True
      thememapbp(B - 1, c) = True
      coun = coun + 1
   End If
End If


End Sub



Private Sub Image6_Click(Index As Integer)
Dim A As Integer: Dim B As Integer: Dim c As Integer: Dim d As Integer

If gamestart = 0 Then
  If Index = coun - 1 Then
     coun = coun - 1
     A = coun + 1
     If PointAttriJudge(coun).NowTeam = "R" Then '红方节点
        Image6(Index).Picture = LoadPicture(App.Path & "\picture\forchosen2.jpg")
     Else
        Image6(Index).Picture = LoadPicture(App.Path & "\picture\forchosen1.jpg")
     End If
     B = mapbp(coun, 0) '记录上一次bp的地图序号
     c = mapbp(coun, 1)
     d = mapbp(coun, 2)
     mapbp(coun, 0) = 100 'bp撤回
     mapbp(coun, 1) = 100
     mapbp(coun, 2) = 100
     indexbp(thememap2index(B, c)) = False
     thememapbp(B, c) = False
     If searchmode <> 1 Then '非搜索模式下
       If mapindex < theme + 1 Then '非全图显示
          If B = mapindex Then
             Image21(d).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B + 1 & "_" & c & ".jpg")
             Image21(d).Enabled = True
          End If
       Else '全图显示
          Image21(thememap2index(B, c)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B + 1 & "_" & c & ".jpg")
          Image21(thememap2index(B, c)).Enabled = True
       End If
     Else '搜索模式下
        Image21(thememap2searchindex(B, c)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\no" & B + 1 & "_" & c & ".jpg")
        Image21(thememap2searchindex(B, c)).Enabled = True
     End If
  Else
     Beep
  End If
ElseIf gamestart = 1 Then
    Image21(7).Picture = Image6(Index).Picture
    maprace(racecoun, 0) = mapbp(Index, 0)
    maprace(racecoun, 1) = mapbp(Index, 1)
    maprace(racecoun, 2) = Index
    Label6.Caption = "Select the winner"
    Label10.Enabled = True
    Label12.Enabled = True
    B = maprace(racecoun, 0) + 1
    c = maprace(racecoun, 1)
    Image6(maprace(racecoun, 2)).Picture = LoadPicture(App.Path & "\" & PicturePath & "\noo" & B & "_" & c & ".jpg")
    Image6(maprace(racecoun, 2)).Enabled = False
    racecoun = racecoun + 1
    Image6(7).Enabled = False
    Image6(8).Enabled = False
    Image6(11).Enabled = False
    Image6(16).Enabled = False
    Image6(19).Enabled = False
    Image6(6).Enabled = False
    Image6(9).Enabled = False
    Image6(10).Enabled = False
    Image6(17).Enabled = False
    Image6(18).Enabled = False
    recorder = recorder + 1
    Text4.Text = recorder
End If
   
End Sub

Private Sub Label10_Click()
Dim i As Integer
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
labelclick = labelclick + 1
Image6(7).Enabled = True
Image6(8).Enabled = True
Image6(11).Enabled = True
Image6(16).Enabled = True
Image6(19).Enabled = True
Image6(6).Enabled = True
Image6(9).Enabled = True
Image6(10).Enabled = True
Image6(17).Enabled = True
Image6(18).Enabled = True
Label10.Enabled = False
Label12.Enabled = False
For i = 0 To racecoun - 1
    Image6(maprace(i, 2)).Enabled = False
    Label6.Visible = True
Next

If labelclick = 1 Then
   Image21(7).Picture = LoadPicture("")
   Label6.Visible = True
Else
   Label10.Caption = Label10.Caption + 1
End If
If Label10.Caption = 5 Then
   Label13.Visible = True
   Label13.Caption = Label9.Caption & " Win"
   Label13.ForeColor = &HFF0000
   Label10.Enabled = False
   Label12.Enabled = False
   For i = 0 To 9
    Image6(RRRA(i)).Enabled = False
   Next
   Label6.Caption = ""
   Label6.Visible = False
   Command4.Visible = False
   ctxNineButton9.Visible = True
Else
   Label6.Caption = "Select the next map"
End If
recorder = recorder + 1
Text4.Text = recorder
End Sub

Private Sub Label12_Click()
Dim i As Integer

Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
labelclick = labelclick + 1

Image6(7).Enabled = True
Image6(8).Enabled = True
Image6(11).Enabled = True
Image6(16).Enabled = True
Image6(19).Enabled = True
Image6(6).Enabled = True
Image6(9).Enabled = True
Image6(10).Enabled = True
Image6(17).Enabled = True
Image6(18).Enabled = True
For i = 0 To racecoun - 1
    Image6(maprace(i, 2)).Enabled = False
Next


Label10.Enabled = False
Label12.Enabled = False
If labelclick = 1 Then
   Image21(7).Picture = LoadPicture("")
   Label6.Visible = True
Else
   Label12.Caption = Label12.Caption + 1
End If
If Label12.Caption = 5 Then
   Label13.Visible = True
   Label13.Caption = Label11.Caption & " Win"
   Label13.ForeColor = &HFF&
   Label10.Enabled = False
   Label12.Enabled = False
   For i = 0 To 9
    Image6(RRRA(i)).Enabled = False
   Next
   Label6.Caption = ""
   Label6.Visible = False
   Command4.Visible = False
   ctxNineButton9.Visible = True
Else
   Label6.Caption = "Select the next map"
End If
recorder = recorder + 1
Text4.Text = recorder
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Hover

End Sub

Private Sub Picture4_Click()

'保持滚动条焦点
VScroll1.SetFocus

End Sub

Private Sub Text2_Change()
Label9.Caption = Text2.Text
End Sub

Private Sub Text3_Change()
Label11.Caption = Text3.Text
End Sub

Private Sub Text4_Change()

Dim i As Integer
Dim j As Integer
PCoun = PCoun + 1
For i = 0 To 9
     IE(i, PCoun) = Image6(MPI(i)).Enabled
Next
LV(0, PCoun) = Line1.Visible: LV(1, PCoun) = Line5.Visible
L10E(PCoun) = Label10.Enabled
L12E(PCoun) = Label12.Enabled
L10C(PCoun) = Label10.Caption
L12C(PCoun) = Label12.Caption
L6C(PCoun) = Label6.Caption
rc(PCoun) = racecoun
LC(PCoun) = labelclick
For i = 0 To 2
    For j = 0 To 9
        MR(j, i, PCoun) = maprace(j, i)
    Next
Next

End Sub



Private Sub Text5_Change()
Dim i As Integer
Dim j As Integer

i = 0
j = 0
searchsum = 0
For j = 0 To theme
    Do Until Maporder(j, i) = 0
       If keywordsearch(Text5.Text, Mapname(j, i)) = True Then
          searchsum = searchsum + 1
       End If
    i = i + 1
    Loop
    i = 0
Next

Label14.Caption = searchsum & "results found"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

If KeyAscii >= 65 And KeyAscii <= 90 Then
   KeyAscii = KeyAscii + 32
ElseIf KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
ElseIf KeyAscii = 13 Then
   ctxNineButton8_Click
   KeyAscii = 0
Else
   KeyAscii = 0
End If

End Sub

Private Sub Timer1_Timer()

If anticlock > 0 Then
    anticlock = anticlock - 1
    Label8.Caption = anticlock
ElseIf anticlock = 0 Then
    
    Dim i As Integer: Dim j As Integer: Dim TemCoun As Integer
    Dim t As String: Dim TemIndex As Integer
    
    i = 0: j = 0: TemCoun = coun
    t = PointAttriJudge(TemCoun).NowTeam
    TemIndex = TemCoun + i
    
    ctxNineButton3_Click
    
    Do Until t <> PointAttriJudge(TemIndex).NowTeam
        Do Until Image21(j).Enabled = True
            j = j + 1
        Loop
        Image21_Click (j)
        i = i + 1
        TemIndex = TemCoun + i
        If TemIndex > AllBPTurns + 1 Then '防止越界
            Exit Do
        End If
    Loop
    
    If PointAttriJudge(coun).NowTeam = "R" Then
        ctxNineButton6_Click
    Else
        ctxNineButton7_Click
    End If
   
    Label8.Caption = anticlock
    
End If


End Sub


Private Sub Timer5_Timer()

time = time + TimeInterval

If time <= StopTime Then
    Frame1.Height = InitialPosition(0) + Amplitude(0) * AnimationXofT(time / Period)
    Frame2.Width = InitialPosition(1) + Amplitude(1) * AnimationXofT(time / Period)
    Frame3.Height = Frame1.Height
    Frame4.Width = Frame2.Width
    Frame4.Left = 19935 - Frame4.Width
Else
    Timer5.Enabled = False
    Timer1.Enabled = True
End If


End Sub

Private Sub VScroll1_Change()
Picture4.Top = -CLng(1050) * VScroll1.Value
If VScroll1.Visible = True Then
   VScroll1.SetFocus
End If
End Sub


Private Sub VScroll1_Scroll()
Picture4.Top = -CLng(1050) * VScroll1.Value
If VScroll1.Visible = True Then
   VScroll1.SetFocus
End If
End Sub


Private Sub VScroll2_Change()
Picture2.Top = -CLng(1695) * VScroll2.Value - 120
If VScroll2.Visible = True Then
   VScroll2.SetFocus
End If
End Sub


Private Sub VScroll2_Scroll()
Picture2.Top = -CLng(1695) * VScroll2.Value - 120
If VScroll2.Visible = True Then
   VScroll2.SetFocus
End If
End Sub

Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)

If NewState = 1 Then
    WindowsMediaPlayer1.Controls.play
End If

End Sub
