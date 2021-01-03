VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Warning"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nexa-Bold"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Style           =   3
      AnimationDuration=   0.2
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Manteka"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "BP_manage_warning.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()
Unload Form12
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form11.Enabled = True
End Sub
