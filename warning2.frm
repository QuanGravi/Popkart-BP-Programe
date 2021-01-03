VERSION 5.00
Begin VB.Form Form08 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Warning"
   ClientHeight    =   1260
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   600
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The amount of maps must be greater than 6"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "warning2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()
Unload Form08
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form03.Enabled = True
End Sub
