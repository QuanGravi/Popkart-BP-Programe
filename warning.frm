VERSION 5.00
Begin VB.Form Form07 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Warning"
   ClientHeight    =   1845
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin AoxLeague_Ver.ctxNineButton ctxNineButton1 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
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
      Caption         =   $"warning.frx":0000
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "warning.frx":00A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctxNineButton1_Click()
Unload Form07
End Sub
