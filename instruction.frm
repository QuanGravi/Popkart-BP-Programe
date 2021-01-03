VERSION 5.00
Begin VB.Form Form06 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   11205
   ClientLeft      =   2265
   ClientTop       =   390
   ClientWidth     =   18405
   Icon            =   "instruction.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   11205
   ScaleWidth      =   18405
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Instruction"
      BeginProperty Font 
         Name            =   "Manteka"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   17595
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   11415
      Left            =   0
      Picture         =   "instruction.frx":048A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18375
   End
End
Attribute VB_Name = "Form06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1.Caption = vbCrLf & "1.The order of BanPick is an analog to that of LOL." & vbCrLf & _
vbCrLf & "2.The time for BanPick is limited. If you have not confirmed your selection in the given time, the programme will select the map randomly." & vbCrLf & _
vbCrLf & "3.Please click the 'Confirm' button every time after you finish the map selection. Before that, you can click the map you have chosen to remove it from your choice. Once the 'Confirm' button is clicked, your choice can not be changed." & vbCrLf & _
vbCrLf & "4.The name of your team can be filled in the textbox at the top of this programe." & vbCrLf & _
vbCrLf & "5.When BP is over, a 'clear' button will show up, which is used to restart the BP." & vbCrLf & _
vbCrLf & "6.Other Functions:" & vbCrLf & "(1) The Search box and 'Search' button will appear after BanPick is started. Type in the Chinese phonetic alphabet of the map to search the map you want (Keyword searching is supported). The programe will help you select the map when 'results found' is 1. In this situation, type 'Enter' to confirm your choice." & vbCrLf _
& "(2) 'BP SETTINGS' is used to set the map pool, BanPick time or BP mechanism. Instruction for this module is shown in itself. These settings will be saved after the programme is closed." & vbCrLf _
& "(3) 'MANAGE MAP POOL' is used to create a map pool. Instruction for this module is shown in itself." & vbCrLf & _
vbCrLf & "7.'MANAGE BP' is used to create a new BP choice. Some rules need to be clarified here:" & vbCrLf _
& "(1) Take 5P5B as an example. For Blue team, the code '1B-2B-1P-2P-2B-2P' means Blue team will Ban 1 map - Ban 2 map - Pick 1 map - Pick 2 map ...following the oreder. For Red team, the code '2B-1B-2P-1P1B-1B1P-1P' means Red team will Ban 2 map - Ban 1 map - Pick 2 map - Pick 1 map and Ban 1 map... So the BP process will follow like this: Blue Ban 1 map - Red Ban 2 map - Blue Ban 2 map - Red Ban 1 map -  Blue Pick 1 map - Red pick 2 map... The Format is: Number + B/P or Number + B/P + Number + B/P, splited with '-'. " & vbCrLf _
& "(2) Following the rules above, you can create a new BP choice you need. Besides, the textbox only allows you to type in number, -, B and P( Capital mode needed). Other input can not be taken into the programe." & vbCrLf _
& "(3) Before saving a new BP choice, some requirement need to be satisfied:" & vbCrLf _
& " First, number of turns Blue teams take to Ban and Pick should equal to that of Red Teams. For example, Blue team goes through 6 turns in 5P5B, which equals to that of Red team." & vbCrLf _
& " Second, the number of maps Blue teams pick should equal to that of Red teams, and it could not be zero. Similarly, the number of banned maps should also be same, but this number can be zero." & vbCrLf _
& " If one of these requirement are not satisfied, your BP choice can not pass the programe check and can not be saved successfully."
'ttttttttttttttttttttttt

End Sub

