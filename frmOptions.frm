VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   172
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.HScrollBar speed 
      Height          =   195
      LargeChange     =   5
      Left            =   840
      Max             =   10
      Min             =   1
      TabIndex        =   0
      Top             =   240
      Value           =   1
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saved Game:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
If MsgBox("Are you sure you want to delete the currently saved game?", vbYesNo, "Saved Game") = vbNo Then Exit Sub

Kill App.path + "\save.dat"
Open App.path + "\save.dat" For Output As #1
Write #1, , , , , "NO SAVE FILE"
Close #1
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
speed.Value = 1100 - gSpeed
End Sub

Private Sub Speed_Change()
If gSpeed > 1100 Then gSpeed = 1100
gSpeed = 1100 - (speed.Value * 100)
SaveSetting "MAGICRPG", "GEN", "SPEED", gSpeed
End Sub
