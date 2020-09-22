VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "King Of Fighters"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "<-- Back"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLICK HERE FOR KOF 2000 ROM (PROJECT - 0)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2400
      Picture         =   "Form4.frx":0000
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Picture         =   "Form4.frx":1318
      Top             =   2040
      Width           =   1350
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form4
Form2.Show
End Sub

Private Sub Command5_Click()
Call Shell("Start.exe " & "http://www.chinawolf.com/~neozero/main.html", 0)
End Sub

Private Sub Image1_Click()
Call Shell("Start.exe " & "http://www.chinawolf.com/~neozero/main.html", 0)
End Sub

Private Sub Image2_Click()
Call Shell("Start.exe " & "Http://www.KofPerfect.com", 0)
End Sub
