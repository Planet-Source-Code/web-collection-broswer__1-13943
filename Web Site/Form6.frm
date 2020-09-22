VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Music Page"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   3030
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EFinder"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "MP3 Songs"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3765
      Left            =   240
      Picture         =   "Form6.frx":0000
      Top             =   120
      Width           =   8265
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Shell("Start.exe " & "Http://www.emp3finder.com", 0)
End Sub

Private Sub Command2_Click()
Unload Form6
Form2.Show
End Sub

Private Sub Command5_Click()
Call Shell("Start.exe " & "Http://www.MP3.com", 0)
End Sub

Private Sub Image1_Click()
MsgBox "Sorry, No Data Detected!!"
End Sub
