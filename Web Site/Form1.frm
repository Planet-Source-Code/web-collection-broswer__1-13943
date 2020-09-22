VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Web Site Searching Ver 1.01"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Start Loading"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Internet FreeWare"
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "You Can Now Enjoy This Program With Lots Of Fun Sites!!"
      BeginProperty Font 
         Name            =   "MS Song"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   4665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "This Program Has Popular Websites"
Unload Form1
Form2.Show
End Sub

Private Sub Command2_Click()
ProgressBar1.Visible = True
Label1.Visible = True
End Sub

Private Sub ProgressBar1_click()
Label1.Visible = False
ProgressBar1.Visible = False
Command2.Visible = False
Label2.Visible = True
Command1.Enabled = True
End Sub

Private Sub Timer1_Timer()

If ProgressBar1.Value >= 200 Then
Let ProgressBar1.Value = 200

End If

If ProgressBar1.Value < 200 Then
Let ProgressBar1.Value = ProgressBar1.Value + 1

ElseIf ProgressBar1.Value > 200 Then
Let ProgressBar1.Value = 200
ElseIf ProgressBar1.Value = 200 Then
Command1.Enabled = True
End If

If ProgressBar1.Value = 1 Then
Label1.Caption = "Loading Web Data"
End If

If ProgressBar1.Value = 90 Then
Label1.Caption = "Loading Dll"
End If

If ProgressBar1.Value = 95 Then
Label1.Caption = "Checking Data"
End If

If ProgressBar1.Value = 100 Then
Label1.Caption = "Loading Visual Basic Codes"
End If

If ProgressBar1.Value = 140 Then
Label1.Caption = "Loading Image"
End If

If ProgressBar1.Value = 180 Then
Label1.Caption = "Optimzed Your PC For Better Performer"
End If

If ProgressBar1.Value = 200 Then
Label1.Visible = False
ProgressBar1.Visible = False
Command2.Visible = False
Label2.Visible = True
End If
End Sub

