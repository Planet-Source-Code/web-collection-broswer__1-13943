VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Search Engine Data"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "<<-- EXIT -->>"
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
      Left            =   4200
      TabIndex        =   11
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Cheat And Tip On This Program !!!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Music -->"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose One Image (CLICK)"
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.Image Image3 
         Height          =   930
         Left            =   240
         Picture         =   "Form2.frx":0000
         Top             =   1440
         Width           =   3210
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   480
         Picture         =   "Form2.frx":0E8F
         Top             =   2880
         Width           =   2115
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   120
         Picture         =   "Form2.frx":12DA
         Top             =   480
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose One"
      Height          =   2655
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command4 
         Caption         =   "Yahoo Mail"
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
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hotbot"
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
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Metacrawler"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Yahoo"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cheat Code  -->"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "KOF 2000 Site-->"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Visual Basic -->"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Shell("Start.exe " & "Http://www.yahoo.com", 0)
End Sub

Private Sub Command10_Click()
MsgBox "Thanks For Using My Program!! BYE BYE "
End
End Sub

Private Sub Command2_Click()
Call Shell("Start.exe " & "http://www.metacrawler.com", 0)
End Sub

Private Sub Command3_Click()
Call Shell("Start.exe " & "http://www.hotbot.com", 0)
End Sub

Private Sub Command4_Click()
Call Shell("Start.exe " & "http://mail.yahoo.com/", 0)
End Sub

Private Sub Command5_Click()
Form3.Show
Unload Form2
End Sub

Private Sub Command6_Click()
Form4.Show
Unload Form2
End Sub

Private Sub Command7_Click()
Form5.Show
Unload Form2
End Sub

Private Sub Command8_Click()
Form6.Show
Unload Form2
End Sub

Private Sub Command9_Click()
MsgBox "To Have An Fast StartUp, Click On The Progress Bar" & Chr(13) & " After Clicking Start Loading!!! Try It And See It!!!"
MsgBox "Thank For Downloading My Program, Please Rate It!!"
MsgBox "All Image & Picture Here Can Be Clicked!!"
End Sub

Private Sub Image1_Click()
Call Shell("Start.exe " & "Http://www.yahoo.com", 0)
End Sub

Private Sub Image2_Click()
Call Shell("Start.exe " & "http://mail.yahoo.com/", 0)
End Sub

Private Sub Image3_Click()
Call Shell("Start.exe " & "http://www.metacrawler.com", 0)
End Sub
