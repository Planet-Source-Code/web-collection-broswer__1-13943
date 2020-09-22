VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Visual Basic Sites"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3960
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
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
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Free Visual Basic Code"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "vBulletin Forum"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ezboard Forum"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Free Forum"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Programmer Heaven"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Planet - Source - Code (Newest Code Version)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   1200
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   2160
      Picture         =   "Form3.frx":1D55
      Top             =   3240
      Width           =   2520
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Shell("Start.exe " & "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1", 0)
End Sub

Private Sub Command2_Click()
Call Shell("Start.exe " & "http://www.programmersheaven.com/", 0)
End Sub

Private Sub Command3_Click()
Call Shell("Start.exe " & "http://www.coolboard.com/boardshow.cfm/mb=213298131085905", 0)
End Sub

Private Sub Command4_Click()
Call Shell("Start.exe " & "http://pub13.ezboard.com/bvisualbasicexplorer", 0)
End Sub

Private Sub Command5_Click()
Call Shell("Start.exe " & "http://forums.vb-world.net/", 0)
End Sub

Private Sub Command6_Click()
Call Shell("Start.exe " & "http://www.freevbcode.com/forum_home.asp", 0)
End Sub

Private Sub Command7_Click()
Form2.Show
Unload Form3
End Sub

Private Sub Image1_Click()
Call Shell("Start.exe " & "http://forums.vb-world.net/", 0)
End Sub

Private Sub Image3_Click()
Call Shell("Start.exe " & "http://www.planet-source-code.com/", 0)
End Sub
