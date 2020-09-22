VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vote For This Program"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Vote"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Poor"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Below Average"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Average"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Good"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Excellent"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Vt, Nr
If Option1.Value = True Then Vt = 5
If Option2.Value = True Then Vt = 4
If Option3.Value = True Then Vt = 3
If Option4.Value = True Then Vt = 2
If Option5.Value = True Then Vt = 1
Nr = 9929
Shell "start http://www.planetsourcecode.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&optCodeRatingValue=" & Vt & "&txtCodeId=" & Nr & "&txtCodeName=Hover%20Buttons&intUserRatingTotal=0&intNumOfUserRatings=0", vbMinimizedNoFocus
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
