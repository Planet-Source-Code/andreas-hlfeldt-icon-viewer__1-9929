VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon View"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Icon View - About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "HERE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1140
         TabIndex        =   7
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":0000
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4048
      Arrange         =   2
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   1650
      Left            =   360
      Pattern         =   "*.ico"
      TabIndex        =   3
      Top             =   360
      Width           =   2055
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer

Private Sub Command1_Click()
End
End Sub

Private Sub Dir1_Change()
On Error GoTo 99
File1 = Dir1
Dim A, P, T
If Len(Dir1) = 3 Then P = ""
If Len(Dir1) <> 3 Then P = "\"
ListView1.ListItems.Clear
T = 1
For A = 0 To File1.ListCount - 1
ImageList1.ListImages.Add I, , LoadPicture(Dir1 & P & File1.List(A))
ListView1.ListItems.Add T, , File1.List(A), I, I
I = I + 1
T = T + 1
Next A
99
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
I = 1
File1 = Dir1
Dim A, P
If Len(Dir1) = 3 Then P = ""
If Len(Dir1) <> 3 Then P = "\"
I = 1
ListView1.ListItems.Clear
ImageList1.ListImages.Clear
For A = 0 To File1.ListCount - 1
ImageList1.ListImages.Add I, , LoadPicture(Dir1 & P & File1.List(A))
ListView1.ListItems.Add I, , File1.List(A), I, I
I = I + 1
Next A
End Sub

Private Sub Label2_Click()
Form2.Show
End Sub
