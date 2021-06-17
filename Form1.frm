VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
      Min             =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1800
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   11160
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "SCHOOL LIBRARY SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
Me.Width = 17000
Me.Height = 13000
Me.Show


End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label2.Caption = "loading...."
Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
Form2.Show

End If
End Sub
