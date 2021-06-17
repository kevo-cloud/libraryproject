VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form9"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11760
      Top             =   7920
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\1profile.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\1profile.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 1profile"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   29
      Top             =   10200
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   28
      Top             =   9360
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   27
      Top             =   8520
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PRINT REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   26
      Top             =   7800
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":0000
      Height          =   2535
      Left            =   240
      TabIndex        =   25
      Top             =   8400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "status book 3"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11880
      TabIndex        =   24
      Top             =   6840
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "status book 2"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11880
      TabIndex        =   23
      Top             =   6000
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "status book 1"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11760
      TabIndex        =   22
      Top             =   5400
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CDL 
      Left            =   6120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      DataField       =   "book 3"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   7560
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "book 2"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      DataField       =   "book 1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   14040
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "admission no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "student class"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "student name"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD PHOTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label12 
      Caption         =   "STATUS BOOK 3"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   6840
      Width           =   3495
   End
   Begin VB.Label Label11 
      Caption         =   "STATUS BOOK 2"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7800
      TabIndex        =   20
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   "STATUS BOOK 1"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "RETURNED BOOKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   18
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "BOOK 3"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "BOOK 2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "BOOK 1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "BOOKS BORROWED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ADMISSION NO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "STUDENT CLASS"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "SUDENT NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STUDENT PROFILE PAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CDL.Filter = "picture file / *.jpg"
CDL.ShowOpen
If CDL.FileName <> "" Then
Image1.Picture = LoadPicture(CDL.FileName)
End If

End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from 1profile"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Update
MsgBox "successfully updated"

End Sub

Private Sub Command7_Click()
Unload Me
Form5.Show

End Sub

Private Sub Form_Load()
Me.Width = 17000
Me.Height = 13000
Me.Show
End Sub
