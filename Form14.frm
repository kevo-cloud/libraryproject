VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form14"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      DataField       =   "status book 3"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   12000
      TabIndex        =   31
      Top             =   7320
      Width           =   3495
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "status book 2"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11880
      TabIndex        =   30
      Top             =   6360
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "status book 1"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11760
      TabIndex        =   29
      Top             =   5760
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12240
      Top             =   8160
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\5profile.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\5profile.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 5profile"
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
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   28
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   15600
      TabIndex        =   27
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   26
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   25
      Top             =   10320
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   24
      Top             =   9720
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   9000
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   8160
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form14.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   8160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
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
   Begin MSComDlg.CommonDialog CDL 
      Left            =   6240
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      DataField       =   "book 3"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   7320
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      DataField       =   "book 2"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   6600
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "book 1"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "age"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "admission no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      DataField       =   "student form"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "student name"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD PHOTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 3"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "RETURNED  BOOKS"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   4680
      Width           =   5535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "BOOK 3"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   3855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "BOOK 2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   3735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "BOOK 1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label Label6 
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
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "AGE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3840
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ADMISSION NO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "STUDENT FORM"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "STUDENT NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
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
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "Form14"
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

Private Sub Command3_Click()
Adodc1.Recordset.Delete
MsgBox "successfully deleted"


End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Update
MsgBox "successfully updated"
End Sub

Private Sub Command5_Click()
Unload Me
Form13.Show
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select * from 5profile"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.AddNew
MsgBox "successfully added"

End Sub

Private Sub Form_Load()
Me.Width = 17000
Me.Height = 13000
Me.Show
End Sub
