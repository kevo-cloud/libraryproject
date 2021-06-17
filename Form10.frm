VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDL 
      Left            =   5280
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12000
      Top             =   9840
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\2profile.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\KEVOKE\Desktop\LIBRARY PROJECT\2profile.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 2profile"
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
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   29
      Top             =   9120
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
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
      Left            =   8280
      TabIndex        =   28
      Top             =   10560
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
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
      Left            =   8280
      TabIndex        =   27
      Top             =   9840
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8280
      TabIndex        =   26
      Top             =   9120
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form10.frx":0000
      Height          =   1935
      Left            =   240
      TabIndex        =   25
      Top             =   9000
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
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
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   12960
      TabIndex        =   24
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "status book 3"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11520
      TabIndex        =   23
      Top             =   8160
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "status book 2"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11520
      TabIndex        =   22
      Top             =   7440
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "status book 1"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11520
      TabIndex        =   21
      Top             =   6720
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      DataField       =   "book 3"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   8160
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      DataField       =   "book 2"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   7440
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      DataField       =   "book 1"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   15360
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEARCH"
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
      Left            =   12600
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "admission no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      DataField       =   "student class"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "student name"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   8760
      TabIndex        =   5
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPLOAD PHOTO"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   360
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 3"
      DataSource      =   "Adodc1"
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
      Left            =   7200
      TabIndex        =   20
      Top             =   8280
      Width           =   3975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 2"
      DataSource      =   "Adodc1"
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
      Left            =   7200
      TabIndex        =   19
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "STATUS BOOK 1"
      DataSource      =   "Adodc1"
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
      Left            =   7200
      TabIndex        =   18
      Top             =   6720
      Width           =   3615
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
      Left            =   7920
      TabIndex        =   17
      Top             =   5760
      Width           =   4815
   End
   Begin VB.Label Label8 
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
      Left            =   240
      TabIndex        =   13
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label Label7 
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
      Left            =   240
      TabIndex        =   12
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   11
      Top             =   6600
      Width           =   2655
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
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Width           =   4215
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
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "STUDENT CLASS"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "STUDENT NAME"
      DataSource      =   "Adodc1"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FORM 2 PROFILE PAGE"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "Form10"
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
Adodc1.RecordSource = "select * from 2profile"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
MsgBox "successfully added"


End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Delete
MsgBox "successfully deleted"

End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Update
MsgBox "successfully updated"

End Sub

Private Sub Command7_Click()
Unload Me
Form6.Show

End Sub

Private Sub Form_Load()
Me.Width = 17000
Me.Height = 13000
Me.Show
End Sub
