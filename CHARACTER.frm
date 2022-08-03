VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   Caption         =   "Character Certificate"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17955
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   17955
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Date"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1320
      TabIndex        =   23
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   35454976
      CurrentDate     =   43863
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11880
      Top             =   7560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Character1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Character1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
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
   Begin VB.TextBox Text7 
      DataField       =   "Nationality"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   13
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Sr no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Year"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   4440
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Father Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3600
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   8160
      Picture         =   "CHARACTER.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "nationality. He/She is not related to me."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6120
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "year. He/She bears a good moral character and is of"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Label Label7 
      Caption         =   "Sr. No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "from the last"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Resident of"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Son/daughter of Shri"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Certified that I khow Mr./Ms./"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Character Certificate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"
End Sub

Private Sub cmddelete_Click()
Adodc1.Recordset.Delete
MsgBox "Record Deleted Successfully"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdnew_Click()
Adodc1.Recordset.MoveLast
a = Text6.Text
Adodc1.Recordset.AddNew
Text6.Text = a + 1
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text7.Enabled = True
End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdprevious_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text7.Enabled = False
End Sub
