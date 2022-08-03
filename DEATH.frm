VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Death Certificate"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17925
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   17925
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataField       =   "Aadhar No"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8400
      TabIndex        =   34
      Top             =   7200
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "DATE of death"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1920
      TabIndex        =   32
      Top             =   3720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
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
      Format          =   126222336
      CurrentDate     =   43863
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Date"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2040
      TabIndex        =   31
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   126222336
      CurrentDate     =   43863
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   12600
      Top             =   7680
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Death1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Death1.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text12 
      DataField       =   "Sr No"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2040
      TabIndex        =   30
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   28
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   27
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   25
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   24
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   22
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   21
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      DataField       =   "Certified Docter Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   20
      Top             =   8040
      Width           =   4695
   End
   Begin VB.TextBox Text10 
      DataField       =   "Name of funeral"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   7200
      Width           =   4575
   End
   Begin VB.TextBox Text9 
      DataField       =   "State"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8400
      TabIndex        =   16
      Top             =   6360
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      DataField       =   "District"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      DataField       =   "Place"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   6360
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Age"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2640
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   9600
      Picture         =   "DEATH.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Aadhar No of Deceased"
      Height          =   495
      Left            =   6840
      TabIndex        =   33
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Sr No"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Certifier Docter Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Name of Funeral Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "City or Town"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Place Where Death Occurred If unknown, give last place of residence"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date of Death"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Full Name of Deceased"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Death Certificate"
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
      Left            =   9240
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
If Option1.Value = True Then
Adodc1.Recordset.Fields("gender").Value = "Male"
Else
Adodc1.Recordset.Fields("gender").Value = "Female"
End If
Adodc1.Recordset.Update
MsgBox "Record Inserted Successfully!!"
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
a = Text12.Text
Adodc1.Recordset.AddNew
Text12.Text = a + 1
Text1.Enabled = True
Text2.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdprevious_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
DTPicker2.Value = Date
Text1.Enabled = False
Text2.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
End Sub
