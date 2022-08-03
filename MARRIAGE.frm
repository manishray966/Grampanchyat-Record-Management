VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Marriage Certificate"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18000
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   18000
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      DataField       =   "Aadhar No w"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7440
      TabIndex        =   39
      Top             =   6240
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Aadhar no H"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   38
      Top             =   6240
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      DataField       =   "DOB W"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7440
      TabIndex        =   36
      Top             =   3600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   126091264
      CurrentDate     =   43863
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "DOB of H"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   35
      Top             =   3600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   126091264
      CurrentDate     =   43863
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "Date of marriage"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1560
      TabIndex        =   34
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   126091264
      CurrentDate     =   43863
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   13080
      Top             =   8160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Marriage1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Marriage1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      DataField       =   "Sr No"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1560
      TabIndex        =   33
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      DataField       =   "Address2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   31
      Top             =   9360
      Width           =   4455
   End
   Begin VB.TextBox Text18 
      DataField       =   "Address1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   30
      Top             =   9360
      Width           =   4455
   End
   Begin VB.TextBox Text17 
      DataField       =   "Witness Name2"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   28
      Top             =   8520
      Width           =   4455
   End
   Begin VB.TextBox Text16 
      DataField       =   "Witness Name1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   27
      Top             =   8520
      Width           =   4455
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
      TabIndex        =   24
      Top             =   6000
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
      TabIndex        =   23
      Top             =   6840
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
      TabIndex        =   22
      Top             =   4320
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
      Left            =   14400
      TabIndex        =   21
      Top             =   5160
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
      Left            =   14400
      TabIndex        =   20
      Top             =   6840
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
      TabIndex        =   19
      Top             =   5160
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
      Left            =   14400
      TabIndex        =   18
      Top             =   6000
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
      Left            =   14400
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      DataField       =   "Father Name W"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   7080
      Width           =   4455
   End
   Begin VB.TextBox Text13 
      DataField       =   "Father Name H"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   7080
      Width           =   4455
   End
   Begin VB.TextBox Text12 
      DataField       =   "Residaece W"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   5280
      Width           =   4455
   End
   Begin VB.TextBox Text11 
      DataField       =   "Residance H"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   5280
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      DataField       =   "Birth Place W"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      DataField       =   "Birth Place H"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   4440
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      DataField       =   "Wife Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      DataField       =   "Husband Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   7560
      Picture         =   "MARRIAGE.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Aadhar No"
      Height          =   495
      Left            =   240
      TabIndex        =   37
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Sr no"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Name of the Witness"
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
      TabIndex        =   26
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Witnesses to the Marriage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Date of Marriage"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Name of the Father or Guardan"
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
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Residance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Birth Place"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Wife"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Husband"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Marriage Certificate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
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
a = Text7.Text
Adodc1.Recordset.AddNew
Text7.Text = a + 1
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
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
DTPicker3.Value = Date
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
End Sub
