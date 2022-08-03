VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "Birth Certificate"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18045
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   18045
   WindowState     =   2  'Maximized
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
      Height          =   540
      Left            =   13080
      TabIndex        =   38
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "Mother's Aadhar No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7680
      MaxLength       =   12
      TabIndex        =   37
      Top             =   8520
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "Father's Aadhar No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7680
      MaxLength       =   12
      TabIndex        =   36
      Top             =   7680
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "Date of resignation"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7200
      TabIndex        =   33
      Top             =   6840
      Width           =   3615
      _ExtentX        =   6376
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
      Format          =   126812160
      CurrentDate     =   43868
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "DOB"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7920
      TabIndex        =   32
      Top             =   3600
      Width           =   3255
      _ExtentX        =   5741
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
      Format          =   126812160
      CurrentDate     =   43868
      MaxDate         =   43951
      MinDate         =   367
   End
   Begin VB.TextBox Text4 
      DataField       =   "District"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4680
      TabIndex        =   31
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox Text14 
      DataField       =   "place of birth"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1680
      TabIndex        =   30
      Top             =   7560
      Width           =   4095
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
      Left            =   15240
      TabIndex        =   28
      Top             =   4560
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
      Left            =   15240
      TabIndex        =   27
      Top             =   5400
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
      Left            =   15240
      TabIndex        =   26
      Top             =   2880
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
      Left            =   13080
      TabIndex        =   25
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdpre 
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
      Left            =   15240
      TabIndex        =   24
      Top             =   3720
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
      Left            =   13080
      TabIndex        =   23
      Top             =   4560
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
      Left            =   13080
      TabIndex        =   22
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      DataField       =   "Resignation no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   20
      Top             =   6840
      Width           =   3615
   End
   Begin VB.TextBox Text11 
      DataField       =   "State"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   18
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox Text9 
      DataField       =   "place"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   6120
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      DataField       =   "Mother name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   4920
      Width           =   9855
   End
   Begin VB.TextBox Text7 
      DataField       =   "Father name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   4200
      Width           =   9855
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
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3600
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
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   9855
   End
   Begin VB.TextBox Text2 
      DataField       =   "Srno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13800
      Top             =   7200
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\birth1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\birth1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   9840
      Picture         =   "BIRTH.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label5 
      Caption         =   "Mother;s Aadhar No"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   35
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Father's Aadhar No"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   34
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Place of Birth(Name and Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Date of Registration"
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
      Left            =   5640
      TabIndex        =   21
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Registration No"
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
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label14 
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
      Height          =   255
      Left            =   8880
      TabIndex        =   17
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label13 
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
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Place"
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
      Left            =   2400
      TabIndex        =   15
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label11 
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
      Left            =   240
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Mother's Name"
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
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Father's Name"
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
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Date of Birth"
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
      Left            =   6000
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "S No,:"
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
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Birth Certificate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
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
MsgBox "Record Saved Successfully!!"
End Sub

Private Sub cmddelete_Click()
Adodc1.Recordset.Delete
MsgBox "Record Deleted Successfully!!"
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
a = Text2.Text
Adodc1.Recordset.AddNew
Text2.Text = a + 1
Text3.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text4.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text14.Enabled = True
Text1.Enabled = True
Text5.Enabled = True
End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdpre_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Form_Load()
'Adodc1.Visible = False
Text3.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text4.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text14.Enabled = False
Text1.Enabled = False
Text5.Enabled = False
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Public Function birth()
Text3.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text4.Text = ""
Text11.Text = ""
Text12.Text = ""
Text14.Text = ""
Text1.Text = ""
Text5.Text = ""
End Function

