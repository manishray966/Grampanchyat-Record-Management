VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "User Generator"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7440
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   6000
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Logindb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\Gramapanchyat\Datebase\Logindb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "logindb"
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
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "New User"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Conform Password"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Lable1 
      Caption         =   "Username"
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Adodc1.Recordset.MoveLast
Text1.Enabled = True
Text2.Enabled = True
Text3.Text = True
Adodc1.Recordset.AddNew
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.MoveLast
If Text2.Text = Text3.Text Then
Adodc1.Recordset.Update
MsgBox "User created successfully"
Else
MsgBox "Reenter password"
Text3.SetFocus
End If
End Sub

Private Sub Form_Load()
Adodc1.Visible = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.Enabled = False
Text2.Enabled = False
Text3.Text = False
End Sub
