VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H0080FFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   9360
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu UG 
      Caption         =   "User Generator"
   End
   Begin VB.Menu Frm 
      Caption         =   "Form"
      Begin VB.Menu BC 
         Caption         =   "Birth Certificare"
      End
      Begin VB.Menu DC 
         Caption         =   "Death Certificate"
      End
      Begin VB.Menu MC 
         Caption         =   "Marriage Certificate"
      End
      Begin VB.Menu RC 
         Caption         =   "Residence Certificate"
      End
      Begin VB.Menu CC 
         Caption         =   "Character Certificate"
      End
   End
   Begin VB.Menu RG 
      Caption         =   "Report Generator"
      Begin VB.Menu BCa 
         Caption         =   "Birth Certificate"
      End
      Begin VB.Menu DCa 
         Caption         =   "Death Certificate"
      End
      Begin VB.Menu MCa 
         Caption         =   "Marriage Certificate"
      End
      Begin VB.Menu RCa 
         Caption         =   "Resideance Certificate"
      End
      Begin VB.Menu CCa 
         Caption         =   "Character Certificate"
      End
   End
   Begin VB.Menu AU 
      Caption         =   "About Us"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

End Sub

Private Sub HOME_Click()

End Sub

Private Sub AU_Click()
Form10.Show
End Sub

Private Sub BC_Click()
Form2.Show
End Sub

Private Sub BCa_Click()
Form8.Show
End Sub

Private Sub CC_Click()
Form6.Show
End Sub

Private Sub CCa_Click()
Form13.Show
End Sub

Private Sub DC_Click()
Form3.Show
End Sub

Private Sub DCa_Click()
Form9.Show
End Sub

Private Sub MC_Click()
Form4.Show
End Sub

Private Sub MCa_Click()
Form11.Show
End Sub

Private Sub RC_Click()
Form5.Show
End Sub

Private Sub RCa_Click()
Form12.Show
End Sub

Private Sub UG_Click()
Form7.Show
End Sub
