VERSION 5.00
Begin VB.MDIForm MD 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Result Processing System"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "MD.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu studOper 
      Caption         =   "Student Results"
      Begin VB.Menu studOperDet 
         Caption         =   "View Result"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
    
    Unload Me
    AdminEndLogin.Show
    
End Sub

Private Sub studOperDet_Click()

    StudentDetails.Show

End Sub
