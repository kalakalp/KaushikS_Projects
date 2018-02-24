VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16395
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu File1 
      Caption         =   "File"
      Begin VB.Menu sd 
         Caption         =   "Select DataBase"
      End
      Begin VB.Menu up 
         Caption         =   "Update"
      End
      Begin VB.Menu cp 
         Caption         =   "Change Password"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sd_Click()
    Admindb.Show
End Sub
