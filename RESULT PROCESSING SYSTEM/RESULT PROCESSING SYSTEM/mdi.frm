VERSION 5.00
Begin VB.Form Adminmdi 
   Caption         =   "Administrator"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "mdi.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   15240
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu sd 
         Caption         =   "Select Database"
      End
      Begin VB.Menu up 
         Caption         =   "Update"
      End
      Begin VB.Menu cp 
         Caption         =   "Change password"
      End
      Begin VB.Menu ex 
         Caption         =   "Logout"
      End
   End
End
Attribute VB_Name = "Adminmdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cp_Click()
    
    Unload Me
    adminpassword.Show
    
End Sub

Private Sub ex_Click()
    
    Unload Me
    AdminEndLogin.Show

End Sub

Private Sub Form_Load()

    WindowState = 2

End Sub

Private Sub sd_Click()

    Unload Me
    Admindb.Show
    
End Sub

Private Sub up_Click()
    Unload Me
    Adminupdate.Show
End Sub
