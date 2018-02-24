VERSION 5.00
Begin VB.Form Admin 
   BackColor       =   &H00FF8080&
   Caption         =   "Admin"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Admin.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton back 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Admin_load()

    WindowState = 2
    Text1.Text = ""
    Text.Text = ""
    
End Sub

Private Sub Back_Click()

    Unload Me
    AdminEndLogin.Show

End Sub

Private Sub CLEAR_Click()
    
    Text1.Text = ""
    Text2.Text = ""
    
End Sub

Private Sub OK_Click()

    Dim rs
    
    
    Call pathsub
    Set rs = PathOb.Execute("SELECT * FROM path1 WHERE c1= '" & 2 & "' ")
    
    If (Text1.Text = rs!c2) And (Text2.Text = rs!c3) Then
        Adminmdi.Show
        Unload Me
    Else
        MsgBox ("Invalid Username or Password")
    End If
    
End Sub

