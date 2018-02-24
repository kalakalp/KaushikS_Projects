VERSION 5.00
Begin VB.Form adminpassword 
   Caption         =   "password"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "adminpassword.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      MaskColor       =   &H0080FFFF&
      TabIndex        =   8
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   8280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-enter password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the new password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the current password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   3735
   End
End
Attribute VB_Name = "adminpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs
Private Sub Command1_Click()

    If Text2.Text = Text3.Text Then
        Set rs = PathOb.Execute("update path1 set c3='" & Text2.Text & "' where c1= '" & 2 & "' ")
        MsgBox ("Password updated")
        Unload Me
        AdminEndLogin.Show
    Else
        MsgBox ("Passwords doesnot match")
    End If
    
End Sub

Private Sub Command2_Click()

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
End Sub

Private Sub Command3_Click()
     Unload Me
    Adminmdi.Show
End Sub

Private Sub Form_Load()

    WindowState = 2
    Call pathsub
    
End Sub

Private Sub Text2_GOTFOCUS()

    Set rs = PathOb.Execute("SELECT * FROM path1 WHERE c1= '" & 2 & "' ")
    
    If Text1.Text = rs!c3 Then
        Text2.Enabled = True
        Text3.Enabled = True
    ElseIf flag = False Then
        MsgBox ("Invalid Password")
        Text2.Enabled = False
        Text3.Enabled = False
    End If
     
End Sub
