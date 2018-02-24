VERSION 5.00
Begin VB.Form AdminEndLogin 
   BackColor       =   &H00FF8080&
   Caption         =   "AdminEndLogin"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "AdminEndLogin1.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.OptionButton Lecturer 
      BackColor       =   &H0000C000&
      Caption         =   "Lecturer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   5640
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton LAbort 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10440
      MaskColor       =   &H00800000&
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton LOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   2
      Top             =   7080
      Width           =   1575
   End
   Begin VB.OptionButton EndUserbtn 
      BackColor       =   &H0000C000&
      Caption         =   "Student"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   5640
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   1
      Top             =   4080
      Width           =   3255
   End
   Begin VB.OptionButton Administrator 
      BackColor       =   &H0000C000&
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   5640
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   0
      Top             =   2640
      Width           =   3255
   End
End
Attribute VB_Name = "AdminEndLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   EndUserbtn.Value = True
    WindowState = 2
    Call pathsub
    Set rs = PathOb.Execute("select * from path1")
    var.path = rs!c2
    
End Sub

Private Sub LAbort_Click()

    End
    
End Sub


Private Sub LOK_Click()

    If Administrator.Value = True Then
        Unload Me
        Admin.Show
        
    ElseIf Lecturer.Value = True Then
        Unload Me
        LectHome.Show
    Else
        Unload Me
        MDIForm1.Show
    End If
    
End Sub


