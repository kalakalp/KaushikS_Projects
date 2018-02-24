VERSION 5.00
Begin VB.Form AdminEndLogin 
   BackColor       =   &H00FF8080&
   Caption         =   "AdminEndLogin"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   Begin VB.OptionButton Lecturer 
      BackColor       =   &H00FF8080&
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
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton LAbort 
      Caption         =   "&Abort"
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
      Left            =   12360
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
      Left            =   5640
      TabIndex        =   2
      Top             =   7080
      Width           =   1575
   End
   Begin VB.OptionButton EndUserbtn 
      BackColor       =   &H00FF8080&
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
      Height          =   540
      Left            =   7560
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin VB.OptionButton Administrator 
      BackColor       =   &H00FF8080&
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
      Height          =   540
      Left            =   7560
      TabIndex        =   0
      Top             =   2640
      Width           =   3615
   End
End
Attribute VB_Name = "AdminEndLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

   EndUserbtn.Value = True
    
End Sub

Private Sub LAbort_Click()

    End
    
End Sub


Private Sub LOK_Click()

    If Administrator.Value = True Then
        Admin.Show
        Unload Me
    ElseIf Lecturer.Value = True Then
        LectHome.Show
        Unload Me
    Else
        MDIForm1.Show
        Unload Me
    End If
    
End Sub


