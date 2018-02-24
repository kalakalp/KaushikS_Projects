VERSION 5.00
Begin VB.Form Adminupdate 
   Caption         =   "Update"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Adminupdate.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Submit"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00004000&
      Height          =   480
      ItemData        =   "Adminupdate.frx":2681C
      Left            =   4440
      List            =   "Adminupdate.frx":26826
      TabIndex        =   1
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the option"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   3135
   End
End
Attribute VB_Name = "Adminupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Combo1.Text = "Lecturer Details" Then
        Unload Me
        LectUpdate.Show
    Else
        Unload Me
        SubUpdate.Show
    End If
    
End Sub

Private Sub Command2_Click()

    Unload Me
    Adminmdi.Show
    
End Sub

Private Sub Form_Load()

    WindowState = 2
    
End Sub
