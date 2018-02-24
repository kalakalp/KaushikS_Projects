VERSION 5.00
Begin VB.Form STUDENT_RESULT_PROCESSING_SYSTEM 
   BackColor       =   &H00000040&
   Caption         =   "STUDENT_RESULT_PROCESSING_SYSTEM"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "STUDENT_RESULT_PROCESSING_SYSTEM.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton Login_cmd 
      BackColor       =   &H00800000&
      Cancel          =   -1  'True
      Caption         =   " &LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      MaskColor       =   &H00000040&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton Abort_cmd 
      BackColor       =   &H00800000&
      Caption         =   "&ABORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label welcome 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME YOU WITH PRIDE TO ""RESULT PROCESSING SYSTEM"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4440
      Width           =   11655
   End
End
Attribute VB_Name = "STUDENT_RESULT_PROCESSING_SYSTEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, rs As Object
Private Sub Abort_cmd_Click()
    End
End Sub


Private Sub Form_Load()
    
    WindowState = 2
    
End Sub

Private Sub Login_cmd_Click()
    photo.Show
    Unload Me
End Sub

