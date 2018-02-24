VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Admindb 
   BackColor       =   &H00FF8080&
   Caption         =   "Admindb"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   Begin MSComDlg.CommonDialog Browse 
      Left            =   2040
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   13440
      TabIndex        =   4
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse...."
      Height          =   495
      Left            =   17160
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   3960
      Width           =   12735
   End
End
Attribute VB_Name = "Admindb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

        With Browse
            .Filter = "(*.mdb;)| *.mdb;|(*.All files)|*.*"
            .ShowOpen
            Text1.Text = .FileName
            If .FileName = "" Then Exit Sub
        End With

End Sub

Private Sub Command2_Click()
    var.path = Text1.Text
    MsgBox ("Database connected")
    
End Sub

Private Sub Text1_GotFocus()

    Text1.Locked = True
    
End Sub
