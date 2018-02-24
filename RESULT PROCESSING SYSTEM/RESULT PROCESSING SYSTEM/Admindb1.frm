VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Admindb 
   BackColor       =   &H00FFFFFF&
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
   Picture         =   "Admindb1.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin MSComDlg.CommonDialog Browse 
      Left            =   2040
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Back"
      Height          =   495
      Left            =   8760
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   4
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   3
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Submit"
      Height          =   495
      Left            =   1440
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   2
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse...."
      Height          =   495
      Left            =   13440
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   12735
   End
End
Attribute VB_Name = "Admindb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Object
Private Sub Command1_Click()

        With Browse
            .Filter = "(*.mdb;)| *.mdb;|(*.All files)|*.*"
            .ShowOpen
            Text1.Text = .FileName
            If .FileName = "" Then Exit Sub
        End With

End Sub

Private Sub Command2_Click()
    Call pathsub
    Set rs = PathOb.Execute("update path1 set c2='" & Text1.Text & "' WHERE c1= '" & 1 & "' ")
    MsgBox ("Database connected")
    Text1.Text = ""
     Unload Me
    Adminmdi.Show
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
End Sub

Private Sub Command4_Click()
    
    Unload Me
    Adminmdi.Show

End Sub

Private Sub Text1_GotFocus()

    Text1.Locked = True
    
End Sub

Private Sub Form_Load()
    WindowState = 2
End Sub
