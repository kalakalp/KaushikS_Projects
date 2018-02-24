VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form StudentDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "Student Result Processing............."
   ClientHeight    =   5640
   ClientLeft      =   5340
   ClientTop       =   4035
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "StudentDetails.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   6735
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   10
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14040
      Top             =   5400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   7800
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "StudentDetails.frx":29FAF
      Left            =   6120
      List            =   "StudentDetails.frx":29FB1
      TabIndex        =   7
      Text            =   "Select Semester"
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show Result"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register Number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   2325
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   5400
      Width           =   1320
   End
End
Attribute VB_Name = "StudentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()

   Unload Me
   MDIForm1.Show

'Label5.Caption = "lc=" & Combo1.ListCount
'Label5.Caption = Label5.Caption & " li=" & Combo1.ListIndex
'Label5.Caption = Label5.Caption & " id=" & Combo1.ItemData(ListIndex)

End Sub

Public Sub Command1_Click()

If Text1.Text = "" Or Combo1.Text = "" Then
    Text2.Text = ""
    Text3.Text = ""
   GoTo a
Else
    Timer1.Enabled = True
    ProgressBar1.Visible = True
    Text1.Locked = True
    Combo1.Locked = True
    Command1.Caption = "Please Wait...."
    var.reg = Text1.Text
    var.sem = Combo1.Text
    var.dept = Text3.Text
    Module1.lectdep = Text3.Text
    Module1.lectsem = Combo1.Text
End If
 
    If 1 < 0 Then
a:                 MsgBox ("Sorry! Enter Register Number Or Select Semester..")
    End If
    
End Sub


Private Sub Form_Load()
    WindowState = 2
End Sub

Private Sub Text1_LOSTFOCUS()
    
    Combo1.clear
    Call openconnection
    Set rs = con.Execute("select * from RESULTS where REGNO='" & Text1.Text & "' ")
    
    On Error GoTo a

    If Not (Text1.Text = "") Or Not (Text2.Text = "") Or Not (Text3.Text = "") Then
        Text2.Text = rs!Name
        Text3.Text = rs!COURSECODE
        
        While Not rs.EOF
            Combo1.AddItem (rs!sem)
            rs.MoveNext
        Wend
    Else
    End If
         
    If 1 < 0 Then
a:                 MsgBox ("Sorry! Enter Register Number Or Select Semester..")
    End If
    
End Sub
Private Sub Timer1_Timer()
    
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If ProgressBar1.Value = 100 Then
        If Combo1.ListIndex + 1 = Combo1.ListCount Then
            Unload Me
            StudentResult.Show
        Else
            Unload Me
            Resultbk.Show
        End If
    End If
End Sub
