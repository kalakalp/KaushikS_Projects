VERSION 5.00
Begin VB.Form Lectdetails 
   Caption         =   "Lectdetails"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00C0FFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Lectdetails.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton Command10 
      Caption         =   "&CANCEL"
      Height          =   975
      Left            =   6960
      TabIndex        =   23
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton Save 
      Caption         =   "&SAVE"
      Height          =   975
      Left            =   840
      TabIndex        =   22
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   4200
      Width           =   3210
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   585
      Left            =   1920
      TabIndex        =   11
      Top             =   3720
      Width           =   3210
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   585
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
      Width           =   3210
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   585
      Left            =   1920
      TabIndex        =   9
      Top             =   2760
      Width           =   3210
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   585
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   3210
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      ForeColor       =   &H00004000&
      Height          =   585
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   3210
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   3210
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   3210
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   1920
      TabIndex        =   4
      Text            =   " "
      Top             =   360
      Width           =   3210
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1920
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      X1              =   1680
      X2              =   1680
      Y1              =   2040
      Y2              =   4560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subjects handling:"
      ForeColor       =   &H00004000&
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dept:"
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desig:"
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Lectdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Object

Private Sub Command1_Click()
    Text1.Enabled = True
End Sub

Private Sub Command10_Click()
    MsgBox ("ROLLBACKED!!")
    Unload Me
    LectUpdate.Show
End Sub

Private Sub Command2_Click()
    Text2.Enabled = True
End Sub

Private Sub Command3_Click()
    Text3.Enabled = True
End Sub

Private Sub Command4_Click()
    Text4.Enabled = True
End Sub

Private Sub Command5_Click()
    Text5.Enabled = True
End Sub

Private Sub Command6_Click()
    Text6.Enabled = True
End Sub

Private Sub Command7_Click()
    Text7.Enabled = True
End Sub

Private Sub Command8_Click()
    Text8.Enabled = True
End Sub

Private Sub Command9_Click()
    Text9.Enabled = True
End Sub

Private Sub Form_Load()

    WindowState = 2
    
     Call sopenconnection

   If Module1.lectdep = "CS" Then
        Set rs = scon.Execute("select * from lcs where Lname= '" & Module1.lectn & "' ")
        Call fill
   End If

   If Module1.lectdep = "EC" Then
        Set rs = scon.Execute("select * from lec where Lname= '" & Module1.lectn & "' ")
        Call fill
   End If

   If Module1.lectdep = "ME" Then
        Set rs = scon.Execute("select * from lme where Lname= '" & Module1.lectn & "' ")
        Call fill
   End If
    
   If Module1.lectdep = "SC" Then
        Set rs = scon.Execute("select * from lsc where Lname= '" & Module1.lectn & "' ")
        Call fill
   End If

End Sub



Private Sub fill()
        Text1.Text = rs!lname
        Text2.Text = rs!desg
        Text3.Text = rs!dept
        Text4.Text = rs!sname1
        Text5.Text = rs!sname2
        Text6.Text = rs!sname3
        Text7.Text = rs!sname4
        Text8.Text = rs!sname5
        Text9.Text = rs!sname6
End Sub

Private Sub Save_Click()
    
    Call sopenconnection
    
    If Module1.lectdep = "CS" Then
            Set Module1.rs = scon.Execute("UPDATE LCS Set LNAME= '" & Text1.Text & "',DESG= '" & Text2.Text & "',DEPT='" & Text3.Text & "',SNAME1='" & Text4.Text & "',SNAME2='" & Text5.Text & "',SNAME3='" & Text6.Text & "',SNAME4='" & Text7.Text & "',SNAME5='" & Text8.Text & "',SNAME6='" & Text9.Text & "' Where LNAME= '" & Module1.lectn & "' ")
    End If
    
    If Module1.lectdep = "EC" Then
            Set Module1.rs = scon.Execute("UPDATE LEC Set LNAME= '" & Text1.Text & "',DESG= '" & Text2.Text & "',DEPT='" & Text3.Text & "',SNAME1='" & Text4.Text & "',SNAME2='" & Text5.Text & "',SNAME3='" & Text6.Text & "',SNAME4='" & Text7.Text & "',SNAME5='" & Text8.Text & "',SNAME6='" & Text9.Text & "' Where LNAME= '" & Module1.lectn & "' ")
    End If
    
    If Module1.lectdep = "ME" Then
            Set Module1.rs = scon.Execute("UPDATE LME Set LNAME= '" & Text1.Text & "',DESG= '" & Text2.Text & "',DEPT='" & Text3.Text & "',SNAME1='" & Text4.Text & "',SNAME2='" & Text5.Text & "',SNAME3='" & Text6.Text & "',SNAME4='" & Text7.Text & "',SNAME5='" & Text8.Text & "',SNAME6='" & Text9.Text & "' Where LNAME= '" & Module1.lectn & "' ")
    End If
    
    If Module1.lectdep = "SC" Then
            Set Module1.rs = scon.Execute("UPDATE LSC Set LNAME= '" & Text1.Text & "',DESG= '" & Text2.Text & "',DEPT='" & Text3.Text & "',SNAME1='" & Text4.Text & "',SNAME2='" & Text5.Text & "',SNAME3='" & Text6.Text & "',SNAME4='" & Text7.Text & "',SNAME5='" & Text8.Text & "',SNAME6='" & Text9.Text & "' Where LNAME= '" & Module1.lectn & "' ")
    End If
                
    MsgBox ("ATTRIBUTES UPDATED!!")
    Unload Me
    LectUpdate.Show
    
End Sub

        
Private Sub Text1_LOSTFOCUS()
    Text1.Enabled = False
End Sub
Private Sub Text2_LOSTFOCUS()
    Text2.Enabled = False
End Sub
Private Sub Text3_LOSTFOCUS()
    Text3.Enabled = False
End Sub
Private Sub Text4_LOSTFOCUS()
    Text4.Enabled = False
End Sub
Private Sub Text5_LOSTFOCUS()
    Text5.Enabled = False
End Sub
Private Sub Text6_LOSTFOCUS()
    Text6.Enabled = False
End Sub
Private Sub Text7_LOSTFOCUS()
    Text7.Enabled = False
End Sub
Private Sub Text8_LOSTFOCUS()
    Text8.Enabled = False
End Sub
Private Sub Text9_LOSTFOCUS()
    Text9.Enabled = False
End Sub

