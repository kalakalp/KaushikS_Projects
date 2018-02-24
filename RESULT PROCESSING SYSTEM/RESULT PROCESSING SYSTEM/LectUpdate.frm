VERSION 5.00
Begin VB.Form LectUpdate 
   Caption         =   "Update"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "LectUpdate.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton Command1 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8640
      TabIndex        =   6
      Tag             =   "Private Sub Form_Load()"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton LectBack 
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4680
      TabIndex        =   4
      Tag             =   "Private Sub Form_Load()"
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton LectNext 
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6720
      TabIndex        =   3
      Tag             =   "Private Sub Form_Load()"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox SubCombo 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Lecturer name"
      Top             =   7200
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.ComboBox DepCombo 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "LectUpdate.frx":2C0D6
      Left            =   8760
      List            =   "LectUpdate.frx":2C0E6
      TabIndex        =   0
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Dep"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label ENTRSUB 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Lecturer name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Tag             =   "Private Sub Form_Load()"
      Top             =   6480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label ENTRDEPT 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Tag             =   "Private Sub Form_Load()"
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Tag             =   "Private Sub Form_Load()"
      X1              =   0
      X2              =   15240
      Y1              =   5520
      Y2              =   5520
   End
End
Attribute VB_Name = "LectUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Object, q As Integer

Private Sub Command1_Click()

Call sopenconnection
        If DepCombo.Text = "CS" Then
            q = InputBox("Enter 1 to delete and 0 to exit")
            
            If (q) Then
                Set rs = scon.Execute("delete * from lcs where lname= '" & SubCombo.Text & "'")
                MsgBox ("Record Removed")
                Unload Me
                LectUpdate.Show
            End If
        
        ElseIf DepCombo.Text = "EC" Then
            q = InputBox("Enter 1 to delete and 0 to exit")
            
            If (q) Then
                Set rs = scon.Execute("delete * from lcs where lname= '" & SubCombo.Text & "'")
                MsgBox ("Record Removed")
                Unload Me
                LectUpdate.Show

            End If
        
        ElseIf DepCombo.Text = "ME" Then
            q = InputBox("Enter 1 to delete and 0 to exit")
            
            If (q) Then
                Set rs = scon.Execute("delete * from lcs where lname= '" & SubCombo.Text & "'")
                MsgBox ("Record Removed")
                Unload Me
                LectUpdate.Show
            
            End If
        
        ElseIf DepCombo.Text = "SC" Then
            q = InputBox("Enter 1 to delete and 0 to exit")
            
            If (q) Then
                Set rs = scon.Execute("delete * from lcs where lname= '" & SubCombo.Text & "'")
                MsgBox ("Record Removed")
                Unload Me
                LectUpdate.Show
            
            End If
        
        End If
End Sub

Private Sub DepCombo_Click()
   SubCombo.Visible = True
   ENTRSUB.Visible = True
  ' LectBack.Visible = True
   LectNext.Visible = True
   Command1.Visible = True
   SubCombo.clear
   
   If DepCombo.Text = "CS" Then
        Call sopenconnection
        Set rs = scon.Execute("select lname from lcs")
   
        While Not rs.EOF
            SubCombo.AddItem (rs!lname)
            rs.MoveNext
        Wend
   End If
   
   If DepCombo.Text = "EC" Then
        Call sopenconnection
        Set rs = scon.Execute("select lname from lec")
   
        While Not rs.EOF
            SubCombo.AddItem (rs!lname)
            rs.MoveNext
        Wend
    End If
   
   If DepCombo.Text = "ME" Then
        Call sopenconnection
        Set rs = scon.Execute("select lname from lme")
   
        While Not rs.EOF
            SubCombo.AddItem (rs!lname)
            rs.MoveNext
        Wend
    End If
    
    If DepCombo.Text = "SC" Then
        Call sopenconnection
        Set rs = scon.Execute("select lname from lsc")
   
        While Not rs.EOF
            SubCombo.AddItem (rs!lname)
            rs.MoveNext
        Wend
        
   End If
   
   
   
End Sub
Private Sub Form_Load()

    WindowState = 2
    
End Sub

Private Sub LectBack_Click()

    Unload Me
    Adminupdate.Show

End Sub

Private Sub LectNext_Click()
    Module1.lectn = SubCombo.Text
    Module1.lectdep = DepCombo.Text
    Unload Me
    Lectdetails.Show
    
End Sub
