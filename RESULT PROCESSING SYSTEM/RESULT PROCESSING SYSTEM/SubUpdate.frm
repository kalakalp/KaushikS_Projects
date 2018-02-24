VERSION 5.00
Begin VB.Form SubUpdate 
   BackColor       =   &H00FF8080&
   Caption         =   "Update"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "SubUpdate.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton Command2 
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
      Left            =   12480
      TabIndex        =   13
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
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
      Left            =   10920
      TabIndex        =   12
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00400000&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   510
      Left            =   9000
      TabIndex        =   7
      Tag             =   "Private Sub Form_Load()"
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00400000&
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   510
      Left            =   9000
      TabIndex        =   6
      Tag             =   "Private Sub Form_Load()"
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox SubCombo 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   5
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Sub"
      Top             =   6840
      Visible         =   0   'False
      Width           =   12735
   End
   Begin VB.ComboBox SemCombo 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      ItemData        =   "SubUpdate.frx":8DC77
      Left            =   9120
      List            =   "SubUpdate.frx":8DC79
      TabIndex        =   4
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Sem"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox DepCombo 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      ItemData        =   "SubUpdate.frx":8DC7B
      Left            =   6600
      List            =   "SubUpdate.frx":8DC88
      TabIndex        =   3
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Dep"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
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
      Left            =   2640
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.CommandButton LectNext 
      Caption         =   "&Ok"
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
      Left            =   9720
      TabIndex        =   1
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   3120
      TabIndex        =   0
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the syllabus"
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
      Left            =   3360
      TabIndex        =   11
      Tag             =   "Private Sub Form_Load()"
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label ENTRSUB 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Tag             =   "Private Sub Form_Load()"
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Tag             =   "Private Sub Form_Load()"
      X1              =   0
      X2              =   15240
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label ENTRDEPT 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the Department and Semester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Tag             =   "Private Sub Form_Load()"
      Top             =   2760
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the new subject name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   5055
   End
End
Attribute VB_Name = "SubUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Unload Me
    subadd.Show

End Sub

Private Sub Command2_Click()
    
    Unload Me
    Del.Show

End Sub

Private Sub DepCombo_Click()
    SubCombo.Clear
    Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    subject
End Sub

Private Sub Form_Load()

    Dim i As Integer
    WindowState = 2
    Option1.Value = True

    For i = 1 To 6
        SemCombo.AddItem (i)
    Next
    
End Sub

Private Sub LectBack_Click()

     Unload Me
     Adminupdate.Show
    
End Sub

Private Sub LectNext_Click()

    Module1.SLNO = SubCombo.ListIndex + 1
    
        
    If Option1.Value = True Then
        Call openconnectionOLD
     Else
        Call openconnection1
     End If
    
        If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE ISEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE ISEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE ISEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE IISEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE IISEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE IISEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE IIISEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 3 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE IIISEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE IIISEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE IVSEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE IVSEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE IVSEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 5 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE VSEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE VSEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE VSEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("UPDATE VISEMCS Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("UPDATE VISEMEC Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("UPDATE VISEMME Set SUBNAME= '" & Text1.Text & "' Where SUBNAME= '" & SubCombo.Text & "' ")
        End If


     
   MsgBox ("Subject updated")
   Unload Me
   Adminupdate.Show
   
End Sub

Private Sub SemCombo_Click()
    SubCombo.Clear
     Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    subject
    ENTRSUB.Visible = True
    SubCombo.Visible = True
    LectBack.Visible = True
    LectNext.Visible = True
    Label1.Visible = True
    Text1.Visible = True
    
End Sub

Public Sub subject()
    
    
    If Option1.Value = True Then
        Call openconnectionOLD
     Else
        Call openconnection1
     End If
     
    If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from ISEMCS")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
    End If
If Module1.lectsem = 1 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from ISEMEC ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 1 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from ISEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IISEMCS ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IISEMEC")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IISEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IIISEMCS ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IIISEMEC ")
   SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IIISEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IVSEMCS ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IVSEMEC ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IVSEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from VSEMCS ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from VSEMEC ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from VSEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from VISEMCS ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from VISEMEC ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from VISEMME ")
    SubCombo.Clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
        
    SubCombo.Enabled = True
    Module1.subj = SubCombo.Text
    
End Sub

