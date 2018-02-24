VERSION 5.00
Begin VB.Form DELETE 
   Caption         =   "delete"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "subadd.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   4680
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
      TabIndex        =   6
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton DEL 
      Caption         =   "&Delete"
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
      Left            =   10560
      TabIndex        =   5
      Tag             =   "Private Sub Form_Load()"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
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
      ItemData        =   "subadd.frx":8DC77
      Left            =   6600
      List            =   "subadd.frx":8DC84
      TabIndex        =   4
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Dep"
      Top             =   4080
      Width           =   1455
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
      ItemData        =   "subadd.frx":8DC94
      Left            =   9120
      List            =   "subadd.frx":8DC96
      TabIndex        =   3
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Sem"
      Top             =   4080
      Width           =   1935
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
      TabIndex        =   2
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Sub"
      Top             =   6840
      Visible         =   0   'False
      Width           =   12735
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
      TabIndex        =   1
      Tag             =   "Private Sub Form_Load()"
      Top             =   0
      Width           =   1215
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
      TabIndex        =   0
      Tag             =   "Private Sub Form_Load()"
      Top             =   480
      Width           =   1215
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
   Begin VB.Line Line1 
      BorderWidth     =   3
      Tag             =   "Private Sub Form_Load()"
      X1              =   0
      X2              =   15240
      Y1              =   5520
      Y2              =   5520
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
      TabIndex        =   8
      Tag             =   "Private Sub Form_Load()"
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
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
      TabIndex        =   7
      Tag             =   "Private Sub Form_Load()"
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "DELETE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    subadd.Show

End Sub

Private Sub DepCombo_Click()
    SubCombo.clear
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

Private Sub DEL_Click()

    'Module1.SLNO = SubCombo.ListIndex + 1
    
        
    If Option1.Value = True Then
        Call openconnectionOLD
     Else
        Call openconnection1
     End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM ISEMCS Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM ISEMEC Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM ISEMME Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM IISEMCS Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM IISEMEC Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM IISEMME Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM IIISEMCS Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 3 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM IIISEMEC Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM IIISEMME Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM IVSEMCS Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM IVSEMEC Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM IVSEMME Where SUBNAME= '" & SubCombo.Text & "' ")
        End If
        
        If Module1.lectsem = 5 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM VSEMCS Where SUBNAME= '" & SubCombo.Text & "'  ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM VSEMEC Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM VSEMME Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("DELETE FROM VISEMCS Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("DELETE FROM VISEMEC Where SUBNAME= '" & SubCombo.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("DELETE FROM VISEMME Where SUBNAME= '" & SubCombo.Text & "' ")
        End If


     
   MsgBox ("Subject DELETEd")
   Unload Me
   SUBDELETE.Show
   
End Sub

Private Sub SemCombo_Click()
    SubCombo.clear
     Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    subject
    ENTRSUB.Visible = True
    SubCombo.Visible = True
    LectBack.Visible = True
    DEL.Visible = True
    'Label1.Visible = True
    'Text1.Visible = True
    
End Sub

Public Sub subject()
    
    
    If Option1.Value = True Then
        Call openconnectionOLD
     Else
        Call openconnection1
     End If
     
    If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from ISEMCS")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
    End If
If Module1.lectsem = 1 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from ISEMEC ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 1 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from ISEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IISEMCS ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IISEMEC")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 2 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IISEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IIISEMCS ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IIISEMEC ")
   SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 3 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IIISEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from IVSEMCS ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from IVSEMEC ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 4 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from IVSEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from VSEMCS ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from VSEMEC ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 5 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from VSEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from VISEMCS ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "EC" Then
    Set Module1.rs = con1.Execute("select * from VISEMEC ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
If Module1.lectsem = 6 And Module1.lectdep = "ME" Then
    Set Module1.rs = con1.Execute("select * from VISEMME ")
    SubCombo.clear
    While Not Module1.rs.EOF
        SubCombo.AddItem (Module1.rs!SUBNAME)
        Module1.rs.MoveNext
    Wend
End If
        
    SubCombo.Enabled = True
    Module1.subj = SubCombo.Text
    
End Sub
