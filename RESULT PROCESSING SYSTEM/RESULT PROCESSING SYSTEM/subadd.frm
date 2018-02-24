VERSION 5.00
Begin VB.Form subadd 
   Caption         =   "ADD"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "subadd.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   10095
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
      Left            =   8880
      TabIndex        =   5
      Tag             =   "Private Sub Form_Load()"
      Top             =   720
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
      Left            =   8880
      TabIndex        =   4
      Tag             =   "Private Sub Form_Load()"
      Top             =   240
      Width           =   1215
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
      ItemData        =   "subadd.frx":8DC77
      Left            =   9000
      List            =   "subadd.frx":8DC79
      TabIndex        =   3
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Sem"
      Top             =   4320
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
      ItemData        =   "subadd.frx":8DC7B
      Left            =   6480
      List            =   "subadd.frx":8DC88
      TabIndex        =   2
      Tag             =   "Private Sub Form_Load()"
      Text            =   "Dep"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton ADD 
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
      Left            =   10440
      TabIndex        =   1
      Tag             =   "Private Sub Form_Load()"
      Top             =   10200
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
      Left            =   3000
      TabIndex        =   0
      Tag             =   "Private Sub Form_Load()"
      Top             =   10200
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
      Left            =   3240
      TabIndex        =   8
      Tag             =   "Private Sub Form_Load()"
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label ENTRSUB 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the new subject"
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
      Left            =   6600
      TabIndex        =   7
      Tag             =   "Private Sub Form_Load()"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Tag             =   "Private Sub Form_Load()"
      X1              =   -120
      X2              =   15120
      Y1              =   5760
      Y2              =   5760
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
      Left            =   3960
      TabIndex        =   6
      Tag             =   "Private Sub Form_Load()"
      Top             =   3000
      Width           =   9135
   End
End
Attribute VB_Name = "subadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer

Private Sub ADD_Click()
A = 0

    If Option1.Value = True Then
        Call openconnectionOLD
     Else
        Call openconnection1
     End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
        
            Set Module1.rs = con1.Execute("SELECT * FROM  ISEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO ISEMCS (SLNO,SUBNAME) VALUES ('" & A & "','" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "EC" Then
            Set Module1.rs = con1.Execute("SELECT * FROM  ISEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO ISEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 1 And Module1.lectdep = "ME" Then
            Set Module1.rs = con1.Execute("SELECT * FROM  ISEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO ISEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "CS" Then
            Set Module1.rs = con1.Execute("SELECT * FROM  IISEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IISEMCS SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "EC" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IISEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IISEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 2 And Module1.lectdep = "ME" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IISEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IISEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "CS" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IIISEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IIISEMCS SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 3 And Module1.lectdep = "EC" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IIISEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IIISEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 3 And Module1.lectdep = "ME" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IIISEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IIISEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "CS" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IVSEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IVSEMCS SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "EC" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IVSEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IVSEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 4 And Module1.lectdep = "ME" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  IVSEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO IVSEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If
        
        If Module1.lectsem = 5 And Module1.lectdep = "CS" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VSEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VSEMCS SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "EC" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VSEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VSEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 5 And Module1.lectdep = "ME" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VSEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
            rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VSEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "CS" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VISEMCS")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VISEMCS SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "EC" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VISEMEC")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VISEMEC SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If

        If Module1.lectsem = 6 And Module1.lectdep = "ME" Then
               Set Module1.rs = con1.Execute("SELECT * FROM  VISEMME")
            While Not Module1.rs.EOF
                If A < Module1.rs!SLNO Then
                    A = Module1.rs!SLNO
                End If
                rs.MoveNext
            Wend
            A = A + 1
            Set Module1.rs = con1.Execute("INSERT INTO VISEMME SLNO,SNAME VALUES " & A & ", '" & Text1.Text & "' ")
        End If


     
   MsgBox ("Subject Added")
   Unload Me
   subadd.Show
   
End Sub

Private Sub DepCombo_Click()
    'SubCombo.clear
    Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    'subject
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
     SubUpdate.Show
    
End Sub

Private Sub DEL_Click()

    'Module1.SLNO = SubCombo.ListIndex + 1
    
        
End Sub

Private Sub SemCombo_Click()
    'SubCombo.clear
     Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    'SUBJECT
    ENTRSUB.Visible = True
    'SubCombo.Visible = True
    LectBack.Visible = True
    ADD.Visible = True
    'Label1.Visible = True
    Text1.Visible = True
    
End Sub

