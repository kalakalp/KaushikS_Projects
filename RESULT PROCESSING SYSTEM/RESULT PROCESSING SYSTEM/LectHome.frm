VERSION 5.00
Begin VB.Form LectHome 
   BackColor       =   &H00FF8080&
   Caption         =   "LectHome"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "New"
      Height          =   510
      Left            =   6720
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Old"
      Height          =   510
      Left            =   6720
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton LectNext 
      Caption         =   "&Next"
      Height          =   510
      Left            =   12600
      TabIndex        =   6
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox SubCombo 
      Enabled         =   0   'False
      Height          =   630
      Left            =   1560
      TabIndex        =   4
      Text            =   "Sub"
      Top             =   7680
      Visible         =   0   'False
      Width           =   12735
   End
   Begin VB.CommandButton LectBack 
      Caption         =   "&BACK"
      Height          =   510
      Left            =   2280
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox SemCombo 
      Height          =   630
      ItemData        =   "LectHome.frx":0000
      Left            =   9120
      List            =   "LectHome.frx":0016
      TabIndex        =   1
      Text            =   "Sem"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.ComboBox DepCombo 
      Height          =   630
      ItemData        =   "LectHome.frx":002C
      Left            =   6600
      List            =   "LectHome.frx":0039
      TabIndex        =   0
      Text            =   "Dep"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the syllabus"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label ENTRSUB 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the subject"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   15240
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label ENTRDEPT 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select the Department and Semester"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   9135
   End
End
Attribute VB_Name = "LectHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Back_Click()
    AdminEndLogin.Show
End Sub
Private Sub DepCombo_Click()
    SubCombo.Clear
    Module1.lectdep = DepCombo.Text
    Module1.lectsem = Val(SemCombo.Text)
    subject
End Sub

Private Sub Form_Load()

    Option1.Value = True

End Sub

Private Sub LectNext_Click()

     LectRes.Show
    
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
    
End Sub

Public Sub subject()
    
    Call openconnection1
    
    If Module1.lectsem = 1 And Module1.lectdep = "CS" Then
    Set Module1.rs = con1.Execute("select * from ISEMCS ")
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
    Set Module1.rs = con1.Execute("select * from IISEMEC ")
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
