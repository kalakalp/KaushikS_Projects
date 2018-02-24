VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form LectRes 
   Caption         =   "LectRes"
   ClientHeight    =   10740
   ClientLeft      =   210
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "LectRes1.frx":0000
   ScaleHeight     =   537
   ScaleMode       =   2  'Point
   ScaleWidth      =   756
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   7920
      Top             =   8040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   7
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&BACK"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&CLICK TO VIEW SEM VISE RESULT"
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00C0FFFF&
      Picture         =   "LectRes1.frx":4DA51
      TabIndex        =   2
      Top             =   10320
      Width           =   5055
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   4815
      Left            =   6840
      OleObjectBlob   =   "LectRes1.frx":54F43
      TabIndex        =   0
      Top             =   2640
      Width           =   7935
   End
   Begin MSChart20Lib.MSChart MSChart1 
      DragIcon        =   "LectRes1.frx":57922
      Height          =   9975
      Left            =   0
      OleObjectBlob   =   "LectRes1.frx":5EE14
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13320
      TabIndex        =   6
      Top             =   9720
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   4
      Top             =   1200
      Width           =   11415
   End
End
Attribute VB_Name = "LectRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs, rs1 As Object, SLNO, big As String, i, a, d, d1, total1, total, fc1, fc, s, s1, p, p1, f1, f As Integer

Private Sub Command1_Click()
    SEM_RESULT.Show
End Sub

Private Sub Command2_Click()
    
    Unload Me
    LectHome.Show

End Sub

Private Sub Form_Load()
 Timer1.Enabled = True
 Call pathsub
    Set rs = PathOb.Execute("select * from path1")
    var.path = rs!c2
    

    d = 0
    fc = 0
    s = 0
    p = 0
    f = 0
    a = 0
    
    WindowState = 2
    
    SLNO = "EX" & Module1.SLNO
    
    Call openconnection
    Set rs = con.Execute("select * from RESULTS where COURSECODE='" & Module1.lectdep & "' and SEM= '" & Module1.lectsem & "' ")
    
    
    If SLNO = "EX1" Then
    
        While Not rs.EOF
            If rs!EX1 < 35 Then
                f = f + 1
            ElseIf rs!EX1 >= 35 And rs!EX1 < 50 Then
                p = p + 1
            ElseIf rs!EX1 >= 50 And rs!EX1 < 60 Then
                s = s + 1
            ElseIf rs!EX1 >= 60 And rs!EX1 < 75 Then
                fc = fc + 1
            ElseIf rs!EX1 >= 75 And rs!EX1 <= 100 Then
                d = d + 1
            ElseIf rs!EX1 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
               
        
    ElseIf SLNO = "EX2" Then
        While Not rs.EOF
            If rs!EX2 <= 35 Then
                f = f + 1
            ElseIf rs!EX2 >= 35 And rs!EX2 < 50 Then
                p = p + 1
            ElseIf rs!EX2 >= 50 And rs!EX2 < 60 Then
                s = s + 1
            ElseIf rs!EX2 >= 60 And rs!EX2 < 75 Then
                fc = fc + 1
            ElseIf rs!EX2 >= 75 And rs!EX2 <= 100 Then
                d = d + 1
            ElseIf rs!EX2 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
        
    ElseIf SLNO = "EX3" Then
        While Not rs.EOF
            If rs!EX3 < 35 Then
                f = f + 1
            ElseIf rs!EX3 >= 35 And rs!EX3 < 50 Then
                p = p + 1
            ElseIf rs!EX3 >= 50 And rs!EX3 < 60 Then
                s = s + 1
            ElseIf rs!EX3 >= 60 And rs!EX3 < 75 Then
                fc = fc + 1
            ElseIf rs!EX3 >= 75 And rs!EX3 <= 100 Then
                d = d + 1
            ElseIf rs!EX3 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
    
    ElseIf SLNO = "EX4" Then
        While Not rs.EOF
            If rs!EX4 < 35 Then
                f = f + 1
            ElseIf rs!EX4 >= 35 And rs!EX4 < 50 Then
                p = p + 1
            ElseIf rs!EX4 >= 50 And rs!EX4 < 60 Then
                s = s + 1
            ElseIf rs!EX4 >= 60 And rs!EX4 < 75 Then
                fc = fc + 1
            ElseIf rs!EX4 >= 75 And rs!EX4 <= 100 Then
                d = d + 1
            ElseIf rs!EX4 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
    
    ElseIf SLNO = "EX5" Then
        While Not rs.EOF
            If rs!EX5 < 35 Then
                f = f + 1
            ElseIf rs!EX5 >= 35 And rs!EX5 < 50 Then
                p = p + 1
            ElseIf rs!EX5 >= 50 And rs!EX5 < 60 Then
                s = s + 1
            ElseIf rs!EX5 >= 60 And rs!EX5 < 75 Then
                fc = fc + 1
            ElseIf rs!EX5 >= 75 And rs!EX5 <= 100 Then
            d = d + 1
            ElseIf rs!EX5 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
    
    ElseIf SLNO = "EX6" Then
        While Not rs.EOF
            If rs!EX6 < 35 Then
                f = f + 1
            ElseIf rs!EX6 >= 35 And rs!EX6 < 50 Then
                p = p + 1
            ElseIf rs!EX6 >= 50 And rs!EX6 < 60 Then
                s = s + 1
            ElseIf rs!EX6 >= 60 And rs!EX6 < 75 Then
                fc = fc + 1
            ElseIf rs!EX6 >= 75 And rs!EX6 <= 100 Then
                d = d + 1
            ElseIf rs!EX6 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
    
    ElseIf SLNO = "EX7" Then
        While Not rs.EOF
            If rs!EX7 < 35 Then
                f = f + 1
            ElseIf rs!EX7 >= 35 And rs!EX7 < 50 Then
                p = p + 1
            ElseIf rs!EX7 >= 50 And rs!EX7 < 60 Then
                s = s + 1
            ElseIf rs!EX7 >= 60 And rs!EX7 < 75 Then
                fc = fc + 1
            ElseIf rs!EX7 >= 75 And rs!EX7 <= 100 Then
                d = d + 1
            ElseIf rs!EX7 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
    
    ElseIf SLNO = "EX8" Then
    
        While Not rs.EOF
            If rs!EX8 < 35 Then
                f = f + 1
            ElseIf rs!EX8 >= 35 And rs!EX8 < 50 Then
                p = p + 1
            ElseIf rs!EX8 >= 50 And rs!EX8 < 60 Then
                s = s + 1
            ElseIf rs!EX8 >= 60 And rs!EX8 < 75 Then
                fc = fc + 1
            ElseIf rs!EX8 >= 75 And rs!EX8 <= 100 Then
                d = d + 1
            ElseIf rs!EX8 = "AB" Then
                a = a + 1
            Else
            End If
            rs.MoveNext
        Wend
   
   End If
   
   total = f + p + s + fc + d + a
   
   
   d1 = d / total * 100
   fc1 = fc / total * 100
   s1 = s / total * 100
   f1 = f / total * 100
   a1 = a / total * 100
   p1 = p / total * 100
 

        MSChart2.Column = 1
        MSChart2.Data = d1
         big = d
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart2.Column = 2
        MSChart2.Data = fc1
         big = fc
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart2.Column = 3
        MSChart2.Data = s1
        big = s
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart2.Column = 4
        MSChart2.Data = p1
        big = p
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart2.Column = 5
        MSChart2.Data = f1
        big = f
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart2.Column = 6
        MSChart2.Data = a1
        big = a
        MSChart2.ColumnLabel = MSChart2.ColumnLabel + "   " + MSChart2.Data + "%   (" + big + ")"
        MSChart1.Column = 1
        MSChart1.Data = d1 + fc1 + s1 + p1
        big = total - f - a
        MSChart1.ColumnLabel = MSChart1.ColumnLabel + "   " + MSChart1.Data + "%   (" + big + ")"
        If Module1.lectdep = "CS" Then
            MSChart2.TitleText = "COMPUTER SCIENCE " + Module1.lectsem + "  SEM"
            MSChart2.Footnote = Module1.lectsubj
        ElseIf Module1.lectdep = "ME" Then
            MSChart2.TitleText = "MEHANICAL " + Module1.lectsem + "  SEM"
            MSChart2.Footnote = Module1.lectsubj
        ElseIf Module1.lectdep = "EC" Then
            MSChart2.TitleText = "ELECTRONICS " + Module1.lectsem + "  SEM"
            MSChart2.Footnote = Module1.lectsubj
        End If


Call openconnection
Set rs = con.Execute("select regno from RESULTS")
    reg = 0
While Not rs.EOF
    If rs!regno > reg Then
        reg = rs!regno
    End If
    rs.MoveNext
Wend
 
reg = Mid$(reg, 6, 2)
Label1.Caption = "Showing results of academic year 20" & reg & "-0" & reg + 1

End Sub

Private Sub Timer1_Timer()

    Label4.Caption = Date
    Label5.Caption = Time
    
End Sub

Private Sub Command3_Click()
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    PrintForm
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
        
    Command1.Visible = True
    Command2.Visible = True
    Command3.Visible = True
    Timer2.Enabled = False

End Sub
