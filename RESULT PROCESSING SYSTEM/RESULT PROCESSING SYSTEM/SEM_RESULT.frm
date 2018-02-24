VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SEM_RESULT 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SEM_RESULT"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "SEM_RESULT.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdclose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12480
      TabIndex        =   1
      Top             =   9600
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid list 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   12726
      _Version        =   393216
      Rows            =   25
      Cols            =   25
      FixedCols       =   0
      BackColor       =   8454143
      BackColorFixed  =   8454143
      BackColorBkg    =   12648447
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      BorderStyle     =   0
      FormatString    =   $"SEM_RESULT.frx":27D41
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "SEM_RESULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim co, rs As Object
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   WindowState = 2
   Call openconnection
    
   Set rs = con.Execute("select * from RESULTS where COURSECODE='" & Module1.lectdep & "'and SEM='" & Module1.lectsem & "' ")
   
   Dim intRC As Integer
    
    intRC = 0
    ReDim rsRecordArr(intRC)
    With list
    ReDim Preserve rsRecordArr(intRC + 1)
        .Rows = intRC + 1
        .Cols = 25
        .Row = intRC
    End With
    intRC = intRC + 1
    
    While Not rs.EOF
       
       With list
            ReDim Preserve rsRecordArr(intRC + 1)
            .Rows = intRC + 1
            .Cols = 25
            .Row = intRC
            .Col = 0
            .Text = rs("INSTCODE").Value
            .Col = 1
            .Text = rs("COURSECODE").Value
            .Col = 2
            .Text = rs("REGNO").Value
            .Col = 3
            .Text = rs("NAME").Value
            .Col = 4
            .Text = rs("SEM").Value
            .Col = 5
            .Text = rs("EX1").Value
            .Col = 6
            .Text = rs("EX2").Value
            .Col = 7
            .Text = rs("EX3").Value
            .Col = 8
            .Text = rs("EX4").Value
            .Col = 9
            .Text = rs("EX5").Value
            .Col = 10
            .Text = rs("EX6").Value
            .Col = 11
            .Text = rs("EX7").Value
            .Col = 12
            .Text = rs("EX8").Value
            .Col = 13
            .Text = rs("EXTOTAL").Value
            .Col = 14
            .Text = rs("IA1").Value
            .Col = 15
            .Text = rs("IA2").Value
            .Col = 16
            .Text = rs("IA3").Value
            .Col = 17
            .Text = rs("IA4").Value
            .Col = 18
            .Text = rs("IA5").Value
            .Col = 19
            .Text = rs("IA6").Value
            .Col = 20
            .Text = rs("IA7").Value
            .Col = 21
            .Text = rs("IA8").Value
            .Col = 22
            .Text = rs("IATOTAL").Value
            .Col = 23
            .Text = rs("GRANDTOTAL").Value
            .Col = 24
            .Text = rs("RESULT").Value
            
            
            
        End With
        
        intRC = intRC + 1
        rs.MoveNext
        
    Wend
End Sub


