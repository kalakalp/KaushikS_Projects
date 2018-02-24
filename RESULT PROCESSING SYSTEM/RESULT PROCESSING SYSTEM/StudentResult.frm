VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form StudentResult 
   BackColor       =   &H00FF8080&
   Caption         =   "Student Results"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "StudentResult.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   12360
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   47
      Top             =   9960
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   -720
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&CLICK TO VIEW SEM VISE RESULT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      MaskColor       =   &H00C0FFFF&
      Picture         =   "StudentResult.frx":3A648
      TabIndex        =   44
      Top             =   9960
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   35
      Top             =   9960
      Width           =   1695
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "StudentResult.frx":41B3A
      TabIndex        =   42
      Top             =   5640
      Width           =   8175
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   46
      Top             =   9000
      Width           =   2175
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   45
      Top             =   9360
      Width           =   2175
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5520
      TabIndex        =   43
      Top             =   240
      Width           =   18855
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   3
      Height          =   1455
      Left            =   8280
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   8280
      X2              =   12120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label GRANDTOTALLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "GRANDTOTAL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   41
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label RESULTLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "RESULT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8400
      TabIndex        =   40
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label TOT8 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   39
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label EX8 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   38
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label SUB8IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   37
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label SUB8LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   36
      Top             =   5280
      Width           =   2835
   End
   Begin VB.Line Line11 
      X1              =   1200
      X2              =   11160
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label REGLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "REG:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   34
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label NAMELBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   33
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label SEMLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "SEM:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4200
      X2              =   4200
      Y1              =   720
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   495
      Left            =   4200
      Top             =   720
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   495
      Left            =   1200
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   495
      Left            =   1200
      Top             =   720
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   3855
      Left            =   1200
      Top             =   1800
      Width           =   9975
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   11160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   11160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   1200
      X2              =   11160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line5 
      X1              =   1200
      X2              =   11160
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line6 
      X1              =   1200
      X2              =   11160
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line7 
      X1              =   1200
      X2              =   11160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   3
      Height          =   615
      Left            =   1200
      Top             =   1200
      Width           =   9975
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   4200
      X2              =   4200
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   7200
      X2              =   7200
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Label SNAMELBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "SUBJECT NAME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   31
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label IAMARKLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "IA MARKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label EXAMARKLBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "EXAM MARKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   29
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label SUB1LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label SUB2LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   27
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label SUB3LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   26
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label SUB4LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   25
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label SUB5LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label SUB6LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   4320
      Width           =   2865
   End
   Begin VB.Label SUB7LBL 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   4800
      Width           =   2835
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   9360
      X2              =   9360
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Label SUB1IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label SUB2IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label SUB3IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label SUB4IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label SUB5IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label SUB6IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label SUB7IA 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label EX1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label EX2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label EX3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label EX4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label EX5 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label EX6 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label EX7 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label IAEX 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "IA + EXAM MARK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label TOT1 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label TOT2 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label TOT3 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label TOT4 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label TOT5 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label TOT6 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label TOT7 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
End
Attribute VB_Name = "StudentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs, rs2 As Object, i, gt, regmid As Integer
Private Sub Command1_Click()
    Unload Me
    StudentDetails.Show
End Sub

Private Sub Command2_Click()

    SEM_RESULT.Show
End Sub

Private Sub Command3_Click()
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    PrintForm
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    
Timer1.Enabled = True
    WindowState = 2
'__________________________Typing distinction or first class ....._____________________________________________________________________
    
    Call openconnection
    Set rs = con.Execute("select * from RESULTS where REGNO='" & var.reg & "'and SEM='" & var.sem & "' ")
    If (rs!Result = "D") Then
        RESULTLBL.Caption = RESULTLBL.Caption & "DISTINCTION"
        ElseIf (rs!Result = "1") Then
        RESULTLBL.Caption = RESULTLBL.Caption & "FIRST CLASS"
        ElseIf (rs!Result = "2") Then
        RESULTLBL.Caption = RESULTLBL.Caption & "SECOND CLASS"
        ElseIf (rs!Result = "F") Then
        RESULTLBL.Caption = RESULTLBL.Caption & "FAIL"
    End If

'________________________________Entering the name and the marks..._______________________________________________________________________________
    NAMELBL.Caption = NAMELBL.Caption & rs!Name
    REGLBL.Caption = REGLBL.Caption & rs!regno
    SEMLBL.Caption = SEMLBL.Caption & rs!sem

    SUB1IA.Caption = SUB1IA.Caption & rs!IA1
    SUB2IA.Caption = SUB2IA.Caption & rs!IA2
    SUB3IA.Caption = SUB3IA.Caption & rs!IA3
    SUB4IA.Caption = SUB4IA.Caption & rs!IA4
    SUB5IA.Caption = SUB5IA.Caption & rs!IA5
    SUB6IA.Caption = SUB6IA.Caption & rs!IA6
    SUB7IA.Caption = SUB7IA.Caption & rs!IA7
    
    EX1.Caption = rs!EX1
    EX2.Caption = rs!EX2
    EX3.Caption = rs!EX3
    EX4.Caption = rs!EX4
    EX5.Caption = rs!EX5
    EX6.Caption = rs!EX6
    EX7.Caption = rs!EX7
    TOT1.Caption = Val(SUB1IA.Caption) + Val(EX1.Caption)
    TOT2.Caption = Val(SUB2IA.Caption) + Val(EX2.Caption)
    TOT3.Caption = Val(SUB3IA.Caption) + Val(EX3.Caption)
    TOT4.Caption = Val(SUB4IA.Caption) + Val(EX4.Caption)
    TOT5.Caption = Val(SUB5IA.Caption) + Val(EX5.Caption)
    TOT6.Caption = Val(SUB6IA.Caption) + Val(EX6.Caption)
    TOT7.Caption = Val(SUB7IA.Caption) + Val(EX7.Caption)
    rs.MoveNext

'______________________________Entering total____________________________________________________________________
gt = Val(TOT1.Caption) + Val(TOT2.Caption) + Val(TOT3.Caption) + Val(TOT4.Caption) + Val(TOT5.Caption) + Val(TOT6.Caption) + Val(TOT7.Caption) + Val(TOT8.Caption)
GRANDTOTALLBL.Caption = GRANDTOTALLBL.Caption & gt

'______________________________Assigning Subject names_________________________________________________________

 regmid = Val(Mid$(var.reg, 6, 2))
 
 If regmid < 9 Then
    Call openconnectionOLD
    
    If var.sem = 1 And var.dept = "CS" Then
       Set rs = con1.Execute("select * from ISEMCS")
       While Not rs.EOF
            Call subject6
       Wend
    End If

    If var.sem = 1 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from ISEMEC ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 1 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from ISEMME ")
        While Not rs.EOF
            Call subject6
        Wend
    End If
    
    If var.sem = 2 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IISEMCS ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 2 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IISEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If
    
    If var.sem = 2 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IISEMME ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 3 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IIISEMCS ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 3 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IIISEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 3 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IIISEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 4 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IVSEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 4 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IVSEMEC ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 4 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IVSEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VSEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VSEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VSEMME ")
        While Not rs.EOF
            Call subject8
        Wend
    End If

    If var.sem = 6 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VISEMCS ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 6 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VISEMEC ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 6 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VISEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If
'____________________________________________________________________________________________
 
 Else
    
    Call openconnection1
        
    If var.sem = 1 And var.dept = "CS" Then
       Set rs = con1.Execute("select * from ISEMCS ")
       While Not rs.EOF
            Call subject7
       Wend
    End If

    If var.sem = 1 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from ISEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 1 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from ISEMME ")
        While Not rs.EOF
            Call subject6
        Wend
    End If
    
    If var.sem = 2 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IISEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 2 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IISEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If
    
    If var.sem = 2 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IISEMME ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 3 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IIISEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 3 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IIISEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 3 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IIISEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 4 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IVSEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 4 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IVSEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 4 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IVSEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VSEMCS ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VSEMEC ")
        While Not rs.EOF
            Call subject7
        Wend
    End If

    If var.sem = 5 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VSEMME ")
        While Not rs.EOF
            Call subject8
        Wend
    End If

    If var.sem = 6 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VISEMCS ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 6 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VISEMEC ")
        While Not rs.EOF
            Call subject6
        Wend
    End If

    If var.sem = 6 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VISEMME ")
        While Not rs.EOF
            Call subject7
        Wend
    End If
    
 End If
 

'________________________________CHART CODE_______________________________________
    
    Set rs2 = con.Execute("select * from RESULTS where REGNO='" & var.reg & "'")
    MSChart1.ColumnCount = 8
    MSChart1.Column = 1
    MSChart1.Data = Val(TOT1.Caption)
    MSChart1.Column = 2
    MSChart1.Data = Val(TOT2.Caption)
    MSChart1.Column = 3
    MSChart1.Data = Val(TOT3.Caption)
    MSChart1.Column = 4
    MSChart1.Data = Val(TOT4.Caption)
    MSChart1.Column = 5
    MSChart1.Data = Val(TOT5.Caption)
    MSChart1.Column = 6
    MSChart1.Data = Val(TOT6.Caption)
    MSChart1.Column = 7
    MSChart1.Data = Val(TOT7.Caption)
    MSChart1.Column = 8
    MSChart1.Data = Val(TOT8.Caption)
    '...............
    MSChart1.RowLabel = NAMELBL.Caption
    If var.sem = 1 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from ISEMCS")
        i = 1
    
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 1 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from ISEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 1 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from ISEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 2 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IISEMCS ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 2 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IISEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 2 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IISEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
           i = i + 1
        Wend
    End If

    If var.sem = 3 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IIISEMCS ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 3 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IIISEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 3 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IIISEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 4 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from IVSEMCS ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 4 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from IVSEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 4 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from IVSEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 5 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VSEMCS ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 5 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VSEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 5 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VSEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 6 And var.dept = "CS" Then
        Set rs = con1.Execute("select * from VISEMCS ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If

    If var.sem = 6 And var.dept = "EC" Then
        Set rs = con1.Execute("select * from VISEMEC ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
    End If
    
    If var.sem = 6 And var.dept = "ME" Then
        Set rs = con1.Execute("select * from VISEMME ")
        i = 1
        While Not rs.EOF
            MSChart1.Column = i
            MSChart1.ColumnLabel = rs!SUBNAME
            rs.MoveNext
            i = i + 1
        Wend
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

Private Sub subject7()
        SUB1LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB2LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB3LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB4LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB5LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB6LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB7LBL.Caption = rs!SUBNAME
        rs.MoveNext
        
End Sub
Private Sub subject6()
        SUB1LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB2LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB3LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB4LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB5LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB6LBL.Caption = rs!SUBNAME
        rs.MoveNext
        
End Sub
'Private Sub subject5()
   '     SUB1LBL.Caption = rs!SUBNAME
 '       rs.MoveNext
  '      SUB2LBL.Caption = rs!SUBNAME
    '    rs.MoveNext
     '   SUB3LBL.Caption = rs!SUBNAME
     '   rs.MoveNext
     '   SUB4LBL.Caption = rs!SUBNAME
     '   rs.MoveNext
     '   SUB5LBL.Caption = rs!SUBNAME
     '   rs.MoveNext
        
'End Sub
Private Sub subject8()
        SUB1LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB2LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB3LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB4LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB5LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB6LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB7LBL.Caption = rs!SUBNAME
        rs.MoveNext
        SUB8LBL.Caption = rs!SUBNAME
        rs.MoveNext
End Sub
Private Sub Timer1_Timer()

    Label4.Caption = Date
    Label5.Caption = Time
    
End Sub

Private Sub Timer2_Timer()
    
    Command1.Visible = True
    Command2.Visible = True
    Command3.Visible = True
    Timer2.Enabled = False
    
End Sub
