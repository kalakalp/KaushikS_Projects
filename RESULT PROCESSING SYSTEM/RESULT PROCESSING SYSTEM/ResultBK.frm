VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Resultbk 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "ResultBK.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
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
      Left            =   5040
      TabIndex        =   2
      Top             =   10680
      Width           =   1695
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
      Left            =   7080
      MaskColor       =   &H00C0FFFF&
      Picture         =   "ResultBK.frx":3A648
      TabIndex        =   1
      Top             =   10680
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   480
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
      Left            =   13080
      TabIndex        =   0
      Top             =   10680
      Width           =   1695
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4095
      Left            =   1080
      OleObjectBlob   =   "ResultBK.frx":41B3A
      TabIndex        =   3
      Top             =   6360
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12960
      TabIndex        =   55
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Label rs8 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   54
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label rs7 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   53
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label rs6 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   52
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label rs5 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   51
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label rs4 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   50
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7200
      TabIndex        =   49
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label rs2 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   48
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label rs1 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   47
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   11040
      X2              =   11040
      Y1              =   1920
      Y2              =   6360
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
      Left            =   9360
      TabIndex        =   46
      Top             =   5520
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
      Left            =   9360
      TabIndex        =   45
      Top             =   5040
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
      Left            =   9360
      TabIndex        =   44
      Top             =   4560
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
      Left            =   9360
      TabIndex        =   43
      Top             =   4080
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
      Left            =   9360
      TabIndex        =   42
      Top             =   3600
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
      Left            =   9360
      TabIndex        =   41
      Top             =   3120
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
      Left            =   9360
      TabIndex        =   40
      Top             =   2640
      Width           =   1575
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
      Left            =   9360
      TabIndex        =   39
      Top             =   2040
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   38
      Top             =   5520
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
      Left            =   7200
      TabIndex        =   37
      Top             =   5040
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
      Left            =   7200
      TabIndex        =   36
      Top             =   4560
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
      Left            =   7200
      TabIndex        =   35
      Top             =   4080
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
      Left            =   7200
      TabIndex        =   34
      Top             =   3600
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
      Left            =   7200
      TabIndex        =   33
      Top             =   3120
      Width           =   1935
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
      Left            =   7200
      TabIndex        =   32
      Top             =   2640
      Width           =   1935
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
      Left            =   4200
      TabIndex        =   31
      Top             =   5520
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
      Left            =   4200
      TabIndex        =   30
      Top             =   5040
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
      Left            =   4200
      TabIndex        =   29
      Top             =   4560
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
      Left            =   4200
      TabIndex        =   28
      Top             =   4080
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
      Left            =   4200
      TabIndex        =   27
      Top             =   3600
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
      Left            =   4200
      TabIndex        =   26
      Top             =   3120
      Width           =   2775
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
      Left            =   4200
      TabIndex        =   25
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   9240
      X2              =   9240
      Y1              =   1920
      Y2              =   6360
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
      Left            =   1200
      TabIndex        =   24
      Top             =   5520
      Width           =   2835
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
      Left            =   1200
      TabIndex        =   23
      Top             =   5040
      Width           =   2865
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
      Left            =   1200
      TabIndex        =   22
      Top             =   4560
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
      Left            =   1200
      TabIndex        =   21
      Top             =   4080
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
      Left            =   1200
      TabIndex        =   20
      Top             =   3600
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
      Left            =   1200
      TabIndex        =   19
      Top             =   3120
      Width           =   2775
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
      Left            =   1200
      TabIndex        =   18
      Top             =   2640
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
      Left            =   7200
      TabIndex        =   17
      Top             =   2040
      Width           =   1815
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
      Left            =   4200
      TabIndex        =   16
      Top             =   2040
      Width           =   2775
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
      Left            =   1200
      TabIndex        =   15
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   7080
      X2              =   7080
      Y1              =   1920
      Y2              =   6360
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   4080
      X2              =   4080
      Y1              =   1920
      Y2              =   6360
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   3
      Height          =   615
      Left            =   1080
      Top             =   1920
      Width           =   11295
   End
   Begin VB.Line Line7 
      X1              =   1080
      X2              =   12360
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line6 
      X1              =   1080
      X2              =   12360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      X1              =   1080
      X2              =   12360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   12360
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   12360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   12360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   3855
      Left            =   1080
      Top             =   2520
      Width           =   11295
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   495
      Left            =   1080
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   495
      Left            =   1080
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   495
      Left            =   4080
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4080
      X2              =   4080
      Y1              =   1440
      Y2              =   1920
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
      Left            =   4200
      TabIndex        =   14
      Top             =   1560
      Width           =   975
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
      Left            =   1200
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
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
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Line Line11 
      X1              =   1080
      X2              =   12360
      Y1              =   5880
      Y2              =   5880
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
      Left            =   1200
      TabIndex        =   11
      Top             =   6000
      Width           =   2835
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
      Left            =   4200
      TabIndex        =   10
      Top             =   6000
      Width           =   2775
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
      Left            =   7200
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
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
      Left            =   9360
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
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
      Left            =   11400
      TabIndex        =   7
      Top             =   2040
      Width           =   735
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
      Left            =   5400
      TabIndex        =   6
      Top             =   960
      Width           =   18855
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
      Height          =   495
      Left            =   12960
      TabIndex        =   5
      Top             =   9960
      Width           =   2055
   End
   Begin VB.Label rs3 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11280
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "Resultbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs, rss As Object, i, gt, regmid As Integer
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
End Sub

Private Sub Form_Load()
    
Timer1.Enabled = True
    WindowState = 2
'__________________________Typing distinction or first class ....._____________________________________________________________________
    
    Call openconnection
    Set rs = con.Execute("select * from RESULTS where REGNO='" & var.reg & "'and SEM='" & var.sem & "' ")
    
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
    TOT1.Caption = Val(SUB1IA.Caption) + Val(rs!EX1)
    TOT2.Caption = Val(SUB2IA.Caption) + Val(EX2.Caption)
    TOT3.Caption = Val(SUB3IA.Caption) + Val(EX3.Caption)
    TOT4.Caption = Val(SUB4IA.Caption) + Val(EX4.Caption)
    TOT5.Caption = Val(SUB5IA.Caption) + Val(EX5.Caption)
    TOT6.Caption = Val(SUB6IA.Caption) + Val(EX6.Caption)
    TOT7.Caption = Val(SUB7IA.Caption) + Val(EX7.Caption)
     
    'rs.MoveNext
'____________________________DISPLAYING SUBJECT VISE RESULT_______________________________________________________
    If (EX1.Caption = "" Or TOT1.Caption = "") Then
        rs1.Caption = ""
    ElseIf (rs!EX1 = "AB") Then
        rs1.Caption = "AB"
    ElseIf (rs!EX1 < 35) Then
        rs2.Caption = "FAIL"
    ElseIf (TOT1.Caption >= 35 And TOT1.Caption < 50) Then
        rs1.Caption = "PASS"
    ElseIf (TOT1.Caption >= 50 And TOT1.Caption < 60) Then
        rs1.Caption = "SECOND CLASS"
    ElseIf (TOT1.Caption >= 60 And TOT1.Caption < 75) Then
        rs1.Caption = "FIRST CLASS"
    ElseIf (TOT1.Caption >= 75) Then
        rs1.Caption = "DISTINCTION"
    End If
'___________________________________________________________________________________
    If (EX2.Caption = "" Or TOT2.Caption = "") Then
        rs2.Caption = ""
    ElseIf (rs!EX2 = "AB") Then
        rs2.Caption = "AB"
    ElseIf (rs!EX2 < 35) Then
        rs2.Caption = "FAIL"
    ElseIf (TOT2.Caption >= 35 And TOT2.Caption < 50) Then
        rs2.Caption = "PASS"
    ElseIf (TOT2.Caption >= 50 And TOT2.Caption < 60) Then
        rs2.Caption = "SECOND CLASS"
    ElseIf (TOT2.Caption >= 60 And TOT2.Caption < 75) Then
        rs2.Caption = "FIRST CLASS"
    ElseIf (TOT2.Caption >= 75) Then
        rs2.Caption = "DISTINCTION"
    End If
    '___________________________________________________________________________________
    If (EX3.Caption = "" Or TOT3.Caption = "") Then
        rs3.Caption = ""
    ElseIf (rs!EX3 = "AB") Then
        rs3.Caption = "AB"
    ElseIf (rs!EX3 < 35) Then
        rs3.Caption = "FAIL"
    ElseIf (TOT3.Caption >= 35 And TOT3.Caption < 50) Then
        rs3.Caption = "PASS"
    ElseIf (TOT3.Caption >= 50 And TOT3.Caption < 60) Then
        rs3.Caption = "SECOND CLASS"
    ElseIf (TOT3.Caption >= 60 And TOT3.Caption < 75) Then
        rs3.Caption = "FIRST CLASS"
    ElseIf (TOT3.Caption >= 75) Then
        rs3.Caption = "DISTINCTION"
    End If
    '___________________________________________________________________________________
    If (rs!EX4 = "" Or TOT4.Caption = "") Then
        rs4.Caption = ""
    ElseIf (rs!EX4 = "AB") Then
        rs4.Caption = "AB"
    ElseIf (rs!EX4 < 35) Then
        rs4.Caption = "FAIL"
    ElseIf (TOT4.Caption >= 35 And TOT4.Caption < 50) Then
        rs4.Caption = "PASS"
    ElseIf (TOT4.Caption >= 50 And TOT4.Caption < 60) Then
        rs4.Caption = "SECOND CLASS"
    ElseIf (TOT4.Caption >= 60 And TOT4.Caption < 75) Then
        rs4.Caption = "FIRST CLASS"
    ElseIf (TOT4.Caption >= 75) Then
        rs4.Caption = "DISTINCTION"
    End If
    '___________________________________________________________________________________
    If (rs!EX5 = "" Or TOT5.Caption = "") Then
        rs5.Caption = ""
    ElseIf (rs!EX5 = "AB") Then
        rs5.Caption = "AB"
    ElseIf (rs!EX5 < 35) Then
        rs5.Caption = "FAIL"
    ElseIf (TOT5.Caption >= 35 And TOT5.Caption < 50) Then
        rs5.Caption = "PASS"
    ElseIf (TOT5.Caption >= 50 And TOT5.Caption < 60) Then
        rs5.Caption = "SECOND CLASS"
    ElseIf (TOT5.Caption >= 60 And TOT5.Caption < 75) Then
        rs5.Caption = "FIRST CLASS"
    ElseIf (TOT5.Caption >= 75) Then
        rs5.Caption = "DISTINCTION"
    End If
    '___________________________________________________________________________________
    If (rs!EX6 = "" Or TOT6.Caption = "") Then
        rs6.Caption = ""
    ElseIf (rs!EX6 = "AB") Then
        rs6.Caption = "AB"
    ElseIf (rs!EX6 < 35) Then
        rs6.Caption = "FAIL"
    ElseIf (TOT6.Caption >= 35 And TOT6.Caption < 50) Then
        rs6.Caption = "PASS"
    ElseIf (TOT6.Caption >= 50 And TOT6.Caption < 60) Then
        rs6.Caption = "SECOND CLASS"
    ElseIf (TOT6.Caption >= 60 And TOT6.Caption < 75) Then
        rs6.Caption = "FIRST CLASS"
    ElseIf (TOT6.Caption >= 75) Then
        rs6.Caption = "DISTINCTION"
    End If
'___________________________________________________________________________________
    If (rs!EX7 = "" Or TOT7.Caption = "") Then
        rs7.Caption = ""
    ElseIf (rs!EX7 = "AB") Then
        rs7.Caption = "AB"
    ElseIf (rs!EX7 < 35) Then
        rs7.Caption = "FAIL"
    ElseIf (TOT7.Caption >= 35 And TOT7.Caption < 50) Then
        rs7.Caption = "PASS"
    ElseIf (TOT7.Caption >= 50 And TOT7.Caption < 60) Then
        rs7.Caption = "SECOND CLASS"
    ElseIf (TOT7.Caption >= 60 And TOT7.Caption < 75) Then
        rs7.Caption = "FIRST CLASS"
    ElseIf (TOT7.Caption >= 75) Then
        rs7.Caption = "DISTINCTION"
    End If
    '___________________________________________________________________________________
    If (rs!EX8 = "" Or TOT8.Caption = "") Then
        rs8.Caption = ""
     ElseIf (rs!EX8 = "AB") Then
        rs8.Caption = "AB"
    ElseIf (TOT8.Caption = "") Then
        rs8.Caption = ""
    ElseIf (rs!EX8 < 35) Then
        rs8.Caption = "FAIL"
    ElseIf (TOT8.Caption >= 35 And TOT8.Caption < 50) Then
        rs8.Caption = "PASS"
    ElseIf (TOT8.Caption >= 50 And TOT8.Caption < 60) Then
        rs8.Caption = "SECOND CLASS"
    ElseIf (TOT8.Caption >= 60 And TOT8.Caption < 75) Then
        rs8.Caption = "FIRST CLASS"
    ElseIf (TOT8.Caption >= 75) Then
        rs8.Caption = "DISTINCTION"
    End If
'______________________________Entering total____________________________________________________________________
'gt = Val(TOT1.Caption) + Val(TOT2.Caption) + Val(TOT3.Caption) + Val(TOT4.Caption) + Val(TOT5.Caption) + Val(TOT6.Caption) + Val(TOT7.Caption) + Val(TOT8.Caption)
'GRANDTOTALLBL.Caption = GRANDTOTALLBL.Caption & gt

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
    
    Set rss = con.Execute("select * from RESULTS where REGNO='" & var.reg & "'")
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



