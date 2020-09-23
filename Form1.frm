VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Toggle Switch Demo"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin Project1.SimpleButton SimpleButton2 
      Height          =   330
      Left            =   1785
      TabIndex        =   23
      Top             =   3615
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      Caption         =   "Subtract "
      Align           =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicDown         =   "Form1.frx":08CA
      PicUp           =   "Form1.frx":0B83
      Textoffset      =   0   'False
   End
   Begin Project1.SimpleButton SimpleButton1 
      Height          =   750
      Left            =   90
      TabIndex        =   21
      Top             =   3450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   1323
      Caption         =   "Add    "
      Align           =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      PicDown         =   "Form1.frx":0E39
      PicUp           =   "Form1.frx":151E
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch6 
      Height          =   270
      Left            =   1935
      TabIndex        =   19
      Top             =   3015
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   476
      ButSize         =   2
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch5 
      Height          =   375
      Left            =   1935
      TabIndex        =   18
      Top             =   2610
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   661
      ButSize         =   0
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch4 
      Height          =   480
      Left            =   1935
      TabIndex        =   17
      Top             =   2100
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   847
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch3 
      Height          =   270
      Left            =   285
      TabIndex        =   14
      Top             =   3030
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   476
      ButSize         =   2
      ButStyle        =   1
      LED             =   -1  'True
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch2 
      Height          =   375
      Left            =   285
      TabIndex        =   13
      Top             =   2625
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   661
      ButSize         =   0
      ButStyle        =   1
      LED             =   -1  'True
   End
   Begin Project1.HorzToggleSwitch HorzToggleSwitch1 
      Height          =   480
      Left            =   270
      TabIndex        =   12
      Top             =   2115
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   847
      ButStyle        =   1
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch6 
      Height          =   510
      Left            =   2700
      TabIndex        =   9
      Top             =   315
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   900
      ButSize         =   2
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch5 
      Height          =   510
      Left            =   1125
      TabIndex        =   8
      Top             =   330
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   900
      ButSize         =   2
      ButState        =   -1  'True
      ButStyle        =   1
      LED             =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1845
      TabIndex        =   7
      Top             =   1515
      Width           =   1020
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch4 
      Height          =   690
      Left            =   2295
      TabIndex        =   6
      Top             =   315
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1217
      ButSize         =   0
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch3 
      Height          =   885
      Left            =   1815
      TabIndex        =   4
      Top             =   315
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1561
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch2 
      Height          =   690
      Left            =   720
      TabIndex        =   2
      Top             =   330
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1217
      ButSize         =   0
      ButStyle        =   1
      LED             =   -1  'True
   End
   Begin Project1.VertToggleSwitch VertToggleSwitch1 
      Height          =   885
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1561
      ButStyle        =   1
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   22
      Top             =   3690
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      Height          =   240
      Left            =   2010
      TabIndex        =   20
      Top             =   3285
      Width           =   360
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   75
      Y2              =   3465
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MOMENTARY"
      Height          =   210
      Left            =   1890
      TabIndex        =   16
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOGGLE"
      Height          =   225
      Left            =   420
      TabIndex        =   15
      Top             =   45
      Width           =   750
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " LED False"
      Height          =   450
      Left            =   255
      TabIndex        =   11
      Top             =   1485
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OFF"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1110
      TabIndex        =   10
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2115
      TabIndex        =   5
      Top             =   1275
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OFF"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OFF"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1230
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X     As Integer
Private addx As Integer

Private Sub Form_Load()

   If VertToggleSwitch1.ButState Then
      Label1.Caption = "ON"
     Else
      Label1.Caption = "OFF"
   End If
   If VertToggleSwitch2.ButState Then
      Label2.Caption = "ON"
     Else
      Label2.Caption = "OFF"
   End If
   If VertToggleSwitch5.ButState Then
      Label4.Caption = "ON"
     Else
      Label4.Caption = "OFF"
   End If

End Sub

Private Sub HorzToggleSwitch6_Click()
   Unload Me
End Sub

Private Sub SimpleButton1_Click()
addx = addx + 1
Label9.Caption = addx
End Sub

Private Sub SimpleButton2_Click()
   addx = addx - 1
   Label9.Caption = addx
End Sub

Private Sub VertToggleSwitch1_Click()

   If VertToggleSwitch1.ButState Then
      Label1.Caption = "ON"
     Else
      Label1.Caption = "OFF"
   End If

End Sub

Private Sub VertToggleSwitch2_Click()

   If VertToggleSwitch2.ButState Then
      Label2.Caption = "ON"
     Else
      Label2.Caption = "OFF"
   End If

End Sub

Private Sub VertToggleSwitch3_Click()

   X = X + 1
   Label3.Caption = X

End Sub

Private Sub VertToggleSwitch4_Click()

   X = X - 1
   Label3.Caption = X

End Sub

Private Sub VertToggleSwitch5_Click()

   If VertToggleSwitch5.ButState Then
      Label4.Caption = "ON"
     Else
      Label4.Caption = "OFF"
   End If

End Sub

Private Sub VertToggleSwitch6_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)

   Text1.Text = "HELLO"

End Sub

Private Sub VertToggleSwitch6_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

   Text1.Text = ""

End Sub

