VERSION 5.00
Begin VB.UserControl VertToggleSwitch 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   Picture         =   "VertToggleSwitch.ctx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   4230
   ToolboxBitmap   =   "VertToggleSwitch.ctx":05B9
   Begin VB.PictureBox p6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3540
      Picture         =   "VertToggleSwitch.ctx":08CB
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   5
      Top             =   2550
      Width           =   270
   End
   Begin VB.PictureBox p5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3240
      Picture         =   "VertToggleSwitch.ctx":0D26
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   4
      Top             =   2550
      Width           =   270
   End
   Begin VB.PictureBox p4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   2625
      Picture         =   "VertToggleSwitch.ctx":1182
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   2535
      Width           =   375
   End
   Begin VB.PictureBox p3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   2205
      Picture         =   "VertToggleSwitch.ctx":169C
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   2535
      Width           =   375
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1380
      Picture         =   "VertToggleSwitch.ctx":1B77
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   2505
      Width           =   480
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   825
      Picture         =   "VertToggleSwitch.ctx":2144
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   2505
      Width           =   480
   End
End
Attribute VB_Name = "VertToggleSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'by Ken Foster April 2010
'Please use and abuse
'Copyrights = none

Public Enum vSize
    Small = 0
    Med = 1
    Tiny = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Small, Med, Tiny
#End If

Public Enum vStyle
    Momentary = 0
    Toggle = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Momentary, Toggle
#End If

Private bButton                    As Integer
Private sShift                     As Integer
Private posX                       As Single
Private posY                       As Single

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private Const m_def_ButState       As Boolean = False
Private Const m_def_ButSize        As Integer = 1
Private Const m_def_ButStyle       As Integer = 0
Private Const m_def_LED            As Boolean = False
Private Const m_def_Enabled   As Boolean = True

Private m_ButState                 As Boolean
Private m_ButSize                  As Integer
Private m_ButStyle                 As vStyle
Private m_LED                      As Boolean
Private m_Enabled               As Boolean
    
Private Sub UserControl_Initialize()
    LED = m_def_LED
    ButSize = m_def_ButSize
    ButStyle = m_def_ButStyle
    ButState = m_def_ButState
    Enabled = m_def_Enabled
End Sub
 
Private Sub UserControl_InitProperties()
    m_LED = m_def_LED
    m_ButSize = m_def_ButSize
    m_ButStyle = m_def_ButStyle
    m_ButState = m_def_ButState
    m_Enabled = m_def_Enabled
End Sub
   
Private Sub UserControl_Resize()
    With UserControl
    Select Case m_ButSize
        Case 0    'small
                .Picture = p3.Picture
                .Width = p3.Width
                .Height = p3.Height
        Case 1    'med
                .Picture = p1.Picture
                .Width = p1.Width
                .Height = p1.Height
        Case 2   'tiny
                .Picture = p5.Picture
                .Width = p5.Width
                .Height = p5.Height
    End Select
    End With
End Sub
    
Private Sub UserControl_MouseDown(Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)
    
    If m_ButStyle = Toggle Then    '----------------------------toggle
    m_ButState = Not m_ButState   'toggle on/off
    DrawButton
    Else      '-------------------------------------------------momentary
    Select Case m_ButSize
        Case 0    'small
            UserControl.Picture = p4.Picture
        Case 1    'med
            UserControl.Picture = p2.Picture
        Case 2   'tiny
            UserControl.Picture = p6.Picture
    End Select
End If
RaiseEvent Click
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
   
Private Sub UserControl_MouseUp(Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    'if style is toggle the exit
    If m_ButStyle = Toggle Then
        Exit Sub
    End If
    Select Case m_ButSize    'if style is momentary
        Case 0    'small
            UserControl.Picture = p3.Picture
        Case 1    'med
            UserControl.Picture = p1.Picture
        Case 2    'tiny
            UserControl.Picture = p5.Picture
    End Select
End Sub

Private Sub UserControl_Click()
    If m_ButStyle = Toggle Then
        RaiseEvent Click
    End If
End Sub
    
Private Sub UserControl_DblClick()
    UserControl_MouseDown bButton, sShift, posX, posY
End Sub
    
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ButSize = .ReadProperty("ButSize", m_def_ButSize)
        ButState = .ReadProperty("ButState", m_def_ButState)
        ButStyle = .ReadProperty("ButStyle", m_def_ButStyle)
        LED = .ReadProperty("LED", m_def_LED)
        Enabled = .ReadProperty("Enabled", m_def_Enabled)
    End With
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ButSize", m_ButSize, m_def_ButSize
        .WriteProperty "ButState", m_ButState, m_def_ButState
        .WriteProperty "ButStyle", m_ButStyle, m_def_ButStyle
        .WriteProperty "LED", m_LED, m_def_LED
        .WriteProperty "Enabled", m_Enabled, m_def_Enabled
    End With
End Sub
        
Private Sub DrawButton()
    With UserControl
    'style equals momentary
    If m_ButStyle = Momentary Then
        Select Case m_ButSize
            Case 0    'small
                .Picture = p3.Picture
            Case 1    'med
                .Picture = p1.Picture
            Case 2   'tiny
                .Picture = p5.Picture
        End Select
    Else
        'style equals toggle
        If m_ButState Then    'LED = green
        Select Case m_ButSize
            Case 0    'small
                .Picture = p4.Picture
                If m_LED Then
                    UserControl.Line (130, 55)-(250, 55), vbGreen
                    UserControl.Line (130, 70)-(250, 70), vbGreen
                End If
            Case 1    'med
                UserControl.Picture = p2.Picture
                If m_LED Then
                    UserControl.Line (150, 55)-(350, 55), vbGreen
                    UserControl.Line (150, 70)-(350, 70), vbGreen
                End If
            Case 2    'tiny
                UserControl.Picture = p6.Picture
                If m_LED Then
                    UserControl.Line (80, 25)-(200, 25), vbGreen
                    UserControl.Line (80, 45)-(200, 45), vbGreen
                End If
        End Select
        Else                         'LED = dark green
        Select Case m_ButSize
            Case 0    'small
                .Picture = p3.Picture
                If m_LED Then
                    UserControl.Line (130, 53)-(250, 53), &H6C00&
                    UserControl.Line (130, 68)-(250, 68), &H6C00&
                End If
            Case 1    'med
                .Picture = p1.Picture
                If m_LED Then
                    UserControl.Line (150, 55)-(350, 55), &H6C00&
                    UserControl.Line (150, 70)-(350, 70), &H6C00&
                End If
            Case 2    'tiny
                .Picture = p5.Picture
                If m_LED Then
                    UserControl.Line (80, 25)-(200, 25), &H6C00&
                    UserControl.Line (80, 45)-(200, 45), &H6C00&
                End If
        End Select
    End If
End If
End With
End Sub

Public Property Get ButSize() As vSize
    ButSize = m_ButSize
End Property

Public Property Let ButSize(ByVal NewButSize As vSize)
    
    m_ButSize = NewButSize
    With UserControl
    Select Case m_ButSize
        Case 0    'small
                .Picture = p3.Picture
                .Width = p3.Width
                .Height = p3.Height
        Case 1    'med
                .Picture = p1.Picture
                .Width = p1.Width
                .Height = p1.Height
        Case 2   'tiny
                .Picture = p5.Picture
                .Width = p5.Width
                .Height = p5.Height
    End Select
    End With
    PropertyChanged "ButSize"
End Property

Public Property Get ButState() As Boolean
Attribute ButState.VB_Description = "Effective only in Toggle mode"
    ButState = m_ButState
End Property

Public Property Let ButState(ByVal NewButState As Boolean)
    
    m_ButState = NewButState
    PropertyChanged "ButState"
    DrawButton
End Property

Public Property Get ButStyle() As vStyle
    ButStyle = m_ButStyle
End Property

Public Property Let ButStyle(ByVal NewButStyle As vStyle)
    
    m_ButStyle = NewButStyle
    'in momentary style always set to false
    If m_ButStyle = Momentary Then
        m_ButState = False
    End If
    PropertyChanged "ButStyle"
    DrawButton
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get LED() As Boolean
Attribute LED.VB_Description = "Effective only in Toggle mode"
    LED = m_LED
End Property

Public Property Let LED(ByVal NewLED As Boolean)
    
    m_LED = NewLED
    If m_ButStyle = Momentary Then
        m_LED = False
    End If
    PropertyChanged "LED"
    DrawButton
End Property

