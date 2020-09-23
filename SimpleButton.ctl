VERSION 5.00
Begin VB.UserControl SimpleButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   MaskColor       =   &H00FF00FF&
   Picture         =   "SimpleButton.ctx":0000
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   Begin VB.PictureBox pDown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1545
      Picture         =   "SimpleButton.ctx":06F2
      ScaleHeight     =   750
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   1515
      Width           =   1140
   End
   Begin VB.PictureBox pUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   210
      Picture         =   "SimpleButton.ctx":0DC7
      ScaleHeight     =   750
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   1515
      Width           =   1140
   End
End
Attribute VB_Name = "SimpleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_DEFAULT = DT_CENTER Or DT_VCENTER

Public Enum Alignment
    [CenterCenter]
    [CenterTop]
    [CenterBottom]
    [LeftCenter]
    [LeftTop]
    [LeftBottom]
    [RightCenter]
    [RightTop]
    [RightBottom]
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private TxtRect As RECT
Private m_textoffset As Boolean
Private m_Caption As String
Private m_Align As Alignment
Private PrevColor As Long
Private m_PicDown As Picture
Private m_PicUp As Picture

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_DblClick()
    UserControl_MouseDown 0, 0, 0, 0
End Sub
    
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    Set PicDown = pDown.Picture
    Set PicUp = pUp.Picture
    m_Caption = "Button"
    Textoffset = True
End Sub
    
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    UserControl.Picture = pDown.Picture
    Cls
    If Trim(m_Caption) <> "" Then    'draw text with offset
        If Textoffset = True Then
           TxtRect.Left = 1
           TxtRect.Top = 1
           TxtRect.Bottom = ScaleHeight + 1
           TxtRect.Right = ScaleWidth + 1
        Else
           TxtRect.Left = 0
           TxtRect.Top = 0
           TxtRect.Bottom = ScaleHeight
           TxtRect.Right = ScaleWidth
        End If
        PrevColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, m_Caption, Len(m_Caption), TxtRect, GetAlign(m_Align) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = PrevColor
    End If
    Refresh
    RaiseEvent Click
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
    
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    UserControl.Picture = pUp.Picture
    Cls
    If Trim(m_Caption) <> "" Then      'draw text
        TxtRect.Left = 0
        TxtRect.Top = 0
        TxtRect.Bottom = ScaleHeight
        TxtRect.Right = ScaleWidth
        PrevColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, m_Caption, Len(m_Caption), TxtRect, GetAlign(m_Align) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = PrevColor
    End If
    Refresh
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
    
Private Sub UserControl_Resize()
    UserControl.ScaleMode = 1  'set to twips for sizing
    UserControl.Width = pUp.Width
    UserControl.Height = pUp.Height
    UserControl.ScaleMode = 3   'set to pixels for text
    
    Cls
    TxtRect.Left = 0  'set rectangle size for alignment of text
    TxtRect.Top = 0
    TxtRect.Right = ScaleWidth
    TxtRect.Bottom = ScaleHeight
    
    If Trim(m_Caption) <> "" Then     'draw text
        TxtRect.Left = 0
        TxtRect.Top = 0
        TxtRect.Bottom = ScaleHeight
        TxtRect.Right = ScaleWidth
        PrevColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, m_Caption, Len(m_Caption), TxtRect, GetAlign(m_Align) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = PrevColor
    End If
    Refresh
End Sub
    
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim X As String
    m_Align = PropBag.ReadProperty("Align", DT_DEFAULT)
    m_Caption = PropBag.ReadProperty("Caption", "Command")
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set PicDown = PropBag.ReadProperty("PicDown", Nothing)
    Set PicUp = PropBag.ReadProperty("PicUp", Nothing)
    m_textoffset = PropBag.ReadProperty("Textoffset", True)
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, "Command")
    Call PropBag.WriteProperty("Align", m_Align, DT_DEFAULT)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("PicDown", m_PicDown, Nothing)
    Call PropBag.WriteProperty("PicUp", m_PicUp, Nothing)
    Call PropBag.WriteProperty("Textoffset", m_textoffset, True)
End Sub
    
Public Property Get Caption() As String
    Caption = m_Caption
End Property
    
Public Property Let Caption(ByVal newCaption As String)
    m_Caption = newCaption
    PropertyChanged "Caption"
    UserControl_Resize
End Property
    
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
    
Public Property Let Enabled(ByVal newEnabled As Boolean)
    UserControl.Enabled() = newEnabled
    PropertyChanged "Enabled"
    UserControl_Resize
End Property
    
Public Property Get Align() As Alignment
    Align = m_Align
End Property
    
Public Property Let Align(ByVal newAlign As Alignment)
    m_Align = newAlign
    PropertyChanged "Align"
    UserControl_Resize
End Property
    
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
    
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property
    
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property
    
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl_Resize
End Property
    
Public Property Get PicDown() As Picture
    Set PicDown = m_PicDown
End Property

Public Property Set PicDown(ByVal NewValue As Picture)
    Set m_PicDown = NewValue
    pDown.Picture = NewValue
    PropertyChanged "PicDown"
End Property

Public Property Get PicUp() As Picture
    Set PicUp = m_PicUp
End Property

Public Property Set PicUp(ByVal NewValue As Picture)
    UserControl.MaskPicture = Nothing
    Set m_PicUp = NewValue
    pUp.Picture = NewValue
    UserControl.Picture = NewValue
    UserControl.MaskPicture = UserControl.Image
    UserControl.MaskColor = UserControl.Point(0, 0)  'get transparent color
    PropertyChanged "PicUp"
    UserControl_Resize
End Property

Public Property Get Textoffset() As Boolean
Attribute Textoffset.VB_Description = "If false, text is stationary. If true, text moves with mouse click."
    Textoffset = m_textoffset
End Property
    
Public Property Let Textoffset(ByVal newTextoffset As Boolean)
    m_textoffset = newTextoffset
    PropertyChanged "Textoffset"
End Property

Private Function GetAlign(ByVal Alng As Alignment) As Long
    Select Case Alng
        Case 0: GetAlign = DT_CENTER Or DT_VCENTER
        Case 1: GetAlign = DT_CENTER Or DT_TOP
        Case 2: GetAlign = DT_CENTER Or DT_BOTTOM
        Case 3: GetAlign = DT_LEFT Or DT_VCENTER
        Case 4: GetAlign = DT_LEFT Or DT_TOP
        Case 5: GetAlign = DT_LEFT Or DT_BOTTOM
        Case 6: GetAlign = DT_RIGHT Or DT_VCENTER
        Case 7: GetAlign = DT_RIGHT Or DT_TOP
        Case 8: GetAlign = DT_RIGHT Or DT_BOTTOM
    End Select
End Function
    
