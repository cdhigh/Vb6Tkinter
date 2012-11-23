VERSION 5.00
Begin VB.UserControl xpcmdbutton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "xpcmdbutton.ctx":0000
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   480
      Picture         =   "xpcmdbutton.ctx":0312
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   960
      Picture         =   "xpcmdbutton.ctx":0390
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   720
      Picture         =   "xpcmdbutton.ctx":0625
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   240
      Picture         =   "xpcmdbutton.ctx":08C9
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "xpcmdbutton.ctx":0B64
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   120
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "xpcmdbutton"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "xpcmdbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Teh Ming Han (teh_minghan@hotmail.com)
'cdhigh modified in 2012.11.23,
'   1. substitute control 'picture' for 'pictureclip'
'   2. add some features
'   3. fix some bugs
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Enum State_b
    Normal_ = 0
    Default_ = 1
End Enum

Dim m_State As State_b
Dim m_Font As Font

Const m_Def_State = State_b.Normal_

Private Type POINT_API
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim s As Integer
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private rcFocus As RECT
Private m_left As Integer
Private m_bFocused As Boolean
Private m_shortcutKey As String

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_left = Button
    If m_left = 1 Then Call UserControl_MouseDown(Button, Shift, x, y)
    ' Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseMove(Button, Shift, x, y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    
    GetCursorPos pnt
    If WindowFromPoint(pnt.x, pnt.y) <> UserControl.hwnd Then
        Timer1.Enabled = False
        RaiseEvent MouseOut
        If m_bFocused Then
            make_xpbutton 4
        Else
            statevalue_pic
        End If
    End If
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    If m_left = 1 Then RaiseEvent Click
    ' RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
    m_bFocused = True
    make_xpbutton 4
End Sub

Private Sub UserControl_ExitFocus()
    m_bFocused = False
    If Not Timer1.Enabled Then
        statevalue_pic
    End If
End Sub

Private Sub UserControl_Initialize()
    rcFocus.Left = 3
    rcFocus.Top = 3
    m_bFocused = False
    statevalue_pic
End Sub

Private Sub UserControl_InitProperties()
    m_State = m_Def_State
    Enabled = True
    Caption = Ambient.DisplayName
    Set Font = UserControl.Ambient.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_bFocused And KeyAscii = vbKeySpace Then
        make_xpbutton 1
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If m_bFocused Then
        If KeyCode = vbKeySpace Then
            RaiseEvent Click
            make_xpbutton 4
        ElseIf ((Shift And vbAltMask) <> 0) And Len(m_shortcutKey) Then
            If LCase(Chr(KeyCode)) = m_shortcutKey Then
                RaiseEvent Click
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_left = Button
    If m_left = 1 Then
        RaiseEvent MouseDown(Button, Shift, x, y)
        make_xpbutton 1
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer1.Enabled = True
    If x >= 0 And y >= 0 And _
        x <= UserControl.ScaleWidth And y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, x, y)
        If Button = vbLeftButton Then
            make_xpbutton 1
        Else
            make_xpbutton 3
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    statevalue_pic
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State = PropBag.ReadProperty("State", m_Def_State)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    statevalue_pic
    If Enabled = True Then
        lbl.ForeColor = vbBlack
    Else
        lbl.ForeColor = RGB(161, 161, 146)
    End If
End Property

Private Sub UserControl_Resize()
    rcFocus.Right = UserControl.ScaleWidth - 3
    rcFocus.Bottom = UserControl.ScaleHeight - 3
    statevalue_pic
    lbl.Top = (UserControl.ScaleHeight - lbl.Height) / 2
    lbl.Left = (UserControl.ScaleWidth - lbl.Width) / 2
End Sub

Private Sub UserControl_Show()
    statevalue_pic
End Sub

Private Sub UserControl_Terminate()
    statevalue_pic
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("State", m_State, m_Def_State)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lbl.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
End Sub

Public Property Get State() As State_b
    State = m_State
End Property

Public Property Let State(ByVal vNewValue As State_b)
    m_State = vNewValue
    PropertyChanged "State"
    statevalue_pic
End Property

Private Sub statevalue_pic()
    If State = Default_ Then
        s = 4
    ElseIf State = Normal_ Then
        s = 0
    End If
    
    If UserControl.Enabled = True Then
        make_xpbutton s
Else:
        make_xpbutton 2
    End If
End Sub

Private Sub make_xpbutton(z As Integer)
    UserControl.ScaleMode = 3                                                   'Draw in pixels
    Dim brx, bry, bw, bh As Integer
    'Short cuts
    brx = UserControl.ScaleWidth - 3                                            'right x
    bry = UserControl.ScaleHeight - 3                                           'right y
    bw = UserControl.ScaleWidth - 6                                             'border width - corners width
    bh = UserControl.ScaleHeight - 6                                            'border height - corners height
    'Draws button
    'Goes clockwise first for corners(first four)
    'followed by borders(next four) and center(last step).
    UserControl.PaintPicture pc(z).Picture, 0, 0, 3, 3, 0, 0, 3, 3
    UserControl.PaintPicture pc(z).Picture, brx, 0, 3, 3, 15, 0, 3, 3
    UserControl.PaintPicture pc(z).Picture, brx, bry, 3, 3, 15, 18, 3, 3
    UserControl.PaintPicture pc(z).Picture, 0, bry, 3, 3, 0, 18, 3, 3
    UserControl.PaintPicture pc(z).Picture, 3, 0, bw, 3, 3, 0, 12, 3
    UserControl.PaintPicture pc(z).Picture, brx, 3, 3, bh, 15, 3, 3, 15
    UserControl.PaintPicture pc(z).Picture, 0, 3, 3, bh, 0, 3, 3, 15
    UserControl.PaintPicture pc(z).Picture, 3, bry, bw, 3, 3, 18, 12, 3
    UserControl.PaintPicture pc(z).Picture, 3, 3, bw, bh, 3, 3, 12, 15
    
    If m_bFocused Then
        DrawFocusRect UserControl.hdc, rcFocus
    End If
    
End Sub

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    Dim idx As Long, ascii As Long
    lbl.Caption = vNewCaption
    PropertyChanged "Caption"
    
    'Key shortcut
    idx = InStr(1, vNewCaption, "&")
    If idx > 0 Then
        ascii = Asc(Mid$(vNewCaption, idx + 1, 1))
        If (ascii >= Asc("a") And ascii <= Asc("z")) Or (ascii >= Asc("A") And ascii <= Asc("Z")) Then
            m_shortcutKey = LCase(Chr(ascii))
        Else
            m_shortcutKey = ""
        End If
    Else
        m_shortcutKey = ""
    End If
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

