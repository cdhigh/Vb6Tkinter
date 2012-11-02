VERSION 5.00
Begin VB.UserControl RKShadeButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  '透明
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   DefaultCancel   =   -1  'True
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   1080
      Picture         =   "RkButton.ctx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   930
      Picture         =   "RkButton.ctx":02C2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   570
      Picture         =   "RkButton.ctx":0584
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   930
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   750
      Picture         =   "RkButton.ctx":0846
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   2
      Top             =   930
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   6  'Inside Solid
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      Top             =   0
      Width           =   975
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   540
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "RkButton.ctx":0B08
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   930
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Timer TimerPaint 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1320
      Top             =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RKShadeButton v1.4"
      Height          =   180
      Left            =   930
      TabIndex        =   5
      Top             =   630
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "RKShadeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'渐变按钮控件RKButton
'
'蓝色炫影  制作
'http://www.rekersoft.cn/
'
'本代码基于GPL V2协议发布。
'您可以自由用于非商业用途。
'请保留此行版权信息，谢谢。

Option Explicit

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
        ByVal heightSrc As Long, ByVal LBLENDFUNCTION As Long) As Boolean
        
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Enum BlendOp
        AC_SRC_OVER = &H0
        AC_SRC_ALPHA = &H1
End Enum

Private Type BLENDFUNCTION
        BlendOp As Byte
        BlendbtnFlags As Byte
        SourceConstantAlpha As Byte
        AlphaFormat As Byte
End Type

Private Type POINTAPI
        x   As Long
        y   As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Enum btnStatus
        btnNormal = 0
        btnHot
        btnPressed
        btbDefault
        btbDefault2
        btnNoDraw
End Enum

Private Const DT_CALCRECT = &H400
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_INTERNAL = &H1000
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0
Private Const DT_RASDISPLAY = 1
Private Const DT_WORDBREAK = &H10
Private Const DT_EDITCONTROL = &H2000

Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event MouseOut()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Const picW = 6
Private Const picH = 32

Private btnDC(4) As Long
Private btnBMP(4) As Long
Private bTime As Integer
Private oldX As Single, oldY As Single
Private rcCaption As RECT
Private rcFocus As RECT

Private defaultFlag As Boolean
Private btnFlag As btnStatus
Private lastStatus As btnStatus
Private bFocused As Boolean

Private szCaption As String
Private nSpeed As Long
Private bEnabled As Boolean
Private bDefault As Boolean

Private nForeColor As Long, nHotColor As Long, nPressedColor As Long


Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        btnFlag = btnPressed
        bTime = 0
        RefreshPicMain
    End If
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        RaiseEvent Click
    End If
End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Static MouseIn As Boolean
        Static pt As POINTAPI
        
        GetCursorPos pt
        ScreenToClient UserControl.hwnd, pt
        oldX = oldX + 1
        MouseIn = (0 <= pt.x) And (pt.x <= UserControl.ScaleWidth) And (0 <= pt.y) And (pt.y <= UserControl.ScaleHeight)
        bTime = 0
        picMain_MouseMove 0, 0, CSng(pt.x), CSng(pt.y)
    End If
End Sub

Private Sub TimerPaint_Timer()
    
    If btnFlag = btnNoDraw Then
        Exit Sub
    ElseIf btnFlag = btnHot Then
        bTime = bTime + nSpeed
        If picMain.ForeColor <> nHotColor Then picMain.ForeColor = nHotColor
    ElseIf btnFlag = btnPressed Then
        bTime = bTime + nSpeed
        If picMain.ForeColor <> nPressedColor Then picMain.ForeColor = nPressedColor
    ElseIf btnFlag = btnNormal Then
        bTime = bTime + 5
        If picMain.ForeColor <> nForeColor Then
            If bEnabled Then
                picMain.ForeColor = nForeColor
            Else
                picMain.ForeColor = vbGrayText
            End If
        End If
    Else
        If defaultFlag = True Then
            bTime = bTime + 15
        Else
            bTime = bTime + 5
        End If
        If picMain.ForeColor <> nForeColor Then picMain.ForeColor = nForeColor
    End If
    
    If bTime >= 255 Then
        ShowTransparency btnDC(btnFlag), 255
        If btnFlag = btbDefault Then
            btnFlag = btbDefault2
            defaultFlag = False
            bTime = 0
        ElseIf btnFlag = btbDefault2 Then
            btnFlag = btbDefault
            bTime = 0
        Else
            btnFlag = btnNoDraw
        End If
        DrawText picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK
        
        picMain.Picture = picMain.Image
        If bFocused Then
            picMain.ForeColor = vbBlack
            DrawFocusRect picMain.hdc, rcFocus
        End If
    Else
        ShowTransparency btnDC(btnFlag), bTime
        DrawText picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK
        If bFocused Then
            picMain.ForeColor = vbBlack
            DrawFocusRect picMain.hdc, rcFocus
        End If
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    bDefault = Ambient.DisplayAsDefault
    RefreshPicMain
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    Else
        btnFlag = btnNormal
    End If
    bTime = 0
End Sub

Private Sub picMain_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
    bFocused = True
End Sub

Private Sub UserControl_ExitFocus()
    bFocused = False
End Sub

Private Sub UserControl_InitProperties()
    nSpeed = 20
    szCaption = Replace(UserControl.Extender.Name, "RKShade", "")
    bDefault = Ambient.DisplayAsDefault
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    End If
    PropertyChanged
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    szCaption = PropBag.ReadProperty("Caption", Replace(UserControl.Extender.Name, "RKShade", ""))
    bEnabled = PropBag.ReadProperty("Enabled", True)
    nSpeed = PropBag.ReadProperty("Speed", "20")
    nForeColor = PropBag.ReadProperty("ForeColor", 0)
    nHotColor = PropBag.ReadProperty("HotColor", 0)
    nPressedColor = PropBag.ReadProperty("PressedColor", 0)
    bDefault = Ambient.DisplayAsDefault
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    End If
    Enabled = bEnabled
    
    lblCaption.Caption = szCaption
    TimerPaint.Enabled = Ambient.UserMode
    
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", szCaption, Replace(UserControl.Extender.Name, "RKShade", "")
    PropBag.WriteProperty "Enabled", bEnabled, True
    PropBag.WriteProperty "Speed", nSpeed, "20"
    PropBag.WriteProperty "ForeColor", nForeColor, 0
    PropBag.WriteProperty "HotColor", nHotColor, 0
    PropBag.WriteProperty "PressedColor", nPressedColor, 0
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = szCaption
End Property

Public Property Let Caption(szCap As String)
    szCaption = szCap
    PropertyChanged "Caption"
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    Paint
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(en As Boolean)
    If bDefault Then
        btnFlag = btbDefault
        defaultFlag = True
    Else
        btnFlag = btnNormal
    End If
    bTime = 0
    bEnabled = en
    UserControl.Enabled = en
    picMain.Enabled = en
    PropertyChanged "Enabled"
    
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Paint
End Property

Public Property Get Speed() As Long
    Speed = nSpeed
End Property

Public Property Let Speed(nSpd As Long)
    nSpeed = nSpd
    PropertyChanged "Speed"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = nForeColor
End Property

Public Property Let ForeColor(clr As OLE_COLOR)
    nForeColor = clr
    lblCaption.ForeColor = clr
    PropertyChanged "ForeColor"
    If bEnabled Then
        picMain.ForeColor = nForeColor
    Else
        picMain.ForeColor = vbGrayText
    End If
    Paint
End Property

Public Property Get HotColor() As OLE_COLOR
    HotColor = nHotColor
End Property

Public Property Let HotColor(clr As OLE_COLOR)
    nHotColor = clr
    PropertyChanged "HotColor"
End Property

Public Property Get PressedColor() As OLE_COLOR
    PressedColor = nPressedColor
End Property

Public Property Let PressedColor(clr As OLE_COLOR)
    nPressedColor = clr
    PropertyChanged "PressedColor"
End Property

Private Sub UserControl_Initialize()
    Dim i As Byte
    For i = 0 To 4
        btnDC(i) = CreateCompatibleDC(picMain.hdc)
    Next i
    rcFocus.Left = 3
    rcFocus.Top = 3
End Sub

Private Sub Paint()
    Dim i As Long
    Dim W As Long
    Dim H As Long

    W = UserControl.ScaleWidth
    H = UserControl.ScaleHeight
    picMain.Width = W
    picMain.Height = H
    For i = 0 To 4
        With picBtn(i)
            btnBMP(i) = CreateCompatibleBitmap(.hdc, W, H)
            Call SelectObject(btnDC(i), btnBMP(i))
            StretchBlt btnDC(i), 2, 2, W - 4, H - 4, .hdc, 2, 2, picW - 4, picH - 4, vbSrcCopy
            BitBlt btnDC(i), 0, 0, 2, 2, .hdc, 0, 0, vbSrcCopy '左上角
            BitBlt btnDC(i), W - 2, 0, 2, 2, .hdc, picW - 2, 0, vbSrcCopy '右上角
            BitBlt btnDC(i), 0, H - 2, 2, 2, .hdc, 0, picH - 2, vbSrcCopy '左下角
            BitBlt btnDC(i), W - 2, H - 2, 2, 2, .hdc, picW - 2, picH - 2, vbSrcCopy '右下角
            StretchBlt btnDC(i), 2, 0, W - 4, 2, .hdc, 2, 0, picW - 4, 2, vbSrcCopy         '上边框
            StretchBlt btnDC(i), 2, H - 2, W - 4, 2, .hdc, 2, picH - 2, picW - 4, 2, vbSrcCopy '下边框
            StretchBlt btnDC(i), 0, 2, 2, H - 4, .hdc, 0, 2, 2, picH - 4, vbSrcCopy  '左边框
            StretchBlt btnDC(i), W - 2, 2, 2, H - 4, .hdc, picW - 2, 2, 2, picH - 4, vbSrcCopy '右边框
            If i >= 3 Then
                BitBlt btnDC(i), 2, 2, W - 4, 1, btnDC(i), 2, 1, vbSrcCopy '上边框
                BitBlt btnDC(i), 2, H - 3, W - 4, 1, btnDC(i), 2, H - 2, vbSrcCopy '下边框
                BitBlt btnDC(i), 2, 2, 1, H - 4, btnDC(i), 1, 2, vbSrcCopy   '左边框
                BitBlt btnDC(i), W - 3, 2, 1, H - 4, btnDC(i), W - 2, 2, vbSrcCopy  '右边框
            End If
            If i = 4 Then
                BitBlt btnDC(4), 2, 3, W - 4, 1, btnDC(4), 2, 1, vbSrcCopy       '上边框
                BitBlt btnDC(4), 2, H - 4, W - 4, 1, btnDC(4), 2, H - 2, vbSrcCopy '下边框
                BitBlt btnDC(4), 3, 2, 1, H - 4, btnDC(4), 1, 2, vbSrcCopy    '左边框
                BitBlt btnDC(4), W - 4, 2, 1, H - 4, btnDC(4), W - 2, 2, vbSrcCopy   '右边框
            End If
        End With
    Next i

    BitBlt picMain.hdc, 0, 0, W, H, btnDC(0), 0, 0, vbSrcCopy

    Call DrawText(picMain.hdc, szCaption, lstrlen(szCaption), rcCaption, DT_CENTER Or DT_EDITCONTROL Or DT_WORDBREAK)
    
    RefreshPicMain
    For i = 0 To 4
        DeleteObject btnBMP(i)
    Next i
End Sub

Private Sub UserControl_Resize()
    Static j As Double
    Static i As Long
    
    j = picMain.TextWidth(szCaption) / UserControl.ScaleWidth
    If Int(j) <> j Then
        i = (Int(j) + 1) * picMain.TextHeight("1")
    End If
    
    rcCaption.Right = UserControl.ScaleWidth
    rcCaption.Top = (UserControl.ScaleHeight - i) \ 2
    rcCaption.Bottom = rcCaption.Top + i
    
    rcFocus.Right = UserControl.ScaleWidth - 3
    rcFocus.Bottom = UserControl.ScaleHeight - 3
    
    Paint
End Sub

Private Sub UserControl_Terminate()
    Dim i As Byte
    For i = 0 To 2
        DeleteObject btnBMP(i)
        DeleteDC btnDC(i)
    Next i
End Sub

Private Sub RefreshPicMain()
    
    If bFocused Then
        picMain.ForeColor = vbBlack
        DrawFocusRect picMain.hdc, rcFocus
    End If
    picMain.Picture = picMain.Image
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    oldX = x: oldY = y
    RaiseEvent MouseDown(Button, Shift, x, y)
    If Button = 1 Then
        bTime = 0
        btnFlag = btnPressed
        lastStatus = btnPressed
        RefreshPicMain
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If oldX = x And oldY = y Then Exit Sub

    Static MouseIn As Boolean
    
    oldX = x: oldY = y

    MouseIn = (0 <= x) And (x <= UserControl.ScaleWidth) And (0 <= y) And (y <= UserControl.ScaleHeight)
    If MouseIn Then
        'ReleaseCapture
        RaiseEvent MouseMove(Button, Shift, x, y)
        SetCapture picMain.hwnd
        
        If Button = 0 Then
            If lastStatus <> btnHot Then
                bTime = 0
                btnFlag = btnHot
                lastStatus = btnHot
            End If
        ElseIf Button = 1 Then
            If lastStatus <> btnPressed Then
                bTime = 0
                btnFlag = btnPressed
                lastStatus = btnPressed
                RefreshPicMain
            End If
        End If
    Else
        If Button = 0 Then ReleaseCapture
        RaiseEvent MouseOut
        If lastStatus = btnHot Or lastStatus = btnPressed Then
            bTime = 0
            If bDefault Then
                btnFlag = btbDefault
                defaultFlag = True
                lastStatus = btbDefault
            Else
                btnFlag = btnNormal
                lastStatus = btnNormal
            End If
            RefreshPicMain
        End If
    End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static MouseIn As Boolean
    
    MouseIn = (0 <= x) And (x <= UserControl.ScaleWidth) And (0 <= y) And (y <= UserControl.ScaleHeight)
    ReleaseCapture
    RaiseEvent MouseUp(Button, Shift, x, y)
    bTime = 0
    If MouseIn Then
        btnFlag = btnHot
    Else
        If bDefault Then
            btnFlag = btbDefault
            defaultFlag = True
            lastStatus = btbDefault
        Else
            btnFlag = btnNormal
            lastStatus = btnNormal
        End If
    End If
    RefreshPicMain
    oldX = oldX + 1
End Sub

Private Sub ShowTransparency(cSrc As Long, ByVal nLevel As Byte)
    Dim LrProps As Long
    With picMain
        .Cls
        LrProps = nLevel * &H10000

        AlphaBlend .hdc, 0, 0, .ScaleWidth, .ScaleHeight, _
            cSrc, 0, 0, .ScaleWidth, .ScaleHeight, LrProps
    End With
End Sub

