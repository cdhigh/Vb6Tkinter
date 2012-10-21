VERSION 5.00
Begin VB.UserControl GridOcx 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   KeyPreview      =   -1  'True
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ToolboxBitmap   =   "GridOcx.ctx":0000
End
Attribute VB_Name = "GridOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'查询API的网址http://vbworld.sxnw.gov.cn/vbapi/
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'XP
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_SHOWDROPDOWN = &H14F

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
'download by http://www.codefans.net
Private Const DFC_BUTTON As Long = &H4

Private Const DFCS_FLAT As Long = &H4000
Private Const DFCS_BUTTONCHECK As Long = &H0
Private Const DFCS_CHECKED As Long = &H400

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128                                                '  Maintenance string for PSS usage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UPPERLEFT As Long
    LOWERRIGHT As Long
End Type

'Subclassing
Private Enum eMsgWhen
    [MSG_AFTER] = 1
    [MSG_BEFORE] = 2
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE
End Enum

Private Const GWL_WNDPROC      As Long = -4                                     'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88                                     'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93                                     'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132                                    'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137                                    'Table A (after) entry count patch offset

Private Type tSubData                                                           'Subclass data type
    hWnd                       As Long                                          'Handle of the window being subclassed
    nAddrSub                   As Long                                          'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long                                          'The address of the pre-existing WndProc
    nMsgCntA                   As Long                                          'Msg after table entry count
    nMsgCntB                   As Long                                          'Msg before table entry count
    aMsgTblA()                 As Long                                          'Msg after table array
    aMsgTblB()                 As Long                                          'Msg Before table array
End Type

Private sc_aSubData()          As tSubData                                      'Subclass data array
Private sc_aBuf(1 To 200) As Byte                                               'Code buffer byte array
Private sc_pCWP                As Long                                          'Address of the CallWindowsProc
Private sc_pEbMode             As Long                                          'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long                                          'Address of the SetWindowsLong function

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEHOVER     As Long = &H2A1
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_ACTIVATEAPP    As Long = &H1C
                                
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize          As Long
    dwFlags         As TRACKMOUSEEVENT_FLAGS
    hwndTrack       As Long
    dwHoverTime     As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean

'API Scroll Bars
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hWnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hWnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hWnd As Long) As Long

Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = 1&
    Style_Flat = 0&
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = 0                                                           'SB_HORZ
    efsVertical = 1                                                             'SB_VERT
End Enum

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&

'滚动条结构体
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private m_bInitialised      As Boolean
Private m_eOrientation      As ScrollBarOrienationEnum
Private m_eStyle            As ScrollBarStyleEnum
Private m_hWnd              As Long
Private m_lSmallChangeHorz  As Long
Private m_lSmallChangeVert  As Long
Private m_bEnabledHorz      As Boolean
Private m_bEnabledVert      As Boolean
Private m_bVisibleHorz      As Boolean
Private m_bVisibleVert      As Boolean
Private m_bNoFlatScrollBars As Boolean

'枚举
Private Enum lgFlagsEnum
    lgFLChecked = 2
    lgFLSelected = 4
    lgFLChanged = 8
    lgFLFontBold = 16
    lgFLFontItalic = 32
    lgFLFontUnderline = 64
    lgFLWordWrap = 128
End Enum

'枚举
Private Enum lgCellFormatEnum
    lgCFBackColor = 2
    lgCFForeColor = 4
    lgCFImage = 8
End Enum

'枚举
Private Enum lgHeaderStateEnum
    lgNormal = 1
    lgHot = 2
    lgDown = 3
End Enum

'枚举
Private Enum lgRectTypeEnum
    lgRTColumn = 0
    lgRTCheckBox = 1
    lgRTImage = 2
End Enum

'枚举
Public Enum lgAllowResizingEnum
    NotResize = 0
    Resize = 1
End Enum

'枚举
Public Enum lgAlignmentEnum
    lgAlignLeftTop = DT_LEFT Or DT_TOP
    lgAlignLeftCenter = DT_LEFT Or DT_VCENTER
    lgAlignLeftBottom = DT_LEFT Or DT_BOTTOM
    lgAlignCenterTop = DT_CENTER Or DT_TOP
    lgAlignCenterCenter = DT_CENTER Or DT_VCENTER
    lgAlignCenterBottom = DT_CENTER Or DT_BOTTOM
    lgAlignRightTop = DT_RIGHT Or DT_TOP
    lgAlignRightCenter = DT_RIGHT Or DT_VCENTER
    lgAlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

'枚举
Public Enum lgBorderStyleEnum
    无 = 0
    边框 = 1
End Enum

Public Enum lgDataTypeEnum
    lgString = 0
    lgNumeric = 1
    lgDate = 2
    lgBoolean = 3
    lgProgressBar = 4
    lgCustom = 5
End Enum

'枚举
Public Enum lgEditTypeEnum
    None = 0
    EnterKey = 2
    F2Key = 4
    MouseClick = 8
    MouseDblClick = 16
End Enum

'枚举
Public Enum SelectModeEnum
    无 = 0
    行 = 1
    列 = 2
End Enum

'枚举
Public Enum FocusStyleEnum
    Light = 0
    Heavy = 1
End Enum

'枚举
Public Enum lgMoveControlEnum
    lgBCNone = 0
    lgBCHeight = 1
    lgBCWidth = 2
    lgBCLeft = 4
    lgBCTop = 8
End Enum

'选择模式枚举
Public Enum lgSearchModeEnum
    lgSMEqual = 0
    lgSMGreaterEqual = 1
    lgSMLike = 2
    lgSMNavigate = 4
End Enum

'排序方式枚举
Public Enum lgSortTypeEnum
    lgSTAscending = 0
    lgSTDescending = 1
End Enum

#If False Then
Private lgFLChecked, lgFLSelected, lgFLChanged, lgFLFontBold, lgFLFontItalic, lgFLFontUnderline, lgFLWordWrap
Private lgNormal, lgHot, lgDown
Private NotResize, Resize
Private lgAlignLeftTop, lgAlignLeftCenter, lgAlignLeftBottom, lgAlignCenterTop, lgAlignCenterCenter, lgAlignCenterBottom, lgAlignRightTop, lgAlignRightCenter, lgAlignRightBottom
Private lgString, lgNumeric, lgDate, lgBoolean, lgProgressBar, lgCustom
Private None, EnterKey, F2Key, MouseClick, MouseDblClick
Private None, lgRow, lgCol
Private lgFRLight, lgFRHeavy
Private lgSMEqual, lgSMGreaterEqual, lgSMLike, lgSMNavigate
Private lgSTAscending, lgSTDescending
#End If

'列的结构体
Private Type udtColumn
    EditCtrl As Object
    dCustomWidth As Single
    lWidth As Long
    lX As Long
    nAlignment As lgAlignmentEnum
    nImageAlignment As lgAlignmentEnum
    nSortOrder As lgSortTypeEnum
    nType As Integer
    nFlags As Integer
    MoveControl As Integer
    bVisible As Boolean
    sCaption As String
    sFormat As String
    sTag As String
End Type

'单元格的结构体
Private Type udtCell
    nAlignment As Integer
    nFormat As Integer
    nFlags As Integer
    sValue As String
End Type

'行的结构体
Private Type udtItem
    lHeight As Long
    lImage As Long
    lItemData As Long
    nFlags As Integer
    sTag As String
    Cell() As udtCell
End Type

'格式的结构体
Private Type udtFormat
    lBackColor As Long
    lForeColor As Long
    nImage As Integer
    nCount As Long
End Type

'整体渲染的结构体
Private Type udtRender
    DTFlag As Long
    CheckBoxSize As Long                                                        '复选框的大小
    ImageSpace As Long                                                          '图片的空白所占的像素
    ImageHeight As Long                                                         '图片的高度
    ImageWidth As Long                                                          '图片的宽度
    LeftImage As Long                                                           '图片的左边位置
    LeftText As Long                                                            '文本的左边位置
    HeaderHeight As Long                                                        '表头的高度
    TextHeight As Long                                                          '文本的高度
End Type

Private WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1

'Data & Columns
Private mCols() As udtColumn
Private mItems() As udtItem
Private mColPtr() As Long
Private mRowPtr() As Long
Private mCF() As udtFormat

Private mItemCount As Long
Private mItemsVisible As Long
Private mSortColumn As Long
Private mSortSubColumn As Long

Private mEditCol As Long
Private mEditRow As Long
Private mCol As Long
Private mRow As Long
Private mMouseCol As Long
Private mMouseRow As Long
Private mMouseDownCol As Long
Private mMouseDownRow As Long
Private mMouseDownX As Long
Private mSelectedRow As Long

Private mR As udtRender
Private mEditPending As Boolean
Private mMouseDown As Boolean
Private mDragCol As Long
Private mResizeCol As Long
Private mEditParent As Long

'Appearance Properties
Private mSelectBackColor As Long
Private mForeColor As Long
Private mHeadForeColor As Long
Private mSelectForeColor As Long
Private mForeColorTotals As Long

Private mFocusColor As Long
Private mGridColor As Long

Private mAlphaBlendSelection As Boolean
Private mBorderStyle As lgBorderStyleEnum
Private mDisplayEllipsis As Boolean
Private mSelectMode As SelectModeEnum
Private mFocusStyle As FocusStyleEnum
Private mFont As Font
Private mGridLines As Boolean
Private mGridLineWidth As Long

'Behaviour Properties
Private mAllowResizing As lgAllowResizingEnum
Private mCheckboxes As Boolean
Private mColumnDrag As Boolean
Private mColumnHeaders As Boolean
Private mColumnSort As Boolean
Private mEditable As Boolean
Private mEditType As lgEditTypeEnum
Private mFullRowSelect As Boolean
Private mHotHeaderTracking As Boolean
Private mMultiSelect As Boolean
Private mRedraw As Boolean
Private mScrollTrack As Boolean
Private mTrackEdits As Boolean

'Miscellaneous Properties
Private mCacheIncrement As Long
Private mEnabled As Boolean
Private mLocked As Boolean
Private mRowHeight As Long

Private mImageList As Object
Private mImageListScaleMode As Integer

'Control State Variables
Private mInCtrl As Boolean
Private mInFocus As Boolean
Private mWindowsNT As Boolean
Private mWindowsXP As Boolean
Private mUnicode As Boolean

Private mPendingRedraw As Boolean
Private mPendingScrollBar As Boolean

Private mClipRgn As Long
Private hTheme As Long
Private mScrollAction As Long
Private mScrollTick As Long
Private mHotColumn As Long
Private mIgnoreKeyPress As Boolean

'Events - Standard VB
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'Events - Control Specific
Public Event CellImageClick(ByVal Row As Long, ByVal Col As Long)
Public Event ColumnClick(Col As Long)
Public Event ColumnSizeChanged(Col As Long, MoveControl As lgMoveControlEnum)
Public Event CustomSort(Ascending As Boolean, Col As Long, Value1 As String, Value2 As String, Swap As Boolean)
Public Event ItemChecked(Row As Long)
Public Event ItemCountChanged()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event RowColChanged()
Public Event Scroll()
Public Event SelectionChanged()
Public Event SortComplete()
Public Event ThemeChanged()
Public Event EnterCell()
Public Event RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)

Private Function IsColumnTruncated(Col As Long) As Boolean
    If (mR.LeftText > 3) And (Col = 0) Then IsColumnTruncated = True
End Function

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Dim eBar As EFSScrollBarConstants
    Dim lV As Long, lSC As Long
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim zDelta As Long
    Dim lHSB As Long
    Dim lVSB As Long
    Dim bRedraw As Boolean
    
    Select Case uMsg
    Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
        lScrollCode = (wParam And &HFFFF&)
        lHSB = SBValue(efsHorizontal)
        lVSB = SBValue(efsVertical)
        
        Select Case uMsg
        Case WM_HSCROLL                                                         ' Get the scrollbar type
            eBar = efsHorizontal
        Case WM_VSCROLL
            eBar = efsVertical
        Case Else                                                               'WM_MOUSEWHEEL
            eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
            lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
        End Select
        bRedraw = True
        Select Case lScrollCode
        Case SB_THUMBTRACK
            ' Is vertical/horizontal?
            pSBGetSI eBar, tSI, SIF_TRACKPOS
            SBValue(eBar) = tSI.nTrackPos
            bRedraw = mScrollTrack
        Case SB_LEFT, SB_BOTTOM
            SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))
        Case SB_RIGHT, SB_TOP
            SBValue(eBar) = SBMin(eBar)
        Case SB_LINELEFT, SB_LINEUP
            If SBVisible(eBar) Then
                lV = SBValue(eBar)
                If (eBar = efsHorizontal) Then
                    lSC = m_lSmallChangeHorz
                Else
                    lSC = m_lSmallChangeVert
                End If
                
                If (lV - lSC < SBMin(eBar)) Then
                    SBValue(eBar) = SBMin(eBar)
                Else
                    SBValue(eBar) = lV - lSC
                End If
                
            End If
            
        Case SB_LINERIGHT, SB_LINEDOWN
            If SBVisible(eBar) Then
                
                lV = SBValue(eBar)
                
                If (eBar = efsHorizontal) Then
                    lSC = m_lSmallChangeHorz
                Else
                    lSC = m_lSmallChangeVert
                End If
                
                If (lV + lSC > SBMax(eBar)) Then
                    SBValue(eBar) = SBMax(eBar)
                Else
                    SBValue(eBar) = lV + lSC
                End If
            End If
            
        Case SB_PAGELEFT, SB_PAGEUP
            SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)
            
        Case SB_PAGERIGHT, SB_PAGEDOWN
            SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)
            
        Case SB_ENDSCROLL
            If Not mScrollTrack Then DrawGrid True
        End Select
        
        If (lHSB <> SBValue(efsHorizontal)) Or (lVSB <> SBValue(efsVertical)) Then
            UpdateCell
            If bRedraw Then DrawGrid True
            RaiseEvent Scroll
        End If
        
    Case WM_MOUSEWHEEL
        
    Case WM_MOUSEMOVE
        If Not mInCtrl Then
            mInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            RaiseEvent MouseEnter
        End If
        
    Case WM_MOUSELEAVE
        If mInCtrl Then
            mInCtrl = False
            DrawHeaderRow
            UserControl.Refresh
            RaiseEvent MouseLeave
        End If
        
    Case WM_SETFOCUS
        If mEnabled Then
            If Not mInFocus Then
                'Debug.Print "WM_SETFOCUS"
                mInFocus = True
                DrawGrid True
            End If
        End If
        
    Case WM_KILLFOCUS
        If lng_hWnd = UserControl.hWnd Then
            If mEnabled Then
                If mInFocus Then
                    'Debug.Print "WM_KILLFOCUS"
                    mInFocus = False
                    DrawGrid True
                End If
            End If
        ElseIf Not mInCtrl Then
            UpdateCell
        End If
        
    Case WM_THEMECHANGED
        DrawGrid True
        RaiseEvent ThemeChanged
        
    End Select
End Sub

Public Function AddColumn(Optional Caption As String, Optional Width As Single = 1000, Optional Alignment As lgAlignmentEnum = lgAlignLeftCenter, Optional DataType As lgDataTypeEnum = lgString, Optional Format As String, Optional ImageAlignment As lgAlignmentEnum = lgAlignLeftCenter, Optional WordWrap As Boolean, Optional Index As Long = 0) As Long
    Dim lCount As Long
    Dim lNewCol As Long
    
    If mCols(0).nAlignment <> 0 Then
        lNewCol = UBound(mCols) + 1
        ReDim Preserve mCols(lNewCol)
        ReDim Preserve mColPtr(lNewCol)
    End If
    
    If (Index > 0) And (Index < lNewCol) Then
        If lNewCol > 1 Then
            For lCount = lNewCol To Index + 1 Step -1
                mColPtr(lCount) = mColPtr(lCount - 1)
            Next lCount
            mColPtr(Index) = lNewCol
        End If
        
        AddColumn = Index
    Else
        mColPtr(lNewCol) = lNewCol
        AddColumn = lNewCol
    End If
    
    With mCols(lNewCol)
        .sCaption = Caption
        .dCustomWidth = Width
        
        'lWidth is always Pixels (because thats what API functions require) and
        'is calculated to prevent repeated Width Scaling calculations
        .lWidth = ScaleX(.dCustomWidth, vbTwips, vbPixels)
        
        .nAlignment = Alignment
        .nImageAlignment = ImageAlignment
        .nSortOrder = lgSTAscending
        .nType = DataType
        .sFormat = Format
        
        If WordWrap Then .nFlags = lgFLWordWrap
        .bVisible = True
    End With
    
    DisplayChange
End Function

Public Function AddItem(Optional ByVal Item As String, Optional Index As Long = 0, Optional Checked As Boolean) As Long
    Dim lCol As Long
    Dim lCount As Long
    Dim sText() As String
    
    mItemCount = mItemCount + 1
    If mItemCount > UBound(mItems) Then
        ReDim Preserve mItems(mItemCount + mCacheIncrement)
        ReDim Preserve mRowPtr(mItemCount + mCacheIncrement)
    End If
    
    If (Index > 0) And (Index < mItemCount) Then
        If mItemCount > 1 Then
            For lCount = mItemCount To Index + 1 Step -1
                mRowPtr(lCount) = mRowPtr(lCount - 1)
            Next lCount
            mRowPtr(Index) = mItemCount
        End If
        
        AddItem = Index
    Else
        mRowPtr(mItemCount) = mItemCount
        AddItem = mItemCount
    End If
    
    If mRowHeight > 0 Then
        mItems(mItemCount).lHeight = ScaleY(mRowHeight, vbTwips, vbPixels)
    Else
        mItems(mItemCount).lHeight = 300
    End If
    
    ReDim mItems(mItemCount).Cell(UBound(mCols))
    
    For lCount = LBound(mCols) To UBound(mCols)
        With mItems(mItemCount).Cell(lCount)
            .nAlignment = mCols(lCount).nAlignment
            .nFormat = -1
            .nFlags = mCols(lCount).nFlags
        End With
        
        ApplyCellFormat mItemCount, lCount, lgCFBackColor, vbWhite
        ApplyCellFormat mItemCount, lCount, lgCFForeColor, mForeColor
    Next lCount
    
    If UBound(mCols) > 0 Then
        lCol = 0
        sText() = Split(Item, vbTab)
        For lCount = LBound(sText) To UBound(sText)
            With mItems(mItemCount).Cell(lCol)
                .sValue = sText(lCount)
            End With
            
            lCol = lCol + 1
            If lCol > UBound(mCols) Then
                Exit For
            End If
        Next lCount
    Else
        mItems(mItemCount).Cell(0).sValue = Item
    End If
    
    If Checked Then
        SetFlag mItems(mItemCount).nFlags, lgFLChecked, True
    End If
    
    DisplayChange
    
    RaiseEvent ItemCountChanged
End Function

Public Property Get AllowResizing() As lgAllowResizingEnum
Attribute AllowResizing.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowResizing = mAllowResizing
End Property

Public Property Let AllowResizing(ByVal NewValue As lgAllowResizingEnum)
    mAllowResizing = NewValue
    
    PropertyChanged "AllowResizing"
End Property

Private Sub ApplyCellFormat(ByVal Row As Long, ByVal Col As Long, Apply As lgCellFormatEnum, ByVal NewValue As Long)
    Dim lBackColor As Long
    Dim lForeColor As Long
    Dim nImage As Integer
    
    Dim lCount As Long
    Dim nIndex As Integer
    Dim nFreeIndex As Integer
    Dim nNewIndex As Integer
    Dim bMatch As Boolean
    
    nIndex = mItems(Row).Cell(Col).nFormat
    
    If nIndex >= 0 Then
        'Get current properties
        With mCF(nIndex)
            lBackColor = .lBackColor
            lForeColor = .lForeColor
            nImage = .nImage
        End With
    Else
        'Set default properties
        lBackColor = vbWhite
        lForeColor = mForeColor
    End If
    
    Select Case Apply
    Case lgCFBackColor
        lBackColor = NewValue
    Case lgCFForeColor
        lForeColor = NewValue
    Case lgCFImage
        nImage = NewValue
    End Select
    
    nFreeIndex = -1
    For lCount = 0 To UBound(mCF)
        If (mCF(lCount).lBackColor = lBackColor) And (mCF(lCount).lForeColor = lForeColor) And (mCF(lCount).nImage = nImage) Then
            'Existing Entry matches what we required
            bMatch = True
            nNewIndex = lCount
            Exit For
        ElseIf (mCF(lCount).nCount = 0) And (nFreeIndex = -1) Then
            'An unused entry
            nFreeIndex = lCount
        End If
    Next lCount
    
    'No existing matches
    If Not bMatch Then
        'Is there an unused Entry?
        If nFreeIndex >= 0 Then
            nNewIndex = nFreeIndex
        Else
            nNewIndex = UBound(mCF) + 1
            ReDim Preserve mCF(nNewIndex + 9)
        End If
        
        With mCF(nNewIndex)
            .lBackColor = lBackColor
            .lForeColor = lForeColor
            .nImage = nImage
        End With
    End If
    
    'Has the Format Entry Index changed?
    If (nIndex <> nNewIndex) Then
        'Increment reference count for new entry
        mCF(nNewIndex).nCount = mCF(nNewIndex).nCount + 1
        
        If nIndex >= 0 Then
            'Decrement reference count for previous entry
            mCF(nIndex).nCount = mCF(nIndex).nCount - 1
        End If
    End If
    
    mItems(Row).Cell(Col).nFormat = nNewIndex
End Sub

Public Property Get SelectBackColor() As OLE_COLOR
    SelectBackColor = mSelectBackColor
End Property

Public Property Let SelectBackColor(ByVal NewValue As OLE_COLOR)
    mSelectBackColor = NewValue
    DisplayChange
    
    PropertyChanged "SelectBackColor"
End Property

Public Sub BindControl(ByVal Col As Long, Ctrl As Object, Optional MoveControl As lgMoveControlEnum = lgBCHeight Or lgBCLeft Or lgBCTop Or lgBCWidth)
    Set mCols(Col).EditCtrl = Ctrl
    mCols(Col).MoveControl = MoveControl
End Sub

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    Dim lSrcR  As Long
    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    
    lCFrom = oColorFrom
    lCTo = oColorTo
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
End Function


Public Property Get BorderStyle() As lgBorderStyleEnum
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As lgBorderStyleEnum)
    mBorderStyle = NewValue
    UserControl.BorderStyle = mBorderStyle
    
    PropertyChanged "BorderStyle"
End Property

Public Property Get CacheIncrement() As Long
    CacheIncrement = mCacheIncrement
End Property

Public Property Let CacheIncrement(ByVal NewValue As Long)
    If NewValue < 0 Then
        mCacheIncrement = 1
    Else
        mCacheIncrement = NewValue
    End If
    
    PropertyChanged "CacheIncrement"
End Property

Public Property Let CellAlignment(ByVal Row As Long, ByVal Col As Long, NewValue As lgAlignmentEnum)
    mItems(mRowPtr(Row)).Cell(Col).nAlignment = NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellAlignment(ByVal Row As Long, ByVal Col As Long) As lgAlignmentEnum
    CellAlignment = mItems(mRowPtr(Row)).Cell(Col).nAlignment
End Property

Public Property Let CellBackColor(ByVal Row As Long, ByVal Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFBackColor, NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get CellBackColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellBackColor = mCF(mItems(mRowPtr(Row)).Cell(Col).nFormat).lBackColor
End Property

Public Property Let CellChecked(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLChecked, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellChecked(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellChecked = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLChecked
End Property

Public Property Let CellChanged(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLChanged, NewValue
End Property

Public Property Get CellChanged(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellChanged = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLChanged
End Property

Public Property Let CellFontBold(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLFontBold, NewValue
    DrawGrid mRedraw
End Property

Public Property Get TrackEdits() As Boolean
    TrackEdits = mTrackEdits
End Property

Public Property Let TrackEdits(ByVal NewValue As Boolean)
    mTrackEdits = NewValue
    
    PropertyChanged "TrackEdits"
End Property

Public Property Get CellFontBold(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontBold = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLFontBold
End Property

Public Property Let CellFontItalic(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLFontItalic, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellFontItalic(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontItalic = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLFontItalic
End Property

Public Property Let CellFontUnderline(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLFontUnderline, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellFontUnderline(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontUnderline = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLFontUnderline
End Property

Public Property Let CellForeColor(ByVal Row As Long, ByVal Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFForeColor, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellForeColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellForeColor = mCF(mItems(mRowPtr(Row)).Cell(Col).nFormat).lForeColor
End Property

Public Property Let CellImage(ByVal Row As Long, ByVal Col As Long, NewValue As Variant)
    Dim nImage As Integer
    
    On Local Error GoTo ItemImageError
    
    If IsNumeric(NewValue) Then
        nImage = NewValue
    Else
        nImage = -mImageList.ListImages(NewValue).Index
    End If
    
    ApplyCellFormat Row, Col, lgCFImage, nImage
    DrawGrid mRedraw
    Exit Property
    
ItemImageError:
    ApplyCellFormat Row, Col, lgCFImage, 0
End Property

Public Property Get CellImage(ByVal Row As Long, ByVal Col As Long) As Variant
    Dim nImage As Integer
    
    nImage = mCF(mItems(mRowPtr(Row)).Cell(Col).nFormat).nImage
    
    If nImage >= 0 Then
        CellImage = nImage
    Else
        CellImage = mImageList.ListImages(Abs(nImage)).Key
    End If
End Property

Public Property Let CellProgressValue(ByVal Row As Long, ByVal Col As Long, NewValue As Integer)
    If mCols(Col).nType = lgProgressBar Then
        If NewValue > 100 Then
            NewValue = 100
        ElseIf NewValue < 0 Then
            NewValue = 0
        End If
        
        mItems(mRowPtr(Row)).Cell(Col).nFlags = NewValue
        DrawGrid mRedraw
    End If
End Property

Public Property Get CellProgressValue(ByVal Row As Long, ByVal Col As Long) As Integer
    If mCols(Col).nType = lgProgressBar Then
        CellProgressValue = mItems(mRowPtr(Row)).Cell(Col).nFlags
    End If
End Property

Public Property Let CellText(ByVal Row As Long, ByVal Col As Long, NewValue As String)
    mItems(mRowPtr(Row)).Cell(Col).sValue = NewValue
    'SetRowSize Row
    
    If mTrackEdits Then CellChanged(Row, Col) = True
    DrawGrid mRedraw
End Property

Public Property Get CellText(ByVal Row As Long, ByVal Col As Long) As String
    CellText = mItems(mRowPtr(Row)).Cell(Col).sValue
End Property

Public Property Let CellWordWrap(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(Col).nFlags, lgFLWordWrap, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellWordWrap(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellWordWrap = mItems(mRowPtr(Row)).Cell(Col).nFlags And lgFLFontItalic
End Property

Public Property Get CheckBoxes() As Boolean
Attribute CheckBoxes.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CheckBoxes = mCheckboxes
End Property

Public Property Let CheckBoxes(ByVal NewValue As Boolean)
    mCheckboxes = NewValue
    DisplayChange
    
    PropertyChanged "CheckBoxes"
End Property

Public Function CheckedCount() As Long
    '#############################################################################################################################
'Purpose: Return Count of Checked Items
    '#############################################################################################################################
    
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgFLChecked Then
            CheckedCount = CheckedCount + 1
        End If
    Next lCount
End Function

Public Sub Clear()
    ReDim mItems(0)
    ReDim mRowPtr(0)
    ReDim mCF(0)
    
    mMouseDownCol = -1
    mMouseDownRow = -1
    
    mCol = -1
    mRow = -1
    mSelectedRow = -1
    
    mHotColumn = -1
    mDragCol = -1
    mResizeCol = -1
    
    mSortColumn = -1
    mSortSubColumn = -1
    
    mScrollAction = 0
    mItemCount = -1
    
    DrawGrid True
End Sub

Public Property Get Col() As Long
    Col = mCol
End Property

Public Property Let Col(ByVal NewValue As Long)
    If SetRowCol(mRow, NewValue) Then
        DrawGrid mRedraw
    End If
End Property

Public Property Get ColAlignment(ByVal Index As Long) As lgAlignmentEnum
    ColAlignment = mCols(Index).nAlignment
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal NewValue As lgAlignmentEnum)
    mCols(Index).nAlignment = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColImageAlignment(ByVal Index As Long) As lgAlignmentEnum
    ColImageAlignment = mCols(Index).nImageAlignment
End Property

Public Property Let ColImageAlignment(ByVal Index As Long, ByVal NewValue As lgAlignmentEnum)
    mCols(Index).nImageAlignment = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColFormat(ByVal Index As Long) As String
    ColFormat = mCols(Index).sFormat
End Property

Public Property Let ColFormat(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sFormat = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColHeading(ByVal Index As Long) As String
    ColHeading = mCols(Index).sCaption
End Property

Public Property Let ColHeading(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sCaption = NewValue
    
    DrawGrid mRedraw
End Property

Public Function ColLeft(ByVal Index As Long) As Long
    Dim r As RECT
    
    SetColRect Index, r
    ColLeft = r.Left
End Function

Public Property Get Cols() As Long
    Cols = UBound(mCols)
End Property

Public Property Let Cols(ByVal NewValue As Long)
    ReDim mCols(NewValue)
End Property

Public Property Get ColType(ByVal Index As Long) As lgDataTypeEnum
    ColType = mCols(Index).nType
End Property

Public Property Let ColType(ByVal Index As Long, ByVal NewValue As lgDataTypeEnum)
    mCols(Index).nType = NewValue
End Property

Public Property Get ColWordWrap(ByVal Index As Long) As Boolean
    ColWordWrap = mCols(Index).nFlags And lgFLWordWrap
End Property

Public Property Let ColWordWrap(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mCols(Index).nFlags, lgFLWordWrap, NewValue
End Property

Public Property Get ColumnDrag() As Boolean
    ColumnDrag = mColumnDrag
End Property

Public Property Let ColumnDrag(ByVal NewValue As Boolean)
    mColumnDrag = NewValue
    
    PropertyChanged "ColumnDrag"
End Property

Public Property Get ColumnSort() As Boolean
    ColumnSort = mColumnSort
End Property

Public Property Let ColumnSort(ByVal NewValue As Boolean)
    mColumnSort = NewValue
    
    PropertyChanged "ColumnSort"
End Property

Public Property Get ColTag(ByVal Index As Long) As String
    ColTag = mCols(Index).sTag
End Property

Public Property Let ColTag(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sTag = NewValue
End Property

Public Property Get ColVisible(ByVal Index As Long) As Boolean
    ColVisible = mCols(Index).bVisible
End Property

Public Property Let ColVisible(ByVal Index As Long, ByVal NewValue As Boolean)
    mCols(Index).bVisible = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColWidth(ByVal Index As Long) As Single
    ColWidth = mCols(Index).dCustomWidth
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal NewValue As Single)
    'dCustomWidth is in the Units the Control is operating in
    mCols(Index).dCustomWidth = NewValue
    mCols(Index).lWidth = ScaleX(NewValue, vbTwips, vbPixels)
    
    DrawGrid mRedraw
End Property

Private Sub CreateRenderData()
    Dim lCount As Long
    Dim lSize As Long
    
    With mR
        lSize = ScaleY(mRowHeight, vbTwips, vbPixels)
        If lSize > 16 Then
            .CheckBoxSize = 16
        Else
            .CheckBoxSize = lSize - 4
        End If
        
        If mCheckboxes Then
            .LeftText = .CheckBoxSize + 2
        Else
            .LeftImage = 0
            .LeftText = 3
        End If
        
        .LeftImage = .LeftText
        
        If mImageList Is Nothing Then
            .ImageSpace = 0
        Else
            .ImageSpace = ((GetRowHeight() - mImageList.ImageHeight) / 2)
            .ImageHeight = mImageList.ImageHeight
            .ImageWidth = mImageList.ImageWidth
            For lCount = LBound(mItems) To UBound(mItems)
                If mItems(lCount).lImage <> 0 Then
                    .LeftText = .LeftText + mImageList.ImageWidth + 2
                    Exit For
                End If
            Next lCount
        End If
        
        .HeaderHeight = GetColumnHeadingHeight()
        .TextHeight = UserControl.TextHeight("A")
        
        If mDisplayEllipsis Then
            .DTFlag = DT_SINGLELINE Or DT_WORD_ELLIPSIS
        Else
            .DTFlag = DT_SINGLELINE
        End If
    End With
End Sub

Private Sub DisplayChange()
    If mRedraw Then
        Refresh
    Else
        mPendingRedraw = True
        mPendingScrollBar = True
    End If
End Sub

Public Property Get DisplayEllipsis() As Boolean
Attribute DisplayEllipsis.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DisplayEllipsis = mDisplayEllipsis
End Property

Public Property Let DisplayEllipsis(ByVal NewValue As Boolean)
    mDisplayEllipsis = NewValue
    DisplayChange
    
    PropertyChanged "DisplayEllipsis"
End Property

Public Sub Sort(Optional Sort As Long = -1, Optional SortType As lgSortTypeEnum = -1, Optional SubSort As Long = -1, Optional SubSortType As lgSortTypeEnum = -1)
    '#############################################################################################################################
'Purpose: Sort Grid based on current Sort Columns.
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lRowIndex As Long
    
    If UpdateCell() Then
        'Set new Columns if specified
        If Sort <> -1 Then
            mSortColumn = Sort
        End If
        
        If SubSort <> -1 Then
            mSortSubColumn = SubSort
        End If
        
        'Validate Sort Columns
        If (mSortColumn = -1) And (mSortSubColumn <> -1) Then
            mSortColumn = mSortSubColumn
            mSortSubColumn = -1
        ElseIf mSortColumn = mSortSubColumn Then
            mSortSubColumn = -1
        End If
        
        'Set Sort Order if specified - otherwise inverse last Sort Order
        With mCols(mSortColumn)
            If SortType = -1 Then
                If .nSortOrder = lgSTAscending Then
                    .nSortOrder = lgSTDescending
                Else
                    .nSortOrder = lgSTAscending
                End If
            Else
                .nSortOrder = SortType
            End If
        End With
        
        If mSortSubColumn <> -1 Then
            With mCols(mSortSubColumn)
                If SubSortType = -1 Then
                    If .nSortOrder = lgSTAscending Then
                        .nSortOrder = lgSTDescending
                    Else
                        .nSortOrder = lgSTAscending
                    End If
                Else
                    .nSortOrder = SubSortType
                End If
            End With
        End If
        
        'Note previously selected Row
        If mRow > -1 Then
            lRowIndex = mRowPtr(mRow)
        End If
        
        SortArray LBound(mItems), mItemCount, mSortColumn, mCols(mSortColumn).nSortOrder
        SortSubList
        
        For lCount = LBound(mRowPtr) To mItemCount
            If mRowPtr(lCount) = lRowIndex Then
                mRow = lCount
                Exit For
            End If
        Next lCount
        
        DrawGrid True
        
        RaiseEvent SortComplete
    End If
End Sub



Private Sub DrawGrid(bRedraw As Boolean)
    '#############################################################################################################################
'Purpose: The Primary Rendering routine. Draws Columns & Rows
    '#############################################################################################################################
    
    Dim IR As RECT                                                              '定义一个矩形
    Dim r As RECT                                                               '定义一个矩形
    Dim lX As Long
    Dim lY As Long
    
    Dim lCol As Long
    Dim lRow As Long
    Dim lMaxRow As Long
    Dim lStartCol As Long
    Dim lColumnsWidth As Long
    Dim lBottomEdge As Long
    Dim lGridColor As Long
    Dim lImageLeft As Long
    Dim lRowWrapSize As Long
    Dim lStart As Long
    Dim lValue As Long
    Dim nImage As Integer
    Dim bLockColor As Boolean
    Dim sText As String
    Dim bBold As Boolean
    Dim bItalic As Boolean
    Dim bUnderLine As Boolean
    
    '如果重画
    If bRedraw Then
        lStartCol = SBValue(efsHorizontal)                                      '记录滚动条方向
        lGridColor = TranslateColor(mGridColor)                                 '记录Grid颜色
        
        lY = mR.HeaderHeight                                                    '记录标题的高度
        mItemsVisible = ItemsVisible()
        lRowWrapSize = (mR.TextHeight * 2)
        
        With UserControl
            .Cls
            
            bBold = .FontBold                                                   '设置字体粗细
            bItalic = .FontItalic                                               '设置字体斜体
            bUnderLine = .FontUnderline                                         '设置字体下划线
            
            lColumnsWidth = DrawHeaderRow()                                     '调用画表头函数
            
            lMaxRow = (SBValue(efsVertical) + mItemsVisible)                    '滚动条方向和可见行数
            If lMaxRow > mItemCount Then                                        '如果大于总行数，则令最行为总行数
                lMaxRow = mItemCount
            End If
            
            '取出滚动条在垂直方向上的位置作为初始值，总行数作为结束值
            lStart = SBValue(efsVertical)
            '双层循环，外层是行循环，内层是列循环
            For lRow = lStart To lMaxRow
                '如果是多选模式或行选模式，并且行标记和选择标记
                If (mMultiSelect Or mFullRowSelect) And (mItems(mRowPtr(lRow)).nFlags And lgFLSelected) Then
                    '锁定颜色
                    bLockColor = True
                    '如果起始列为零，先设置起始列的矩形，然后画起始列的矩形
                    If lStartCol = 0 Then                                       ' ensure 1st column is visible
                        '如果第一列的宽度小于要呈现的左边文本
                        If mCols(0).lWidth < mR.LeftText Then
                            'SetRect是API函数用来设置格式化矩形的，格式化矩形来管理编辑控件文本显示区域的大小，
                            '默认和窗口客户区一样大，但可以用SetRect来设置。
                            '指定格式化矩形的新的尺寸。
                            '使用SetRect函数设置一个对多行编辑控件的格式化矩形。此格式化矩形为文本的边界矩形，与编辑控件窗口的大小无关。
                            '参数： lpRect 是一个CRect或一个指向RECT的指针。它表明了格式化矩形的新的界线。
                            '       左上角坐标X1,KY1;右下角坐标X2,Y2。即先画|再画_，最终成|_
                            SetRect r, 0, lY + 1, mCols(0).lWidth, lY + (mItems(mRowPtr(lRow)).lHeight) = 1
                        Else
                            SetRect r, 0, lY + 1, mR.LeftText, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                        End If
                        DrawRect .hdc, r, TranslateColor(vbWhite), True
                    Else
                        r.Right = 0
                    End If
                    
                    '紧挨R的右边接着在话一个矩形
                    SetRect r, r.Right, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                    
                    If mAlphaBlendSelection Then
                        lValue = BlendColor(TranslateColor(mSelectBackColor), TranslateColor(vbWhite), 120)
                    Else
                        '将颜色值转换为整型
                        lValue = TranslateColor(mSelectBackColor)
                    End If
                    
                    DrawRect .hdc, r, lValue, True
                    
                    .ForeColor = mSelectForeColor
                Else
                    '置锁定颜色标记为假，格式化矩形，然后画矩形
                    bLockColor = False
                    SetRect r, 0, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                    DrawRect .hdc, r, TranslateColor(vbWhite), True
                End If
                
                lX = 0
                For lCol = lStartCol To UBound(mCols)
                    '
                    If mCols(mColPtr(lCol)).bVisible Then
                        '【API函数--SetRectRgn】SelectClipRgn
                        '功能：设置区域为X1，Y1和X2，Y2描述的矩形;它是设置一个已存在区域而不是创建一个新区域,矩形的底和右边不包括在区域内
                        '参数：hRgn Long，该区域将被设置为指定矩形
                        '      X1,Y1 Long，矩形左上角X，Y坐标
                        '      X2,Y2 Long，矩形右下角X，Y坐标
                        '【API函数--SelectClipRgn】
                        '功能：为指定设备场景选择新的剪裁区
                        '参数：
                        '     hdc Long，将设置新剪裁区的设备场景
                        '     hRgn Long，将为设备场景设置剪裁区的句柄，该区域使用设备坐标
                        '返回值：
                        '    Long，下列常数之一，以描述当前剪裁区：
                        '    COMPLEXREGION：该区域有互相交叠的边界；SIMPLEREGION：该区域边界没有互相交叠；NULLREGION：区域为空
                        '
                        SetRectRgn mClipRgn, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight
                        SelectClipRgn .hdc, mClipRgn
                        
                        Call SetRect(r, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight)
                        
                        If Not bLockColor Then
                            'If mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor <> vbWhite Then
                            'DrawRect .hdc, R, TranslateColor(mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor),  True
                            'End If
                            .ForeColor = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lForeColor
                        End If
                        
                        '如果正在画第0列
                        If lCol = 0 Then
                            '如果被设置了复选框
                            If mCheckboxes Then
                                Call SetRect(r, 3, lY, mR.CheckBoxSize, lY + mItems(mRowPtr(lRow)).lHeight)
                                
                                If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                    lValue = 5
                                Else
                                    lValue = 0
                                End If
                                '【API函数 DrawFrameControl】
                                '功 能：用于描绘一个标准控件。
                                '参 数：
                                '     hDC Long，要在其中作画的设备场景
                                '     lpRect RECT，指定帧的位置及大小的一个矩形
                                '     un1 Long，指定帧类型的一个常数。这些常数包括DFC_BUTTON，DFC_CAPTION，DFC_MENU，以及DFC_SCROLL
                                '     un2 Long，一个常数，指定欲描绘的帧的状态。由带有前缀DFCS_的一个常数构成
                                '返回值：Long，非零表示成功，零表示失败。会设置GetLastError
                                If Not DrawTheme("Button", 3, lValue, r) Then
                                    If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                        Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                    Else
                                        Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                    End If
                                End If
                            End If
                            
                            '如果包含图片空白
                            If mR.ImageSpace > 0 Then
                                'If we have an Image Index then Draw it
                                If mItems(mRowPtr(lRow)).lImage <> 0 Then
                                    'Calculate Image offset (using ScaleMode of ImageList)
                                    If lImageLeft = 0 Then
                                        '标准函数-ScaleX：
                                        '功能：用以将 Form，PictureBox 或 Printer 的宽度或高度值从一种 ScaleMode 属性的度量单位转换到另一种。
                                        lImageLeft = ScaleX(mR.LeftImage, vbPixels, mImageListScaleMode)
                                    End If
                                    
                                    '标准函数-Draw：
                                    '功能：在一幅图象上执行了一次图形操作后，把该图象绘制到某个目标设备描述体中，例如 PictureBox 控件中。
                                    '参数：object 必需的。对象表达式，其值是:ListImage 对象、ListImages 集合
                                    '      hDC 必需的。一个设置为目标对象的 hDC 属性的值。
                                    '      x,y 可选的。用来指定设备描述体内绘制图象的位置坐标。如果不指定这些，图象将被绘制在设备描述体的起点。
                                    '      style 可选的。它指定了在图象上进行的操作，“设置值”中有详细说明。
                                    
                                    '备注：hDC 属性是 Windows 操作系统用来作内部引用到对象的句柄（数值）。
                                    '      可以在任何有 hDC 属性的控件内部区域画图。在 Visual Basic 中，上述控件包括 Form 对象、PictureBox 控件和 Printer 对象。
                                    If bLockColor Then
                                        mImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 2
                                    Else
                                        mImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 1
                                    End If
                                End If
                            End If
                            
                            Call SetRect(r, mR.LeftText + 3, lY, (lX + mCols(mColPtr(lCol)).lWidth) - 3, lY + mItems(mRowPtr(lRow)).lHeight)
                        Else
                            Call SetRect(r, lX + 3, lY, (lX + mCols(mColPtr(lCol)).lWidth) - 3, lY + mItems(mRowPtr(lRow)).lHeight)
                        End If
                        
                        '判断列的类型
                        Select Case mCols(mColPtr(lCol)).nType
                        Case lgBoolean                                          '布尔型
                            SetItemRect mRowPtr(lRow), mColPtr(lCol), lY, r, lgRTCheckBox
                            
                            If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                lValue = 5
                            Else
                                lValue = 0
                            End If
                            
                            '【API函数--DrawFocusRect】
                            '    函数原型：
                            '    函数功能：
                            '    参    数：
                            '　　返 回 值：
                            '    备    注：
                            If Not DrawTheme("Button", 3, lValue, r) Then
                                If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                Else
                                    Call DrawFrameControl(.hdc, r, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                End If
                            End If
                            
                        Case lgProgressBar                                      '进度条型
                            If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags > 0 Then
                                lValue = ((mCols(mColPtr(lCol)).lWidth - 2) / 100) * mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags
                                
                                SetRect r, lX + 2, lY + 2, lX + lValue, (lY + mItems(mRowPtr(lRow)).lHeight) - 2
                                DrawRect .hdc, r, TranslateColor(&H8080FF), True '画进度条
                            End If
                            
                        Case Else                                               '如果是其它类型的列
                            With mItems(mRowPtr(lRow)).Cell(mColPtr(lCol))
                                UserControl.FontBold = .nFlags And lgFLFontBold
                                UserControl.FontItalic = .nFlags And lgFLFontItalic
                                UserControl.FontUnderline = .nFlags And lgFLFontUnderline
                                
                                If Len(mCols(mColPtr(lCol)).sFormat) > 0 Then
                                    sText = Format$(.sValue, mCols(mColPtr(lCol)).sFormat)
                                Else
                                    sText = .sValue
                                End If
                                
                                If mItems(mRowPtr(lRow)).lHeight > lRowWrapSize Then
                                    If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLWordWrap Then
                                        lValue = .nAlignment Or DT_WORDBREAK
                                    Else
                                        lValue = .nAlignment Or mR.DTFlag
                                    End If
                                Else
                                    lValue = .nAlignment Or mR.DTFlag
                                End If
                                
                                nImage = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).nImage
                                If nImage <> 0 Then
                                    SetItemRect mRowPtr(lRow), mColPtr(lCol), lY, IR, lgRTImage
                                    
                                    If IR.Left >= 0 Then
                                        If bLockColor Then
                                            mImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, mImageListScaleMode), ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 2
                                        Else
                                            mImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, mImageListScaleMode), ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 1
                                        End If
                                    End If
                                    
                                    'Adjust Text Rect
                                    Select Case mCols(mColPtr(lCol)).nImageAlignment
                                    Case lgAlignLeftTop, lgAlignLeftCenter, lgAlignLeftBottom
                                        r.Left = r.Left + (IR.Right - IR.Left)
                                        
                                    Case lgAlignRightTop, lgAlignRightCenter, lgAlignRightBottom
                                        r.Right = r.Right - (IR.Right - IR.Left)
                                    End Select
                                End If
                                
                                '【API函数--DrawFocusRect】
                                '    函数原型：
                                '    函数功能：
                                '    参    数：
                                '　　返 回 值：
                                '    备    注：
                                Call DrawText(UserControl.hdc, sText, -1, r, lValue)
                            End With
                        End Select
                        
                        lX = lX + mCols(mColPtr(lCol)).lWidth
                    End If
                Next lCol
                
                SelectClipRgn .hdc, 0&
                
                'Display Horizontal Lines
                If mGridLines Then
                    DrawLine .hdc, 0, lY, lColumnsWidth, lY, lGridColor, mGridLineWidth
                End If
                
                lY = lY + mItems(mRowPtr(lRow)).lHeight
            Next lRow
            
            '#############################################################################################################################
            'Display Vertical Lines
            '显示垂直线
            If mGridLines Then
                lBottomEdge = r.Bottom
                
                lX = 0
                For lCol = lStartCol To UBound(mCols)
                    If mCols(mColPtr(lCol)).bVisible Then
                        DrawLine .hdc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth
                        
                        lX = lX + mCols(mColPtr(lCol)).lWidth
                    End If
                Next lCol
            End If
            
            '#############################################################################################################################
            'Display Focus Rectangle
            '显示焦点矩形
            If (mSelectMode <> SelectModeEnum.无) And (mRow >= 0) Then
                If Not mInFocus Then
                    lY = RowTop(mRow)
                    If lY >= 0 Then
                        r.Right = 0
                        If mSelectMode = 列 Then
                            SetColRect mCol, r
                            r.Top = lY + 1
                            r.Bottom = lY + mItems(mRowPtr(mRow)).lHeight
                        Else                                                    'If mFullRowSelect Then
                            SetRect r, mR.LeftText, lY + 1, lColumnsWidth, lY + mItems(mRowPtr(mRow)).lHeight
                        End If
                        
                        '【API函数--DrawFocusRect】
                        '    函数原型：
                        '    函数功能：
                        '    参    数：
                        '　　返 回 值：
                        '    备    注：
                        If r.Right > 0 Then
                            Select Case mFocusStyle
                            Case Light
                                Call DrawFocusRect(.hdc, r)
                            Case Heavy
                                DrawRect .hdc, r, TranslateColor(mFocusColor), False
                            End Select
                        End If
                    End If
                End If
            End If
            
            .Refresh
            
            .FontBold = bBold
            .FontItalic = bItalic
            .FontUnderline = bUnderLine
        End With
        
        'Debug.Print "Drawgrid mRedraw " & Timer
        
        mPendingRedraw = False
    Else
        mPendingRedraw = True
    End If
End Sub
Private Function LongToSignedShort(dwUnsigned As Long) As Integer
    If dwUnsigned < 32768 Then
        LongToSignedShort = CInt(dwUnsigned)
    Else
        LongToSignedShort = CInt(dwUnsigned - &H10000)
    End If
End Function
Private Sub FillGradient(lhDC As Long, rRect As RECT, ByVal clrFirst As OLE_COLOR, ByVal clrSecond As OLE_COLOR, Optional ByVal bVertical As Boolean)
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pGradRect       As GRADIENT_RECT
    
    With pVert(0)
        .x = rRect.Left
        .y = rRect.Top
        .Red = LongToSignedShort((clrFirst And &HFF&) * 256)
        .Green = LongToSignedShort(((clrFirst And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrFirst And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pVert(1)
        .x = rRect.Right
        .y = rRect.Bottom
        .Red = LongToSignedShort((clrSecond And &HFF&) * 256)
        .Green = LongToSignedShort(((clrSecond And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrSecond And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pGradRect
        .UPPERLEFT = 0
        .LOWERRIGHT = 1
    End With
    
    GradientFill lhDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, &H0, &H1)
End Sub

Private Sub DrawHeader(lCol As Long, State As lgHeaderStateEnum)
    Dim r As RECT
    
    If lCol > -1 Then
        With UserControl
            .ForeColor = mHeadForeColor
            'Draw the Column Headers
            Call SetRect(r, mCols(mColPtr(lCol)).lX, 0, mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth + 1, mR.HeaderHeight)
            'DrawRect .hdc, R, TranslateColor(BackColorFixed), True
            DrawOfficeXPHeader .hdc, r, State
            
            'Render Sort Arrows
            If mCols(mColPtr(lCol)).lWidth > 8 Then
                If mColPtr(lCol) = mSortColumn Then
                    DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6, 9, 5, mCols(mColPtr(lCol)).nSortOrder
                    
                    Call SetRect(r, mCols(mColPtr(lCol)).lX + 3, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (5 + 8), mR.HeaderHeight)
                ElseIf mColPtr(lCol) = mSortSubColumn Then
                    DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6, 6, 3, mCols(mColPtr(lCol)).nSortOrder
                    
                    Call SetRect(r, mCols(mColPtr(lCol)).lX + 3, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (5 + 8), mR.HeaderHeight)
                Else
                    Call SetRect(r, mCols(mColPtr(lCol)).lX + 3, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (3 * 2), mR.HeaderHeight)
                End If
            Else
                Call SetRect(r, mCols(mColPtr(lCol)).lX + 3, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (3 * 2), mR.HeaderHeight)
            End If
            
            Call DrawText(.hdc, mCols(mColPtr(lCol)).sCaption, -1, r, mCols(mColPtr(lCol)).nAlignment Or mR.DTFlag)
            
        End With
    End If
End Sub

Private Function DrawHeaderRow() As Long
    '#############################################################################################################################
'Purpose: Renders all Column Headers
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lX As Long
    
    mHotColumn = -1
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
        If mCols(mColPtr(lCol)).bVisible Then
            mCols(mColPtr(lCol)).lX = lX
            DrawHeader lCol, lgNormal
            lX = lX + mCols(mColPtr(lCol)).lWidth
        End If
    Next lCol
    
    DrawHeaderRow = lX
End Function

Private Function InvertThisColor(oInsColor As OLE_COLOR)
    '#############################################################################################################################
'Source: Riccardo Cohen
    '#############################################################################################################################
    
    Dim lROut As Long, lGOut As Long, lBOut As Long
    Dim lRGB As Long
    
    lRGB = TranslateColor(oInsColor)
    
    lROut = (255 - (lRGB And &HFF&))
    lGOut = (255 - ((lRGB And &HFF00&) / &H100))
    lBOut = (255 - ((lRGB And &HFF0000) / &H10000))
    InvertThisColor = RGB(lROut, lGOut, lBOut)
End Function


Private Sub DrawLine(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long, lWidth As Long)
    Dim PT As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long
    
    hPen = CreatePen(0, lWidth, lcolor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, PT
    LineTo hdc, X2, Y2
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Sub

Private Sub DrawOfficeXPHeader(lhDC As Long, rRect As RECT, State As lgHeaderStateEnum)
    With rRect
        Select Case State
        Case lgNormal
            Call FillGradient(lhDC, rRect, &HFCE1CB, &HE0A57D, True)
            
            DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
            DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
            
            DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &HCB8C6A, 1
            DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1
            
        Case lgHot
            .Right = .Right - 1
            Call FillGradient(lhDC, rRect, &HDCFFFF, &H5BC0F7, True)
            
            DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
            DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
            
            DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1
            
        Case lgDown
            .Right = .Right - 1
            Call FillGradient(lhDC, rRect, &H87FE8, &H7CDAF7, True)
            
            DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
            DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
            
            DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1
            
        End Select
    End With
End Sub

Private Sub DrawRect(hdc As Long, rc As RECT, lcolor As Long, bFilled As Boolean)
    Dim lNewBrush As Long
    '创建一个画刷
    lNewBrush = CreateSolidBrush(lcolor)
    
    '如果要填充则则填充，否则只画边框
    If bFilled Then
        Call FillRect(hdc, rc, lNewBrush)
    Else
        Call FrameRect(hdc, rc, lNewBrush)
    End If
    
    Call DeleteObject(lNewBrush)
    '知识点：
    '【API函数--FillRect】
    '    函数原型：int FillRect(HDC hdc, CONST RECT *lprc, HBRUSH hbr)；
    '    函数功能：该函数用指定的画刷填充矩形，此函数包括矩形的左上边界，但不包括矩形的右下边界。
    '    参    数：
    '              hdc:     设备环境句柄?
    '              lprc:    指向含有将填充矩形的逻辑坐标的RECT结构的指针?
    '              hbr:     用来填充矩形的画刷的句柄?
    '　　返 回 值：如果函数调用成功，返回值非零；如果函数调用失败，返回值是0。
    '    备    注：
    '          由参数hbr定义的画刷可以是一个逻辑现刷句柄也可以是一个颜色值，如果指定一个逻辑画刷的句柄，
    '          调用下列函数之一来获得句柄；CreateHatchBrush、CreatePatternBrush或CreateSolidBrush。
    '          此外，你可以用GetStockObject来获得一个库存画刷句柄。如果指定一个颜色值，
    '          必须是标准系统颜色（所选择的颜色必须加1）如FillRect(hdc, &rect, (HBRUSH)(COLOR_ENDCOLORS+1))，参见GetSysColor可得到所有标准系统颜色列表。
    '          当填充一个指定矩形时，FillRect不包括矩形的右、下边界。无论当前映射模式如何，GDI填充一个矩形都不包括右边的列和下面的行。
    '【API函数--FrameRect】
    '    函数原型：int frameRect(HDC hdc, CONST RECT *lprc, HBRUSH hbr)；
    '    函数功能：该函数用指定的画刷为指定的矩形画边框。边框的宽和高总是一个逻辑单元。
    '    参    数：
    '             hdc:  将要画边框的设备环境句柄。
    '             lprc: 指向包含矩形左上角和右上角逻辑坐标的结构RECT的指针?
    '             hbr:  用于画边框的画刷句柄?
    '　　返 回 值：如果函数调用成功，返回值非零；如果函数调用失败，返回值是0。
    '    备    注：由参数hbr定义的画刷必须是由CreateHatchBrush、CreatePatternBrush或CreateSolidBrush创建的，或者是由使用GetStockObject获得的。
    '              如果RECT结构中的底部成员的值少于或等于顶部成员，或右部成员少于或等于左部成员，此函数画不了矩形
    '【API函数--DeleteObject】
    '    函数原型：
    '    函数功能：用这个函数删除GDI对象，比如画笔、刷子、字体、位图、区域以及调色板等等。对象使用的所有系统资源都会被释放
    '    参    数：hObject:Long，一个GDI对象的句柄
    '　　返 回 值：Long，非零表示成功，零表示失败
    '    备    注：不要删除一个已选入设备场景的画笔、刷子或位图。如删除以位图为基础的阴影（图案）刷子，位图不会由这个函数删除——只有刷子被删掉
    
End Sub

Private Sub DrawSortArrow(lX As Long, lY As Long, lWidth As Long, lStep As Long, nOrientation As lgSortTypeEnum)
    Dim hPenOld As Long
    Dim hPen As Long
    Dim lCount As Long
    Dim lVerticalChange As Long
    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    
    hPen = CreatePen(0, 1, TranslateColor(vbButtonShadow))
    hPenOld = SelectObject(hdc, hPen)
    
    If nOrientation = lgSTDescending Then
        lVerticalChange = -1
        lY = lY + lStep - 1
    Else
        lVerticalChange = 1
    End If
    
    X1 = lX
    X2 = lWidth
    Y1 = lY
    
    MoveTo hdc, X1, Y1, ByVal 0&
    
    For lCount = 1 To lStep
        LineTo hdc, X1 + X2, Y1
        X1 = X1 + 1
        Y1 = Y1 + lVerticalChange
        X2 = X2 - 2
        MoveTo hdc, X1, Y1, ByVal 0&
    Next lCount
    
    Call SelectObject(hdc, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)
    If mWindowsNT Then
        DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hdc, lpString, nCount, lpRect, wFormat
    End If
    '【知识点】
    '【API函数--DrawTextW】
    '    函数原型：
    '    函数功能：
    '    参    数：
    '　　返 回 值：
    '    备    注：
    '【API函数--DrawTextA】
    '    函数原型：
    '    函数功能：
    '    参    数：
    '　　返 回 值：
    '    备    注：
End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
    Dim lResult As Long
    
    On Error GoTo DrawThemeError
    
    If mWindowsXP Then
        hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
        If (hTheme) Then
            
            lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, rtRect, rtRect)
            DrawTheme = (lResult = 0)
        Else
            DrawTheme = False
        End If
        
        If CloseTheme Then Call CloseThemeData(hTheme)
    End If
    Exit Function
    '知识点：
    '【API函数--OpenThemeData】
    '    函数原型：
    '    函数功能：
    '    参    数：
    '　　返 回 值：
    '    备    注：
    '【API函数--DrawThemeBackground】
    '    函数原型：
    '    函数功能：
    '    参    数：
    '　　返 回 值：
    '    备    注：
    '【API函数--CloseThemeData】
    '    函数原型：
    '    函数功能：
    '    参    数：
    '　　返 回 值：
    '    备    注：
DrawThemeError:
    DrawTheme = False
End Function

Public Property Get Editable() As Boolean
Attribute Editable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
    
    PropertyChanged "Editable"
End Property

Public Sub EditCell(ByVal Row As Long, ByVal Col As Long)
    '#############################################################################################################################
'Purpose: Used to start an Edit. Note the RequestEdit event. This event allows
    'the Edit to be cancelled before anything visible occurs by setting the Cancel
    'flag.
    '#############################################################################################################################
    
    Dim bCancel As Boolean
    
    If mEditPending Then
        If Not UpdateCell() Then
            Exit Sub
        End If
    End If
    
    If IsEditable() And (mCols(mColPtr(Col)).nType <> lgBoolean) Then
        RaiseEvent RequestEdit(Row, Col, bCancel)
        If Not bCancel Then
            mEditCol = Col
            mEditRow = Row
            
            MoveEditControl mCols(mColPtr(mEditCol)).MoveControl
            
            'Check if an external Control is used.
            If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
                'Using internal TextBox
                With txtEdit
                    Select Case mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nAlignment
                    Case lgAlignCenterBottom, lgAlignCenterCenter, lgAlignCenterTop
                        .Alignment = vbCenter
                    Case lgAlignLeftBottom, lgAlignLeftCenter, lgAlignLeftTop
                        .Alignment = vbLeftJustify
                    Case Else
                        .Alignment = vbRightJustify
                    End Select
                    
                    '.BackColor = mBackColorEdit
                    .FontBold = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontBold
                    .FontItalic = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontItalic
                    .FontUnderline = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontUnderline
                    
                    .Text = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .Visible = True
                    .SetFocus
                End With
            Else
                On Local Error Resume Next
                
                With mCols(mColPtr(mEditCol)).EditCtrl
                    If UserControl.ContainerHwnd <> .Container.hWnd Then
                        mEditParent = UserControl.ContainerHwnd
                        SetParent .hWnd, UserControl.ContainerHwnd
                    Else
                        mEditParent = 0
                    End If
                    .Enabled = True
                    .Visible = True
                    .ZOrder
                    
                    Subclass_Start .hWnd
                    Call Subclass_AddMsg(.hWnd, WM_KILLFOCUS, MSG_AFTER)
                    
                    If TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox Then
                        SendMessageAsLong mCols(mColPtr(mEditCol)).EditCtrl.hWnd, CB_SHOWDROPDOWN, 1&, 0&
                    End If
                    
                    .SetFocus
                End With
                
                On Local Error GoTo 0
            End If
            
            mEditPending = True
        End If
    End If
End Sub

Public Property Get EditType() As lgEditTypeEnum
Attribute EditType.VB_ProcData.VB_Invoke_Property = ";Behavior"
    EditType = mEditType
End Property

Public Property Let EditType(ByVal NewValue As lgEditTypeEnum)
    mEditType = NewValue
    
    PropertyChanged "EditType"
End Property

'功能：从指定列查找与所要搜索的文本相匹配的单元格
Public Function FindItem(ByVal SearchText As String, Optional ByVal SearchColumn As Long = 0, Optional SearchMode As lgSearchModeEnum = lgSMEqual, Optional MatchCase As Boolean) As Long
    'Search the specified Column for a Cell that matches the search text
    Dim lCount As Long
    Dim sCellText As String
    
    FindItem = -1
    
    '如果指定了搜索列和搜索内容
    If (SearchColumn >= 0) And (Len(SearchText) > 0) Then
        If Not MatchCase Then
            SearchText = UCase$(SearchText)
        End If
        
        For lCount = LBound(mItems) To mItemCount
            '如果严格匹配，则取出单元格的原始文本；否则取出原始文本后转换为大些
            If MatchCase Then
                sCellText = mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue
            Else
                sCellText = UCase$(mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue)
            End If
            
            '查找模式
            Select Case SearchMode
            Case lgSMEqual                                                      '相等
                If sCellText = SearchText Then
                    FindItem = lCount
                    Exit For
                End If
                
            Case lgSMGreaterEqual                                               '大于等于
                If sCellText >= SearchText Then
                    FindItem = lCount
                    Exit For
                End If
                
            Case lgSMLike                                                       '类似
                If sCellText Like SearchText & "*" Then
                    FindItem = lCount
                    Exit For
                End If
                
            Case lgSMNavigate                                                   '导航
                If Len(sCellText) > 0 Then
                    '单元格内容大于搜索内容且第一个字符相等
                    If (sCellText >= SearchText) And ((Mid$(sCellText, 1, 1)) = Mid$(SearchText, 1, 1)) Then
                        FindItem = lCount
                        Exit For
                    End If
                End If
            End Select
        Next lCount
    End If
End Function

Public Property Let FocusColor(ByVal NewValue As OLE_COLOR)
    mFocusColor = NewValue
    
    PropertyChanged "FocusColor"
End Property

Public Property Get FocusColor() As OLE_COLOR
Attribute FocusColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusColor = mFocusColor
End Property

Public Property Get SelectMode() As SelectModeEnum
Attribute SelectMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SelectMode = mSelectMode
End Property

Public Property Let SelectMode(ByVal NewValue As SelectModeEnum)
    mSelectMode = NewValue
    DisplayChange
    
    PropertyChanged "SelectMode"
End Property

Public Property Get FocusStyle() As FocusStyleEnum
Attribute FocusStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusStyle = mFocusStyle
End Property

Public Property Let FocusStyle(ByVal NewValue As FocusStyleEnum)
    mFocusStyle = NewValue
    DisplayChange
    
    PropertyChanged "FocusStyle"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    Set UserControl.Font = mFont
    
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    PropertyChanged "ForeColor"
End Property

Public Property Get SelectForeColor() As OLE_COLOR
Attribute SelectForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SelectForeColor = mSelectForeColor
End Property

Public Property Let SelectForeColor(ByVal lNewValue As OLE_COLOR)
    mSelectForeColor = lNewValue
    DisplayChange
    
    PropertyChanged "SelectForeColor"
End Property

Public Property Get ForeColorTotals() As OLE_COLOR
Attribute ForeColorTotals.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorTotals = mForeColorTotals
End Property

Public Property Let ForeColorTotals(ByVal NewValue As OLE_COLOR)
    mForeColorTotals = NewValue
    DisplayChange
    
    PropertyChanged "ForeColorTotals"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FullRowSelect = mFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal NewValue As Boolean)
    mFullRowSelect = NewValue
    DisplayChange
    
    PropertyChanged "FullRowSelect"
End Property

Private Function GetColFromX(x As Single) As Long
    Dim lX As Long
    Dim lCol As Long
    
    GetColFromX = -1
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
        With mCols(mColPtr(lCol))
            If .bVisible Then
                If (x > lX) And (x <= lX + .lWidth) Then
                    GetColFromX = lCol
                    Exit For
                End If
                
                lX = lX + .lWidth
            End If
        End With
    Next lCol
End Function

Private Function GetColumnHeadingHeight() As Long
    '#############################################################################################################################
'Purpose: Return Height of Header Row
    '#############################################################################################################################
    
    Dim lHeight As Long
    
    With UserControl
        lHeight = .TextHeight("A") + 4
        If GetRowHeight() > lHeight Then
            GetColumnHeadingHeight = GetRowHeight()
        Else
            GetColumnHeadingHeight = lHeight
        End If
    End With
End Function

Private Function GetFlag(ByVal nFlags As Integer, nFlag As lgFlagsEnum) As Boolean
    '#############################################################################################################################
'Purpose: Gets information by bit flags
    '#############################################################################################################################
    
    If nFlags And nFlag Then
        GetFlag = True
    End If
End Function

Private Function GetRowFromY(y As Single) As Long
    '#############################################################################################################################
'Purpose: Return Row from mouse position
    '#############################################################################################################################
    
    Dim lColumnHeadingHeight As Long
    Dim lRow As Long
    Dim lStart As Long
    Dim lY As Long
    
    'Are we below Header?
    If mColumnHeaders Then
        lColumnHeadingHeight = GetColumnHeadingHeight()
        If y <= lColumnHeadingHeight Then
            GetRowFromY = -1
            Exit Function
        End If
    End If
    
    lY = lColumnHeadingHeight
    lStart = SBValue(efsVertical)
    
    For lRow = lStart To mItemCount
        lY = lY + mItems(mRowPtr(lRow)).lHeight
        
        If lY >= y Then
            Exit For
        End If
    Next lRow
    
    If lRow <= mItemCount Then
        GetRowFromY = lRow
    Else
        GetRowFromY = -1
    End If
End Function


Private Function GetRowHeight() As Long
    '#############################################################################################################################
'Purpose: Return Row Height
    '#############################################################################################################################
    
    With UserControl
        If mRowHeight > 0 Then
            GetRowHeight = .ScaleY(mRowHeight, vbTwips, vbPixels)
        Else
            GetRowHeight = 300
        End If
    End With
End Function

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColor = mGridColor
End Property

Public Property Let GridColor(ByVal NewValue As OLE_COLOR)
    mGridColor = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "GridColor"
End Property

Public Property Get GridLines() As Boolean
Attribute GridLines.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridLines = mGridLines
End Property

Public Property Let GridLines(ByVal NewValue As Boolean)
    mGridLines = NewValue
    DisplayChange
    
    PropertyChanged "GridLines"
End Property

Public Property Let GridLineWidth(NewValue As Long)
    mGridLineWidth = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridLineWidth() As Long
Attribute GridLineWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridLineWidth = mGridLineWidth
End Property

Public Property Get HeadForeColor() As OLE_COLOR
    HeadForeColor = mHeadForeColor
End Property

Public Property Let HeadForeColor(ByVal NewValue As OLE_COLOR)
    mHeadForeColor = NewValue
    
    PropertyChanged "HeadForeColor"
End Property


Private Function IsValidRowCol(Row As Long, Col As Long) As Boolean
    IsValidRowCol = (Row > -1) And (Col > -1)
End Function

Public Property Get HotHeaderTracking() As Boolean
    HotHeaderTracking = mHotHeaderTracking
End Property

Public Property Let HotHeaderTracking(ByVal NewValue As Boolean)
Attribute HotHeaderTracking.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    mHotHeaderTracking = NewValue
    
    If Not NewValue Then
        DrawHeaderRow
    End If
    
    PropertyChanged "HotHeaderTracking"
End Property

Public Property Get ImageList() As Object
    Set ImageList = mImageList
End Property

Public Property Let ImageList(ByVal NewValue As Object)
    Set mImageList = NewValue
    If Not mImageList Is Nothing Then
        mImageListScaleMode = mImageList.Parent.ScaleMode
    End If
    
    DisplayChange
End Property

Private Function IsEditable() As Boolean
    If Not mLocked And mEditable Then IsEditable = (mItemCount >= 0)
End Function

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod        As Long
    Dim bLibLoaded  As Boolean
    
    If mUnicode Then
        hMod = GetModuleHandleW(sModule)
    Else
        hMod = GetModuleHandleA(sModule)
    End If
    
    If hMod = 0 Then
        If mUnicode Then
            hMod = LoadLibraryW(sModule)
        Else
            hMod = LoadLibraryA(sModule)
        End If
        If hMod Then
            bLibLoaded = True
        End If
    End If
    
    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    
    If bLibLoaded Then
        Call FreeLibrary(hMod)
    End If
End Function

Public Property Let ItemBackColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellBackColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Get ItemChecked(ByVal Index As Long) As Boolean
    ItemChecked = mItems(mRowPtr(Index)).nFlags And lgFLChecked
End Property

Public Property Let ItemChecked(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mRowPtr(Index)).nFlags, lgFLChecked, NewValue
    DrawGrid mRedraw
End Property

Public Property Get ItemCount() As Long
    ItemCount = mItemCount + 1
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = mItems(mRowPtr(Index)).lItemData
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal NewValue As Long)
    mItems(mRowPtr(Index)).lItemData = NewValue
End Property

Public Property Let ItemFontBold(ByVal Index As Long, ByVal NewValue As Boolean)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellFontBold(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Let ItemForeColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellForeColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Let ItemImage(ByVal Index As Long, NewValue As Variant)
    On Local Error GoTo ItemImageError
    
    If IsNumeric(NewValue) Then
        mItems(mRowPtr(Index)).lImage = NewValue
    Else
        mItems(mRowPtr(Index)).lImage = -mImageList.ListImages(NewValue).Index
    End If
    
    DrawGrid mRedraw
    Exit Property
    
ItemImageError:
    mItems(mRowPtr(Index)).lImage = 0
End Property

Public Property Get ItemImage(ByVal Index As Long) As Variant
    If mItems(mRowPtr(Index)).lImage >= 0 Then
        ItemImage = mItems(mRowPtr(Index)).lImage
    Else
        ItemImage = mImageList.ListImages(Abs(mItems(mRowPtr(Index)).lImage)).Key
    End If
End Property

Public Property Get ItemSelected(ByVal Index As Long) As Boolean
    ItemSelected = mItems(mRowPtr(Index)).nFlags And lgFLSelected
End Property

Public Property Let ItemSelected(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mRowPtr(Index)).nFlags, lgFLSelected, NewValue
    DrawGrid mRedraw
End Property

Public Property Get ItemTag(ByVal Index As Long) As String
    ItemTag = mItems(mRowPtr(Index)).sTag
End Property

Public Property Let ItemTag(ByVal Index As Long, ByVal NewValue As String)
    mItems(mRowPtr(Index)).sTag = NewValue
End Property

Public Function ItemsVisible() As Long
    Dim lBorderWidth As Long
    
    If mBorderStyle = 边框 Then lBorderWidth = 2
    With UserControl
        ItemsVisible = (.ScaleHeight - GetColumnHeadingHeight() - (lBorderWidth * 2)) / GetRowHeight()
    End With
End Function

Public Property Get MouseCol() As Long
    MouseCol = mMouseCol
End Property

Public Property Get MouseRow() As Long
    MouseRow = mMouseRow
End Property

'功能：移动编辑控件
'参数：lgMoveControlEnum
Private Sub MoveEditControl(ByVal MoveControl As lgMoveControlEnum)
    Dim r As RECT
    Dim lBorderWidth As Long
    Dim nScaleMode As ScaleModeConstants                                        '声明一个度量单位常量
    Dim lHeight As Long
    
    '给编辑列加上边框
    SetColRect mEditCol, r
    
    '如果列没有被截断，则令外边框的左边位置移至Grid的边线处
    If Not IsColumnTruncated(mEditCol) Then
        r.Left = r.Left + mGridLineWidth
    End If
    
    On Local Error Resume Next
    
    'Check if an external Control is used.
    '检查是否使用了外部控件（即除TextBox外的控件）
    If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
        'Using internal TextBox
        With txtEdit
            .Left = r.Left
            '顶部位置=所编辑行的行的顶部+表格线的宽度，即在表格线的下方
            .Top = RowTop(mEditRow) + mGridLineWidth
            '高度=所编行的高度-表格线的宽度
            .Height = mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth
            .Width = (r.Right - r.Left)
        End With
    Else
        nScaleMode = UserControl.Parent.ScaleMode
        If mBorderStyle = 边框 Then
            lBorderWidth = 2
        End If
        
        '如果所编辑列的编辑控件是ComboBox
        If (TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox) Then
            '设置控件的Left、Top、Width、Height
            With mCols(mColPtr(mEditCol)).EditCtrl
                'UserControl.Extender.Left 属性，可读可写的整型值，它使用容器的刻度单位指定控件左边缘相对于容器左边缘的位置。
                'UserControl.Extender.Top  属性，可读可写的整型值，它使用容器的刻度单位指定控件的上边缘相对于容器上边缘的位置。
                'ScaleX将宽度的度量单位从一种转换到另一种
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
                    .Left = ScaleX(r.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
                    .Top = ScaleY(RowTop(mEditRow) + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
                    .Width = ScaleX((r.Right - r.Left), vbPixels, nScaleMode)
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
'TwipsPerPixel:     返回水平 (TwipsPerPixelX) 或垂直 (TwipsPerPixelY) 度量的对象的每一像素中的缇数。
                    lHeight = mRowHeight / Screen.TwipsPerPixelX - mGridLineWidth - 4
                    Call SendMessageAsLong(.hWnd, CB_SETITEMHEIGHT, -1, ByVal lHeight)
                    Call SendMessageAsLong(.hWnd, CB_SETITEMHEIGHT, 0, ByVal lHeight)
                End If
            End With
        Else
            With mCols(mColPtr(mEditCol)).EditCtrl
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
                    .Left = ScaleX(r.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
                    .Top = ScaleY(RowTop(mEditRow) + mGridLineWidth + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
                    .Height = ScaleY(mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth, vbPixels, nScaleMode)
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
                    .Width = ScaleX((r.Right - r.Left), vbPixels, nScaleMode)
                End If
            End With
        End If
    End If
    
    On Local Error GoTo 0
End Sub

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MultiSelect = mMultiSelect
End Property

Public Property Let MultiSelect(ByVal NewValue As Boolean)
    mMultiSelect = NewValue
    
    If Not NewValue Then
        SetSelection False
        DisplayChange
    End If
    
    PropertyChanged "MultiSelect"
End Property

Private Function NavigateDown() As Long
    If mRow < mItemCount Then
        NavigateDown = mRow + 1
    Else
        NavigateDown = mRow
    End If
End Function

Private Function NavigateLeft() As Long
    If mCol > 0 Then
        NavigateLeft = mCol - 1
    Else
        NavigateLeft = mCol
    End If
End Function

Private Function NavigateRight() As Long
    If mCol < UBound(mCols) Then
        NavigateRight = mCol + 1
    Else
        NavigateRight = mCol
    End If
End Function

Private Function NavigateUp() As Long
    If mRow > 0 Then
        NavigateUp = mRow - 1
    Else
        NavigateUp = mRow
    End If
End Function

Private Property Get Orientation() As ScrollBarOrienationEnum
    SBOrientation = m_eOrientation
End Property

'@@@@@@@@@@@@@@@@@@@@@@[滚动条部分]@@@@@@@@@@@@@@@@@@@@-Start
Private Sub pSBClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If
        
        On Error GoTo 0
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pSBCreateScrollBar()
    Dim lR As Long
    Dim hParent As Long
    
    On Error Resume Next
    lR = InitialiseFlatSB(m_hWnd)
    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        SBStyle = m_eStyle
    End If
End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long
    
    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, Lo, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, Lo, tSI
    End If
End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long
    
    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, Lo, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
    End If
End Sub

Private Sub pSBSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

Private Property Get SBCanBeFlat() As Boolean
    SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub SBCreate(ByVal hWndA As Long)
    pSBClearUp
    m_hWnd = hWndA
    pSBCreateScrollBar
End Sub

Private Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBEnabled = m_bEnabledHorz
    Else
        SBEnabled = m_bEnabledVert
    End If
End Property

Private Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
    Dim Lo As Long
    Dim lF As Long
    
    Lo = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, Lo, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, Lo, lF
    End If
    
End Property

Private Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_PAGE
    SBLargeChange = tSI.nPage
End Property

Private Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO
    
    pSBGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    SBMax = tSI.nMax                                                            ' - tSI.nPage
End Property

Private Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = iMax + SBLargeChange(eBar)
    tSI.nMin = SBMin(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE
    SBMin = tSI.nMin
End Property

Private Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
    m_eOrientation = eOrientation
    pSBSetOrientation
End Property

Private Sub SBRefresh()
    EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
End Sub

Private Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SBSmallChange = m_lSmallChangeHorz
    Else
        SBSmallChange = m_lSmallChangeVert
    End If
End Property

Private Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Private Property Get SBStyle() As ScrollBarStyleEnum
    SBStyle = m_eStyle
End Property

Private Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
    Dim lR As Long
    If (m_bNoFlatScrollBars) Then
        ' can't do it..
        'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
        Exit Property
    Else
        If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
        End If
        If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
        End If
        'Debug.Print lR
        m_eStyle = eStyle
    End If
    
End Property

Private Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_POS
    SBValue = tSI.nPos
End Property

Private Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    
    If SBVisible(eBar) Then
        If (iValue <> SBValue(eBar)) Then
            tSI.nPos = iValue
            pSBLetSI eBar, tSI, SIF_POS
        End If
    End If
End Property

Private Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBVisible = m_bVisibleHorz
    Else
        SBVisible = m_bVisibleVert
    End If
End Property

Private Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Private Sub ScrollList(nDirection As Integer)
    Dim lCount As Long
    Dim lItemsVisible As Long
    
    mScrollAction = nDirection
    
    Do While mScrollAction = nDirection
        mScrollTick = GetTickCount()
        
        If nDirection = 1 Then
            If SBValue(efsVertical) > SBMin(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
                If mMultiSelect Then
                    SetFlag mItems(mRowPtr(SBValue(efsVertical))).nFlags, lgFLSelected, True
                Else
                    mRow = SBValue(efsVertical)
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        Else
            If SBValue(efsVertical) < SBMax(efsVertical) Then
                lItemsVisible = ItemsVisible()
                
                SBValue(efsVertical) = SBValue(efsVertical) + 1
                If mMultiSelect Then
                    For lCount = SBValue(efsVertical) To SBValue(efsVertical) + lItemsVisible
                        If lCount > mItemCount Then
                            Exit For
                        Else
                            SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, True
                        End If
                    Next lCount
                Else
                    mRow = SBValue(efsVertical) + (lItemsVisible - 1)
                    If mRow > mItemCount Then
                        mRow = mItemCount
                    End If
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        End If
        
        RaiseEvent SelectionChanged
        DrawGrid mRedraw
        RaiseEvent Scroll
        
        'Sleep 25
        'DoEvents
    Loop
End Sub

Public Property Get ScrollTrack() As Boolean
    ScrollTrack = mScrollTrack
End Property

Public Property Let ScrollTrack(ByVal NewValue As Boolean)
    mScrollTrack = NewValue
    
    PropertyChanged "ScrollTrack"
End Property
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-End


Public Property Get Redraw() As Boolean
Attribute Redraw.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal NewValue As Boolean)
    mRedraw = NewValue
    
    If mRedraw Then
        If mPendingScrollBar Then
            SetScrollBars
        End If
        If mPendingRedraw Then
            CreateRenderData
            DrawGrid mRedraw
        End If
    Else
        mPendingScrollBar = False
        mPendingRedraw = False
    End If
    
    PropertyChanged "Redraw"
End Property

Public Sub Refresh()
    CreateRenderData
    SetScrollBars
    DrawGrid True
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    Dim lCount As Long
    Dim lPosition As Long
    Dim bSelected As Boolean
    
    '#############################################################################################################################
    'See AddItem for details of the Arrays used
    '#############################################################################################################################
    
    'Note selected state before deletion
    '在删除前标记选中状态
    bSelected = mItems(mRowPtr(Index)).nFlags And lgFLSelected
    
    'Decrement the reference count on each cells format Entry
    '根据每个单元格的格式化入口，减少引用
    If mItemCount >= 0 Then
        For lCount = 0 To UBound(mCols)
            If mItems(Index).Cell(Count).nFormat >= 0 Then
                mCF(mItems(Index).Cell(lCount).nFormat).nCount = mCF(mItems(Index).Cell(lCount).nFormat).nCount - 1
            End If
        Next lCount
    End If
    
    lPosition = mRowPtr(Index)
    
    'Reset Item Data重置数据
    For lCount = mRowPtr(Index) To mItemCount - 1
        mItems(lCount) = mItems(lCount + 1)
    Next lCount
    
    'Adjust Index调整索引
    For lCount = Index To mItemCount - 1
        mRowPtr(lCount) = mRowPtr(lCount + 1)
    Next lCount
    
    'Validate Indexes for Items after deleted Item
    '删除选项后验证索引
    For lCount = 0 To mItemCount - 1
        If mRowPtr(lCount) > lPosition Then
            mRowPtr(lCount) = mRowPtr(lCount) - 1
        End If
    Next lCount
    
    mItemCount = mItemCount - 1
    
    If mItemCount < 0 Then
        Clear
    Else
        If (mItemCount + mCacheIncrement) < UBound(mItems) Then
            '在保留原数据的情况下重新定义数组
            ReDim Preserve mItems(mItemCount)
            ReDim Preserve mRowPtr(mItemCount)
        End If
        
        If bSelected Then
            If mMultiSelect Then
                RaiseEvent SelectionChanged
            ElseIf Index > mItemCount Then
                SetFlag mItems(mRowPtr(mItemCount)).nFlags, lgFLSelected, True
            ElseIf mItemCount >= 0 Then
                SetFlag mItems(mRowPtr(Index)).nFlags, lgFLSelected, True
            End If
        End If
        
        If Index > mItemCount Then
            SetRowCol mRow - 1, mCol
        End If
    End If
    
    DisplayChange
    
    RaiseEvent ItemCountChanged
End Sub

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal NewValue As Long)
    If SetRowCol(NewValue, mCol) Then DrawGrid mRedraw
End Property

Public Property Get RowHeight() As Long
    RowHeight = mRowHeight
End Property

Public Property Let RowHeight(ByVal NewValue As Long)
    mRowHeight = NewValue
    DisplayChange
    
    PropertyChanged "RowHeight"
End Property

Public Function RowTop(Index As Long) As Long
    Dim lRow As Long
    Dim lStart As Long
    Dim lY As Long
    
    lStart = SBValue(efsVertical)
    
    If Index >= lStart Then
        lY = GetColumnHeadingHeight()
        
        For lRow = lStart To Index - 1
            lY = lY + mItems(mRowPtr(lRow)).lHeight
        Next lRow
    Else
        lY = -1
    End If
    
    RowTop = lY
End Function



Public Property Get AlphaBlendSelection() As Boolean
    AlphaBlendSelection = mAlphaBlendSelection
End Property

Public Property Let AlphaBlendSelection(ByVal NewValue As Boolean)
    mAlphaBlendSelection = NewValue
    DisplayChange
    
    PropertyChanged "AlphaBlendSelection"
End Property


Public Function SelectedCount() As Long
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgFLSelected Then SelectedCount = SelectedCount + 1
    Next lCount
End Function

Private Function SetColRect(ByVal Index As Long, r As RECT)
    Dim lCol As Long
    Dim lCount As Long
    Dim lScrollValue As Long
    Dim lX As Long
    
    lScrollValue = SBValue(efsHorizontal)
    
    If Index < lScrollValue Then
        r.Left = -1
    Else
        For lCol = lScrollValue To Index - 1
            If mCols(mColPtr(lCol)).bVisible Then
                lX = lX + mCols(mColPtr(lCol)).lWidth
                lCount = lCount + 1
            End If
        Next lCol
        
        If IsColumnTruncated(Index) Then
            r.Left = mR.LeftText
            r.Right = r.Left + (mCols(mColPtr(Index)).lWidth - mR.LeftText)
        Else
            r.Left = lX
            r.Right = r.Left + mCols(mColPtr(Index)).lWidth
        End If
    End If
End Function

Private Sub SetItemRect(ByVal Row As Long, ByVal Col As Long, lY As Long, r As RECT, ItemType As lgRectTypeEnum)
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim nAlignment As lgAlignmentEnum
    
    Select Case ItemType
    Case lgRTColumn
        nAlignment = mCols(Col).nAlignment
        
    Case lgRTCheckBox
        nAlignment = mCols(Col).nAlignment
        lHeight = mR.CheckBoxSize
        lWidth = mR.CheckBoxSize
        
    Case lgRTImage
        nAlignment = mCols(Col).nImageAlignment
        lHeight = mR.ImageHeight
        lWidth = mR.ImageWidth
    End Select
    
    Select Case nAlignment
    Case lgAlignLeftTop
        lLeft = mCols(Col).lX + 1
        lTop = lY + 2
    Case lgAlignLeftCenter
        lLeft = mCols(Col).lX + 1
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
    Case lgAlignLeftBottom
        lLeft = mCols(Col).lX + 1
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
    Case lgAlignCenterTop
        lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
        lTop = lY + 2
    Case lgAlignCenterCenter
        lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
    Case lgAlignCenterBottom
        lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
    Case lgAlignRightTop
        lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
        lTop = lY + 2
    Case lgAlignRightCenter
        lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
    Case lgAlignRightBottom
        lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
        lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
    End Select
    
    Call SetRect(r, lLeft, lTop, lLeft + lWidth, lTop + lHeight)
End Sub

Private Sub SetFlag(nFlags As Integer, nFlag As lgFlagsEnum, bValue As Boolean)
    If bValue Then
        nFlags = (nFlags Or nFlag)
    Else
        nFlags = (nFlags And Not (nFlag))
    End If
End Sub

Private Sub SetRedrawState(bState As Boolean)
    '#############################################################################################################################
'Purpose: Used to prevent Internal Redraws while preserving User Controlled Redraw state
    '
    'bDrawLocked used to prevent nested Calls to Lock Redraw
    '#############################################################################################################################
    
    Static bDrawLocked As Boolean
    Static bOriginalRedraw As Boolean
    
    If bState Then
        bDrawLocked = False
        mRedraw = bOriginalRedraw
    ElseIf Not bDrawLocked Then
        bDrawLocked = True
        bOriginalRedraw = mRedraw
        mRedraw = False
    End If
End Sub


Private Function SetRowCol(lRow As Long, lCol As Long, Optional bSetScroll As Boolean) As Boolean
    '#############################################################################################################################
'Purpose: To update current Row/Col and fire Events if necessary
    '#############################################################################################################################
    
    Dim r As RECT
    Dim lCount As Long
    
    If (mCol <> lCol) Or (mRow <> lRow) Then
        mCol = lCol
        mRow = lRow
        
        RaiseEvent RowColChanged
        
        'Do we need to change Bars?
        If bSetScroll Then
            SetColRect mCol, r
            
            'Scroll to make Column visible
            If r.Left < 0 Then
                For lCount = SBValue(efsHorizontal) To SBMin(efsHorizontal) Step -1
                    If r.Left > 0 Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) - 1
                    SetColRect mCol, r
                Next lCount
            Else
                For lCount = SBValue(efsHorizontal) To SBMax(efsHorizontal)
                    If r.Left + mCols(mCol).lWidth < UserControl.ScaleWidth Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) + 1
                    SetColRect mCol, r
                Next lCount
            End If
            
            If SBValue(efsHorizontal) = SBMin(efsHorizontal) Then
                SetScrollBars
            End If
            
            If mRow < SBValue(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
            ElseIf mRow > SBValue(efsVertical) + (ItemsVisible() - 1) Then
                SBValue(efsVertical) = SBValue(efsVertical) + 1
            End If
            
            RaiseEvent Scroll
        End If
        
        SetRowCol = True
    End If
End Function

Private Sub SetScrollBars()
    '#############################################################################################################################
'Purpose: Sets the visibilty of scroll bars and sets max scroll values
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lRow As Long
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lVSB As Long
    Dim bHVisible As Boolean
    Dim bVVisible As Boolean
    
    If m_hWnd <> 0 Then
        '#############################################################################################################################
        'Calculate total width of columns
        For lCol = LBound(mCols) To UBound(mCols)
            If mCols(mColPtr(lCol)).bVisible Then
                lWidth = lWidth + mCols(mColPtr(lCol)).lWidth
            End If
        Next lCol
        
        If (lWidth > UserControl.ScaleWidth) Then
            SBMax(efsHorizontal) = UBound(mCols) - 1
            bHVisible = True
        Else
            SBMax(efsHorizontal) = UBound(mCols)
            bHVisible = (SBValue(efsHorizontal) > SBMin(efsHorizontal))
        End If
        
        '#############################################################################################################################
        'Calculate total height of rows
        lHeight = GetColumnHeadingHeight()
        For lRow = LBound(mItems) To mItemCount
            lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
        Next lRow
        
        If lHeight > UserControl.ScaleHeight Then
            'Adjust scrollbar to best-fit Rows to Grid
            lHeight = GetColumnHeadingHeight()
            For lRow = mItemCount To LBound(mItems) Step -1
                lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
                
                If lHeight > UserControl.ScaleHeight Then
                    Exit For
                End If
                
                lVSB = lVSB + 1
            Next lRow
            
            SBMax(efsVertical) = mItemCount - lVSB
            bVVisible = True
        Else
            SBMax(efsVertical) = mItemCount
        End If
        
        '#############################################################################################################################
        'If SBVisible(efsHorizontal) <> bHVisible Then
        SBVisible(efsHorizontal) = bHVisible
        'End If
        'If SBVisible(efsVertical) <> bVVisible Then
        SBVisible(efsVertical) = bVVisible
        'End If
    End If
End Sub

Private Function SetSelection(bState As Boolean, Optional lFromRow As Long = -1, Optional lToRow As Long = -1) As Boolean
    Dim lCount As Long
    Dim lStep As Long
    Dim bSelectionChanged As Boolean
    
    If lFromRow = -1 Then
        lFromRow = LBound(mItems)
    End If
    
    If lToRow = -1 Then
        lToRow = UBound(mItems)
    End If
    
    If lFromRow >= lToRow Then
        lStep = -1
    Else
        lStep = 1
    End If
    
    For lCount = lFromRow To lToRow Step lStep
        If (mItems(mRowPtr(lCount)).nFlags And lgFLSelected) <> bState Then
            SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, bState
            bSelectionChanged = True
        End If
    Next lCount
    
    SetSelection = bSelectionChanged
End Function

'@@@@@@@@@@@@@@@@@@@@@@[值排序部分]@@@@@@@@@@@@@@@@@@@@-Start
Private Sub SortArrayString(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub
    
    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst
    
    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue > mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
        Else
            bSwap = mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue < mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex
    
    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayString lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayString lBoundary + 1, lLast, lSortColumn, nSortType
End Sub

Private Sub SortArrayDate(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bIsDate(1) As Boolean
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub
    
    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst
    
    For lIndex = lFirst + 1 To lLast
        bIsDate(0) = IsDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue)
        bIsDate(1) = IsDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        
        If nSortType = 0 Then
            If Not bIsDate(0) Then
                bSwap = False
            ElseIf Not bIsDate(1) Then
                bSwap = True
            Else
                bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If
        Else
            If Not bIsDate(0) Then
                bSwap = True
            ElseIf Not bIsDate(1) Then
                bSwap = False
            Else
                bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex
    
    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayDate lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayDate lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayNumeric(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub
    
    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst
    
    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = Val(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > Val(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        Else
            bSwap = Val(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < Val(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex
    
    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayNumeric lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayNumeric lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayCustom(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub
    
    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst
    
    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            RaiseEvent CustomSort(True, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
        Else
            RaiseEvent CustomSort(False, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex
    
    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayCustom lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayCustom lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayBool(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub
    
    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst
    
    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) > GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
        Else
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) < GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex
    
    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayBool lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayBool lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    Select Case mCols(lSortColumn).nType
    Case lgBoolean
        SortArrayBool lFirst, lLast, lSortColumn, nSortType
    Case lgDate
        SortArrayDate lFirst, lLast, lSortColumn, nSortType
    Case lgNumeric
        SortArrayNumeric lFirst, lLast, lSortColumn, nSortType
    Case lgCustom
        SortArrayCustom lFirst, lLast, lSortColumn, nSortType
    Case Else
        SortArrayString lFirst, lLast, lSortColumn, nSortType
    End Select
End Sub
Private Sub SortSubList()
    Dim lCount As Long
    Dim lStartSort As Long
    Dim bDifferent As Boolean
    Dim sMajorSort As String
    
    If mSortSubColumn > -1 Then
        lStartSort = LBound(mItems)
        For lCount = LBound(mItems) To mItemCount
            bDifferent = mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue <> sMajorSort
            If bDifferent Or lCount = mItemCount Then
                If lCount > 1 Then
                    If lCount - lStartSort > 1 Then
                        If lCount = mItemCount And Not bDifferent Then
                            SortArray lStartSort, lCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        Else
                            SortArray lStartSort, lCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        End If
                    End If
                    lStartSort = lCount
                End If
                
                sMajorSort = mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue
            End If
        Next lCount
    End If
End Sub
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-End

Private Sub SwapLng(Value1 As Long, Value2 As Long)
    Static lTemp As Long
    
    lTemp = Value1
    Value1 = Value2
    Value2 = lTemp
End Sub

Private Function ToggleEdit() As Boolean
    If IsEditable() Then
        ToggleEdit = True
        
        If mEditPending Then
            UpdateCell
        ElseIf (mRow <> -1) And (mCol <> -1) Then
            EditCell mRow, mCol
        End If
    End If
End Function

'属性：设置最顶行
Public Property Let TopRow(ByVal NewValue As Long)
    If NewValue > SBMax(efsVertical) Then
        SBValue(efsVertical) = SBMax(efsVertical)
    Else
        SBValue(efsVertical) = NewValue
    End If
    
    SetRowCol NewValue, mCol, True
    DrawGrid mRedraw
End Property
'功能：跟踪鼠标离开
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub
'功能：颜色转换
Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then TranslateColor = &HFFFF
End Function

Private Function UpdateCell() As Boolean
    Dim bCancel As Boolean
    Dim bRequestUpdate As Boolean
    Dim sNewValue As String
    
    If mEditPending Then
        If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
            bRequestUpdate = (mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue <> txtEdit.Text)
            sNewValue = txtEdit.Text
        Else
            bRequestUpdate = True
        End If
        
        If bRequestUpdate Then
            RaiseEvent RequestUpdate(mEditRow, mEditCol, sNewValue, bCancel)
        End If
        
        If Not bCancel Then
            SetRedrawState False
            
            If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
                txtEdit.Visible = False
            Else
                On Local Error Resume Next
                With mCols(mColPtr(mEditCol)).EditCtrl
                    If mEditParent <> 0 Then SetParent .hWnd, mEditParent
                    Subclass_Stop .hWnd
                    .Visible = False
                End With
                On Local Error GoTo 0
            End If
            
            mEditPending = False
            If bRequestUpdate Then
                mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue = sNewValue
                SetFlag mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags, lgFLChanged, True
                DisplayChange
            End If
            
            SetRedrawState True
            DrawGrid True
        End If
    End If
    
    UpdateCell = Not bCancel
End Function

'@@@@@@@@@@@@@@@@@@@@@@[用户控件部分]@@@@@@@@@@@@@@@@@@@@-Start
Private Sub UserControl_Click()
    If (mEditType And MouseClick) And (mMouseRow > -1) Then ToggleEdit
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If (mEditType And MouseDblClick) And (mMouseRow > -1) Then ToggleEdit
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
    mClipRgn = CreateRectRgn(0, 0, 0, 0)
    mUnicode = IsWindowUnicode(UserControl.hWnd)
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    mWindowsNT = ((OS.dwPlatformId And 2) = 2)
    If (OS.dwMajorVersion > 5) Then
        mWindowsXP = True
    ElseIf (OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1) Then
        mWindowsXP = True
    End If
    
    Set txtEdit = UserControl.Controls.Add("VB.TextBox", "txtEdit")
    With txtEdit
        .BorderStyle = 0
        .Visible = False
    End With
    
    ReDim mCols(0)
    ReDim mColPtr(0)
    Clear
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font
    
    'Appearance Properties
    mSelectBackColor = &HFCE1CB
    mForeColor = vbWindowText
    mHeadForeColor = vbWindowText
    mSelectForeColor = vbHighlightText
    mForeColorTotals = vbRed
    
    mFocusColor = vbBlue
    mGridColor = &HC0C0C0
    
    mAlphaBlendSelection = False
    mDisplayEllipsis = True
    mSelectMode = [行]
    mFocusStyle = Heavy
    mGridLines = True
    mGridLineWidth = 1
    
    'Behaviour Properties
    mAllowResizing = Resize
    mBorderStyle = 边框
    mCheckboxes = False
    mColumnDrag = False
    mColumnHeaders = True
    mColumnSort = False
    mEditable = False
    mEditType = EnterKey
    mFullRowSelect = True
    mHotHeaderTracking = True
    mMultiSelect = False
    mRedraw = True
    mScrollTrack = True
    mTrackEdits = False
    
    'Miscellaneous Properties
    mCacheIncrement = 10
    mEnabled = True
    mLocked = False
    mRowHeight = 300
    
    UserControl.BorderStyle = mBorderStyle
    CreateRenderData
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lNewCol As Long
    Dim lNewRow As Long
    Dim bClearSelection As Boolean
    Dim bRedraw As Boolean
    
    lNewCol = mCol
    lNewRow = mRow
    
    SetRedrawState False
    
    'Used to determine if selected Items need to be cleared
    bClearSelection = True
    
    Select Case KeyCode
    Case vbKeyReturn, vbKeyEscape                                               'Allow escape to abort editing
        bClearSelection = False
        
        If (mEditType And EnterKey) Then
            If KeyCode = vbKeyEscape Then
                txtEdit.Visible = False
                mEditPending = False
            Else
                If ToggleEdit() Then KeyCode = 0
            End If
        End If
    Case vbKeyF2
        bClearSelection = False
        
        If (mEditType And F2Key) Then
            If ToggleEdit() Then KeyCode = 0
        End If
    Case vbKeySpace
        bClearSelection = False
        If mCheckboxes Then
            mIgnoreKeyPress = True
            
            SetFlag mItems(mRowPtr(mRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mRow)).nFlags, lgFLChecked)
            RaiseEvent ItemChecked(mRow)
            KeyCode = 0
        End If
    Case vbKeyA
        bClearSelection = False
        
        If (Shift And vbCtrlMask) And mMultiSelect Then
            mIgnoreKeyPress = True
            
            SetSelection True
            RaiseEvent SelectionChanged
            KeyCode = 0
        End If
    Case vbKeyUp
        If (Shift And vbShiftMask) And mMultiSelect Then bClearSelection = False
        If UpdateCell() Then
            lNewRow = NavigateUp()
            KeyCode = 0
        End If
    Case vbKeyDown
        If (Shift And vbShiftMask) And mMultiSelect Then bClearSelection = False
        If UpdateCell() Then
            lNewRow = NavigateDown()
            KeyCode = 0
        End If
    Case vbKeyLeft
        If Not mEditPending Then
            lNewCol = NavigateLeft()
            KeyCode = 0
        End If
    Case vbKeyRight
        If Not mEditPending Then
            lNewCol = NavigateRight()
            KeyCode = 0
        End If
    Case vbKeyPageUp
        If UpdateCell() Then
            If mRow > 0 Then
                lNewRow = (mRow - ItemsVisible()) + 1
                If lNewRow < 0 Then lNewRow = 0
                SBValue(efsVertical) = lNewRow
            End If
            KeyCode = 0
        End If
    Case vbKeyPageDown
        If UpdateCell() Then
            If mRow < mItemCount Then
                lNewRow = (mRow + ItemsVisible()) - 1
                If lNewRow > mItemCount Then lNewRow = mItemCount
                SBValue(efsVertical) = lNewRow
            End If
            
            KeyCode = 0
        End If
    Case vbKeyHome
        If Shift And vbShiftMask Then
            If UpdateCell() Then
                If mMultiSelect Then
                    bClearSelection = False
                    
                    SetSelection False
                    SetSelection True, 1, mRow
                    RaiseEvent SelectionChanged
                End If
                
                lNewRow = 0
                
                SBValue(efsVertical) = SBMin(efsVertical)
                KeyCode = 0
            End If
        ElseIf Shift And vbCtrlMask Then
            If UpdateCell() Then
                lNewRow = 0
                
                SBValue(efsVertical) = SBMin(efsVertical)
                KeyCode = 0
            End If
        ElseIf Not mEditPending Then
            lNewCol = 0
            
            SBValue(efsHorizontal) = SBMin(efsHorizontal)
            KeyCode = 0
        End If
    Case vbKeyEnd
        If Shift And vbShiftMask Then
            If UpdateCell() Then
                If mMultiSelect Then
                    bClearSelection = False
                    
                    SetSelection False
                    SetSelection True, mRow, mItemCount
                    RaiseEvent SelectionChanged
                End If
                
                lNewRow = mItemCount
                
                SBValue(efsVertical) = SBMax(efsVertical)
                KeyCode = 0
            End If
        ElseIf Shift And vbCtrlMask Then
            If UpdateCell() Then
                lNewRow = mItemCount
                
                SBValue(efsVertical) = SBMax(efsVertical)
                KeyCode = 0
            End If
        ElseIf Not mEditPending Then
            lNewCol = UBound(mCols)
            
            SBValue(efsHorizontal) = SBMax(efsHorizontal)
            KeyCode = 0
        End If
    End Select
    SetRedrawState True
    If KeyCode = 0 Then
        'Do we want to clear selection?
        If bClearSelection And (mRow <> lNewRow) Then bRedraw = SetSelection(False)
        If Not mItems(mRowPtr(lNewRow)).nFlags And lgFLSelected Then
            bRedraw = True
            SetFlag mItems(mRowPtr(lNewRow)).nFlags, lgFLSelected, True
            RaiseEvent SelectionChanged
        End If
        If bRedraw Or SetRowCol(lNewRow, lNewCol, True) Then DrawGrid mRedraw
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Static lTime As Long
    Static sCode As String
    
    Dim lResult As Long
    Dim bEatKey As Boolean
    
    If mEnabled Then
        'Used to prevent a beep
        If (mEditType And EnterKey) And (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
            KeyAscii = 0
            bEatKey = True
        End If
        
        If Not bEatKey Then RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mIgnoreKeyPress = False
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As RECT
    Dim lMCol As Long
    Dim bCancel As Boolean
    Dim bProcessed As Boolean
    Dim bRedraw As Boolean
    Dim bSelectionChanged As Boolean
    Dim bState As Boolean
    
    If Not mLocked And (Button <> 0) And (mItemCount >= 0) Then
        mScrollAction = 0
        
        lMCol = GetColFromX(x)
        
        mMouseDownRow = GetRowFromY(y)
        mMouseDownX = x
        
        If Button = vbLeftButton Then
            Call SetCapture(UserControl.hWnd)
            mMouseDown = True
            
            If y < mR.HeaderHeight Then
                If (UserControl.MousePointer <> vbSizeWE) Then
                    mMouseDownCol = lMCol
                    If mMouseDownCol <> -1 Then
                        With UserControl
                            DrawHeader mMouseCol, lgDown
                            .Refresh
                        End With
                    End If
                End If
            ElseIf mMouseDownRow > -1 Then
                If UpdateCell() Then
                    If mCheckboxes And (x <= 15) Then
                        bRedraw = True
                        mMouseDown = False
                        
                        SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked)
                        RaiseEvent ItemChecked(mMouseDownRow)
                    Else
                        If lMCol > -1 Then
                            If IsEditable() And mCols(mColPtr(lMCol)).nType = lgBoolean Then
                                SetItemRect mRowPtr(mMouseDownRow), mColPtr(lMCol), RowTop(mMouseDownRow), r, lgRTCheckBox
                                
                                If (x >= r.Left) And (y >= r.Top) And (x <= r.Left + mR.CheckBoxSize) And (y <= r.Top + mR.CheckBoxSize) Then
                                    bRedraw = True
                                    RaiseEvent RequestEdit(mMouseDownRow, lMCol, bCancel)
                                    
                                    If Not bCancel Then
                                        bState = (mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lMCol)).nFlags And lgFLChecked)
                                        SetFlag mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lMCol)).nFlags, lgFLChecked, Not bState
                                    End If
                                End If
                            End If
                        End If
                        
                        If Not bProcessed Then
                            bState = (mItems(mRowPtr(mMouseDownRow)).nFlags And lgFLSelected)
                            
                            If mMultiSelect Then
                                If (Shift And vbShiftMask) Then
                                    bSelectionChanged = SetSelection(False) Or SetSelection(True, mRow, mMouseDownRow)
                                ElseIf Shift And vbCtrlMask Then
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                                    bSelectionChanged = True
                                Else
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                                    bSelectionChanged = True
                                End If
                            Else
                                If Shift And vbCtrlMask Then
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                                    bSelectionChanged = True
                                ElseIf Not bState Then
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                                    bSelectionChanged = True
                                End If
                            End If
                        End If
                        
                        bRedraw = bRedraw Or SetRowCol(mMouseDownRow, lMCol)
                    End If
                    
                    If bRedraw Then DrawGrid mRedraw
                End If
            End If
        Else                                                                    ' Right Button
            If mMouseDownRow > -1 Then
                If UpdateCell() Then
                    SetRowCol mMouseDownRow, lMCol
                    bSelectionChanged = SetSelection(False) Or SetSelection(True, mMouseDownRow, mMouseDownRow)
                    DrawGrid mRedraw
                End If
            End If
        End If
        
        If bSelectionChanged Then RaiseEvent SelectionChanged
    End If
    
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static lResizeX As Long
    
    Dim r As RECT
    Dim lCount As Long, lWidth As Long
    Dim nMove As lgMoveControlEnum
    Dim nPointer As Integer
    Dim bSelectionChanged As Boolean
    
    If Not mLocked And (mItemCount >= 0) Then
        mMouseCol = GetColFromX(x)
        mMouseRow = GetRowFromY(y)
        'Header button tracking
        If mMouseDownCol <> -1 Then
            If (mMouseDownCol = mMouseCol) And (MouseRow = -1) Then
                DrawHeader mMouseCol, lgDown
            Else
                DrawHeader mMouseDownCol, lgNormal
            End If
            UserControl.Refresh
        End If
        
        'Hot tracking
        If mHotHeaderTracking And (Button = 0) Then
            If y < mR.HeaderHeight Then
                'Do we need to draw a new "hot" header?
                If (mMouseCol <> mHotColumn) Then
                    DrawHeaderRow
                    DrawHeader mMouseCol, lgHot
                    mHotColumn = mMouseCol
                End If
            ElseIf (mHotColumn <> -1) Then
                'We have a previous "hot" header to clear
                DrawHeaderRow
            End If
        End If
        
        If (Button = vbLeftButton) Then
            If (mResizeCol >= 0) Then
                'We are resizing a Column
                lWidth = (x - lResizeX)
                If lWidth > 1 Then
                    mCols(mColPtr(mResizeCol)).lWidth = lWidth
                    mCols(mColPtr(mResizeCol)).dCustomWidth = ScaleX(mCols(mColPtr(mResizeCol)).lWidth, vbPixels, vbTwips)
                    
                    DrawGrid mRedraw
                    
                    nMove = mCols(mColPtr(mResizeCol)).MoveControl
                    RaiseEvent ColumnSizeChanged(mResizeCol, nMove)
                    
                    If mEditPending Then MoveEditControl nMove
                End If
            ElseIf (mMouseDownRow = -1) Then
                If mColumnDrag Then
                    DrawHeaderRow
                    
                    If (mMouseDownCol > -1) And (mDragCol = -1) Then mDragCol = mMouseDownCol
                    If (mDragCol <> -1) Then mCols(mColPtr(mDragCol)).lX = mCols(mColPtr(mDragCol)).lX - (mMouseDownX - x)
                End If
            Else
                If mMouseDown And y < 0 Then
                    'Mouse has been dragged off off the control
                    ScrollList 1
                ElseIf mMouseDown And y > UserControl.ScaleHeight Then
                    'Mouse has been dragged off off the control
                    ScrollList 2
                ElseIf mMouseDown And (Shift = 0) And (mMouseRow > -1) Then
                    If mScrollAction = 0 Then
                        bSelectionChanged = SetSelection(False)
                        
                        If mMultiSelect Then
                            SetSelection True, mMouseDownRow, mMouseRow
                        Else
                            SetSelection True, mMouseRow, mMouseRow
                        End If
                        
                        If SetRowCol(mMouseRow, mMouseCol) Then
                            RaiseEvent SelectionChanged
                            DrawGrid mRedraw
                        End If
                    Else
                        mScrollAction = 0
                    End If
                End If
            End If
        ElseIf (Button = 0) Then
            nPointer = vbDefault
            'Only check for resize cursor if no buttons depressed
            If (mMouseRow = -1) Then
                lResizeX = 0
                mResizeCol = -1
                
                If (mAllowResizing = Resize) Then
                    For lCount = SBValue(efsHorizontal) To UBound(mCols)
                        lWidth = lWidth + mCols(mColPtr(lCount)).lWidth
                        
                        If (x < lWidth + 4) And (x > lWidth - 4) Then
                            nPointer = vbSizeWE
                            mResizeCol = lCount
                            Exit For
                        End If
                        
                        lResizeX = lResizeX + mCols(mColPtr(lCount)).lWidth
                    Next lCount
                End If
            End If
            
            With UserControl
                If .MousePointer <> nPointer Then .MousePointer = nPointer
            End With
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As RECT
    Dim lCurrentMouseCol As Long, lCurrentMouseRow As Long, lTemp As Long
    
    If (Button = vbLeftButton) Then
        Call ReleaseCapture
        
        lCurrentMouseCol = GetColFromX(x)
        lCurrentMouseRow = GetRowFromY(y)
        
        If (mDragCol >= 0) Then
            'We moved a Column
            If lCurrentMouseCol > -1 Then
                lTemp = mColPtr(mDragCol)
                mColPtr(mDragCol) = mColPtr(lCurrentMouseCol)
                mColPtr(lCurrentMouseCol) = lTemp
            End If
            DrawGrid True
        ElseIf (mResizeCol >= 0) Then
            'We resized a Column so reset Scrollbars
            SetScrollBars
            DrawGrid mRedraw
            UserControl.MousePointer = vbDefault
        ElseIf (lCurrentMouseRow = -1) Then
            'Sort requested from Column Header click
            If (lCurrentMouseCol = mMouseDownCol) And (mMouseDownCol <> -1) Then
                If mColumnSort Then
                    If (Shift And vbCtrlMask) And (mSortColumn <> -1) Then
                        If mSortSubColumn <> mColPtr(mMouseDownCol) Then mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTAscending
                        mSortSubColumn = mColPtr(mMouseDownCol)
                        Sort , mCols(mColPtr(mSortColumn)).nSortOrder
                    Else
                        If mSortColumn <> mColPtr(mMouseDownCol) Then
                            mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTAscending
                            mSortSubColumn = -1
                        End If
                        mSortColumn = mColPtr(mMouseDownCol)
                        
                        If mSortSubColumn <> -1 Then
                            Sort , , , mCols(mColPtr(mSortSubColumn)).nSortOrder
                        Else
                            Sort
                        End If
                    End If
                Else
                    DrawHeaderRow
                    RaiseEvent ColumnClick(mMouseDownCol)
                End If
            End If
        ElseIf mMouseDownRow > -1 Then
            If IsValidRowCol(mMouseRow, mMouseCol) Then
                If SetRowCol(mMouseRow, mMouseCol) Then DrawGrid mRedraw
                
                If mCF(mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).nFormat).nImage <> 0 Then
                    SetItemRect mRowPtr(mMouseRow), mMouseCol, RowTop(mMouseRow), r, lgRTImage
                    If (x >= r.Left) And (y >= r.Top) And (x <= r.Left + mR.ImageWidth) And (y <= r.Top + mR.ImageHeight) Then
                        RaiseEvent CellImageClick(mMouseRow, mMouseCol)
                    End If
                End If
            End If
        Else
            DrawHeaderRow
        End If
    End If
    
    mMouseDown = False
    mMouseDownCol = -1
    mDragCol = -1
    mResizeCol = -1
    mScrollAction = 0
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        'Appearance Properties
        mSelectBackColor = .ReadProperty("SelectBackColor", &HFCE1CB)
        mForeColor = .ReadProperty("ForeColor", vbWindowText)
        mHeadForeColor = .ReadProperty("HeadForeColor", vbWindowText)
        mSelectForeColor = .ReadProperty("SelectForeColor", vbHighlightText)
        mForeColorTotals = .ReadProperty("ForeColorTotals", vbRed)
        
        mGridColor = .ReadProperty("GridColor", &HC0C0C0)
        
        mAlphaBlendSelection = .ReadProperty("AlphaBlendSelection", False)
        mBorderStyle = .ReadProperty("BorderStyle", 边框)
        mDisplayEllipsis = .ReadProperty("DisplayEllipsis", True)
        mFocusColor = .ReadProperty("FocusColor", vbBlue)
        mSelectMode = .ReadProperty("SelectMode", [行])
        mFocusStyle = .ReadProperty("FocusStyle", Heavy)
        mGridLines = .ReadProperty("GridLines", True)
        mGridLineWidth = .ReadProperty("GridLineWidth", 1)
        
        'Behaviour Properties
        mAllowResizing = .ReadProperty("AllowResizing", Resize)
        mCheckboxes = .ReadProperty("Checkboxes", False)
        mColumnDrag = .ReadProperty("ColumnDrag", False)
        mColumnHeaders = .ReadProperty("ColumnHeaders", True)
        mColumnSort = .ReadProperty("ColumnSort", False)
        mEditable = .ReadProperty("Editable", False)
        mEditType = .ReadProperty("EditType", EnterKey)
        mFullRowSelect = .ReadProperty("FullRowSelect", True)
        mHotHeaderTracking = .ReadProperty("HotHeaderTracking", True)
        mMultiSelect = .ReadProperty("MultiSelect", False)
        mRedraw = .ReadProperty("Redraw", True)
        mScrollTrack = .ReadProperty("ScrollTrack", True)
        mTrackEdits = .ReadProperty("TrackEdits", False)
        
        'Miscellaneous Properties
        mCacheIncrement = .ReadProperty("CacheIncrement", 10)
        mEnabled = .ReadProperty("Enabled", True)
        mLocked = .ReadProperty("Locked", False)
        mRowHeight = .ReadProperty("RowHeight", 300)
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    UserControl.BorderStyle = mBorderStyle
    
    CreateRenderData
    'SetColors
    
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then bTrack = False
        End If
        
        With UserControl
            Call Subclass_Start(.hWnd)
            Call Subclass_AddMsg(.hWnd, WM_KILLFOCUS, MSG_AFTER)
            'Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hWnd, WM_VSCROLL, MSG_AFTER)
            
            If mWindowsXP Then Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED)
        End With
        
        SBCreate UserControl.hWnd
        SBStyle = Style_Regular
        
        SBLargeChange(efsHorizontal) = 5
        SBSmallChange(efsHorizontal) = 1
        
        SBLargeChange(efsVertical) = 5
        SBSmallChange(efsVertical) = 1
    End If
End Sub

Private Sub UserControl_Resize()
    If m_hWnd <> 0 Then Refresh
End Sub

Private Sub UserControl_Terminate()
    On Local Error GoTo UCTError
    If Not mClipRgn = 0 Then DeleteObject mClipRgn
    pSBClearUp
    Call Subclass_Stop(UserControl.hWnd)
UCTError: Exit Sub
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", mFont, Ambient.Font)
        'Appearance Properties
        Call .WriteProperty("SelectBackColor", mSelectBackColor, &HFCE1CB)
        Call .WriteProperty("ForeColor", mForeColor, vbWindowText)
        Call .WriteProperty("HeadForeColor", mHeadForeColor, vbWindowText)
        Call .WriteProperty("SelectForeColor", mSelectForeColor, vbHighlightText)
        Call .WriteProperty("ForeColorTotals", mForeColorTotals, vbRed)
        
        Call .WriteProperty("GridColor", mGridColor, &HC0C0C0)
        
        Call .WriteProperty("AlphaBlendSelection", mAlphaBlendSelection, False)
        Call .WriteProperty("BorderStyle", mBorderStyle, 边框)
        Call .WriteProperty("DisplayEllipsis", mDisplayEllipsis, True)
        Call .WriteProperty("SelectMode", mSelectMode, [行])
        Call .WriteProperty("FocusColor", mFocusColor, vbBlue)
        Call .WriteProperty("FocusStyle", mFocusStyle, Heavy)
        Call .WriteProperty("GridLines", mGridLines, True)
        Call .WriteProperty("GridLineWidth", mGridLineWidth, 1)
        
        'Behaviour Properties
        Call .WriteProperty("AllowResizing", mAllowResizing, Resize)
        Call .WriteProperty("Checkboxes", mCheckboxes, False)
        Call .WriteProperty("ColumnDrag", mColumnDrag, False)
        Call .WriteProperty("ColumnHeaders", mColumnHeaders, True)
        Call .WriteProperty("ColumnSort", mColumnSort, False)
        Call .WriteProperty("Editable", mEditable, False)
        Call .WriteProperty("EditType", mEditType, EnterKey)
        Call .WriteProperty("FullRowSelect", mFullRowSelect, True)
        Call .WriteProperty("HotHeaderTracking", mHotHeaderTracking, True)
        Call .WriteProperty("MultiSelect", mMultiSelect, False)
        Call .WriteProperty("Redraw", mRedraw, True)
        Call .WriteProperty("ScrollTrack", mScrollTrack, True)
        Call .WriteProperty("TrackEdits", mTrackEdits, False)
        
        'Miscellaneous Properties
        Call .WriteProperty("CacheIncrement", mCacheIncrement, 10)
        Call .WriteProperty("Enabled", mEnabled, True)
        Call .WriteProperty("Locked", mLocked, False)
        Call .WriteProperty("RowHeight", mRowHeight, 300)
    End With
End Sub
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-End

'@@@@@@@@@@@@@@@@@@@@@@[子类化部分]@@@@@@@@@@@@@@@@@@@@-Start
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        If (When And eMsgWhen.MSG_AFTER) Then Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End With
End Sub

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    Dim i As Long, j As Long
    Dim nSubIdx As Long
    Dim sSubCode As String
    
    Const PAGE_EXECUTE_READWRITE As Long = &H40&
    Const PATCH_01 As Long = 18
    Const PATCH_02 As Long = 68
    Const PATCH_03 As Long = 78
    Const PATCH_06 As Long = 116
    Const PATCH_07 As Long = 121
    Const PATCH_0A As Long = 186
    Const FUNC_CWPA As String = "CallWindowProcA"
    Const FUNC_CWPW As String = "CallWindowProcW"
    Const FUNC_SWLA As String = "SetWindowLongA"
    Const FUNC_SWLW As String = "SetWindowLongW"
    Const MOD_USER As String = "user32"
    'If it's the first time through here..
    
    If (sc_aBuf(1) = 0) Then
        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
        "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
        "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
        "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
        Hex$(&HA4 + (0 * 12)) & "070000C3"
        
        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < 200
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                     'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop
        
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))
        If IsWindowUnicode(lng_hWnd) Then
            sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWPW)
            sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWLW)
        Else
            sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWPA)
            sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWLA)
        End If
        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        Subclass_Start = nSubIdx
    End If
    
    With sc_aSubData(nSubIdx)
        .nAddrSub = GlobalAlloc(0, 200)
        Call VirtualProtect(ByVal .nAddrSub, 200, PAGE_EXECUTE_READWRITE, i)
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), 200)
        
        .hWnd = lng_hWnd
        'Store the hWnd
        If IsWindowUnicode(lng_hWnd) Then
            .nAddrOrig = SetWindowLongW(.hWnd, GWL_WNDPROC, .nAddrSub)
        Else
            .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        End If
        'Set our WndProc in place
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)
    End With
End Function

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    With sc_aSubData(zIdx(lng_hWnd))
        If IsWindowUnicode(.hWnd) Then
            Call SetWindowLongW(.hWnd, GWL_WNDPROC, .nAddrOrig)
        Else
            Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        End If
        'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With
End Sub
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-End

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long, nOff1 As Long, nOff2 As Long
    
    If (uMsg = -1) Then
        nMsgCnt = -1
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop
        
        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If
    
    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
    Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If
    
    If (uMsg <> -1) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then
                If (Not bAdd) Then Exit Function
            ElseIf (.hWnd = 0) Then
                If (bAdd) Then Exit Function
            End If
        End With
        zIdx = zIdx - 1
    Loop
    
    If (Not bAdd) Then Debug.Assert False
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub
'download by http://www.codefans.net
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub
