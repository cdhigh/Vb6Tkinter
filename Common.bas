Attribute VB_Name = "Common"
Option Explicit

Public VbeInst As VBE

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long

Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPSPERINCH = 1440
Private Type Size
    cx As Long
    cy As Long
End Type

'注册表API声明
'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Public Const REG_SZ = 1
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WOW64_64KEY = &H100

'这些用于获取系统默认字体
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const DEFAULT_GUI_FONT = 17
Private Const LF_FACESIZE = 32
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Const WTOP = "top" '设置tkinter中顶层窗体名字

Public g_DefaultFontName As String '暂存系统默认字体名，避免每次查询
Public g_Comps() As Object '当前窗体的控件列表，第一项为窗体对象实例

Public g_bUnicodePrefixU As Boolean '是否在UNICODE字符串前加前缀u
Public g_PythonExe As String '用于GUI预览，保存python.exe全路径
Public g_AppVerString As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const OFFICIAL_SITE As String = "https://github.com/cdhigh/Vb6Tkinter"
Public Const OFFICIAL_RELEASES As String = "https://github.com/cdhigh/Vb6Tkinter/releases"
Public Const OFFICIAL_UPDATE_INFO As String = "https://api.github.com/repos/cdhigh/Vb6Tkinter/releases"

'PYTHON中UNICODE字符串前缀的处理函数，如果字符串中存在双字节字符，则根据选项增加适当的前缀
'否则，只是简单的增加单引号，即使空串也增加一对单引号
Public Function U(s As String) As String
    
    Dim nLen As Long
    s = Replace(s, vbCrLf, "\n")
    nLen = Len(s)
    
    If lstrlen(s) > nLen Then  '存在双字节字符
        If g_bUnicodePrefixU Then
            U = IIf(isQuoted(s), "u" & s, "u'" & s & "'")
        Else
            U = IIf(isQuoted(s), s, "'" & s & "'")
        End If
    ElseIf nLen Then
        U = IIf(isQuoted(s), s, "'" & s & "'")
    Else
        U = "''"
    End If
    
End Function

'判断字符串是否已经有单引号或双引号括起来了
Public Function isQuoted(s As String) As Boolean
    isQuoted = (Left$(s, 1) = "'" Or Left$(s, 1) = Chr$(34)) And (Right$(s, 1) = "'" Or Right$(s, 1) = Chr$(34))
End Function

'如果字符串有引号，则自动去掉
Public Function UnQuote(s As String) As String
    If isQuoted(s) Then
        UnQuote = Mid(s, 2, Len(s) - 2)
    Else
        UnQuote = s
    End If
End Function

'自动将字符串使用单引号或双引号括起来
Public Function Quote(s As String) As String
    If isQuoted(s) Then
        Quote = s
    ElseIf InStr(1, s, "'") >= 1 Then '字符串里面有单引号，则使用双引号括起来
        Quote = Chr$(34) & s & Chr$(34)
    Else
        Quote = "'" & s & "'"
    End If
End Function

'直接去掉字符串的第一个字符和最后一个字符（假定为引号）
Public Function UnQuoteFast(s As String) As String
    UnQuoteFast = Mid(s, 2, Len(s) - 2)
End Function

'直接将字符串使用单引号括起来
Public Function QuoteFast(s As String) As String
    QuoteFast = "'" & s & "'"
End Function

'提取文件名，包括扩展名，不包括路径名
Public Function FileFullName(ByVal sF As String) As String
    Dim ns As Long
    
    ns = InStrRev(sF, "\")
    If ns <= 0 Then
        FileFullName = sF
    Else
        FileFullName = Right$(sF, Len(sF) - ns)
    End If
End Function

'提取文件扩展名
Public Function FileExt(sF As String) As String
    Dim sFName As String, ns As Long
    sFName = FileFullName(sF)
    ns = InStrRev(sFName, ".")
    If ns > 0 Then
        FileExt = Right$(sFName, Len(sFName) - ns)
    End If
End Function

'提取路径名，包括最后的"\"
Public Function PathName(sF As String) As String
    Dim ns As Long
    
    ns = InStrRev(sF, "\")
    If ns <= 0 Then
        PathName = ""
    Else
        PathName = Left$(sF, ns)
    End If
    
End Function

Function getDPI(bX As Boolean) As Integer                                       '获取屏幕分辨率
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    If bX = True Then
        getDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        getDPI = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    RetVal = ReleaseDC(0, hdc)
End Function
Function Twip2PixelX(x As Long) As Long                                         '水平方向Twip转Pixel
    Twip2PixelX = x / TWIPSPERINCH * getDPI(True)
End Function
Function Twip2PixelY(x As Long) As Long                                         '垂直方向Twip转Pixel
    Twip2PixelY = x / TWIPSPERINCH * getDPI(False)
End Function
Function Point2PixelX(x As Long) As Long                                        '水平方向Point转Pixel
    Point2PixelX = Twip2PixelX(x * 20)
End Function
Function Point2PixelY(x As Long) As Long                                        '垂直方向Point转Pixel
    Point2PixelY = Twip2PixelY(x * 20)
End Function
Function getScreenX() As Long                                                   '获取屏幕宽
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    getScreenX = GetDeviceCaps(hdc, HORZRES)
    RetVal = ReleaseDC(0, hdc)
End Function
Function getScreenY() As Long                                                   '获取屏幕高
    Dim hdc As Long, RetVal As Long
    hdc = GetDC(0)
    getScreenY = GetDeviceCaps(hdc, VERTRES)
    RetVal = ReleaseDC(0, hdc)
End Function

Public Function CharWidth() As Long                '获取默认字体字符宽度(像素)
    Dim hdc As Long, RetVal As Long
    Dim typSize     As Size
    Dim lngX     As Long
    Dim lngY     As Long
    
    hdc = GetDC(0)
    RetVal = GetTextExtentPoint32(hdc, "ABli", 4, typSize)
    CharWidth = typSize.cx / 4
    RetVal = ReleaseDC(0, hdc)
End Function

'VB颜色转Python的RGB颜色
'要使用调色板的颜色才能转换为RGB颜色，使用系统颜色无法转换
'Public Function ColorToRGBStr(ByVal dwColor As Long) As String
'    Dim clrHex As String
'    If dwColor > 0 Then
'        clrHex = Replace(Format(Hex$(dwColor), "@@@@@@"), " ", "0")
'        ColorToRGBStr = "'#" & Mid$(clrHex, 5, 2) & Mid$(clrHex, 3, 2) & Mid$(clrHex, 1, 2) & "'"
'    End If
'End Function

'VB颜色转Python的RGB颜色
'不管使用调色板还是系统颜色，都可以转换为RGB颜色
Public Function TranslateColor(ByVal dwColor As OLE_COLOR) As String
    Dim nColor As Long, hPalette As Long, clrHex As String
    If OleTranslateColor(dwColor, hPalette, nColor) = 0 Then
        clrHex = Replace(Format(Hex$(nColor), "@@@@@@"), " ", "0")
        TranslateColor = "'#" & Mid$(clrHex, 5, 2) & Mid$(clrHex, 3, 2) & Mid$(clrHex, 1, 2) & "'"
    End If
End Function

' 获取系统中所有安装的Python路径
Public Function GetAllInstalledPython() As String()
    Dim nRe As Long, nHk As Long, nHk2 As Long, i As Long, nLen As Long
    Dim sVer As String, sAllPath As String, sBuff As String, sPythonExe As String
    Dim saVer() As String, nVerNum As Long
    
    nRe = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Python\PythonCore", 0, KEY_READ Or KEY_WOW64_64KEY, nHk)
    If nRe <> 0 Then
        GetAllInstalledPython = Split("")
        Exit Function
    End If
    
    i = 0
    nVerNum = 0
    nLen = 255
    sBuff = String$(255, 0)
    Do While (RegEnumKeyEx(nHk, i, sBuff, nLen, 0, vbNullString, ByVal 0&, ByVal 0&) = 0)
        If nLen > 1 Then
            sBuff = Left$(sBuff, InStr(1, sBuff, Chr(0)) - 1)
            
            ReDim Preserve saVer(nVerNum) As String
            saVer(nVerNum) = sBuff
            nVerNum = nVerNum + 1
        End If
        i = i + 1
        nLen = 255
        sBuff = String$(255, 0)
    Loop
    RegCloseKey nHk
    
    '查询具体安装路径
    For i = 1 To nVerNum
        nRe = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Python\PythonCore\" & saVer(i - 1) & "\InstallPath", 0, KEY_READ Or KEY_WOW64_64KEY, nHk2)
        If nRe = 0 Then
            nLen = 255
            sBuff = String$(255, 0)
            nRe = RegQueryValueEx(nHk2, "", 0&, REG_SZ, sBuff, nLen)  '查询子键默认字符串值
            If nRe = 0 And nLen > 1 Then
                sBuff = Left$(sBuff, InStr(1, sBuff, Chr(0)) - 1)
                
                sPythonExe = sBuff & IIf(Right$(sBuff, 1) = "\", "", "\") & "python.exe"
                sPythonExe = sPythonExe & "," & sBuff & IIf(Right$(sBuff, 1) = "\", "", "\") & "pythonw.exe"
                sAllPath = sAllPath & IIf(Len(sAllPath), ",", "") & sPythonExe
            End If
            RegCloseKey nHk2
        End If
    Next
    
    GetAllInstalledPython = Split(sAllPath, ",")
End Function

'获取系统默认字体名
Public Function GetDefaultFontName() As String
    Dim hFont As Long, lfont As LOGFONT
    
    If Len(g_DefaultFontName) Then
        GetDefaultFontName = g_DefaultFontName
    Else
        hFont = GetStockObject(DEFAULT_GUI_FONT)
        If hFont <> 0 Then
            GetObject hFont, Len(lfont), lfont
            DeleteObject hFont
            GetDefaultFontName = StrConv(lfont.lfFaceName, vbUnicode)
            If InStr(1, GetDefaultFontName, Chr(0)) > 0 Then
                GetDefaultFontName = Left$(GetDefaultFontName, InStr(1, GetDefaultFontName, Chr(0)) - 1)
            End If
            g_DefaultFontName = GetDefaultFontName  '暂存，下一次就不用API查询了
        End If
    End If
End Function

'获取当前窗体的所有控件列表，返回字符为使用|分割的名字和类型名
Public Function GetAllComps() As String()
    Dim nCnt As Long, i As Long, sa() As String
    On Error Resume Next
    nCnt = UBound(g_Comps)
    On Error GoTo 0
    If nCnt <= 0 Then
        GetAllComps = Split("")
        Exit Function
    End If
    
    ReDim sa(nCnt) As String
    For i = 0 To nCnt
        sa(i) = g_Comps(i).Name & "|" & TypeName(g_Comps(i))
    Next
    GetAllComps = sa
End Function

'根据依赖关系排序控件，简单的冒泡排序，这需要在生成代码之前调用
'基本原理是顶层控件先生成代码，父控件先生成代码，最后是子控件
Public Sub SortWidgets(ByRef aCompsSorted() As Object, ByVal cnt As Long)
    Dim idx1 As Long, idx2 As Long
    Dim tmp4exchange As Object
    
    For idx1 = 0 To cnt - 2
        For idx2 = idx1 + 1 To cnt - 1
            If aCompsSorted(idx1).Compare(aCompsSorted(idx2)) > 0 Then '重者沉底
                Set tmp4exchange = aCompsSorted(idx1)
                Set aCompsSorted(idx1) = aCompsSorted(idx2)
                Set aCompsSorted(idx2) = tmp4exchange
            End If
        Next
    Next
    
End Sub

'将版本号字符串前面和后面非数字部分都删除掉，比如："v1.6.8 test" 将返回 "1.6.8"
Private Function purifyVerStr(txt As String) As String
    Dim maxCnt As Integer, idx As Integer, startIdx As Integer, endIdx As Integer
    Dim ch As String
    txt = Trim(txt)
    maxCnt = Len(txt)
    startIdx = 1
    endIdx = maxCnt
    '开头部分
    For idx = 1 To maxCnt
        ch = Mid(txt, idx, 1)
        If (ch >= "0") And (ch <= "9") Then
            startIdx = idx
            Exit For
        End If
    Next
    '结尾部分
    For idx = maxCnt To 1 Step -1
        ch = Mid(txt, idx, 1)
        If (ch >= "0") And (ch <= "9") Then
            endIdx = idx
            Exit For
        End If
    Next
    
    If startIdx <= endIdx Then
        purifyVerStr = Mid(txt, startIdx, endIdx - startIdx + 1)
    Else
        purifyVerStr = ""
    End If
End Function

'比较两个版本号，确定新版本号是否比老版本号更新，
'版本号格式为：1.1.0
Public Function isVersionNewerThan(newVer As String, currVer As String) As Boolean
    Dim newArr As Variant, currArr As Variant, idx As Integer, maxCnt As Integer
    Dim vn As Integer, vc As Integer
    newVer = purifyVerStr(newVer)
    currVer = purifyVerStr(currVer)
    If Len(newVer) = 0 Or Len(currVer) = 0 Then
        isVersionNewerThan = False
        Exit Function
    End If
    
    newArr = Split(newVer, ".")
    currArr = Split(currVer, ".")
    maxCnt = UBound(newArr)
    If UBound(currArr) < maxCnt Then '两个数组最小的一个
        maxCnt = UBound(currArr)
    End If
    
    For idx = 0 To maxCnt
        vn = Val(newArr(idx))
        vc = Val(currArr(idx))
        If vn > vc Then
            isVersionNewerThan = True
            Exit Function
        ElseIf vn < vc Then
            isVersionNewerThan = False
            Exit Function
        End If
    Next
    
    '如果前面都一样，则长的一个为大
    If UBound(newArr) > UBound(currArr) Then
        isVersionNewerThan = True
    Else
        isVersionNewerThan = False
    End If
End Function
