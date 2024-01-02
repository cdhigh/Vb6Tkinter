Attribute VB_Name = "MultiLanguage"
'多语言支持模块
'语言文件： Vb6Tkinter.lng
' 文件格式：
'    [语言名称]
'    控件名=字符串
'    其他字符名=字符串    '这个用于内部字符串，比如帮助信息等，在字符串中使用\n表示回车
'
'ChangeLanguage(语言名)   : 切换控件显示语言，这个函数也会一次性缓存对应语种的所有字符串到内存
'L(名字,默认字符串)       : 获取指定字符串
'L_F(名字,默认字符串,其他参数) : 类似Python的{0}{1}格式化字符串
'GetAllLanguageName()     : Vb6Tkinter.lng中所有语种的名称，字符串数组

Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Const LanguageFile = "Vb6Tkinter.lng"
Private m_Lng As New Dictionary                                                 '对应语种的{名称,字符串}字典
Public Const DEF_LNG = "English(&E)"

'根据当前语种设置获取一个字符串
Public Function L(sKey As String, ByVal sDefault As String) As String
    sDefault = Replace(sDefault, "\n", vbCrLf)
    L = GetString("", sKey, sDefault)
End Function

'根据当前语种设置获取一个字符串
'支持类似python的{0}{1}格式化字符串，从{0}开始编号，不支持{}形式(没有数字索引)
Public Function L_F(sKey As String, ByVal sDefault As String, ParamArray v() As Variant) As String
    
    Dim s As String, I As Long
    
    s = L(sKey, sDefault)
    
    For I = 0 To UBound(v)
        s = Replace(s, "{" & I & "}", CStr(v(I)))
    Next
    
    L_F = s
    
End Function

Public Function GetAllLanguageName() As String()
    On Error Resume Next
    Dim s As String, ns As Long
    s = vbNullString
    If LngFileExist() Then
        s = Space(1000)
        ns = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, s, 1000, LngFile())
        GetAllLanguageName = Split(Trim(Replace(Left(s, ns), Chr(0) & Chr(0), "")), Chr(0))
    Else
        GetAllLanguageName = Split(DEF_LNG)                                     '没有语言文件则默认为英文
    End If
    s = ""
End Function

Public Function ChangeLanguage(Language As String) As Boolean
    Dim I As Long, Ctrl As Control, s As String, ns As Long, sa() As String
    
    '先缓存对应语种的语言字符串
    s = Space(10000)
    ns = GetPrivateProfileString(Language, vbNullString, vbNullString, s, 10000, LngFile())
    sa = Split(Trim(Replace(Left(s, ns), Chr(0) & Chr(0), "")), Chr(0))
    m_Lng.RemoveAll
    For I = 0 To UBound(sa)
        s = Space(256)
        ns = GetPrivateProfileString(Language, sa(I), "", s, 256, LngFile())
        s = Trim(Replace(Replace(Left(s, ns), Chr(0), ""), "\n", vbCrLf))
        If Len(s) Then m_Lng.Add sa(I), s
    Next
    
    '切换所有控件的语言
    For I = 0 To Forms.Count - 1
        For Each Ctrl In Forms(I).Controls
            ChangeControlLanguage Ctrl, Language
        Next
    Next I
    
    ChangeLanguage = ns > 0
    
End Function

Public Sub ChangeControlLanguage(ctl As Control, Language As String)
    Select Case TypeName(ctl)
    Case "Label", "CommandButton", "CheckBox", "OptionButton", "Frame", "Menu", "xpcmdbutton"
        ctl.Caption = GetString(Language, ctl.Name, ctl.Caption)
    End Select
End Sub

Private Function LngFile() As String
    LngFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & LanguageFile
End Function

Public Function LngFileExist() As Boolean
    LngFileExist = IIf(Dir(LngFile()) = LanguageFile, True, False)
End Function

Private Function GetString(Language As String, Key As String, sDefault As String) As String
    If m_Lng.Exists(Key) Then
        GetString = m_Lng.Item(Key)
    Else
        GetString = sDefault
    End If
End Function

