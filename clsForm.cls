VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'窗体类的抽象，保存VB窗体属性
Private m_Base As clsBaseControl
Private m_Name As String
Private m_Title As String
Private m_Geometry As String
Private m_WindowState As String
Private m_left As Long
Private m_Top As Long
Private m_Topmost As Boolean
Private m_Alpha As String
Private m_Resizable As Boolean
Private m_IconFile As String
Private m_IconEmbedding As Boolean
Private m_Protocol As String
Private m_Bindcommand As String
Private m_FrmPath As String          '保存窗体所在目录
Private m_ScaleMode As Long
Private m_IconFilesInPath As String '在窗体所在目录下所有的图标文件，使用逗号分隔
Private m_StartupPosition As String

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean)
    
    If oop Then
        sOut.Append "        self.master.title(" & U(m_Title) & ")"
        If m_StartupPosition = "Manual" Then  '手工指定窗口初始位置
            sOut.Append "        self.master.geometry('" & m_Geometry & CoordinateOnStart() & "')"
        ElseIf m_StartupPosition <> "Center" Then
            sOut.Append "        self.master.geometry('" & m_Geometry & "')"
        End If
        If Not m_Resizable Then sOut.Append "        self.master.resizable(0,0)"
        If Len(m_WindowState) > 0 And m_WindowState <> "'normal'" Then
            sOut.Append "        self.master.wm_state(" & m_WindowState & ")"
        End If
        If m_Topmost Then
            '增加wait函数是因为http://code.activestate.com/lists/python-list/385778/，有网友反馈偶尔会出现窗体不显示的情况，
            'wait函数可以规避这个问题
            sOut.Append "        self.master.wait_visibility(self.master)"
            sOut.Append "        self.master.wm_attributes('-topmost', 1)"
        End If
        If Len(m_Alpha) Then
            sOut.Append "        self.master.wm_attributes('-alpha', " & m_Alpha & ")"
        End If
    Else
        sOut.Append "    " & m_Name & " = Tk()"
        sOut.Append "    " & m_Name & ".title(" & U(m_Title) & ")"
        If m_StartupPosition = "Manual" Then  '手工指定窗口初始位置
            sOut.Append "    " & m_Name & ".geometry('" & m_Geometry & CoordinateOnStart() & "')"
        ElseIf m_StartupPosition = "Center" Then
            sOut.Append "    " & m_Name & ".geometry('" & m_Geometry & "')"
        End If
        If Not m_Resizable Then sOut.Append "    " & m_Name & ".resizable(0,0)"
        If Len(m_WindowState) > 0 And m_WindowState <> "'normal'" Then
            sOut.Append "    " & m_Name & ".wm_state(" & m_WindowState & ")"
        End If
        If m_Topmost Then
            sOut.Append "    " & m_Name & ".wait_visibility(" & m_Name & ")"
            sOut.Append "    " & m_Name & ".wm_attributes('-topmost', 1)"
        End If
        If Len(m_Alpha) Then
            sOut.Append "    " & m_Name & ".wm_attributes('-alpha', " & m_Alpha & ")"
        End If
    End If
    
    ProcessIcon sOut, oop  '处理窗体图标
    ProcessBindCommand sOut, sCmdFunc, oop ' 处理bindcommand
    ProcessProtocol sOut, sCmdFunc, oop '窗口消息拦截，如果有的话
    
    If m_StartupPosition = "Center" Then
        CenterWindow sOut, oop  '如果需要窗体在屏幕中间
    End If
    
    If Not oop Then
        sOut.Append "    gComps['" & m_Name & "'] = " & m_Name & vbCrLf
    End If
End Sub

Private Sub ProcessIcon(sOut As cStrBuilder, oop As Boolean)
    
    Dim s As String, sBase64 As String
    
    If Len(m_IconFile) = 0 Then Exit Sub
    
    s = UCase(FileExt(m_IconFile))
    
    '处理窗体图标
    If oop Then
        If s = "GIF" And m_IconEmbedding Then
            sBase64 = CreateIconBase64(Space(12)) '使用BASE64将GIF图片嵌入源码
            If Len(sBase64) Then
                sOut.Append "        self.icondata = " & String(3, Chr(34)) & vbCrLf & sBase64 & String(3, Chr(34))
                sOut.Append "        self.iconimg = PhotoImage(data=self.icondata)"
                sOut.Append "        self.master.tk.call('wm', 'iconphoto', self.master._w, self.iconimg)"
            End If
        ElseIf Len(s) = 0 Or s = "ICO" Then '如果没有后缀名，则假定为ico格式
            sOut.Append "        self.master.iconbitmap(default=r'" & FileFullName(m_IconFile) & "')"
        Else
            sOut.Append "        self.iconimg = PhotoImage(file=r'" & FileFullName(m_IconFile) & "')"
            sOut.Append "        self.master.tk.call('wm', 'iconphoto', self.master._w, self.iconimg)"
        End If
    Else
        If s = "GIF" And m_IconEmbedding Then
            sBase64 = CreateIconBase64(Space(8))
            If Len(sBase64) Then
                sOut.Append "    icondata = " & String(3, Chr(34)) & vbCrLf & sBase64 & String(3, Chr(34))
                sOut.Append "    iconimg = PhotoImage(data=icondata)"
                sOut.Append "    " & m_Name & ".tk.call('wm', 'iconphoto', " & m_Name & "._w, iconimg)"
            End If
        ElseIf Len(s) = 0 Or s = "ICO" Then '如果图标没有后缀名，则假定为ico格式
            sOut.Append "    " & m_Name & ".iconbitmap(default=r'" & FileFullName(m_IconFile) & "')"
        Else
            sOut.Append "    iconimg = PhotoImage(file=r'" & FileFullName(m_IconFile) & "')"
            sOut.Append "    " & m_Name & ".tk.call('wm', 'iconphoto', " & m_Name & "._w, iconimg)"
        End If
    End If
End Sub

Private Sub ProcessBindCommand(sOut As cStrBuilder, sCmdFunc As cStrBuilder, oop As Boolean)
    
    Dim i As Long, s() As String, sCmdName As String
    
    If Len(m_Bindcommand) = 0 Then Exit Sub
    
    s = Split(m_Bindcommand, ",")
    
    For i = 0 To UBound(s)
        s(i) = Trim(s(i))
        If Left$(s(i), 1) = "<" And Right$(s(i), 1) = ">" Then
            
            sCmdName = m_Name & "_" & Replace(Replace(Replace(s(i), "<", ""), ">", ""), "-", "_")
            
            If oop Then
                sOut.Append Space(8) & "self.master.bind('" & s(i) & "', self." & sCmdName & ")"
                sCmdFunc.Append m_Base.CreateFuncDefOOP(sCmdName, "event")
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If Right$(s(i), 3) >= "-a>" And Right$(s(i), 3) <= "-z>" Then
                    s(i) = Left(s(i), Len(s(i)) - 2) & UCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '变大写
                    sOut.Append Space(8) & "self.master.bind('" & s(i) & "', self." & sCmdName & ")"
                ElseIf Right$(s(i), 3) >= "-A>" And Right$(s(i), 3) <= "-Z>" Then
                    s(i) = Left(s(i), Len(s(i)) - 2) & LCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '变小写
                    sOut.Append Space(8) & "self.master.bind('" & s(i) & "', self." & sCmdName & ")"
                End If
            Else
                sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                sCmdFunc.Append m_Base.CreateFuncDef(sCmdName, "event")
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If Right(s(i), 3) >= "-a>" And Right(s(i), 3) <= "-z>" Then
                    s(i) = Left(s(i), Len(s(i)) - 2) & UCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '对应字母变大写
                    sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                ElseIf Right(s(i), 3) >= "-A>" And Right(s(i), 3) <= "-Z>" Then
                    s(i) = Left(s(i), Len(s(i)) - 2) & LCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '对应字母变小写
                    sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                End If
            End If
            
        End If
    Next
    
End Sub

'窗口消息拦截，如果有的话
Private Sub ProcessProtocol(sOut As cStrBuilder, sCmdFunc As cStrBuilder, oop As Boolean)
    
    Dim sa() As String, i As Long, s As String
    
    If Len(m_Protocol) = 0 Then Exit Sub
    
    sa = Split(m_Protocol, ",")
    
    If oop Then
        For i = 0 To UBound(sa)
            sOut.Append "        self.master.protocol('" & sa(i) & "', self.EV_" & sa(i) & ")"
            If sa(i) = "WM_DELETE_WINDOW" Then
                s = "        if askyesno(message='Are you sure to quit?'):" & vbCrLf & Space(12) & "self.master.destroy()"
                sCmdFunc.Append m_Base.CreateFuncDefOOP("EV_" & sa(i), "event=None", s)
            Else
                sCmdFunc.Append m_Base.CreateFuncDefOOP("EV_" & sa(i), "event=None")
            End If
        Next
    Else
        For i = 0 To UBound(sa)
            sOut.Append "    " & m_Name & ".protocol('" & sa(i) & "', EV_" & sa(i) & ")"
            If sa(i) = "WM_DELETE_WINDOW" Then
                s = "    if askyesno(message='Are you sure to quit?'):" & vbCrLf & "        gComps['" & m_Name & "'].destroy()"
                sCmdFunc.Append m_Base.CreateFuncDef("EV_" & sa(i), "event=None", s)
            Else
                sCmdFunc.Append m_Base.CreateFuncDef("EV_" & sa(i), "event=None")
            End If
        Next
    End If
    
End Sub

'将主窗体放置在屏幕中间
Private Sub CenterWindow(sOut As cStrBuilder, oop As Boolean)
    Dim W As String, H As String
    
    W = Left$(m_Geometry, InStr(1, m_Geometry, "x") - 1)
    H = Mid$(m_Geometry, InStr(1, m_Geometry, "x") + 1)
    
    If oop Then
        sOut.Append Space(8) & "# To center the window on the screen."
        sOut.Append Space(8) & "ws = self.master.winfo_screenwidth()"
        sOut.Append Space(8) & "hs = self.master.winfo_screenheight()"
        sOut.Append Space(8) & "x = (ws / 2) - (" & W & " / 2)"
        sOut.Append Space(8) & "y = (hs / 2) - (" & H & " / 2)"
        sOut.Append Space(8) & "self.master.geometry('%dx%d+%d+%d' % (" & W & "," & H & ",x,y))"
    Else
        sOut.Append Space(4) & "# To center the window on the screen."
        sOut.Append Space(4) & "ws = top.winfo_screenwidth()"
        sOut.Append Space(4) & "hs = top.winfo_screenheight()"
        sOut.Append Space(4) & "x = (ws / 2) - (" & W & " / 2)"
        sOut.Append Space(4) & "y = (hs / 2) - (" & H & " / 2)"
        sOut.Append Space(4) & "top.geometry('%dx%d+%d+%d' % (" & W & "," & H & ",x,y))"
    End If
End Sub

Private Function CreateIconBase64(sPrexSpace As String) As String

    Dim sIconFile As String, abContent() As Byte
    
    CreateIconBase64 = ""
    If Len(m_IconFile) <= 0 Then Exit Function
    
    sIconFile = IIf(InStr(1, m_IconFile, "\") Or InStr(1, m_IconFile, "/"), "", m_FrmPath) & m_IconFile
    
    On Error GoTo DirErr
    If Dir(sIconFile) = "" Then
        MsgBox L_F("l_msgFileNotExist", "文件{0}不存在，请重新选择文件。", sIconFile), vbInformation
        Exit Function
    End If
    
    If ReadFileBinaryContent(sIconFile, abContent) = 0 Then
        MsgBox L_F("l_msgReadFileError", "读取文件{0}出错。", sIconFile), vbInformation
        Exit Function
    End If
    
    Base64Encode abContent, CreateIconBase64, sPrexSpace
    Exit Function
DirErr:
    MsgBox L_F("l_msgFileNotExist", "文件{0}不存在，请重新选择文件。", sIconFile), vbInformation
End Function


'创建对象后要马上调用这个函数初始化各参数, 窗体大小以像素为单位
Public Sub InitConfig(o As Object, nWidth As Long, nHeight As Long, dMethods As Dictionary)
    
    Dim s As String, cAttr As String, cValue As String, sa() As String, i As Long, Idx As Long
    
    '获取窗体所在目录
    On Error Resume Next
    s = o.FileNames(1)
    If Err.Number = 0 And Len(s) Then
        m_FrmPath = PathName(s)
    Else
        m_FrmPath = ""
    End If
    Err.Clear
    On Error GoTo 0
    
    AutoDetectIconFile
    
    m_Name = o.Properties("Name")
    m_Title = o.Properties("Caption")
    m_Geometry = nWidth & "x" & nHeight
    m_WindowState = IIf(o.Properties("WindowState") = vbMaximized, "'zoomed'", IIf(o.Properties("WindowState") = vbMinimized, "'iconic'", "'normal'"))
    m_Resizable = IIf(o.Properties("MaxButton"), True, False)
    m_left = m_Base.toPixelX(o.Properties("Left"))
    m_Top = m_Base.toPixelY(o.Properties("Top"))
    m_IconEmbedding = True
    m_Protocol = ""
    If InStr(1, m_IconFilesInPath, ",") <= 0 Then
        m_IconFile = m_IconFilesInPath
    Else
        m_IconFile = ""
    End If
    m_Bindcommand = ""
    m_Topmost = False
    m_Alpha = ""
    If o.Properties("StartUpPosition") = 1 Or o.Properties("StartUpPosition") = 2 Then
        m_StartupPosition = "Center"
    ElseIf o.Properties("StartUpPosition") = 0 Then
        m_StartupPosition = "Manual"
    Else
        m_StartupPosition = "Default"
    End If
    
    
    '使用控件的tag属性保存一些额外默认要保存的属性，
    '开始字符为：p@，后接（属性=值）对，每个属性使用@隔开，对于窗体来说，属性值必选
    s = Trim(o.Properties("Tag"))
    If Len(s) Then
        sa = Split(s, "@")
        If UCase(sa(0)) = "P" Then
            For i = 1 To UBound(sa)
                s = sa(i)
                Idx = InStr(2, s, "=")
                If Idx > 0 Then '有值
                    cAttr = Left$(s, Idx - 1)
                    cValue = Mid$(s, Idx + 1)
                    Select Case cAttr
                    Case "topmost"
                        m_Topmost = IIf((cValue = "0" Or UCase(cValue) = "FALSE"), False, True)
                    Case "alpha"
                        m_Alpha = cValue
                    Case "icon"
                        m_IconFile = cValue
                    Case "iconembedding"
                        m_IconEmbedding = IIf((cValue = "0" Or UCase(cValue) = "FALSE"), False, True)
                    Case "bindcommand"
                        m_Bindcommand = cValue
                    Case "protocol"
                        m_Protocol = cValue
                    End Select
                End If
            Next
        End If
    End If
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    If sAttr = "resizable" Or sAttr = "iconembedding" Or sAttr = "topmost" Then
        sa = Split("1,0", ",")
    ElseIf sAttr = "windowstate" Then
        sa = Split("'normal','iconic','zoomed'", ",")
    ElseIf sAttr = "icon" And Len(m_IconFilesInPath) > 0 Then
        sa = Split(m_IconFilesInPath, ",")
        GetAttrValueList = 2
    ElseIf sAttr = "protocol" Then
        sa = Split("WM_DELETE_WINDOW", ",")
        GetAttrValueList = 2
    ElseIf sAttr = "alpha" Then
        sa = Split("0.0,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0", ",")
        GetAttrValueList = 2
    ElseIf sAttr = "position" Then
        sa = Split("Default,Center,Manual", ",")
    Else
        GetAttrValueList = 0
    End If
End Function

Public Function Tips(sAttr As String) As String
    Tips = sAttr & vbCrLf
    Select Case sAttr:
        Case "title"
            Tips = Tips & L("l_TipFrmTitle", "窗体的标题。")
        Case "geometry"
            Tips = Tips & L("l_TipFrmGeometry", "窗体的大小（宽x高），单位为像素。")
        Case "left"
            Tips = Tips & L("l_TipFrmLeft", "窗体的初始位置坐标，单位为像素。")
        Case "top"
            Tips = Tips & L("l_TipFrmTop", "窗体的初始位置坐标，单位为像素。")
        Case "resizable"
            Tips = Tips & L("l_TipFrmResizable", "用户是否可以改变窗体的大小。可选值：1/0。")
        Case "windowstate"
            Tips = Tips & L("l_TipFrmState", "窗体启动时的状态。'iconic'：最小化，'zoomed'：最大化，'normal'：正常状态。")
        Case "topmost"
            Tips = Tips & L("l_TipFrmTopmost", "窗体是否置顶。")
        Case "alpha"
            Tips = Tips & L("l_TipFrmAlpha", "窗体透明度。取值范围为0-1，0为全透明，1为不透明。")
        Case "icon"
            Tips = Tips & L("l_TipFrmIcon", "图标文件名，支持ico,gif,ppm,pgm格式，省略路径名则在当前VB窗体目录下查找，建议icon文件放在py文件同一目录下。")
        Case "iconembedding"
            Tips = Tips & L("l_TipFrmIconEmbedding", "是否将图标文件转换为base64编码嵌入源代码，仅支持GIF格式嵌入。")
        Case "bindcommand"
            Tips = Tips & L("l_TipBindCommand", "使用bind()绑定的事件处理列表，绑定多个则使用逗号分隔(比如：<Control-C>,<F8>,<Alt-A>)。\n双击文本框查看详情。")
        Case "protocol"
            Tips = Tips & L("l_TipFrmProtocol", "处理窗口消息，比如设置为'WM_DELETE_WINDOW'则处理窗口关闭事件，如果要拦截多个消息，请使用逗号分隔，比如'WM_DELETE_WINDOW,WM_SAVE'")
        Case "position"
            Tips = Tips & L("l_TipFrmPosition", "窗体启动时位置。")
        Case Else:
            Tips = Tips & L("l_TipUnknown", "未知属性")
    End Select
End Function

'将窗体目录下的所有图标文件存入m_IconFilesInPath，以备后续选择
Private Function AutoDetectIconFile() As String
    Dim sOut As New cStrBuilder, s As String
    m_IconFilesInPath = ""
    If Len(m_FrmPath) Then
        s = Join(SearchFiles(m_FrmPath, "*.ico", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.gif", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.ppm", True), ",")
        If Len(s) Then sOut.Append s
        s = Join(SearchFiles(m_FrmPath, "*.pgm", True), ",")
        If Len(s) Then sOut.Append s
        m_IconFilesInPath = sOut.toString(",")
    End If
End Function

'计算和生成窗体启动坐标
Private Function CoordinateOnStart() As String
    CoordinateOnStart = IIf(m_left >= 0, "+" & CStr(m_left), CStr(m_left))
    CoordinateOnStart = CoordinateOnStart & IIf(m_Top >= 0, "+" & CStr(m_Top), CStr(m_Top))
End Function

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    Dim re As Collection
    
    Set re = New Collection
    re.Add "geometry|" & m_Geometry & "|1"
    re.Add "title|" & m_Title & "|1"
    re.Add "left|" & m_left & "|1"
    re.Add "top|" & m_Top & "|1"
    re.Add "resizable|" & IIf(m_Resizable, 1, 0) & "|1"
    re.Add "windowstate|" & m_WindowState & "|1"
    re.Add IIf(m_Topmost, "topmost|1|1", "topmost||0")
    re.Add IIf(m_Alpha <> "", "alpha|" & m_Alpha & "|1", "alpha||0")
    re.Add "position|" & m_StartupPosition & "|1"
    re.Add "icon|" & m_IconFile & "|1"
    re.Add "iconembedding|" & IIf(m_IconEmbedding, 1, 0) & "|1"
    re.Add IIf(m_Bindcommand <> "", "bindcommand|" & m_Bindcommand & "|1", "bindcommand||0")
    re.Add IIf(m_Protocol <> "", "protocol|" & m_Protocol & "|1", "protocol||0")
    
    Set Allitems = re
End Function

Public Sub SetConfig(sAttrs As String)
    Dim sa() As String, i As Long
    
    m_Title = ""
    m_Resizable = True
    m_WindowState = ""
    m_Topmost = False
    m_Alpha = ""
    m_IconFile = ""
    m_IconEmbedding = False
    m_Bindcommand = ""
    m_Protocol = ""
    m_StartupPosition = ""
    m_left = 0
    m_Top = 0
    
    sa = Split(sAttrs, "|")
    
    For i = 0 To UBound(sa) - 1 Step 2
        SetSingleConfig sa(i) & "|" & sa(i + 1)
    Next
    
End Sub

Public Sub SetSingleConfig(sAttr As String)
    Dim sa() As String
    
    sa = Split(sAttr, "|")
    'Debug.Assert (UBound(sa) Mod 1 = 0)
    
    If UBound(sa) >= 1 Then
        Select Case sa(0)
            Case "title"
                m_Title = UnQuote(Trim(sa(1)))
            Case "geometry"
                m_Geometry = UnQuote(Trim(sa(1)))
            Case "left"
                m_left = C2Lng(Trim(sa(1)))
            Case "top"
                m_Top = C2Lng(Trim(sa(1)))
            Case "resizable"
                m_Resizable = IIf(Trim(sa(1)) = "1", True, False)
            Case "windowstate"
                m_WindowState = Trim(sa(1))
            Case "topmost"
                m_Topmost = IIf(Trim(sa(1)) = "1", True, False)
            Case "alpha"
                m_Alpha = Trim(sa(1))
            Case "icon"
                m_IconFile = UnQuote(Trim(sa(1)))
            Case "iconembedding"
                m_IconEmbedding = IIf(Trim(sa(1)) = "1", True, False)
            Case "bindcommand"
                m_Bindcommand = Trim(Replace(Replace(Replace(sa(1), Chr(34), ""), "'", ""), " ", ""))
            Case "protocol"
                m_Protocol = Trim(Replace(Replace(Replace(sa(1), Chr(34), ""), "'", ""), " ", ""))
            Case "position"
                m_StartupPosition = Trim(sa(1))
        End Select
    End If
End Sub

Private Function C2Lng(sV As String) As Long
    On Error Resume Next
    C2Lng = CLng(sV)
End Function

Public Property Get Parent() As String
    Parent = ""
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(s As String)
    m_Name = s
End Property

'用于改变其默认对应的widget类型，修改widget类型后注意属性列表的合法性
Public Function SetWidgetType(sType As String, sStyleName As String)
    'm_Base.ctlType = sType
    'm_Base.StyleName = sStyleName
End Function

'确定主处理函数能否调用其toString()来产生代码，默认为True，设置为False说明由其他对象来调用处理
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = True
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    'm_CanbeOutByMainForm = bEnable
End Property

Public Property Get Description() As String
    Description = L("l_DescForm", "窗体对象，设置顶层窗口的属性。")
End Property

Public Property Let ScaleMode(nV As Long)
    m_ScaleMode = nV
End Property

Private Sub Class_Initialize()
    Set m_Base = New clsBaseControl
End Sub

Private Sub Class_Terminate()
    Set m_Base = Nothing
End Sub

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

