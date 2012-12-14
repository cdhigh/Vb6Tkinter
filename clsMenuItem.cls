VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'菜单项类，表示每一个菜单项

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_Visible As Boolean
Private m_IsSeparator As Boolean
Private m_IsCheckBox As Boolean

Private m_Childs() As Object
Private m_numChilds As Long
Private m_IdxCurChild As Long

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean)
    
    If oop Then
        toStringOOP sOut, sCmd, rel
    Else
        toStringStructure sOut, sCmd, rel
    End If
    
End Sub

'输出结构化代码
Public Sub toStringStructure(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, rel As Boolean)
    
    Dim i As Long, extra As String, var As String, sTmp As String, s() As String, sr2 As String, sr3 As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    If m_numChilds > 0 Then
        sOut.Append vbCrLf & Space(4) & m_Base.Name & " = Menu(" & m_Base.Parent & extra & ")"
        For i = 0 To m_numChilds - 1                                            '递归生成代码
            m_Childs(i).toStringStructure sOut, sCmd, rel
        Next
        
        sOut.Append Space(4) & "gComps['" & m_Base.Name & "'] = " & m_Base.Name
        If m_Visible Then                                                       '如果Visible=0，一般说明此菜单设置为右键弹出菜单
            m_Base("variable") = ""
            m_Base("command") = ""
            extra = GetExtraParams(False)
            sOut.Append Space(4) & m_Base.Parent & ".add_cascade(menu=" & m_Base.Name & IIf(Len(extra), ", " & extra, "") & ")"
        End If
    ElseIf m_IsSeparator Then                                                   '菜单分隔符
        sOut.Append Space(4) & m_Base.Parent & ".add_separator()"
    ElseIf m_IsCheckBox Then                                                    ' 菜单增加选择框
        If m_Base("variable") <> "" Then
            sOut.Append Space(4) & m_Base("variable") & " = StringVar()"
            sOut.Append Space(4) & m_Base("variable") & ".set(1)"
        End If
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams(False)
        sOut.Append Space(4) & m_Base.Parent & ".add_checkbutton(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDef(m_Base("command"), "event=None")
        End If
    Else                                                                        '子菜单
        var = m_Base("variable")
        m_Base("variable") = ""
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams(False)
        m_Base("variable") = var
        sOut.Append Space(4) & m_Base.Parent & ".add_command(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDef(m_Base("command"), "event=None")
        End If
    End If
    
    '有需要使用bind语句绑定的其他事件处理
    If m_numChilds = 0 And Not m_IsSeparator And Len(m_Base("bindcommand")) Then
        sTmp = UnQuote(m_Base("bindcommand")) '自动去掉括号，如果有的话
        
        s = Split(sTmp, ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If Left$(s(i), 1) = "<" And Right$(s(i), 1) = ">" Then
                sOut.Append Space(4) & WTOP & ".bind_all('" & s(i) & "', " & m_Base("command") & ")"
                
                sr2 = Mid$(s(i), Len(s(i)) - 1, 1)
                sr3 = Mid$(s(i), Len(s(i)) - 2, 1)
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If sr3 = "-" Then
                    If sr2 >= "a" And sr2 <= "z" Then
                        s(i) = Left$(s(i), Len(s(i)) - 2) & UCase(sr2) & ">"
                        sOut.Append Space(4) & WTOP & ".bind_all('" & s(i) & "', " & m_Base("command") & ")"
                    ElseIf sr2 >= "A" And sr2 <= "Z" Then
                        s(i) = Left$(s(i), Len(s(i)) - 2) & LCase(sr2) & ">"
                        sOut.Append Space(4) & WTOP & ".bind_all('" & s(i) & "', " & m_Base("command") & ")"
                    End If
                End If
            End If
        Next
    End If
    
End Sub

'输出面向对象代码
Public Sub toStringOOP(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, rel As Boolean)
    
    Dim i As Long, extra As String, var As String, sTmp As String, s() As String, sr2 As String, sr3 As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    If m_numChilds > 0 Then
        sOut.Append vbCrLf & Space(8) & "self." & m_Base.Name & " = Menu(self." & m_Base.Parent & extra & ")"
        For i = 0 To m_numChilds - 1                                            '递归生成代码
            m_Childs(i).toStringOOP sOut, sCmd, rel
        Next
        
        If m_Visible Then                                                       '如果Visible=0，一般说明此菜单设置为右键弹出菜单
            m_Base("variable") = ""
            m_Base("command") = ""
            extra = GetExtraParams(True)
            sOut.Append Space(8) & "self." & m_Base.Parent & ".add_cascade(menu=" & "self." & m_Base.Name & IIf(Len(extra), ", " & extra, "") & ")"
        End If
    ElseIf m_IsSeparator Then                                                   '菜单分隔符
        sOut.Append Space(8) & "self." & m_Base.Parent & ".add_separator()"
    ElseIf m_IsCheckBox Then                                                    ' 菜单增加选择框
        If m_Base("variable") <> "" Then
            sOut.Append Space(8) & "self." & m_Base("variable") & " = StringVar()"
            sOut.Append Space(8) & "self." & m_Base("variable") & ".set(1)"
        End If
        
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams(True)
        sOut.Append Space(8) & "self." & m_Base.Parent & ".add_checkbutton(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDefOOP(m_Base("command"), "event=None")
        End If
    Else
        var = m_Base("variable")
        m_Base("variable") = ""
        If m_Base("command") = "" Then m_Base("command") = m_Base.Name & "_Cmd"
        extra = GetExtraParams(True)
        m_Base("variable") = var
        sOut.Append Space(8) & "self." & m_Base.Parent & ".add_command(" & extra & ")"
        If Len(m_Base("command")) Then
            sCmd.Append m_Base.CreateFuncDefOOP(m_Base("command"), "event=None")
        End If
    End If
    
    '有需要使用bind语句绑定的其他事件处理
    If m_numChilds = 0 And Not m_IsSeparator And Len(m_Base("bindcommand")) Then
        sTmp = UnQuote(m_Base("bindcommand")) '自动去掉括号，如果有的话
        
        s = Split(sTmp, ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If Left(s(i), 1) = "<" And Right(s(i), 1) = ">" Then
                sOut.Append Space(8) & "self." & WTOP & ".bind_all('" & s(i) & "', " & "self." & m_Base("command") & ")"
                
                sr2 = Mid$(s(i), Len(s(i)) - 1, 1)
                sr3 = Mid$(s(i), Len(s(i)) - 2, 1)
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If sr3 = "-" Then
                    If sr2 >= "a" And sr2 <= "z" Then
                        s(i) = Left$(s(i), Len(s(i)) - 2) & UCase$(sr2) & ">"
                        sOut.Append Space(8) & "self." & WTOP & ".bind_all('" & s(i) & "', " & "self." & m_Base("command") & ")"
                    ElseIf sr2 >= "A" And sr2 <= "Z" Then
                        s(i) = Left$(s(i), Len(s(i)) - 2) & LCase$(sr2) & ">"
                        sOut.Append Space(8) & "self." & WTOP & ".bind_all('" & s(i) & "', " & "self." & m_Base("command") & ")"
                    End If
                End If
            End If
        Next
    End If
    
End Sub

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, Optional parentWidth As Long, Optional parentHeight As Long, Optional dMethods As Dictionary)
    
    Dim s As String
    
    m_Base.Name = o.Properties("Name")
    
    m_Visible = o.Properties("Visible")
    m_IsSeparator = (o.Properties("Caption") = "-")
    m_IsCheckBox = (o.Properties("Checked") = True)
    
    '这些是所有的默认值
    m_dicTotal("label") = Replace(o.Properties("Caption"), "&", "")
    m_dicTotal("fg") = ""
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("tearoff") = "0"
    m_dicTotal("relief") = "RAISED"
    m_dicTotal("state") = IIf(o.Properties("Enabled"), "'normal'", "'disabled'")
    m_dicTotal("underline") = IIf(InStr(1, o.Properties("Caption"), "&"), InStr(1, o.Properties("Caption"), "&") - 1, "-1")
    m_dicTotal("variable") = o.Properties("Name") & "Var"
    m_dicTotal("font") = ""
    m_dicTotal("accelerator") = TransShortcut(o.Properties("Shortcut"))
    m_dicTotal("command") = o.Properties("Name") & "_Cmd"
    m_dicTotal("postcommand") = ""
    m_dicTotal("bindcommand") = ""
    
    m_Base("tearoff") = m_dicTotal("tearoff")
    m_Base("variable") = m_dicTotal("variable")
    m_Base("command") = m_dicTotal("command")
    m_Base("label") = m_dicTotal("label")
    If m_dicTotal("state") <> "'normal'" Then m_Base("state") = m_dicTotal("state")
    If m_dicTotal("underline") <> "-1" Then m_Base("underline") = m_dicTotal("underline")
    
    If m_dicTotal("accelerator") <> "" Then                                     '快捷键绑定
        s = m_dicTotal("accelerator")
        m_Base("accelerator") = s
        s = "'<" & Replace(s, "Ctrl", "Control") & ">'"
        m_Base("bindcommand") = s
    End If
    
End Sub

'将VB快捷键的枚举值转换为一个可读的字符串
Private Function TransShortcut(nsc As Long) As String
    Select Case nsc
    Case 0
        TransShortcut = ""
    Case vbextMenuShortcutCtrlA To vbextMenuShortcutCtrlZ                       '1 - 26
        TransShortcut = "Ctrl-" & Chr(nsc + 64)
    Case vbextMenuShortcutF1 To vbextMenuShortcutF12                            ' 27 - 38
        TransShortcut = "F" & CStr(nsc - 26)
    Case vbextMenuShortcutCtrlF1 To vbextMenuShortcutCtrlF12                    ' 39 - 50
        TransShortcut = "Ctrl-F" & CStr(nsc - 38)
    Case vbextMenuShortcutShiftF1 To vbextMenuShortcutShiftF12                  ' 51 - 62
        TransShortcut = "Shift-F" & CStr(nsc - 50)
    Case vbextMenuShortcutCtrlShiftF1 To vbextMenuShortcutCtrlShiftF12          '  63 - 74
        TransShortcut = "Ctrl-Shift-F" & CStr(nsc - 62)
    Case vbextMenuShortcutCtrlIns                                               '= 75
        TransShortcut = "Ctrl-Insert"
    Case vbextMenuShortcutShiftIns                                              '= 76
        TransShortcut = "Shift-Insert"
    Case vbextMenuShortcutDel                                                   '= 77
        TransShortcut = "Delete"
    Case vbextMenuShortcutShiftDel                                              ' = 78
        TransShortcut = "Shift-Delete"
    Case vbextMenuShortcutAltBksp                                               ' = 79
        TransShortcut = "Alt-BackSpace"
    Case Else
        TransShortcut = ""
    End Select
    
End Function


'除了必选参数外，这个函数生成用户选择的其他参数列表
Public Function GetExtraParams(Optional oop As Boolean = False) As String
    
    Dim cfg As Variant, k As Variant, ks As Variant, sValue As String
    
    Set ks = m_Base.Keys
    
    For Each k In ks
        If isExtra(k) And Len(m_Base(k)) Then
            '需要使用引号括起来的属性，如果用户忘了，则在这里自动添加
            If k = "label" Then
                sValue = U(m_Base(k))
            ElseIf InStr(1, " fg, bg, state, accelerator, ", " " & k & ",") Then
                sValue = Quote(m_Base(k))
            Else
                sValue = m_Base(k)
            End If
            GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & k & "=" & sValue
        End If
    Next
    
    If Len(m_Base("command")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "command=" & IIf(oop, "self.", "") & m_Base("command")
    End If
    If Len(m_Base("variable")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "variable=" & IIf(oop, "self.", "") & m_Base("variable")
    End If
    If Len(m_Base("font")) Then
        GetExtraParams = GetExtraParams & IIf(Len(GetExtraParams), ", ", "") & "font=" & IIf(oop, "self.", "") & m_Base.Name & "Font"
    End If
    
End Function

Private Function isExtra(ByVal sK As String) As Boolean
    isExtra = (InStr(1, " tearoff, variable, command, postcommand, bindcommand, font, ", Space(1) & sK & ",") <= 0)
End Function

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If sAttr = "tearoff" Then
        GetAttrValueList = 1
        sa = Split("1,0", ",")
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

Public Function Tips(sAttr As String) As String
    
    If sAttr = "tearoff" Then
        Tips = sAttr & vbCrLf & L("l_TipTearOff", "菜单是否可以变成单独窗口，设置为1（默认）时，菜单第一项为虚线，用户点击这条虚线会将菜单弹出为一个单独的窗口，就像PYTHON默认编辑器IDLE一样。")
    ElseIf sAttr = "postcommand" Then
        Tips = sAttr & vbCrLf & L("l_TipPostCmdMenu", "每次用户点击菜单弹出时调用的回调函数。")
    ElseIf sAttr = "accelerator" Then
        Tips = sAttr & vbCrLf & L("l_TipAcceleratorMenu", "全局快捷键定义。")
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

Private Sub Class_Initialize()
    
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Menu"
    m_Base.StyleName = ""
    m_Base.Parent = "MainMenu"
    Erase m_Childs
    m_numChilds = 0
    m_IdxCurChild = 0
    
End Sub

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    Dim re As Collection, k As Variant, ks As Collection
    
    Set re = New Collection
    
    '标准参数
    Set ks = m_dicTotal.Keys
    For Each k In ks
        If Len(m_Base(k)) Then
            re.Add k & "|" & m_Base(k) & "|1"
        Else
            re.Add k & "|" & m_dicTotal(k) & "|0"
        End If
    Next
    
    '用户增加的自定义参数（如果有的话）
    Set ks = m_Base.Keys
    For Each k In ks
        If Not m_dicTotal.Exists(k) Then
            re.Add k & "|" & m_Base(k) & "|1"
        End If
    Next
    
    Set Allitems = re
    
End Function

Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
End Sub

Public Property Let Parent(s As String)
    m_Base.Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Base.Parent
End Property

Public Property Get Name() As String
    Name = m_Base.Name
End Property

Public Property Let Name(s As String)
    m_Base.Name = s
End Property

'用于改变其默认对应的widget类型，修改widget类型后注意属性列表的合法性
Public Function SetWidgetType(sType As String, sStyleName As String)
    'm_Base.ctlType = sType
    'm_Base.StyleName = sStyleName
End Function

'确定主处理函数能否调用其toString()来产生代码，默认为True，设置为False说明由其他对象来调用处理
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = False
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    'm_CanbeOutByMainForm = bEnable
End Property

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescMenuItem", "菜单项对象，每一个对象对应到Tkinter的Menu响应命令的项目。")
End Property

Public Sub AddChild(o As Object)
    ReDim Preserve m_Childs(m_numChilds) As Object
    Set m_Childs(m_numChilds) = o
    m_numChilds = m_numChilds + 1
End Sub

Public Function GetNextChild(Optional nIdxChild As Long = -1) As Object
    m_IdxCurChild = IIf(nIdxChild >= 0, nIdxChild, m_IdxCurChild)
    If m_IdxCurChild < m_numChilds Then
        Set GetNextChild = m_Childs(m_IdxCurChild)
        m_IdxCurChild = m_IdxCurChild + 1
    Else
        Set GetNextChild = Nothing
        m_IdxCurChild = 0
    End If
End Function

Public Property Get ChildCount() As Long
    ChildCount = m_numChilds
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property

