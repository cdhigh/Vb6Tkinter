VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'多行文本类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_CanbeOutByMainForm As Boolean
Private m_Text As String

'输出PYTHON代码,
'sOut: 输出参数，界面代码
'sCmd: 输出参数，事件处理回调代码
'sI18n: 输出参数，控件文本翻译代码
'rel:是否使用相对坐标,
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    Dim scrlcmd As String
    
    m_Base.toString sOut, sCmd, sI18n, rel, usettk
    
    '如果有初始文本，则插入文本到文本框
    If Len(m_Text) > 0 Then
        sOut.Append "        self." & m_Base.Name & ".insert('1.0'," & m_Text & ")"
    End If
    
    ' 增加滚动条的绑定处理
    scrlcmd = m_Base("xscrollcommand")
    If Right$(scrlcmd, 4) = ".set" Then
        sOut.Append "        self." & Left$(scrlcmd, Len(scrlcmd) - 4) & "['command'] = self." & m_Base.Name & ".xview"
    End If
    
    scrlcmd = m_Base("yscrollcommand")
    If Right$(scrlcmd, 4) = ".set" Then
        sOut.Append "        self." & Left$(scrlcmd, Len(scrlcmd) - 4) & "['command'] = self." & m_Base.Name & ".yview"
    End If
End Sub

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    Dim s As String, sa() As String, I As Long, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    m_Text = Replace(o.Properties("Text"), vbCrLf, "\n")
    If InStr(1, m_Text, Chr(34)) > 0 Or InStr(1, m_Text, "'") > 0 Then
        m_Text = String(3, Chr(34)) & m_Text & String(3, Chr(34))
    Else
        m_Text = "'" & m_Text & "'"
    End If
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("fg") = TranslateColor(o.Properties("ForeColor"))
    m_dicTotal("bg") = TranslateColor(o.Properties("BackColor"))
    m_dicTotal("bd") = "1"
    m_dicTotal("relief") = "SUNKEN"
    m_dicTotal("state") = IIf(o.Properties("Enabled"), "'normal'", "'disabled'")
    m_dicTotal("wrap") = "CHAR"
    m_dicTotal("spacing1") = "0"
    m_dicTotal("spacing2") = "0"
    m_dicTotal("spacing3") = "0"
    m_dicTotal("font") = "(" & U(CStr(o.Properties("Font")("Name"))) & "," & Round(o.Properties("Font")("Size")) & _
                        IIf(o.Properties("Font")("Bold"), ",'bold'", "") & IIf(o.Properties("Font")("Italic"), ",'italic'", "") & _
                        IIf(o.Properties("Font")("Underline"), ",'underline'", "") & IIf(o.Properties("Font")("Strikethrough"), ",'overstrike')", ")")
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("tooltip") = o.Properties("ToolTipText")
    m_dicTotal("cursor") = m_Base.GetCursorName(o.Properties("MousePointer"))
    m_dicTotal("xscrollcommand") = ""
    m_dicTotal("yscrollcommand") = ""
    m_dicTotal("bindcommand") = m_Base.GetBindCommandStr(dMethods)
    
    '这些是默认选择的
    m_Base("x") = m_dicTotal("x")
    m_Base("y") = m_dicTotal("y")
    m_Base("width") = m_dicTotal("width")
    m_Base("height") = m_dicTotal("height")
    m_Base("relx") = m_dicTotal("relx")
    m_Base("rely") = m_dicTotal("rely")
    m_Base("relwidth") = m_dicTotal("relwidth")
    m_Base("relheight") = m_dicTotal("relheight")
    If m_dicTotal("tooltip") <> "" Then m_Base("tooltip") = m_dicTotal("tooltip")
    If o.Properties("ForeColor") <> vbWindowText Then m_Base("fg") = m_dicTotal("fg")
    If o.Properties("BackColor") <> vbWindowBackground Then m_Base("bg") = m_dicTotal("bg")
    If m_dicTotal("state") <> "'normal'" Then m_Base("state") = m_dicTotal("state")
    If m_dicTotal("font") <> "(" & U(GetDefaultFontName()) & ",9)" Then m_Base("font") = m_dicTotal("font")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    If m_dicTotal("cursor") <> "" Then m_Base("cursor") = m_dicTotal("cursor")
    If m_dicTotal("bindcommand") <> "" Then m_Base("bindcommand") = m_dicTotal("bindcommand")
    
    '使用控件的tag属性保存一些额外默认要保存的属性，
    '开始字符为：p@，后接属性名，每个属性使用@隔开，属性值可选，如果有值，则使用k=v格式
    s = Trim(o.Properties("Tag"))
    If Len(s) Then
        sa = Split(s, "@")
        If sa(0) = "p" Or sa(0) = "P" Then
            For I = 1 To UBound(sa)
                s = sa(I)
                idx = InStr(2, s, "=")
                If idx > 0 Then '有值
                    m_dicTotal(Left$(s, idx - 1)) = Mid$(s, idx + 1)
                    m_Base(Left$(s, idx - 1)) = m_dicTotal(Left$(s, idx - 1))
                ElseIf Len(s) > 0 And Len(m_dicTotal(s)) > 0 Then '仅有属性，则判断属性是否合法
                    m_Base(s) = m_dicTotal(s)
                End If
            Next
        End If
    End If
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    Dim aComps() As String, I As Long, satmp() As String, sOut As String
    
    GetAttrValueList = 1
    If sAttr = "state" Then
        sa = Split("'normal','readonly','disabled'", ",")
    ElseIf sAttr = "wrap" Then
        sa = Split("CHAR,WORD,NONE", ",")
    ElseIf sAttr = "xscrollcommand" Or sAttr = "yscrollcommand" Then
        aComps = GetAllComps()
        If UBound(aComps) >= 0 Then
            sOut = ""
            For I = 1 To UBound(aComps) '0为窗体，不需要判断
                satmp = Split(aComps(I), "|")
                If UBound(satmp) > 0 Then
                    If satmp(1) = "clsScrollbar" Then
                        sOut = sOut & IIf(Len(sOut), ",", "") & satmp(0) & ".set"
                    End If
                End If
            Next
            sa = Split(sOut, ",")
            GetAttrValueList = IIf(Len(sOut) > 0, 1, 0)
        Else
            GetAttrValueList = 0
        End If
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

'判断此控件是否存在对应的属性
Public Function hasAttribute(sAttr As String) As Boolean
    hasAttribute = m_Base.hasAttribute(sAttr)
End Function

'获取此控件对应的当前设定的属性值，没有则返回空串
Public Function GetAttrCurrentValue(sAttr As String) As String
    GetAttrCurrentValue = m_Base.GetAttrCurrentValue(sAttr)
End Function

Public Function Tips(sAttr As String) As String
    
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "state"
        Tips = Tips & L("l_TipStateHasRO", "State of the widget. They are 'normal', 'readonly', 'disabled'")
    Case "wrap"
        Tips = Tips & L("l_TipWrapMode", "Controls the display of lines that are too wide.\n'CHAR':will be broken at any character.\n'WORD': will break the line after the last word that will fit.\n'NONE':provide a horizontal scrollbar.")
    Case "spacing1"
        Tips = Tips & L("l_TipSpacing1", "Specifies how much extra vertical space is put above each line of text.")
    Case "spacing2"
        Tips = Tips & L("l_TipSpacing2", "Specifies how much extra vertical space to add between displayed lines of text when a logical line wraps.")
    Case "spacing3"
        Tips = Tips & L("l_TipSpacing3", "Specifies how much extra vertical space is added below each line of text.")
    Case Else
        Tips = m_Base.Tips(sAttr)
    End Select
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Text"
    m_Base.StyleName = ""
    m_CanbeOutByMainForm = True
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

'用于改变其默认对应的widget类型，修改widget类型后注意属性列表的合法性
Public Function SetWidgetType(sType As String, sStyleName As String)
    m_Base.ctlType = sType
    m_Base.StyleName = sStyleName
End Function

'确定主处理函数能否调用其toString()来产生代码，默认为True，设置为False说明由其他对象来调用处理
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescText", "Text widget.")
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property

'用于模拟比较排序的函数，实际上是判断两个对象的依赖关系
'用本对象和另一个对象比较，逻辑结果为'本对象-另一个对象'
'返回值含义：
'<0：表示本对象需要在另一个对象前输出代码
'=0：表示两者将没有依赖关系，代码前后顺序无影响
'>0：另一个对象要先输出代码。
'整体的逻辑结果类似是重的沉底
Public Function Compare(ByRef Obj As Object) As Long
    
    Dim scrlcmd As String, scrl As String
    
    '先判断和滚动条的依赖关系
    scrlcmd = m_Base("xscrollcommand")
    If Len(scrlcmd) And TypeName(Obj) = "clsScrollbar" Then
        scrl = IIf(Right$(scrlcmd, 4) = ".set", Left$(scrlcmd, Len(scrlcmd) - 4), "")
        If scrl = Obj.Name Then
            Compare = 1 ' 滚动条先
            Exit Function
        End If
    End If
    scrlcmd = m_Base("yscrollcommand")
    If Len(scrlcmd) And TypeName(Obj) = "clsScrollbar" Then
        scrl = IIf(Right$(scrlcmd, 4) = ".set", Left$(scrlcmd, Len(scrlcmd) - 4), "")
        If scrl = Obj.Name Then
            Compare = 1 ' 滚动条先
            Exit Function
        End If
    End If
    
    If Parent = Obj.Name Then '父控件先输出代码
        Compare = 1
    ElseIf Obj.Parent = Name Then
        Compare = -1
    ElseIf Parent = WTOP And Obj.Parent <> WTOP Then '顶层控件先输出
        Compare = -1
    ElseIf Parent <> WTOP And Obj.Parent = WTOP Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function



