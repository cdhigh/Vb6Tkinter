VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'菜单类，这个类和其他的控件类有一些不同，因为要处理菜单嵌套等特殊情况

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类

Private m_Childs() As Object
Private m_numChilds As Long
Private m_IdxCurChild As Long
Private m_CanbeOutByMainForm As Boolean

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmd As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    If oop Then
        toString = toStringOOP(sCmd, rel)
    Else
        toString = toStringStructure(sCmd, rel)
    End If
    
End Function

'输出结构化代码
Private Function toStringStructure(ByRef sCmd As cStrBuilder, rel As Boolean) As String
    
    Dim sOut As New cStrBuilder, s() As String, i As Long, extra As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    sOut.Append Space(4) & m_Base.Name & " = Menu(" & m_Base.Parent & extra & ")"
    
    '逐个输出各顶级菜单，各顶级菜单下的子菜单由各顶级菜单负责生成对应的代码
    For i = 0 To m_numChilds - 1
        sOut.Append m_Childs(i).toString(sCmd, rel, False, False)
    Next
    
    sOut.Append Space(4) & m_Base.Parent & "['menu'] = " & m_Base.Name
    sOut.Append Space(4) & "gComps['" & m_Base.Name & "'] = " & m_Base.Name
    
    toStringStructure = sOut.toString(vbCrLf)
    
End Function

'输出面向对象代码
Private Function toStringOOP(ByRef sCmd As cStrBuilder, rel As Boolean) As String
    
    Dim sOut As New cStrBuilder, s() As String, i As Long, extra As String
    
    extra = IIf(Len(m_Base("tearoff")), ", tearoff=" & m_Base("tearoff"), "")
    
    sOut.Append Space(8) & "self." & m_Base.Name & " = Menu(self." & m_Base.Parent & extra & ")"
    
    '逐个输出各顶级菜单，各顶级菜单下的子菜单由各顶级菜单负责生成对应的代码
    For i = 0 To m_numChilds - 1
        sOut.Append m_Childs(i).toString(sCmd, rel, True, False)
    Next
    
    sOut.Append Space(8) & "self." & m_Base.Parent & "['menu'] = self." & m_Base.Name
    
    toStringOOP = sOut.toString(vbCrLf)
    
End Function

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(Optional o As Object, Optional parentWidth As Long, Optional parentHeight As Long, Optional dMethods As Dictionary)
    
    m_Base("tearoff") = "0"
    
    '这些是所有的默认值
    m_dicTotal("tearoff") = "0"
    m_dicTotal("label") = ""
    m_dicTotal("fg") = ""
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("relief") = ""
    m_dicTotal("state") = ""
    m_dicTotal("underline") = "-1"
    m_dicTotal("variable") = m_Base.Name & "Var"
    m_dicTotal("font") = ""
    
End Sub

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
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.Name = "MainMenu"
    m_Base.ctlType = "Menu"
    m_Base.StyleName = ""
    Erase m_Childs
    m_numChilds = 0
    m_IdxCurChild = 0
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
    Erase m_Childs
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
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'对象序列化函数
Public Sub Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Sub

Public Sub Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Sub

Public Property Get Description() As String
    Description = L("l_DescMenu", "主菜单对象，对应到Tkinter的Menu控件。")
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

'用于模拟比较排序的函数，实际上是判断两个对象的依赖关系
'用本对象和另一个对象比较，逻辑结果为'本对象-另一个对象'
'返回值含义：
'<0：表示本对象需要在另一个对象前输出代码
'=0：表示两者将没有依赖关系，代码前后顺序无影响
'>0：另一个对象要先输出代码。
'整体的逻辑结果类似是重的沉底
Public Function Compare(ByRef Obj As Object) As Long
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


