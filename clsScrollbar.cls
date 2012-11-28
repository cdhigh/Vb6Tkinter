VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'滚动条类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_CanbeOutByMainForm As Boolean

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim sCmd As String
    sCmd = m_Base("command")
    
    If oop Then
        sCmdFunc.Append m_Base.CreateFuncDefOOP(sCmd, "method, dir, jump=None")
    Else
        sCmdFunc.Append m_Base.CreateFuncDef(sCmd, "method, dir, jump=None")
    End If
    m_Base("command") = ""
    
    toString = m_Base.toString(sCmdFunc, rel, oop, usettk)
    
    m_Base("command") = sCmd
    
End Function

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    
    m_Base.Name = o.Properties("Name")
    m_Base.StyleName = IIf(o.ClassName = "HScrollBar", "Horizontal", "Vertical") & ".TScrollbar"
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("orient") = IIf(o.ClassName = "HScrollBar", "'horizontal'", "'vertical'")
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("jump") = 0
    m_dicTotal("relief") = "SUNKEN"
    m_dicTotal("activerelief") = "SUNKEN"
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("command") = "" 'o.Properties("Name") & "_Cmd"
    m_dicTotal("bindcommand") = ""
    
    '这些是默认选择的
    m_Base("x") = m_dicTotal("x")
    m_Base("y") = m_dicTotal("y")
    m_Base("width") = m_dicTotal("width")
    m_Base("height") = m_dicTotal("height")
    m_Base("relx") = m_dicTotal("relx")
    m_Base("rely") = m_dicTotal("rely")
    m_Base("relwidth") = m_dicTotal("relwidth")
    m_Base("relheight") = m_dicTotal("relheight")
    'm_Base("command") = m_dicTotal("command")
    m_Base("orient") = m_dicTotal("orient")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    If sAttr = "jump" Then
        sa = Split("1,0", ",")
    ElseIf sAttr = "activerelief" Then
        sa = Split("FLAT,GROOVE,RAISED,RIDGE,SOLID,SUNKEN", ",")
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

Public Function Tips(sAttr As String) As String
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "jump"
        Tips = Tips & L("l_TipJump", "设置为0（默认值）则每次滚动条数值变化就调用一次回调函数，设置为1则等鼠标松开后才调一次回调函数。")
    Case "activerelief"
        Tips = Tips & L("l_TipActiveRelief", "设置鼠标指针移动在滚动条上时的式样，可选值：FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN。")
    Case Else:
        Tips = m_Base.Tips(sAttr)
    End Select
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Scrollbar"
    m_Base.StyleName = "TScrollbar"
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
    Description = L("l_DescScrollbar", "滚动条对象，对应到Tkinter的Scrollbar控件。")
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
    Dim tObj As String
    tObj = TypeName(Obj)
    '如果是可滚动控件和滚动条相比较，则交给可滚动对象来比较
    If tObj = "clsListbox" Or tObj = "clsText" Or tObj = "clsCanvas" Or tObj = "clsTreeview" Then
        Compare = Obj.Compare(Me)
        Compare = Switch(Compare > 0, -1, Compare < 0, 1, Compare = 0, 0) '比较结果要反个个
        Exit Function
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



