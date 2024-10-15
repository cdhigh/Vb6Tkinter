VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboboxAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'组合框适配器类
'因为TK中没有组合框，所以VB的组合框映射到Tk的OptionMenu
'如果启动了TTK后，则映射到TTK的Combobox
'VB窗体上一旦有组合框，则OptionMenu和Combobox都创建
'所以使用了一个中间层做为适配器，选择其中一个

Private m_usettk As Boolean
Private m_OptionMenu As clsOptionMenu
Private m_Combobox As clsCombobox
Private m_CanbeOutByMainForm As Boolean

Public Property Let TTK(usettk As Boolean)
    m_usettk = usettk
End Property

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, usettk As Boolean)
    If usettk Then
        m_Combobox.toString sOut, sCmdFunc, rel, usettk
    Else
        m_OptionMenu.toString sOut, sCmdFunc, rel, usettk
    End If
End Sub

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    m_Combobox.InitConfig o, parentWidth, parentHeight, dMethods
    m_OptionMenu.InitConfig o, parentWidth, parentHeight, dMethods
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If m_usettk Then
        GetAttrValueList = m_Combobox.GetAttrValueList(sAttr, sa)
    Else
        GetAttrValueList = m_OptionMenu.GetAttrValueList(sAttr, sa)
    End If
End Function

'判断此控件是否存在对应的属性
Public Function hasAttribute(sAttr As String) As Boolean
    If m_usettk Then
        hasAttribute = m_Combobox.hasAttribute(sAttr)
    Else
        hasAttribute = m_OptionMenu.hasAttribute(sAttr)
    End If
End Function

'获取此控件对应的当前设定的属性值，没有则返回空串
Public Function GetAttrCurrentValue(sAttr As String) As String
    If m_usettk Then
        GetAttrCurrentValue = m_Combobox.GetAttrCurrentValue(sAttr)
    Else
        GetAttrCurrentValue = m_OptionMenu.GetAttrCurrentValue(sAttr)
    End If
End Function

Public Function Tips(sAttr As String) As String
    If m_usettk Then
        Tips = m_Combobox.Tips(sAttr)
    Else
        Tips = m_OptionMenu.Tips(sAttr)
    End If
End Function

Private Sub Class_Initialize()
    m_usettk = True
    Set m_OptionMenu = New clsOptionMenu
    Set m_Combobox = New clsCombobox
    m_CanbeOutByMainForm = True
End Sub

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    If m_usettk Then
        Set Allitems = m_Combobox.Allitems()
    Else
        Set Allitems = m_OptionMenu.Allitems()
    End If
End Function

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    If m_usettk Then
        m_Combobox.SetConfig (sAttrs)
    Else
        m_OptionMenu.SetConfig (sAttrs)
    End If
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    If m_usettk Then
        m_Combobox.SetSingleConfig (sAttr)
    Else
        m_OptionMenu.SetSingleConfig (sAttr)
    End If
End Sub

Public Property Let Parent(s As String)
    m_Combobox.Parent = s
    m_OptionMenu.Parent = s
End Property

Public Property Get Parent() As String
    If m_usettk Then
        Parent = m_Combobox.Parent
    Else
        Parent = m_OptionMenu.Parent
    End If
End Property

Public Property Get Name() As String
    If m_usettk Then
        Name = m_Combobox.Name
    Else
        Name = m_OptionMenu.Name
    End If
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
Public Function Serializer(vSer As clsSerialization)
    If m_usettk Then
        m_Combobox.Serializer vSer
    Else
        m_OptionMenu.Serializer vSer
    End If
End Function

Public Function Deserializer(vSer As clsSerialization)
    If m_usettk Then
        m_Combobox.Deserializer vSer
    Else
        m_OptionMenu.Deserializer vSer
    End If
End Function

Public Property Get Description() As String
    If m_usettk Then
        Description = m_Combobox.Description
    Else
        Description = m_OptionMenu.Description
    End If
End Property


Public Property Let ScaleMode(nV As Long)
    m_Combobox.ScaleMode = nV
    m_OptionMenu.ScaleMode = nV
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


