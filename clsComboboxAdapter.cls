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

Public Property Let TTK(usettk As Boolean)
    m_usettk = usettk
End Property


'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    If m_usettk Then
        toString = m_Combobox.toString(sCmdFunc, rel, oop, usettk)
    Else
        toString = m_OptionMenu.toString(sCmdFunc, rel, oop, usettk)
    End If
End Function

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    m_Combobox.InitConfig o, parentWidth, parentHeight
    m_OptionMenu.InitConfig o, parentWidth, parentHeight
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



