VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotebookTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'选项卡中的单页，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_Comps() As Object  '保存此选项页中所有的控件
Private m_CanbeOutByMainForm As Boolean
Private m_SubWidgets() As Object
Private m_numSubWidgets As Long
Private m_Name As String
Private m_Caption As String
Private m_Parent  As String

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean)
    
    Dim i As Long
    
    '得先按依赖关系进行排序
    SortWidgets m_SubWidgets, m_numSubWidgets
    
    For i = 0 To m_numSubWidgets - 1
        m_SubWidgets(i).toString sOut, sCmdFunc, rel, oop, usettk
    Next
    
End Sub


'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, Optional dMethods As Dictionary)
    
    m_Name = o.Properties("Name")
    
    On Error Resume Next
    m_Caption = o.Properties("Caption")
    On Error GoTo 0
    
    Erase m_SubWidgets
    m_numSubWidgets = 0
    
End Sub

'增加一个子控件
Public Sub AddSubWidget(oWidget As Object)
    ReDim Preserve m_SubWidgets(m_numSubWidgets) As Object
    Set m_SubWidgets(m_numSubWidgets) = oWidget
    m_numSubWidgets = m_numSubWidgets + 1
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 0
End Function

'判断此控件是否存在对应的属性
Public Function hasAttribute(sAttr As String) As Boolean
    hasAttribute = False
End Function

'获取此控件对应的当前设定的属性值，没有则返回空串
Public Function GetAttrCurrentValue(sAttr As String) As String
    GetAttrCurrentValue = ""
End Function

Public Function Tips(sAttr As String) As String
    Tips = ""
End Function

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    '''''
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    ''''''
End Sub

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    Set Allitems = New Collection
End Function

Private Sub Class_Initialize()
    m_numSubWidgets = 0
    m_CanbeOutByMainForm = True
    m_Parent = WTOP
End Sub

Public Property Let Parent(s As String)
    m_Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Parent
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

'用于改变其默认对应的widget类型，修改widget类型后注意属性列表的合法性
Public Function SetWidgetType(sType As String, sStyleName As String)
    ''''
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
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescNotebookTab", "Tab of widget Notebook. TTK needed.")
End Property

Public Property Let ScaleMode(nV As Long)
    'm_Base.ScaleMode = nV
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


