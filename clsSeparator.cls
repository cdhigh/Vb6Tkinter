VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'分隔条类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_CanbeOutByMainForm As Boolean

'输出PYTHON代码,
'sOut: 输出参数，界面代码
'sCmd: 输出参数，事件处理回调代码
'sI18n: 输出参数，控件文本翻译代码
'rel:是否使用相对坐标,
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmd As cStrBuilder, ByRef sI18n As cStrBuilder, ByVal rel As Boolean, ByVal usettk As Boolean)
    If Len(m_Base("orient")) = 0 Then m_Base("orient") = "'horizontal'"
    m_Base.toString sOut, sCmd, sI18n, rel, usettk
End Sub

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, Optional dMethods As Dictionary)
    
    Dim W As Long, H As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    Dim s As String, sa() As String, I As Long, idx As Long
    
    m_Base.SetVbWidgetInstance o
    
    m_Base.Name = o.Properties("Name")
    
    If o.Properties("X1") <= o.Properties("X2") Then
        X1 = o.Properties("X1")
        Y1 = o.Properties("Y1")
        X2 = o.Properties("X2")
        Y2 = o.Properties("Y2")
    Else
        X1 = o.Properties("X2")
        Y1 = o.Properties("Y2")
        X2 = o.Properties("X1")
        Y2 = o.Properties("Y1")
    End If
    
    m_dicTotal("x") = m_Base.toPixelX(X1)
    m_dicTotal("y") = m_Base.toPixelY(Y1)
    
    '判断是水平还是垂直的，如果LINE控件是斜的，则如果水平投影较长则认为水平，垂直的判断亦然
    W = Abs(X2 - X1)
    H = Abs(Y2 - Y1)
    If W >= H Then
        m_dicTotal("width") = m_Base.toPixelX(W)
        m_dicTotal("height") = o.Properties("BorderWidth")
        m_dicTotal("relx") = Format(X1 / parentWidth, "0.###")
        m_dicTotal("rely") = Format(Y1 / parentHeight, "0.###")
        m_dicTotal("relwidth") = Format(W / parentWidth, "0.###")
        m_dicTotal("relheight") = Format(o.Properties("BorderWidth") / m_Base.toPixelY(parentHeight), "0.####")
        m_dicTotal("orient") = "'horizontal'"
    Else
        m_dicTotal("width") = o.Properties("BorderWidth")
        m_dicTotal("height") = m_Base.toPixelX(H)
        m_dicTotal("relx") = Format(X1 / parentWidth, "0.###")
        m_dicTotal("rely") = Format(Y1 / parentHeight, "0.###")
        m_dicTotal("relwidth") = Format(o.Properties("BorderWidth") / m_Base.toPixelX(parentWidth), "0.####")
        m_dicTotal("relheight") = Format(H / parentHeight, "0.###")
        m_dicTotal("orient") = "'vertical'"
    End If
    
    m_dicTotal("bg") = TranslateColor(o.Properties("BorderColor"))
    
    '这些是默认选择的
    m_Base("x") = m_dicTotal("x")
    m_Base("y") = m_dicTotal("y")
    m_Base("width") = m_dicTotal("width")
    m_Base("height") = m_dicTotal("height")
    m_Base("relx") = m_dicTotal("relx")
    m_Base("rely") = m_dicTotal("rely")
    m_Base("relwidth") = m_dicTotal("relwidth")
    m_Base("relheight") = m_dicTotal("relheight")
    m_Base("orient") = m_dicTotal("orient")
    'If o.Properties("BorderColor") <> vbWindowText Then
    m_Base("bg") = m_dicTotal("bg")
    
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
    GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
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
    Tips = m_Base.Tips(sAttr)
End Function

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
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

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Separator"
    m_Base.StyleName = "TSeparator"
    m_CanbeOutByMainForm = True
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
    Description = L("l_DescSeparator", "Separator widget. It only supports horizontal or vertical style.")
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

