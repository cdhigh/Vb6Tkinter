VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'数值选择条类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim s As String
    If usettk And Len(m_Base("label")) Then
        s = m_Base("label")
        m_Base("label") = ""
    End If
    
    toString = m_Base.toString(sCmdFunc, rel, oop, usettk)
    
    If usettk And Len(s) Then
        m_Base("label") = s
    End If
    
End Function

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    
    Dim nIdx As Variant
    
    m_Base.Name = o.Properties("Name")
    m_Base.StyleName = IIf(o.Properties("Orientation"), "Vertical", "Horizontal") & ".TScale"
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("label") = ""
    m_dicTotal("fg") = ""
    m_dicTotal("bg") = ""
    m_dicTotal("bd") = ""
    m_dicTotal("from_") = o.Properties("Min")
    m_dicTotal("to") = o.Properties("Max")
    m_dicTotal("resolution") = o.Properties("SmallChange")
    m_dicTotal("relief") = "FLAT"
    m_dicTotal("digits") = 0
    m_dicTotal("orient") = IIf(o.Properties("Orientation"), "'vertical'", "'horizontal'")
    m_dicTotal("showvalue") = 1
    m_dicTotal("variable") = ""
    m_dicTotal("sliderlength") = 30
    m_dicTotal("sliderrelief") = "RAISED"
    m_dicTotal("tickinterval") = 0
    m_dicTotal("font") = ""
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("command") = o.Properties("Name") & "_Cmd"
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
    m_Base("label") = m_dicTotal("label")
    'm_Base("command") = m_dicTotal("command")
    m_Base("orient") = m_dicTotal("orient")
    m_Base("from_") = m_dicTotal("from_")
    m_Base("to") = m_dicTotal("to")
    m_Base("resolution") = m_dicTotal("resolution")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    If sAttr = "showvalue" Then
        sa = Split("1,0", ",")
    ElseIf sAttr = "digits" Then
        sa = Split("0,1,2,3,4", ",")
    ElseIf sAttr = "sliderrelief" Then
        sa = Split("FLAT,GROOVE,RAISED,RIDGE,SOLID,SUNKEN", ",")
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

Public Function Tips(sAttr As String) As String
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "from_":
        Tips = Tips & L("l_TipFrom", "设置最小值。")
    Case "to":
        Tips = Tips & L("l_TipTo", "设置最大值。")
    Case "resolution":
        Tips = Tips & L("l_TipResolution", "设置步距值。")
    Case "showvalue":
        Tips = Tips & L("l_TipShowValue", "是否将当前值以文字形式显示在上方，默认为1。")
    Case "digits":
        Tips = Tips & L("l_TipDigits", "控制浮点数的显示位数。")
    Case "sliderlength":
        Tips = Tips & L("l_TipSliderLen", "滑动指示块的长度，默认为30像素。")
    Case "sliderrelief":
        Tips = Tips & L("l_TipSliderRelief", "滑动指示块的外观效果，可选值有：FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN。")
    Case "tickinterval":
        Tips = Tips & L("l_TipTickInterval", "滑动槽中是否显示周期定位点，默认不显示，如果要显示，这个值设置为多少步长一个点，可以为小数。")
    Case Else:
        Tips = m_Base.Tips(sAttr)
    End Select
End Function

Private Sub Class_Initialize()
    
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Scale"
    m_Base.StyleName = "TScale"
    
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

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
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

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescScale", "滑动条对象，对应到Tkinter的Scale控件。")
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property


