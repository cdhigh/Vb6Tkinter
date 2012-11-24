VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'选项卡类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary       '保存全部的属性，包括默认值
Private m_Base As clsBaseControl       '基础控件类
Private m_Tabs() As String             '选项卡标题，下标从1开始的，0实际没用
Private m_nTabs As Long                '选项卡数目

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim i As Long, sOut As New cStrBuilder, s As String
    
    sOut.Append m_Base.toString(sCmdFunc, rel, oop, usettk)
    
    '增加选项卡
    If oop Then
        sOut.Append vbCrLf & Space(8) & L("l_cmtTodoNBFrame", "#TODO，请在各个Frame(代表各选项卡内容)里面添加需要的控件。")
        For i = 1 To m_nTabs  'TAB下标从1开始
            s = "self." & m_Base.Name & "Frame" & i
            sOut.Append Space(8) & s & " = Frame(self." & m_Base.Name & ")"
            sOut.Append Space(8) & s & "Lbl = Label(" & s & ", text='Please add widgets in code.')"
            sOut.Append Space(8) & s & "Lbl.place(relx=0.1,rely=0.5)" & vbCrLf
        Next
        
        For i = 1 To m_nTabs
            s = "self." & m_Base.Name & "Frame" & i
            sOut.Append Space(8) & "self." & m_Base.Name & ".add(" & s & ", text=" & U(m_Tabs(i)) & ")"
        Next
    Else
        sOut.Append vbCrLf & Space(4) & L("l_cmtTodoNBFrame", "#TODO，请在各个Frame(代表各选项卡内容)里面添加需要的控件。")
        For i = 1 To m_nTabs
            s = m_Base.Name & "Frame" & i
            sOut.Append Space(4) & s & " = Frame(" & m_Base.Name & ")"
            sOut.Append Space(4) & s & "Lbl = Label(" & s & ", text='Please add widgets in code.')"
            sOut.Append Space(4) & s & "Lbl.place(relx=0.1,rely=0.5)" & vbCrLf
        Next
        
        For i = 1 To m_nTabs
            s = m_Base.Name & "Frame" & i
            sOut.Append Space(4) & m_Base.Name & ".add(" & s & ", text=" & U(m_Tabs(i)) & ")"
        Next
    End If
    
    toString = sOut.toString(vbCrLf)
    
End Function


'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    
    Dim i As Long, oiTabs As ITabs
    
    m_Base.Name = o.Properties("Name")
    
    Set oiTabs = o.Properties("Tabs").object
    m_nTabs = oiTabs.Count
    
    ReDim m_Tabs(m_nTabs) As String
    For i = 1 To m_nTabs
        m_Tabs(i) = oiTabs.Item(i).Caption  'TAB下标从1开始
    Next
    
    '这些是所有的默认值
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("relief") = "FLAT"
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
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
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
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
    m_Base.ctlType = "Notebook"
    m_Base.StyleName = "TNotebook"
    ReDim m_Tabs(0) As String
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
    Erase m_Tabs
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
    Description = L("l_DescNotebook", "选项卡对象，对应到Tkinter.ttk的Notebook控件。")
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
Public Function Compare(ByRef obj As Object) As Long
    If Parent = "top" And obj.Parent <> "top" Then '顶层控件先输出
        Compare = -1
    ElseIf Parent <> "top" And obj.Parent = "top" Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function



