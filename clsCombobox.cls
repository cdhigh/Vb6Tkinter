VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCombobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'组合框类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_List As String

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim sOut As New cStrBuilder
    
    '先创建下拉列表
    If Len(m_Base("values")) Then
        If oop Then
            sOut.Append Space(8) & "self." & m_Base.Name & "List = " & m_List
        Else
            sOut.Append Space(4) & m_Base.Name & "List = " & m_List
        End If
        If Len(m_Base("textvariable")) Then m_Base("text") = GetFirstItemText()
    End If
    
    sOut.Append m_Base.toString(sCmdFunc, rel, oop, usettk)
    
    '默认启动后组合框显示第一项
    If Len(m_Base("values")) Then
        If Len(m_Base("textvariable")) Then
            m_Base.Remove ("text")
        Else
            If oop Then
                sOut.Append Space(8) & "self." & m_Base.Name & ".set(self." & m_Base.Name & "List[0])"
            Else
                sOut.Append Space(4) & m_Base.Name & ".set(" & m_Base.Name & "List[0])"
            End If
        End If
    End If
    
    toString = sOut.toString(vbCrLf)
    
End Function

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    
    Dim i As Long, oList As Object, sList As String
    
    m_Base.Name = o.Properties("Name")
    
    '保存下拉选项值
    Set oList = o.Properties("List")
    sList = "["
    i = IIf(o.Properties("Style") = 2, 1, 0)                                    ' 2-Dropdown List时第一项IndexedValue为控件名，所以从第二项开始
    Do While True
        If oList.IndexedValue(i) = "" Then Exit Do
        sList = sList & "'" & oList.IndexedValue(i) & "',"
        i = i + 1
    Loop
    m_List = sList & IIf(sList = "[", "'Add items in design or code!',]", "]")
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("displayrows") = ""
    m_dicTotal("textvariable") = m_Base.Name & "Var"
    m_dicTotal("fg") = m_Base.ColorToRGBStr(o.Properties("ForeColor"))
    m_dicTotal("bg") = m_Base.ColorToRGBStr(o.Properties("BackColor"))
    m_dicTotal("bd") = "1"
    m_dicTotal("relief") = "SUNKEN"
    m_dicTotal("justify") = "'left'"
    m_dicTotal("state") = IIf(o.Properties("Enabled"), IIf(o.Properties("Locked") Or o.Properties("Style") = 2, "'readonly'", "'normal'"), "'disabled'")
    m_dicTotal("values") = m_Base.Name & "List"
    m_dicTotal("font") = "size=" & Round(o.Properties("Font")("Size")) & ",weight=" & IIf(o.Properties("Font")("Bold"), "'bold'", "'normal'") & ",slant=" & IIf(o.Properties("Font")("Italic"), "'italic'", "'roman'")
    m_dicTotal("takefocus") = IIf(o.Properties("TabStop"), 1, 0)
    m_dicTotal("exportselection") = ""
    m_dicTotal("postcommand") = ""
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
    m_Base("values") = m_dicTotal("values")
    If m_dicTotal("fg") <> "" Then m_Base("fg") = m_dicTotal("fg")
    If m_dicTotal("bg") <> "" Then m_Base("bg") = m_dicTotal("bg")
    If m_dicTotal("state") <> "'normal'" Then m_Base("state") = m_dicTotal("state")
    If m_dicTotal("font") <> "size=9,weight='normal',slant='roman'" Then m_Base("font") = m_dicTotal("font")
    If m_dicTotal("takefocus") = 0 Then m_Base("takefocus") = m_dicTotal("takefocus")
    
End Sub

Public Function Tips(sAttr As String) As String
    
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "displayrows":
        Tips = Tips & L("l_TipDisplayRows", "指定要显示行数。")
    Case "state":
        Tips = Tips & L("l_TipStateHasRO", "控件状态，请使用括号括起来，可选值有：'normal', 'readonly', 'disabled'")
    Case "values":
        Tips = Tips & L("l_TipValues", "控件的下拉文本列表，此属性值设置为变量名（为一个列表），如果需要运行时添加和删除下拉文本列表，需要重新使用组合框的configure方法设置values属性。")
    Case "postcommand":
        Tips = Tips & L("l_TipPostCmd", "每次下拉列表打开前调用此回调函数。")
    Case "exportselection":
        Tips = Tips & L("l_TipExportSelection", "默认为False，为True时连接至Window Manager selection。")
    Case Else:
        Tips = m_Base.Tips(sAttr)
    End Select
    
End Function

'获取下拉列表第一项的文本
Private Function GetFirstItemText() As String
    If Len(m_List) Then
        GetFirstItemText = Mid(m_List, 3, InStr(1, m_List, ",") - 4)
    Else
        GetFirstItemText = ""
    End If
End Function

Private Sub Class_Initialize()
    
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Combobox"
    m_Base.StyleName = "TCombobox"
    
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
    Description = L("l_DescCombobox", "文本组合框对象，对应到tkinter.ttk的Combobox控件。")
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property



