VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptionMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'下拉选择菜单类，用于设置和保持相应的属性和值，并生成对应PYTHON的TKinter代码

Private m_dicTotal As Dictionary                                                '保存全部的属性，包括默认值
Private m_Base As clsBaseControl                                                '基础控件类
Private m_List As String

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展(如果使用TTK，则VB的组合框对应到TTK的组合框，不会生出这个控件)
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    If Len(m_Base("variable")) = 0 Then
        MsgBox L_F("l_msgMusthasVar", "OptionMenu(ComboBox)必须要选择variable属性，现在将不生成OptionMenu({0})的代码。", m_Base.Name), vbInformation
        Exit Function
    End If
    
    If oop Then
        toString = toStringOOP(sCmdFunc, rel, usettk)
    Else
        toString = toStringStructure(sCmdFunc, rel, usettk)
    End If
    
End Function

'输出结构化代码
Private Function toStringStructure(ByRef sCmdFunc As cStrBuilder, rel As Boolean, usettk As Boolean) As String
    
    Dim sOut As New cStrBuilder, s() As String, i As Long, extra As String
    
    sOut.Append Space(4) & m_Base.Name & "List = " & m_List
    sOut.Append Space(4) & m_Base("variable") & " = StringVar(value=" & m_Base.Name & "List[0])"
    
    extra = m_Base("variable") & ", *" & m_Base.Name & "List"
    
    sOut.Append Space(4) & m_Base.Name & " = OptionMenu(" & m_Base.Parent & IIf(Len(extra), ", ", "") & extra & ")"
    sOut.Append Space(4) & m_Base.Name & ".place(" & m_Base.GetPositionParams(rel) & ")"
    sOut.Append Space(4) & "gComps['" & m_Base.Name & "'] = " & m_Base.Name
    sOut.Append Space(4) & "gComps['" & m_Base("variable") & "'] = " & m_Base("variable")
    
    '创建事件处理函数框架（如果有）
    sCmdFunc.Append m_Base.CreateFuncDef(m_Base("command"))
    
    If Len(m_Base("bindcommand")) Then                                          '有需要使用bind语句绑定的其他事件处理
        s = Split(m_Base("bindcommand"), ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If Left(s(i), 1) = "<" And Right(s(i), 1) = ">" Then
                sOut.Append Space(4) & m_Base.Name & ".bind('" & s(i) & "', " & m_Base.Name & "_" & Replace(Mid(s(i), 2, Len(s(i)) - 2), "-", "_") & ")"
                sCmdFunc.Append m_Base.CreateFuncDef(m_Base.Name & "_" & Replace(Mid(s(i), 2, Len(s(i)) - 2), "-", "_"), "event")
            End If
        Next
    End If
    
    toStringStructure = sOut.toString(vbCrLf)
    
End Function

'输出面向对象代码
Private Function toStringOOP(ByRef sCmdFunc As cStrBuilder, rel As Boolean, usettk As Boolean) As String
    
    Dim sOut As New cStrBuilder, s() As String, i As Long, extra As String
    
    sOut.Append Space(8) & "self." & m_Base.Name & "List = " & m_List
    sOut.Append Space(8) & "self." & m_Base("variable") & " = StringVar(value=self." & m_Base.Name & "List[0])"
    
    extra = "self." & m_Base("variable") & ", *self." & m_Base.Name & "List"
    
    sOut.Append Space(8) & "self." & m_Base.Name & " = OptionMenu(self." & m_Base.Parent & IIf(Len(extra), ", ", "") & extra & ")"
    sOut.Append Space(8) & "self." & m_Base.Name & ".place(" & m_Base.GetPositionParams(rel) & ")"
    
    '创建事件处理函数框架（如果有）
    sCmdFunc.Append m_Base.CreateFuncDefOOP(m_Base("command"))
    
    If Len(m_Base("bindcommand")) Then                                          '有需要使用bind语句绑定的其他事件处理
        s = Split(m_Base("bindcommand"), ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If Left(s(i), 1) = "<" And Right(s(i), 1) = ">" Then
                sOut.Append Space(8) & "self." & m_Base.Name & ".bind('" & s(i) & "', self." & m_Base.Name & "_" & Replace(Mid(s(i), 2, Len(s(i)) - 2), "-", "_") & ")"
                sCmdFunc.Append m_Base.CreateFuncDefOOP(m_Base.Name & "_" & Replace(Mid(s(i), 2, Len(s(i)) - 2), "-", "_"), "event")
            End If
        Next
    End If
    
    toStringOOP = sOut.toString(vbCrLf)
    
End Function


'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long)
    
    Dim i As Long, oList As Object, sList As String
    
    m_Base.Name = o.Properties("Name")
    
    '保存下拉选项值
    Set oList = o.Properties("List")
    sList = "("
    i = IIf(o.Properties("Style") = 2, 1, 0)                                    ' 2-Dropdown List时第一项IndexedValue为控件名，所以从第二项开始
    Do While True
        If oList.IndexedValue(i) = "" Then Exit Do
        sList = sList & "'" & oList.IndexedValue(i) & "',"
        i = i + 1
    Loop
    m_List = sList & IIf(sList = "(", "'Add items in design or code!',)", ")")
    
    m_dicTotal("x") = m_Base.toPixelX(o.Properties("Left"))
    m_dicTotal("y") = m_Base.toPixelY(o.Properties("Top"))
    m_dicTotal("width") = m_Base.toPixelX(o.Properties("Width"))
    m_dicTotal("height") = m_Base.toPixelY(o.Properties("Height"))
    m_dicTotal("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_dicTotal("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_dicTotal("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_dicTotal("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    m_dicTotal("variable") = o.Properties("Name") & "Var"
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
    m_Base("variable") = m_dicTotal("variable")
    
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

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "OptionMenu"
    m_Base.StyleName = "TMenubutton"
End Sub

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
End Sub

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
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
    Description = L("l_DescOptionmenu", "组合框对象，需要预先设置要显示的字符列表，对应到Tkinter的Optionmenu控件。")
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property


