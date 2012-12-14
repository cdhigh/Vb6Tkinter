VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatusbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'状态栏类，因为TK和TTK库都没有状态栏控件，这里自己实现一个状态栏控件

'一个简单的状态栏python控件类定义，因为VB不允许太多的续行符，所以分成两部分，在程序中再连接起来
'此控件的使用方法：
'  statusbar = Statusbar(top)  #创建仅有一个窗格的状态栏
'  statusbar = Statusbar(top, slots=[30,10,]) #创建三个窗格，前两个宽度分别为30,10像素，第三个占据剩余空间
'  statusbar.pack(side=BOTTOM, fill=X)   #放置在窗体最下面
'  statusbar.set(0,'Demo mode')  #设置第一个窗格的内容
Const DefineStatusbar1  As String = "class Statusbar(Frame):" & vbCrLf & _
"    """ & """" & """A Simple Statusbar" & vbCrLf & _
"    Usage:self.status = Statusbar(self.top, panelwidths=(15,5,))" & vbCrLf & _
"          self.status.pack(side=BOTTOM, fill=X)" & vbCrLf & _
"          self.status.set(0,'Demo mode')" & vbCrLf & _
"          self.status.text('Demo mode')" & vbCrLf & _
"    """ & """" & """" & vbCrLf & _
"    def __init__(self, master, **kw):" & vbCrLf & _
"        """ & """" & """Options:" & vbCrLf & _
"        panelwidths - a tuple of width of panels, atual number of panels is len(panelwidths)+1." & vbCrLf & _
"        """ & """" & """" & vbCrLf & _
"        Frame.__init__(self, master)" & vbCrLf & _
"        panelwidths = kw['panelwidths'] if 'panelwidths' in kw else []" & vbCrLf & _
"        self.lbls = []" & vbCrLf & _
"        for pnlwidth in panelwidths:" & vbCrLf & _
"            lbl = Label(self, width=pnlwidth, anchor=W, relief=SUNKEN)" & vbCrLf & _
"            self.lbls.append(lbl)" & vbCrLf & _
"            lbl.pack(side=LEFT, fill=Y)" & vbCrLf & _
"        lbl = Label(self, anchor=W, relief=SUNKEN)" & vbCrLf & _
"        self.lbls.append(lbl)" & vbCrLf & _
"        lbl.pack(fill=BOTH, expand=1)" & vbCrLf & vbCrLf

Const DefineStatusbar2 As String = "    def set(self, panel, format, *args):" & vbCrLf & _
"        if panel >= len(self.lbls): raise IndexError" & vbCrLf & _
"        self.lbls[panel]['text'] = format % args" & vbCrLf & _
"        self.lbls[panel].update_idletasks()" & vbCrLf & vbCrLf & _
"    text = lambda self,format,*args : self.set(0,format,*args)" & vbCrLf & vbCrLf & _
"    def panelwidth(self, panel, width=None):" & vbCrLf & _
"        if panel >= len(self.lbls): raise IndexError" & vbCrLf & _
"        if width is None:" & vbCrLf & _
"            panelwidth = self.lbls[panel]['width']" & vbCrLf & _
"        else:" & vbCrLf & _
"            self.lbls[panel]['width'] = width" & vbCrLf & vbCrLf & _
"    def clear(self):" & vbCrLf & _
"        for panel in self.lbls:" & vbCrLf & _
"            panel.config(text='')" & vbCrLf & _
"            panel.update_idletasks()" & vbCrLf & vbCrLf

Private m_Base As clsBaseControl
Private m_CanbeOutByMainForm As Boolean

'输出状态栏控件的定义
Public Property Get WidgetCode() As String
    WidgetCode = DefineStatusbar1 & DefineStatusbar2
End Property

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean)
    
    If oop Then
        sOut.Append Space(8) & "self." & m_Base.Name & " = Statusbar(self." & m_Base.Parent & _
            IIf(Len(m_Base("panelwidths")), ", panelwidths=" & m_Base("panelwidths"), "") & ")"
        If Len(m_Base("side")) Then
            sOut.Append Space(8) & "self." & m_Base.Name & ".pack(side=" & m_Base("side") & ", fill=X)"
        ElseIf rel Then
            sOut.Append Space(8) & "self." & m_Base.Name & ".place(relx=" & m_Base("relx") & ", rely=" & _
                m_Base("rely") & ", relwidth=" & m_Base("relwidth") & ", relheight=" & m_Base("relheight") & ")"
        Else
            sOut.Append Space(8) & "self." & m_Base.Name & ".place(x=" & m_Base("x") & ", y=" & m_Base("y") & _
                ", width=" & m_Base("width") & ", height=" & m_Base("height") & ")"
        End If
    Else
        sOut.Append Space(4) & m_Base.Name & " = Statusbar(" & m_Base.Parent & _
            IIf(Len(m_Base("panelwidths")), ", panelwidths=" & m_Base("panelwidths"), "") & ")"
        If Len(m_Base("side")) Then
            sOut.Append Space(4) & m_Base.Name & ".pack(side=" & m_Base("side") & ", fill=X)"
        ElseIf rel Then
            sOut.Append Space(4) & m_Base.Name & ".place(relx=" & m_Base("relx") & ", rely=" & m_Base("rely") & _
                ", relwidth=" & m_Base("relwidth") & ", relheight=" & m_Base("relheight") & ")"
        Else
            sOut.Append Space(4) & m_Base.Name & ".place(x=" & m_Base("x") & ", y=" & m_Base("y") & ", width=" & _
                m_Base("width") & ", height=" & m_Base("height") & ")"
        End If
    End If
    
End Sub

'创建对象后要马上调用这个函数初始化各参数
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, dMethods As Dictionary)
    
    Dim i As Long, nCnt As Long, s As String, Panels As IPanels
    
    m_Base.Name = o.Properties("Name")
    
    m_Base("x") = m_Base.toPixelX(o.Properties("Left"))
    m_Base("y") = m_Base.toPixelY(o.Properties("Top"))
    m_Base("width") = m_Base.toPixelX(o.Properties("Width"))
    m_Base("height") = m_Base.toPixelY(o.Properties("Height"))
    m_Base("relx") = Format(o.Properties("Left") / parentWidth, "0.###")
    m_Base("rely") = Format(o.Properties("Top") / parentHeight, "0.###")
    m_Base("relwidth") = Format(o.Properties("Width") / parentWidth, "0.###")
    m_Base("relheight") = Format(o.Properties("Height") / parentHeight, "0.###")
    
    If o.Properties("Align") = vbAlignTop Then
        m_Base("side") = "TOP"
    ElseIf o.Properties("Align") = vbAlignBottom Then
        m_Base("side") = "BOTTOM"
    End If
    
    m_Base("panelwidths") = ""
    Set Panels = o.Properties("Panels").object
    nCnt = Panels.Count
    If o.Properties("Style") = sbrNormal And nCnt > 0 Then
        For i = 1 To nCnt
            s = s & Round(m_Base.toPixelX(Panels(i).MinWidth) / CharWidth()) & ","
        Next
        
        If Len(s) Then
            m_Base("panelwidths") = "(" & s & ")"
        End If
    End If
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    If sAttr = "side" Then
        GetAttrValueList = 2
        sa = Split("TOP,BOTTOM", ",")
    Else
        GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
    End If
End Function

Public Function Tips(sAttr As String) As String
    
    If sAttr = "side" Then
        Tips = sAttr & vbCrLf & L("l_TipSide", "状态栏放置位置\n可选值有：TOP,BOTTOM,留空则可以放置在任何位置。")
    ElseIf sAttr = "panelwidths" Then
        Tips = sAttr & vbCrLf & L("l_TipPanelWidths", "一个元组或列表，状态栏各窗格的宽度（字符数），最后一个窗格总是占满剩余空间，不需要列出。")
    Else
        Tips = m_Base.Tips(sAttr)
    End If
    
End Function

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)

    Dim sa() As String, i As Long
    
    sa = Split(sAttrs, "|")
    Debug.Assert (UBound(sa) Mod 1 = 0)
    
    'm_Base.RemoveAll
    
    For i = 0 To UBound(sa) - 1 Step 2
        SetSingleConfig sa(i) & "|" & sa(i + 1)
    Next
    
End Sub

'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    
    Dim sa() As String, s As String
    
    sa = Split(sAttr, "|")
    Debug.Assert (UBound(sa) = 1)
    
    If sa(0) = "panelwidths" Then
        s = Replace(Replace(sa(1), "'", ""), " ", "")
        s = Replace(Replace(s, "[", "("), "]", ")")
        If Left$(s, 1) <> "(" And Right$(s, 1) <> ")" Then
            If InStr(1, s, ",") Then '使用逗号分隔
                m_Base("panelwidths") = "(" & s & ")"
            ElseIf InStr(1, s, " ") Then '使用空格分隔
                s = Replace(Replace(s, "  ", " "), " ", ",")
                m_Base("panelwidths") = "(" & s & ")"
            Else
                m_Base("panelwidths") = s
            End If
        Else
            m_Base("panelwidths") = s
        End If
    Else
        m_Base(sa(0)) = sa(1)
    End If
    
End Sub

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    Dim re As Collection, k As Variant, ks As Collection
    
    Set re = New Collection
    
    Set ks = m_Base.Keys
    For Each k In ks
        re.Add k & "|" & m_Base(k) & "|1"
    Next
    
    Set Allitems = re
    
End Function

Private Sub Class_Initialize()
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Statusbar"
    m_Base.StyleName = ""
    m_CanbeOutByMainForm = True
End Sub

Private Sub Class_Terminate()
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
    Description = L("l_DescStatusbar", "状态栏对象，为自定义控件，可设置多个窗格，直接使用控件的set()方法设置状态提示。")
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



