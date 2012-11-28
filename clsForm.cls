VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'窗体类的抽象，保存VB窗体属性

Private m_Base As clsBaseControl
Private m_Name As String
Private m_Title As String
Private m_Geometry As String
Private m_Resizable As Boolean
Private m_IconFile As String
Private m_IconEmbedding As Boolean
Private m_Protocol As String
Private m_FrmPath As String          '保存窗体所在目录
Private m_ScaleMode As Long

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim sOut As New cStrBuilder, s As String, sBase64 As String, sa() As String, i As Long
    
    If oop Then
        sOut.Append "        self.master.title(" & U(m_Title) & ")"
        sOut.Append "        self.master.geometry('" & m_Geometry & "')"
        If Not m_Resizable Then sOut.Append "        self.master.resizable(0,0)"
        
        '处理窗体图标
        s = UCase(FileExt(m_IconFile))
        If s = "GIF" And m_IconEmbedding Then
            sBase64 = CreateIconBase64(Space(12)) '使用BASE64将GIF图片嵌入源码
            If Len(sBase64) Then
                sOut.Append "        self.icondata = " & String(3, Chr(34)) & vbCrLf & sBase64 & String(3, Chr(34))
                sOut.Append "        self.iconimg = PhotoImage(data=self.icondata)"
                sOut.Append "        self.master.tk.call('wm', 'iconphoto', self.master._w, self.iconimg)"
            End If
        ElseIf s = "ICO" Then
            sOut.Append "        self.master.iconbitmap(default=r'" & FileFullName(m_IconFile) & "')"
        ElseIf Len(s) Then
            sOut.Append "        self.iconimg = PhotoImage(file=r'" & FileFullName(m_IconFile) & "')"
            sOut.Append "        self.master.tk.call('wm', 'iconphoto', self.master._w, self.iconimg)"
        End If
        
        '窗口消息拦截，如果有的话
        If Len(m_Protocol) Then
            sa = Split(m_Protocol, ",")
            For i = 0 To UBound(sa)
                sOut.Append "        self.master.protocol('" & sa(i) & "', self.EV_" & sa(i) & ")"
                If sa(i) = "WM_DELETE_WINDOW" Then
                    s = "        if askyesno(message='Are you sure to quit?'):" & vbCrLf & Space(12) & "self.master.destroy()"
                    sCmdFunc.Append m_Base.CreateFuncDefOOP("EV_" & sa(i), "event=None", s)
                Else
                    sCmdFunc.Append m_Base.CreateFuncDefOOP("EV_" & sa(i), "event=None")
                End If
            Next
        End If
    Else
        sOut.Append "    " & m_Name & " = Tk()"
        sOut.Append "    " & m_Name & ".title(" & U(m_Title) & ")"
        sOut.Append "    " & m_Name & ".geometry('" & m_Geometry & "')"
        If Not m_Resizable Then sOut.Append "    " & m_Name & ".resizable(0,0)"
        
        '处理窗体图标
        s = UCase(FileExt(m_IconFile))
        If s = "GIF" And m_IconEmbedding Then
            sBase64 = CreateIconBase64(Space(8))
            If Len(sBase64) Then
                sOut.Append "    icondata = " & String(3, Chr(34)) & vbCrLf & sBase64 & String(3, Chr(34))
                sOut.Append "    iconimg = PhotoImage(data=icondata)"
                sOut.Append "    " & m_Name & ".tk.call('wm', 'iconphoto', " & m_Name & "._w, iconimg)"
            End If
        ElseIf s = "ICO" Then
            sOut.Append "    " & m_Name & ".iconbitmap(default=r'" & FileFullName(m_IconFile) & "')"
        ElseIf Len(s) Then
            sOut.Append "    iconimg = PhotoImage(file=r'" & FileFullName(m_IconFile) & "')"
            sOut.Append "    " & m_Name & ".tk.call('wm', 'iconphoto', " & m_Name & "._w, iconimg)"
        End If
        
        '窗口消息拦截，如果有的话
        If Len(m_Protocol) Then
            sa = Split(m_Protocol, ",")
            For i = 0 To UBound(sa)
                sOut.Append "    " & m_Name & ".protocol('" & sa(i) & "', EV_" & sa(i) & ")"
                If sa(i) = "WM_DELETE_WINDOW" Then
                    s = "    if askyesno(message='Are you sure to quit?'):" & vbCrLf & "        gComps['" & m_Name & "'].destroy()"
                    sCmdFunc.Append m_Base.CreateFuncDef("EV_" & sa(i), "event=None", s)
                Else
                    sCmdFunc.Append m_Base.CreateFuncDef("EV_" & sa(i), "event=None")
                End If
            Next
        End If
        
        sOut.Append "    gComps['" & m_Name & "'] = " & m_Name & vbCrLf
    End If
    
    toString = sOut.toString(vbCrLf)
    
End Function

Private Function CreateIconBase64(sPrexSpace As String) As String

    Dim sIconFile As String, abContent() As Byte
    
    CreateIconBase64 = ""
    If Len(m_IconFile) <= 0 Then Exit Function
    
    sIconFile = IIf(InStr(1, m_IconFile, "\") Or InStr(1, m_IconFile, "/"), "", m_FrmPath) & m_IconFile
    
    If Dir(sIconFile) = "" Then
        MsgBox L_F("l_msgFileNotExist", "文件{0}不存在，请重新选择文件。", sIconFile), vbInformation
        Exit Function
    End If
    
    If ReadFileBinaryContent(sIconFile, abContent) = 0 Then
        MsgBox L_F("l_msgReadFileError", "读取文件{0}出错。", sIconFile), vbInformation
        Exit Function
    End If
    
    Base64Encode abContent, CreateIconBase64, sPrexSpace
    
End Function


'创建对象后要马上调用这个函数初始化各参数, 窗体大小以像素为单位
Public Sub InitConfig(o As Object, nWidth As Long, nHeight As Long)
    
    Dim s As String
    
    m_Name = o.Properties("Name")
    m_Title = o.Properties("Caption")
    m_Geometry = nWidth & "x" & nHeight
    m_Resizable = IIf(o.Properties("MaxButton"), True, False)
    m_IconEmbedding = True
    m_IconFile = ""
    m_Protocol = ""
    
    On Error Resume Next
    s = o.FileNames(1)
    If Err.Number = 0 And Len(s) Then
        m_FrmPath = PathName(s)
    Else
        m_FrmPath = ""
    End If
    Err.Clear
    On Error GoTo 0
    
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    If sAttr = "resizable" Or sAttr = "iconembedding" Then
        sa = Split("1,0", ",")
    Else
        GetAttrValueList = 0
    End If
End Function

Public Function Tips(sAttr As String) As String
    Tips = sAttr & vbCrLf
    Select Case sAttr:
        Case "title"
            Tips = Tips & L("l_TipFrmTitle", "窗体的标题。")
        Case "geometry"
            Tips = Tips & L("l_TipFrmGeometry", "窗体的大小（宽x高），单位为像素。")
        Case "resizable"
            Tips = Tips & L("l_TipFrmResizable", "用户是否可以改变窗体的大小。可选值：1/0。")
        Case "icon"
            Tips = Tips & L("l_TipFrmIcon", "图标文件名，支持ico,gif,ppm,pgm格式，省略路径名则在当前VB窗体目录下查找，建议icon文件放在py文件同一目录下。")
        Case "iconembedding"
            Tips = Tips & L("l_TipFrmIconEmbedding", "是否将图标文件转换为base64编码嵌入源代码，仅支持GIF格式嵌入。")
        Case "protocol"
            Tips = Tips & L("l_TipFrmProtocol", "处理窗口消息，比如设置为'WM_DELETE_WINDOW'则处理窗口关闭事件，如果要拦截多个消息，请使用逗号分隔，比如'WM_DELETE_WINDOW,WM_SAVE'")
        Case Else:
            Tips = Tips & L("l_TipUnknown", "未知属性")
    End Select
End Function

'返回一个集合，每个项目三元对"属性名|值|是否默认选择"
'这个函数用于主界面填充属性参数列表框
Public Function Allitems() As Collection
    Dim re As Collection
    
    Set re = New Collection
    re.Add "geometry|" & m_Geometry & "|1"
    re.Add "title|" & m_Title & "|1"
    re.Add "resizable|" & IIf(m_Resizable, 1, 0) & "|1"
    re.Add "icon||1"
    re.Add "iconembedding|" & IIf(m_IconEmbedding, 1, 0) & "|1"
    re.Add "protocol||1"
    
    Set Allitems = re
    
End Function

Public Sub SetConfig(sAttrs As String)
    Dim sa() As String, i As Long
    
    sa = Split(sAttrs, "|")
    Debug.Assert (UBound(sa) Mod 1 = 0)
    
    For i = 0 To UBound(sa) - 1 Step 2
        Select Case sa(i)
            Case "title"
                m_Title = UnQuote(Trim(sa(i + 1)))
            Case "geometry"
                m_Geometry = UnQuote(Trim(sa(i + 1)))
            Case "resizable"
                m_Resizable = IIf(Trim(sa(i + 1)) = "0", False, True)
            Case "icon"
                m_IconFile = UnQuote(Trim(sa(i + 1)))
            Case "iconembedding"
                m_IconEmbedding = IIf(Trim(sa(i + 1)) = "0", False, True)
            Case "protocol"
                m_Protocol = Trim(Replace(Replace(Replace(sa(i + 1), Chr(34), ""), "'", ""), " ", ""))
        End Select
    Next
    
End Sub

Public Sub SetSingleConfig(sAttr As String)
    SetConfig sAttr
End Sub

Public Property Get Parent() As String
    Parent = ""
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(s As String)
    m_Name = s
End Property

'用于改变其默认对应的widget类型，修改widget类型后注意属性列表的合法性
Public Function SetWidgetType(sType As String, sStyleName As String)
    'm_Base.ctlType = sType
    'm_Base.StyleName = sStyleName
End Function

'确定主处理函数能否调用其toString()来产生代码，默认为True，设置为False说明由其他对象来调用处理
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = True
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    'm_CanbeOutByMainForm = bEnable
End Property

Public Property Get Description() As String
    Description = L("l_DescForm", "窗体对象，设置顶层窗口的属性。")
End Property

Public Property Let ScaleMode(nV As Long)
    m_ScaleMode = nV
End Property

Private Sub Class_Initialize()
    Set m_Base = New clsBaseControl
End Sub

Private Sub Class_Terminate()
    Set m_Base = Nothing
End Sub

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

