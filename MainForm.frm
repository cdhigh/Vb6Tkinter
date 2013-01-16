VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Visual Tkinter of Python - cdhigh@sohu.com"
   ClientHeight    =   8134
   ClientLeft      =   56
   ClientTop       =   672
   ClientWidth     =   12964
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   926
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cmbEditCombo 
      CausesValidation=   0   'False
      Height          =   300
      ItemData        =   "MainForm.frx":058A
      Left            =   7560
      List            =   "MainForm.frx":058C
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbEditList 
      Height          =   300
      ItemData        =   "MainForm.frx":058E
      Left            =   6360
      List            =   "MainForm.frx":0590
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VisualTkinter.xpcmdbutton CmdRefsFormsList 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   864
      Caption         =   "刷新窗体列表(&R)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar stabar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   7755
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   660
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbFrms 
      Height          =   300
      ItemData        =   "MainForm.frx":0592
      Left            =   120
      List            =   "MainForm.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TxtTips 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4920
      Width           =   2415
   End
   Begin VB.ListBox LstComps 
      Height          =   3514
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VisualTkinter.GridOcx LstCfg 
      Height          =   6855
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   6015
      _ExtentX        =   10617
      _ExtentY        =   12090
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   840
      Width           =   4095
   End
   Begin VisualTkinter.xpcmdbutton CmdGenCode 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   864
      Caption         =   "生成代码(&G)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VisualTkinter.xpcmdbutton CmdCopyToClipboard 
      Height          =   495
      Left            =   5340
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   864
      Caption         =   "拷贝到剪贴板(&C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VisualTkinter.xpcmdbutton CmdSaveToFile 
      Height          =   495
      Left            =   7950
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   864
      Caption         =   "保存到文件(&F)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VisualTkinter.xpcmdbutton CmdQuit 
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      _ExtentX        =   4039
      _ExtentY        =   864
      Caption         =   "退出(&Q)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuRefreshForms 
         Caption         =   "刷新窗体列表(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenCode 
         Caption         =   "生成代码(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSeparator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveToFile 
         Caption         =   "保存代码到文件(&F)"
         Begin VB.Menu mnuSaveAll 
            Caption         =   "保存全部内容(&A)"
         End
         Begin VB.Menu mnuSaveMainOnly 
            Caption         =   "仅保存main()函数(&M)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSaveUiOnly 
            Caption         =   "仅保存界面生成类(&G)"
         End
      End
      Begin VB.Menu mnuCopyToClipboard 
         Caption         =   "拷贝代码到剪贴板(&C)"
         Begin VB.Menu mnuCopyToClipAll 
            Caption         =   "拷贝全部内容(&A)"
         End
         Begin VB.Menu mnuCopyToClipMainOnly 
            Caption         =   "仅拷贝main()函数(&M)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCopyToClipUiOnly 
            Caption         =   "仅拷贝界面生成类(&G)"
         End
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddProperty 
         Caption         =   "增加一个自定义属性(&P)"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveConfig 
         Caption         =   "保存配置到文件(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRestoreConfig 
         Caption         =   "从文件恢复配置(&L)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "选项(&O)"
      Begin VB.Menu mnuOopCode 
         Caption         =   "生成面向对象代码(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuV2andV3Code 
         Caption         =   "生成Python 2.x/3.x兼容代码(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseTtk 
         Caption         =   "启用TTK主题库(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRelPos 
         Caption         =   "使用相对坐标(&R)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUnicodePrefixU 
         Caption         =   "Unicode字符串增加前缀u(&U)"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPythonExe 
         Caption         =   "设置Python.exe位置(&E)..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuPreview 
         Caption         =   "界面预览(&P)"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "语言(&L)"
      Begin VB.Menu mnuLng 
         Caption         =   "简体中文(&C)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mConnect As Connect

'2012.11.23，为了滚动条和列表框的绑定方便，将其修改为全局变量g_Comps，放在Common.bas中
'Private m_Comps() As Object             '和LstComps行数一样多，对应各组件生成的实例
Private m_MainMenu As clsMenu              '菜单对象
Private m_PrevCompIdx As Long
Private m_curFrm As Object
Private m_prevsf As String
Private m_nLngNum As Long                   ' 语言种类
Private m_HasCommonDialog As Boolean
Private m_saTmpFile() As String
Private m_TxtCodeExpanded As Boolean
Private m_TxtTipsExpanded As Boolean
Private m_BriefCaption As String

'窗体和控件的序列化字符串都用相应的字符串包起来，方便查找和对应
Const REGX_INC_FRM_S = "<<<HFS>>>"
Const REGX_INC_FRM_E = "<<<HFE>>>"
Const REGX_INC_CTL_S = "<<<CTS>>>"
Const REGX_INC_CTL_E = "<<<CTE>>>"
Const SEP_NAME_FROM_CONTENT = "<<<SNFC>>>"

Const REGX_PATTERN_FRM = REGX_INC_FRM_S & "(.*[\s\S\n\r\b]*?)" & REGX_INC_FRM_E
Const REGX_PATTERN_CTL = REGX_INC_CTL_S & "(.*[\s\S\n\r\b]*?)" & REGX_INC_CTL_E
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Sub Form_Load()
    
    Dim s As String
    
    ReDim g_Comps(0) As Object
    g_DefaultFontName = ""
    m_HasCommonDialog = False
    m_TxtCodeExpanded = False
    m_TxtTipsExpanded = False
    
    ReDim m_saTmpFile(0) As String
    
    '多语种支持
    InitMultiLanguage
    
    LstCfg.Redraw = False
    LstCfg.Editable = True
    LstCfg.EditType = EnterKey Or MouseDblClick Or F2Key
    LstCfg.CheckBoxes = True
    LstCfg.AddColumn "Property", 2260, lgAlignCenterCenter
    LstCfg.AddColumn "Value", 3450, lgAlignCenterCenter
    LstCfg.ColAlignment(0) = lgAlignLeftCenter
    LstCfg.ColAlignment(1) = lgAlignLeftCenter
    LstCfg.SelectBackColor = &HFCC597 'vbHighlight
    LstCfg.Redraw = True
    
    m_BriefCaption = "Visual Tkinter of Python - cdhigh@sohu.com - v" & App.Major & "." & App.Minor & IIf(App.Revision > 0, "." & App.Revision, "")
    #If DebugVer Then
        m_BriefCaption = m_BriefCaption & " [Debug Mode] "
    #End If
    Me.Caption = m_BriefCaption
    
    mnuOopCode.Checked = GetSetting(App.Title, "Settings", "OopCode", "1") = "1"
    mnuV2andV3Code.Checked = GetSetting(App.Title, "Settings", "V2andV3Code", "1") = "1"
    mnuUseTtk.Checked = GetSetting(App.Title, "Settings", "UseTtk", "1") = "1"
    mnuRelPos.Checked = GetSetting(App.Title, "Settings", "RelPos", "1") = "1"
    mnuUnicodePrefixU.Checked = GetSetting(App.Title, "Settings", "UnicodePrefix", "0") = "1"
    g_bUnicodePrefixU = mnuUnicodePrefixU.Checked
    
    g_PythonExe = GetSetting(App.Title, "Settings", "PythonExe", "")
    
    mnuSaveMainOnly.Visible = Not mnuOopCode.Checked
    mnuSaveUiOnly.Visible = mnuOopCode.Checked
    mnuCopyToClipMainOnly.Visible = Not mnuOopCode.Checked
    mnuCopyToClipUiOnly.Visible = mnuOopCode.Checked
    
    Set cmbEditList.Font = LstCfg.Font
    Set cmbEditCombo.Font = LstCfg.Font
    
    ResizeInit Me
    
    CmdRefsFormsList_Click
    
End Sub

'多语种支持初始化
Private Sub InitMultiLanguage()
    
    Dim i As Long, s As String, sa() As String
    
    If Not LngFileExist() Then
        m_nLngNum = 0
        mnuLng(0).Checked = True
        Exit Sub
    End If
    
    sa = GetAllLanguageName()
    mnuLng(0).Caption = sa(0)
    m_nLngNum = 1
    For i = 1 To UBound(sa)
        Load mnuLng(i)
        mnuLng(i).Caption = sa(i)
        m_nLngNum = m_nLngNum + 1
    Next
    
    '切换语言，注册表保存的语言优先，其次根据操作系统选择
    s = GetSetting(App.Title, "Settings", "Language", "")
    i = m_nLngNum
    If Len(s) Then                                                              '选择之前保存的语言种类，如果存在的话
        For i = 0 To m_nLngNum - 1
            If mnuLng(i).Caption = s Then
                ChangeLanguage (mnuLng(i).Caption)
                mnuLng(i).Checked = True
                Exit For
            End If
        Next
    End If
    
    '尝试判断操作系统语种
    If i > m_nLngNum - 1 Then
        
        i = GetSystemDefaultLCID()
        If i = &H804 Or i = &H4 Or i = &H1004 Then
            s = "简体中文"
        ElseIf i = &H404 Or i = &HC04 Then
            s = "繁體中文"
        ElseIf i Mod 16 = 9 Then
            s = "English"
        Else                                                                    '其他语言先按英语处理，待软件启动后用户再选择合适的语言
            s = "English"
        End If
        
        For i = 0 To m_nLngNum - 1
            If InStr(1, mnuLng(i).Caption, s) > 0 Then
                ChangeLanguage (mnuLng(i).Caption)
                mnuLng(i).Checked = True
                Exit For
            End If
        Next
        
        ' 无法自动确认语种，默认选择第一个
        If i > m_nLngNum - 1 Then
            ChangeLanguage (mnuLng(0).Caption)
            mnuLng(0).Checked = True
        End If
    End If
    
End Sub

Private Sub CmdQuit_Click()
    mConnect.Hide
End Sub

Private Sub cmbFrms_Click()
    
    Dim frm As Object
    
    '查找到对应的窗体引用
    Set m_curFrm = Nothing
    If Len(cmbFrms.Text) Then
        For Each frm In VBE.ActiveVBProject.VBComponents
            If frm.Type = vbext_ct_VBForm And frm.Name = cmbFrms.Text Then
                Set m_curFrm = frm
                Exit For
            End If
        Next
    End If
    
    m_PrevCompIdx = -1
    
    '将控件添加到列表
    If Not ResetLstComps(m_curFrm) Then
        LstComps.Clear
        LstCfg.Clear
        TxtTips.Text = L("l_TipHasNoControl", "目前窗体上没有任何控件，请先在窗体上添加控件。")
        m_PrevCompIdx = -1
    Else
        LstComps.ListIndex = 0
        LstComps_Click
    End If
    
    If LstComps.ListCount > 0 Then
        CmdGenCode.Enabled = True
        CmdCopyToClipboard.Enabled = True
        CmdSaveToFile.Enabled = True
        mnuSaveToFile.Enabled = True
        mnuCopyToClipboard.Enabled = True
        mnuAddProperty.Enabled = True
        'mnuSaveConfig.Enabled = True
        'mnuRestoreConfig.Enabled = True
        mnuGenCode.Enabled = True
    Else
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub

Private Sub CmdCopyToClipboard_Click()
    Me.PopupMenu mnuCopyToClipboard
End Sub

'更新各个列表，创建对应的控件类实例, 返回false表示初始化失败
Private Function ResetLstComps(frm As Object) As Boolean
    
    Dim Obj As Object, ObjClsModule As Object, i As Long, s As String, j As Long, idx As Long
    Dim nScaleMode As Long, nScaleWidth As Long, nScaleHeight As Long
    Dim CodeMember As Member, CodeMembers As Members, dMethods As New Dictionary
    
    ResetLstComps = False
    If frm Is Nothing Then Exit Function
    
    LstComps.Clear
    'Erase g_Comps
    Set m_MainMenu = Nothing
    
    '创建窗体实例做为列表第一项
    ReDim g_Comps(0) As Object
    Set g_Comps(0) = New clsForm
    
    '因为ScaleX/ScaleY为窗体类独有方法，只能先在这里转换窗体大小为像素单位
    nScaleWidth = Round(ScaleX(frm.Properties("ScaleWidth"), frm.Properties("ScaleMode"), vbPixels))
    nScaleHeight = Round(ScaleY(frm.Properties("ScaleHeight"), frm.Properties("ScaleMode"), vbPixels))
    g_Comps(0).InitConfig frm, nScaleWidth, nScaleHeight, dMethods
    g_Comps(0).Name = WTOP
    LstComps.AddItem g_Comps(0).Name & " (Form)"
    i = 1
    
    m_HasCommonDialog = False
    
    '获取窗体的代码模块中所有的过程函数列表，保存为一个字典对象，传入各类模块，用于自动生成对应的bindcommand
    If Not frm.CodeModule Is Nothing Then
        Set CodeMembers = frm.CodeModule.Members
        If Not CodeMembers Is Nothing Then
            For Each CodeMember In CodeMembers
                If CodeMember.Type = vbext_mt_Method Then
                    idx = InStrRev(CodeMember.Name, "_")
                    If idx > 1 Then
                        s = Left$(CodeMember.Name, idx - 1)
                        If dMethods.Exists(s) Then
                            dMethods.Item(s) = dMethods.Item(s) & "," & CodeMember.Name & "," '使用逗号做为分隔符方便查找
                        Else
                            dMethods.Item(s) = "," & CodeMember.Name & ","
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    
    '将控件添加到列表中
    For Each Obj In frm.Designer.VBControls
        
        CreateObj Obj, ObjClsModule                                             '生成对应类模块实例
        
        If Not ObjClsModule Is Nothing Then
            
            '用于自动单位转换，需要在InitConfig之前设置这个值
            ObjClsModule.ScaleMode = frm.Properties("ScaleMode")
            
            '如果窗体存在菜单控件，则创建主菜单对象，主菜单控件将管理所有的菜单项
            If Obj.ClassName = "Menu" And m_MainMenu Is Nothing Then
                ReDim Preserve g_Comps(i) As Object
                Set m_MainMenu = New clsMenu
                Set g_Comps(i) = m_MainMenu
                LstComps.AddItem m_MainMenu.Name & " (MainMenu)"
                m_MainMenu.InitConfig
                i = i + 1
            End If
            
            '添加控件到控件列表
            ReDim Preserve g_Comps(i) As Object
            Set g_Comps(i) = ObjClsModule
            LstComps.AddItem Obj.Properties("Name") & " (" & Obj.ClassName & ")"
            
            '初始化各控件对应的类模块对象
            If Obj.Container Is frm.Designer Then
                g_Comps(i).InitConfig Obj, frm.Properties("ScaleWidth"), frm.Properties("ScaleHeight"), dMethods
                g_Comps(i).Parent = IIf(Obj.ClassName = "Menu", "MainMenu", WTOP)
            ElseIf Obj.Container.ClassName = "Menu" Then  '子菜单
                g_Comps(i).InitConfig Obj, 0, 0, dMethods
                g_Comps(i).Parent = Obj.Container.Properties("Name")
            Else
                On Error Resume Next
                nScaleMode = Obj.Container.Properties("ScaleMode")
                nScaleWidth = Obj.Container.Properties("ScaleWidth")
                nScaleHeight = Obj.Container.Properties("ScaleHeight")
                If Err.Number Then         'Frame和个别其他容器不支持ScaleWidth属性，则使用Width代替
                    nScaleMode = vbTwips
                    nScaleWidth = Me.ScaleX(Obj.Container.Properties("Width"), frm.Properties("ScaleMode"), vbTwips)
                    nScaleHeight = Me.ScaleY(Obj.Container.Properties("Height"), frm.Properties("ScaleMode"), vbTwips)
                End If
                Err.Clear
                On Error GoTo 0
                g_Comps(i).ScaleMode = nScaleMode
                g_Comps(i).InitConfig Obj, nScaleWidth, nScaleHeight, dMethods
                g_Comps(i).Parent = Obj.Container.Properties("Name")
            End If
            
            i = i + 1
            ResetLstComps = True
        ElseIf Obj.ClassName = "CommonDialog" Then
            m_HasCommonDialog = True
        Else
            MsgBox L_F("l_msgCtlNotSupport", "当前暂不支持'{0}'控件(控件名：{1})\n\n程序将不生成此控件的代码。", Obj.ClassName, Obj.Properties("Name")), vbInformation, App.Title
        End If
    Next  'frm.Designer.VBControls
    
    '生成菜单的树形层次关系，为生成代码建立基础
    CreateMenuHiberarchy
    
    '整理Notebook和其各页签内控件的父子关系
    ArrangeNotebookAndSubWidgets
    
    '尝试自动将滚动条绑定到对应的控件
    TryAssignScrollbar2Widgets
    
    '正确设置ComboboxAdapter的TTK属性
    For i = 1 To UBound(g_Comps)
        If TypeName(g_Comps(i)) = "clsComboboxAdapter" Then g_Comps(i).TTK = mnuUseTtk.Checked
    Next
    
End Function

'生成一个控件字符实例对象:输入ctlobj:控件对象，clsobj:对应的字符串对象
Private Function CreateObj(ByRef ctlobj As Object, ByRef clsobj As Object) As Object
    Dim o As Object, sName As String, idx As Long
    
    Select Case ctlobj.ClassName:
        Case "Label"
            Set clsobj = New clsLabel
        Case "CommandButton"
            Set clsobj = New clsButton
        Case "TextBox"
            If ctlobj.Properties("MultiLine") Then Set clsobj = New clsText Else Set clsobj = New clsEntry
        Case "CheckBox"
            Set clsobj = New clsCheckbutton
        Case "OptionButton"
            Set clsobj = New clsRadiobutton
        Case "ComboBox"
            Set clsobj = New clsComboboxAdapter
            clsobj.TTK = mnuUseTtk.Checked
        Case "ListBox"
            Set clsobj = New clsListbox
        Case "HScrollBar", "VScrollBar"
            Set clsobj = New clsScrollbar
        Case "Slider"
            Set clsobj = New clsScale
        Case "Frame"
            '判断是否是TabStrip控件的一页
            idx = InStr(2, ctlobj.Properties("Name"), "__Tab") '从2开始至少保证__Tab前有一个字符
            If idx > 1 And Not m_curFrm Is Nothing Then
                '循环查询是否有合适的TabStrip控件
                Set clsobj = Nothing
                sName = Left$(ctlobj.Properties("Name"), idx - 1)
                For Each o In m_curFrm.Designer.VBControls
                    If o.ClassName = "TabStrip" And o.Properties("Name") = sName Then
                        Set clsobj = New clsNotebookTab  '使用Tab类来代替Frame
                        Exit For
                    End If
                Next
                If clsobj Is Nothing Then Set clsobj = New clsLabelFrame '没有对应的TabStrip对象
            Else
                Set clsobj = New clsLabelFrame
            End If
        Case "PictureBox"
            '判断是否是TabStrip控件的一页
            idx = InStr(2, ctlobj.Properties("Name"), "__Tab") '从2开始至少保证__Tab前有一个字符
            If idx > 1 And Not m_curFrm Is Nothing Then
                '循环查询是否有合适的TabStrip控件
                Set clsobj = Nothing
                sName = Left$(ctlobj.Properties("Name"), idx - 1)
                For Each o In m_curFrm.Designer.VBControls
                    If o.ClassName = "TabStrip" And o.Properties("Name") = sName Then
                        Set clsobj = New clsNotebookTab  '使用Tab类来代替PictureBox
                        Exit For
                    End If
                Next
                If clsobj Is Nothing Then Set clsobj = New clsCanvas '没有对应的TabStrip对象
            Else
                Set clsobj = New clsCanvas
            End If
        Case "Menu"
            Set clsobj = New clsMenuItem
        Case "ProgressBar"
            Set clsobj = New clsProgressBar                                         '需要启用TTK才支持
            mnuUseTtk.Checked = True
        Case "TreeView"
            Set clsobj = New clsTreeview                                            '需要启用TTK才支持
            mnuUseTtk.Checked = True
        Case "TabStrip"
            Set clsobj = New clsNotebook                                            '需要启用TTK才支持
            mnuUseTtk.Checked = True
        Case "Line"
            Set clsobj = New clsSeparator
            mnuUseTtk.Checked = True
        Case "StatusBar"
            Set clsobj = New clsStatusbar
        Case Else:
            Set clsobj = Nothing
    End Select
    
    Set CreateObj = clsobj
    
End Function

'生成菜单的树形层次关系，为生成代码建立基础
'子类储存父类的名字，父类储存所有子类的引用
Private Sub CreateMenuHiberarchy()

    Dim i As Long, j As Long
    If Not m_MainMenu Is Nothing Then
        For i = 0 To UBound(g_Comps)
            If TypeName(g_Comps(i)) = "clsMenu" Then
                '将所有的顶层菜单做为clsMenu的子控件
                For j = 0 To UBound(g_Comps)
                    If TypeName(g_Comps(j)) = "clsMenuItem" And g_Comps(j).Parent = "MainMenu" Then
                        g_Comps(i).AddChild g_Comps(j)
                    End If
                Next
            ElseIf TypeName(g_Comps(i)) = "clsMenuItem" Then
                '子菜单有可能还有子菜单
                For j = 0 To UBound(g_Comps)
                    If TypeName(g_Comps(j)) = "clsMenuItem" And g_Comps(j).Parent = g_Comps(i).Name Then
                        g_Comps(i).AddChild g_Comps(j)
                    End If
                Next
            End If
        Next
    End If
    
End Sub

'整理选项卡控件和其内部控件的父子关系
Private Sub ArrangeNotebookAndSubWidgets()

    Dim i As Long, j As Long, k As Long, L As Long, idx As Long, ctlNum As Long
    Dim sTabName As String, sNbName As String, sTmp As String
    
    If UBound(g_Comps) <= 0 Then  ' 0固定为顶层窗体
        Exit Sub
    End If
    
    ctlNum = UBound(g_Comps)
    For i = 1 To ctlNum
        If TypeName(g_Comps(i)) = "clsNotebookTab" Then
            sTabName = g_Comps(i).Name
            idx = InStr(2, sTabName, "__Tab")
            If idx > 1 Then
                sNbName = Left$(sTabName, idx - 1) ' Notebook控件名
                For j = 1 To ctlNum
                    If TypeName(g_Comps(j)) = "clsNotebook" And g_Comps(j).Name = sNbName Then
                        '获取TAB号
                        sTmp = Right$(sTabName, 1)
                        If sTmp >= "1" And sTmp <= "9" Then '最多支持9个标签页
                            g_Comps(j).AddTab g_Comps(i), CLng(sTmp)  ' 加入Notebook对象
                            g_Comps(i).EnableOutByMainForm = False
                            
                            '此标签页内所有控件均有clsNotebookTab来接管，不再由主窗口输出代码
                            For k = 1 To ctlNum
                                If g_Comps(k).Parent = sTabName Then
                                    g_Comps(k).EnableOutByMainForm = False
                                    g_Comps(i).AddSubWidget g_Comps(k)
                                    
                                    ' 万一标签页内还有其他容器控件
                                    If TypeName(g_Comps(k)) = "clsCanvas" Or TypeName(g_Comps(k)) = "clsLabelFrame" Then
                                        For L = 1 To ctlNum
                                            If g_Comps(L).Parent = g_Comps(k).Name Then
                                                g_Comps(L).EnableOutByMainForm = False
                                                g_Comps(i).AddSubWidget g_Comps(L)
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                            
                        End If
                    End If
                Next
            End If
        End If
    Next

End Sub

'进行一些分析，尝试将滚动条自动绑定到合适的控件，不一定成功，而且可能误判，只算是尽力而为
Private Sub TryAssignScrollbar2Widgets()
    
    Dim i As Long, ctlNum As Long, Obj As Object, o As Object, oName As String
    Dim vX1 As Long, vY1 As Long, vX2 As Long, vY2 As Long
    Dim oX1 As Long, oY1 As Long, oX2 As Long, oY2 As Long
    Dim thresholdX1 As Long, thresholdY1 As Long
    Dim thresholdX2 As Long, thresholdY2 As Long
    Dim isWidgetScrl As Boolean, Assigned As Boolean
    
    If UBound(g_Comps) <= 0 Or m_curFrm Is Nothing Then  ' 0固定为顶层窗体
        Exit Sub
    End If
    
    '水平方向和垂直方向都使用20个像素做为查找控件的门限
    thresholdX1 = Round(ScaleX(20, vbPixels, m_curFrm.Properties("ScaleMode")))
    thresholdY1 = Round(ScaleY(20, vbPixels, m_curFrm.Properties("ScaleMode")))
    thresholdX2 = Round(ScaleX(5, vbPixels, m_curFrm.Properties("ScaleMode"))) '5个像素是允许控件和滚动条重叠的部分
    thresholdY2 = Round(ScaleY(5, vbPixels, m_curFrm.Properties("ScaleMode")))
    
    ctlNum = UBound(g_Comps)
    For Each Obj In m_curFrm.Designer.VBControls
        If Obj.ClassName = "HScrollBar" Then
            '水平滚动条，则判断其上方有没有需要设置滚动条的控件
            'vX1,vY1,vX2,vY2构成一个矩形，如果其他控件的左下角和右下角落在这个矩形内，则认为滚动条对应
            vX1 = Obj.Properties("Left") - thresholdX1
            If vX1 < 0 Then vX1 = 0
            vY1 = Obj.Properties("Top") - thresholdY1
            If vY1 < 0 Then vY1 = 0
            vX2 = Obj.Properties("Left") + Obj.Properties("Width") + thresholdX1
            vY2 = Obj.Properties("Top") + thresholdY2
            Assigned = False
            For Each o In m_curFrm.Designer.VBControls
                If InStr(1, "PictureBox,ListBox,TreeView,TextBox,", o.ClassName & ",") > 0 Then '只有这些控件可能需要滚动条
                    isWidgetScrl = True
                    If o.ClassName = "TextBox" Then '只有多行文本框才支持滚动
                        If Not o.Properties("MultiLine") Then
                            isWidgetScrl = False
                        End If
                    End If
                    If isWidgetScrl Then
                        oX1 = o.Properties("Left")
                        oY1 = o.Properties("Top") + o.Properties("Height")
                        oX2 = oX1 + o.Properties("Width")
                        oY2 = oY1
                        
                        '第一行为左下角判断，第二行为右下角判断
                        If (oX1 >= vX1 And oX1 <= vX2 And oY1 >= vY1 And oY1 <= vY2) _
                            And (oX2 >= vX1 And oX2 <= vX2 And oY2 >= vY1 And oY2 <= vY2) Then
                            '设置控件的xscrollcommand属性
                            oName = o.Properties("Name")
                            For i = 1 To ctlNum
                                If g_Comps(i).Name = oName Then
                                    g_Comps(i).SetSingleConfig ("xscrollcommand|" & Obj.Properties("Name") & ".set")
                                    Assigned = True
                                    Debug.Print oName & " assigned to " & Obj.Properties("name")
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If Assigned Then
                    Exit For
                End If
            Next
        ElseIf Obj.ClassName = "VScrollBar" Then
            '垂直滚动条，则判断其左方有没有需要设置滚动条的控件
            'vX1,vY1,vX2,vY2构成一个矩形，如果其他控件的右上角和右下角落在这个矩形内，则认为滚动条对应
            vX1 = Obj.Properties("Left") - thresholdX1
            If vX1 < 0 Then vX1 = 0
            vY1 = Obj.Properties("Top") - thresholdY1
            If vY1 < 0 Then vY1 = 0
            vX2 = Obj.Properties("Left") + thresholdX2
            vY2 = Obj.Properties("Top") + Obj.Properties("Height") + thresholdY1
            Assigned = False
            For Each o In m_curFrm.Designer.VBControls
                If InStr(1, "PictureBox,ListBox,TreeView,TextBox,", o.ClassName & ",") > 0 Then '只有这些控件可能需要滚动条
                    isWidgetScrl = True
                    If o.ClassName = "TextBox" Then '只有多行文本框才支持滚动
                        If Not o.Properties("MultiLine") Then
                            isWidgetScrl = False
                        End If
                    End If
                    If isWidgetScrl Then
                        oX1 = o.Properties("Left") + o.Properties("Width")
                        oY1 = o.Properties("Top")
                        oX2 = oX1
                        oY2 = oY1 + o.Properties("Height")
                        
                        '第一行为右上角判断，第二行为右下角判断
                        If (oX1 >= vX1 And oX1 <= vX2 And oY1 >= vY1 And oY1 <= vY2) _
                            And (oX2 >= vX1 And oX2 <= vX2 And oY2 >= vY1 And oY2 <= vY2) Then
                            '设置控件的yscrollcommand属性
                            oName = o.Properties("Name")
                            For i = 1 To ctlNum
                                If g_Comps(i).Name = oName Then
                                    g_Comps(i).SetSingleConfig ("yscrollcommand|" & Obj.Properties("Name") & ".set")
                                    Debug.Print oName & " assigned to " & Obj.Properties("name")
                                    Assigned = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                If Assigned Then
                    Exit For
                End If
            Next
        End If
    Next
    
End Sub

Private Sub CmdGenCode_Click()
    
    Dim i As Long, cnt As Long, o As Object
    Dim strHead As New cStrBuilder, strOut As New cStrBuilder, strCmd As New cStrBuilder, s As String
    Dim OutOnlyV3 As Boolean, OutOOP As Boolean, OutRelPos As Boolean, usettk As Boolean
    Dim bUnicodePrefix As Boolean  '临时保存UNICODE前缀方式
    Dim aCompsSorted() As Object '用于排序的代码输出
    
    If LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Or m_curFrm Is Nothing Then
        Exit Sub
    End If
    
    On Error Resume Next
    s = m_curFrm.Name
    If Err.Number Then
        If MsgBox(L("l_msgGetAttrOfFrmFailed", "获取窗体属性失败，对应VB工程已经关闭？\n请重新刷新窗体列表或重新打开工程再试。" & _
            "\n现在重新刷新窗体列表吗？"), vbInformation + vbYesNo) = vbYes Then
            CmdRefsFormsList_Click
        End If
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    OutOnlyV3 = Not mnuV2andV3Code.Checked
    OutOOP = mnuOopCode.Checked
    OutRelPos = mnuRelPos.Checked
    usettk = mnuUseTtk.Checked
    
    '绝对坐标
'    If Not OutRelPos And m_curFrm.Properties("ScaleMode") <> vbTwips Then
'        '如果使用绝对坐标，则Frame控件仅支持vbTwips模式
'        For Each o In m_curFrm.Designer.VBControls
'            If o.ClassName = "Frame" Then
'                MsgBox L("l_msgFrameNotSupportInAbs", "绝对坐标布局不支持Frame控件，请改用相对坐标或去掉Frame控件。"), vbInformation
'                Exit Sub
'            End If
'        Next
'    End If
'
    bUnicodePrefix = g_bUnicodePrefixU '先暂存，在函数最后恢复
    If OutOnlyV3 Then
        g_bUnicodePrefixU = False  'V3模式下不需要任何前缀
    End If
    
    '在输出代码前先更新一下当前显示的数据
    UpdateCfgtoCls LstComps.ListIndex
    
    strHead.Append "#!/usr/bin/env python"
    strHead.Append "#-*- coding:utf-8 -*-" & vbCrLf
    strHead.Append "import os, sys"
    
    If OutOnlyV3 Then                                                           '输出仅针对PYTHON 3.X的代码
        strHead.Append "from tkinter import *"
        strHead.Append "from tkinter.font import Font"
        If usettk Then strHead.Append "from tkinter.ttk import *"
        strHead.Append "#Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel"
        strHead.Append "from tkinter.messagebox import *"
        strHead.Append "#Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')"
        strHead.Append IIf(m_HasCommonDialog, "", "#") & "import tkinter.filedialog as tkFileDialog"
        strHead.Append IIf(m_HasCommonDialog, "", "#") & "import tkinter.simpledialog as tkSimpleDialog  #askstring()"
        If m_HasCommonDialog Then
            strHead.Append "import tkinter.colorchooser as tkColorChooser  #askcolor()"
        End If
        strHead.Append vbCrLf
    Else
        strHead.Append "try:"
        strHead.Append "    from tkinter import *"
        strHead.Append "except ImportError:  #Python 2.x"
        strHead.Append "    PythonVersion = 2"
        strHead.Append "    from Tkinter import *"
        strHead.Append "    from tkFont import Font"
        If usettk Then
            strHead.Append "    from ttk import *"
        End If
        strHead.Append "    #Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel"
        strHead.Append "    from tkMessageBox import *"
        strHead.Append "    #Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkSimpleDialog"
        If m_HasCommonDialog Then
            strHead.Append "    import tkColorChooser  #askcolor()"
        End If
        strHead.Append "else:  #Python 3.x"
        strHead.Append "    PythonVersion = 3"
        strHead.Append "    from tkinter.font import Font"
        If usettk Then
            strHead.Append "    from tkinter.ttk import *"
        End If
        strHead.Append "    from tkinter.messagebox import *"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.filedialog as tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.simpledialog as tkSimpleDialog    #askstring()"
        If m_HasCommonDialog Then
            strHead.Append "    import tkinter.colorchooser as tkColorChooser  #askcolor()"
        End If
        strHead.Append vbCrLf
    End If
    
    '如果存在状态栏控件，则先输出状态栏控件的类定义
    For i = 1 To UBound(g_Comps)  '0固定为窗体，不用判断
        If TypeName(g_Comps(i)) = "clsStatusbar" Then
            strHead.Append g_Comps(i).WidgetCode()
            Exit For
        End If
    Next
    
    If OutOOP Then
        strCmd.Append vbCrLf
        strCmd.Append "class Application(Application_ui):"
        strCmd.Append "    " & L("l_cmtClsApp", "#这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。")
        strCmd.Append "    def __init__(self, master=None):"
        strCmd.Append "        Application_ui.__init__(self, master)" & vbCrLf
        
        strOut.Append "class Application_ui(Frame):"
        strOut.Append "    " & L("l_cmtClsUi", "#这个类仅实现界面生成功能，具体事件处理代码在子类Application中。")
        strOut.Append "    def __init__(self, master=None):"
        strOut.Append "        Frame.__init__(self, master)"
        g_Comps(0).toString strOut, strCmd, OutRelPos, OutOOP, usettk  'g_Comps(0)固定是Form
        strOut.Append "        self.createWidgets()" & vbCrLf
        strOut.Append "    def createWidgets(self):"
        strOut.Append "        self." & WTOP & " = self.winfo_toplevel()" & vbCrLf
        If usettk Then strOut.Append "        self.style = Style()" & vbCrLf
    Else
        strHead.Append L("l_cmtgComps", "#所有控件和控件绑定变量引用字典，使用这个字典是为了方便在其他函数中引用所有控件。")
        strHead.Append "gComps = {}"
        strHead.Append vbCrLf & vbCrLf
        
        strOut.Append vbCrLf
        strOut.Append "def main(argv):"
        g_Comps(0).toString strOut, strCmd, OutRelPos, OutOOP, usettk  'g_Comps(0)固定是Form
        If usettk Then
            strOut.Append "    style = Style()"
            strOut.Append "    gComps['style'] = style" & vbCrLf
        End If
    End If
    
    '根据依赖关系排序控件的先后顺序
    cnt = 0
    For i = 1 To UBound(g_Comps)
        If g_Comps(i).EnableOutByMainForm Then
            ReDim Preserve aCompsSorted(cnt) As Object
            Set aCompsSorted(cnt) = g_Comps(i)
            cnt = cnt + 1
        End If
    Next
    If cnt > 0 Then
        SortWidgets aCompsSorted, cnt
    End If
    
    '遍历各控件，由各控件自己输出自己的界面生成代码
    For i = 0 To cnt - 1
        aCompsSorted(i).toString strOut, strCmd, OutRelPos, OutOOP, usettk
        strOut.Append ""  '两个控件之间使用一个空行隔开
    Next
    
    '输出到文本框
    If OutOOP Then
        strCmd.Append "if __name__ == ""__main__"":"
        strCmd.Append "    " & WTOP & " = Tk()"
        strCmd.Append "    Application(" & WTOP & ").mainloop()"
        strCmd.Append vbCrLf
        'strCmd.Append "    try: " & WTOP & ".destroy()"
        'strCmd.Append "    except: pass" & vbCrLf
        TxtCode.Text = strHead.toString(vbCrLf) & strOut.toString(vbCrLf) & strCmd.toString(vbCrLf)
    Else
        strOut.Append "    " & WTOP & ".mainloop()"
        'strOut.Append "    try: " & WTOP & ".destroy()"
        'strOut.Append "    except: pass"
        strOut.Append vbCrLf & vbCrLf
        strOut.Append "if __name__ == ""__main__"":"
        strOut.Append "    main(sys.argv)" & vbCrLf
        TxtCode.Text = strHead.toString(vbCrLf) & strCmd.toString(vbCrLf) & strOut.toString(vbCrLf)
    End If
    
    strOut.Reset
    strHead.Reset
    strCmd.Reset
    
    g_bUnicodePrefixU = bUnicodePrefix    '恢复UNICODE前缀模式
    
End Sub

Private Sub CmdRefsFormsList_Click()
    
    Dim frm As Object, nScaleMode As Long
    
    cmbFrms.Clear
    LstComps.Clear
    LstCfg.Clear
    
    If VBE.ActiveVBProject Is Nothing Then
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
        Exit Sub
    End If
    
    Me.Caption = m_BriefCaption & " [" & VBE.ActiveVBProject.Name & "]"
    
    '查找工程中所有的窗体,全部添加到组合框供选择输出
    For Each frm In VBE.ActiveVBProject.VBComponents
        If frm.Type = vbext_ct_VBForm Then
            nScaleMode = frm.Properties("ScaleMode")
            If nScaleMode <> vbTwips And nScaleMode <> vbPoints And nScaleMode <> vbPixels Then
                MsgBox L_F("l_msgFailedScaleMode", "查找到窗体'{0}'，但是ScaleMode={1}，程序仅支持模式1/2/3。", _
                         frm.Properties("Name"), nScaleMode), vbInformation
            Else
                cmbFrms.AddItem frm.Name
            End If
        End If
    Next
    
    If cmbFrms.ListCount >= 1 Then
        cmbFrms.ListIndex = 0      '触发cmbFrms_Click
    Else
        CmdGenCode.Enabled = False
        CmdCopyToClipboard.Enabled = False
        CmdSaveToFile.Enabled = False
        mnuSaveToFile.Enabled = False
        mnuCopyToClipboard.Enabled = False
        mnuAddProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub

Private Sub CmdSaveToFile_Click()
    Me.PopupMenu mnuSaveToFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Long
    If TxtCode.Width = Me.ScaleWidth Then
        TxtCode_DblClick
        Cancel = True
    ElseIf TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
        Cancel = True
    ElseIf UBound(m_saTmpFile) > 0 Then '删除用于预览的临时文件
        On Error Resume Next
        For i = 0 To UBound(m_saTmpFile)
            If Len(m_saTmpFile(i)) Then
                Kill m_saTmpFile(i)
            End If
        Next
        On Error GoTo 0
        ReDim m_saTmpFile(0) As String
    End If
End Sub

Private Sub Form_Resize()
    
    Dim sLstCfgWidth As Single
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    If m_TxtCodeExpanded Then
        TxtCode_DblClick
    ElseIf m_TxtTipsExpanded Then
        TxtTips_DblClick
    End If
    
    If Me.Width < 9000 Then Me.Width = 9000
    If Me.Height < 6750 Then Me.Height = 6750
    
    If LstCfg.ItemCount > 0 Then LstCfg.TopRow = 0  '规避此控件的一个BUG，这个BUG可能会导致部分项目显示不全
    
    ResizeForm Me
    
    '调整配置列表框的列宽度
    sLstCfgWidth = GetOrignalWidth(LstCfg)
    If LstCfg.Width < sLstCfgWidth Then
        LstCfg.ColWidth(0) = 2260
        LstCfg.ColWidth(1) = 3450
    ElseIf sLstCfgWidth > 1 Then '确认宽度值有效
        LstCfg.ColWidth(0) = 2260 * LstCfg.Width / sLstCfgWidth
        LstCfg.ColWidth(1) = 3450 * LstCfg.Width / sLstCfgWidth
    End If
    
End Sub

Private Sub LstCfg_ItemChecked(Row As Long)
    If InStr(1, " x, y, relx, rely, width, height, relwidth, relheight,", " " & LstCfg.CellText(Row, 0) & ",") Then
        LstCfg.ItemChecked(Row) = True
    ElseIf LstComps.ListCount > 0 And LstComps.ListIndex >= 0 Then '额外必选的属性
        Select Case Mid$(LstComps.List(LstComps.ListIndex), InStrRev(LstComps.List(LstComps.ListIndex), " ") + 1)
            Case "(TextBox)"
                If LstCfg.CellText(Row, 0) = "textvariable" Then LstCfg.ItemChecked(Row) = True
            Case "(Line)"
                If LstCfg.CellText(Row, 0) = "orient" Then LstCfg.ItemChecked(Row) = True
        End Select
    End If
    
    '更新列表中的数值到实例对象和数组
    UpdateCfgtoCls m_PrevCompIdx
    
End Sub

Private Sub LstCfg_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = &H41 Or KeyCode = &H61) And Shift = vbCtrlMask Then 'Ctrl+a:全选，主要用于测试目的
        Dim i As Long
        LstCfg.Redraw = False
        For i = 0 To LstCfg.ItemCount - 1
            LstCfg.ItemChecked(i) = True
        Next
        LstCfg.Redraw = True
    End If
End Sub

Private Sub LstCfg_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Or InStr(1, " x, y, relx, rely, width, height, relwidth, relheight, ", " " & LstCfg.CellText(Row, 0) & ",") Then
        Cancel = True
    Else
        '自动添加所需的属性选择项，存在可选列表时，在下拉列表中选择，否则切换成文本框
        FillcmbEdit Row, Col
    End If
    
End Sub

Private Sub LstCfg_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    If NewValue <> "" Then
        LstCfg.ItemChecked(Row) = True
    End If
End Sub

Private Sub LstCfg_RowColChanged()
    If LstComps.ListIndex >= 0 Then
        TxtTips.Text = g_Comps(LstComps.ListIndex).Tips(LstCfg.CellText(LstCfg.Row, 0))
    End If
End Sub

Private Sub LstComps_Click()
    
    Dim ctl As Object, s As String
    
    If LstComps.ListCount = 0 Or LstComps.ListIndex < 0 Then Exit Sub
    
    On Error Resume Next
    s = m_curFrm.Name
    If Err.Number Then
        If MsgBox(L("l_msgGetAttrOfFrmFailed", "获取窗体属性失败，对应VB工程已经关闭？\n请重新刷新窗体列表或重新打开工程再试。" & _
            "\n现在重新刷新窗体列表吗？"), vbInformation + vbYesNo) = vbYes Then
            CmdRefsFormsList_Click
        End If
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    '更新列表中的数值到实例对象
    UpdateCfgtoCls m_PrevCompIdx
    
    FetchCfgFromCls LstComps.ListIndex
    
    m_PrevCompIdx = LstComps.ListIndex
    
    '显示控件描述
    TxtTips.Text = g_Comps(LstComps.ListIndex).Description
    
    '选择对应的控件
    m_curFrm.Designer.SelectedVBControls.Clear
    For Each ctl In m_curFrm.Designer.VBControls
        If ctl.Properties("Name") = Left(LstComps.List(LstComps.ListIndex), InStr(1, LstComps.List(LstComps.ListIndex), " ") - 1) Then
            ctl.InSelection = True
            Exit For
        End If
    Next
    
End Sub

'在对象中获取配置信息到列表框
Private Sub FetchCfgFromCls(idx As Long)
    
    Dim nRow As Long, cfg As Variant, cItms As Collection
    
    If idx < 0 Or idx > UBound(g_Comps) Then Exit Sub
    
    LstCfg.Redraw = False
    If LstCfg.ItemCount > 0 Then LstCfg.TopRow = 0  '这个操作是为了规避GridOcx的滚动条BUG，就是切换控件后有部分项目无法完整显示
    LstCfg.Clear
    Set cItms = g_Comps(idx).Allitems()
    For Each cfg In cItms
        nRow = LstCfg.AddItem(Left(cfg, InStr(1, cfg, "|") - 1))
        LstCfg.CellText(nRow, 1) = Mid(cfg, InStr(1, cfg, "|") + 1, InStrRev(cfg, "|") - InStr(1, cfg, "|") - 1)
        LstCfg.ItemChecked(nRow) = CLng(Mid(cfg, InStrRev(cfg, "|") + 1))
    Next
    LstCfg.Redraw = True
    
End Sub

'更新配置到实例对象,idx表示当前在LstCfg上显示的属性是属于哪个控件的。
Private Sub UpdateCfgtoCls(idx As Long)
    Dim s As String, i As Long
    
    If idx < 0 Or idx > UBound(g_Comps) Then Exit Sub
    
    LstCfg.Refresh
    
    s = ""
    For i = 0 To LstCfg.ItemCount - 1
        If LstCfg.ItemChecked(i) Then
            s = s & IIf(i > 0, "|", "") & LstCfg.CellText(i, 0) & "|" & LstCfg.CellText(i, 1)
        End If
    Next
    
    If Len(s) Then g_Comps(idx).SetConfig s
    
End Sub

'增加自定义配置
Private Sub mnuAddProperty_Click()
    
    Dim s As String, sa() As String, nRow As Long, i As Long
    
    If LstCfg.ItemCount <= 0 Then Exit Sub
    
    s = InputBox(L("l_ProForAddAttr", "请输入属性和数值对，使用'属性=值'的形式，比如：x=20 。\n注意Python是大小写敏感的。"), App.Title)
    s = Trim(s)
    If Len(s) <= 0 Then
        Exit Sub
    End If
    
    sa = Split(s, "=")
    If UBound(sa) < 1 Then Exit Sub
    
    ' 如果输入的属性已经存在，则覆盖原有的值
    sa(0) = Trim(sa(0))
    For i = 0 To LstCfg.ItemCount - 1
        If LstCfg.CellText(i, 0) = sa(0) Then
            LstCfg.CellText(i, 1) = Trim(sa(1))
            Exit For
        End If
    Next
    '新增一个属性
    If i >= LstCfg.ItemCount Then
        i = LstCfg.AddItem(Trim(sa(0)))
        LstCfg.CellText(i, 1) = Trim(sa(1))
    End If
    
    LstCfg.ItemChecked(i) = True
    UpdateCfgtoCls m_PrevCompIdx
    
End Sub

Private Sub mnuCopyToClipAll_Click()
    Clipboard.Clear
    Clipboard.SetText TxtCode.Text
End Sub

Private Sub mnuCopyToClipMainOnly_Click()
    
    Dim s As String, nm As Long, nf As Long
    
    '分析代码，仅提取main(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nm = InStr(1, s, "def main(argv):")
    nf = InStr(1, s, "if __name__")
    If nm > 0 And nf > 0 Then
        Clipboard.Clear
        Clipboard.SetText Mid(s, nm, nf - nm)
    Else
        MsgBox L("l_msgNoMain", "代码中找不到main()函数！"), vbInformation
    End If
    
End Sub

Private Sub mnuCopyToClipUiOnly_Click()
    
    Dim s As String, nui As Long, napp As Long
    
    '分析代码，仅提取Application_ui(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nui = InStr(1, s, "class Application_ui(Frame):")
    napp = InStr(1, s, "class Application(Application_ui):")
    If nui > 0 And napp > 0 Then
        Clipboard.Clear
        Clipboard.SetText Mid(s, nui, napp - nui)
    Else
        MsgBox L("l_msgNoClsUi", "代码中找不到Application_ui类！"), vbInformation
    End If
    
End Sub

Private Sub mnuFile_Click()
    'mnuSaveConfig.Enabled = LstComps.ListCount > 0
    'mnuRestoreConfig.Enabled = LstComps.ListCount > 0
    mnuGenCode.Enabled = LstComps.ListCount > 0
End Sub

Private Sub mnuGenCode_Click()
    CmdGenCode_Click
End Sub

Private Sub mnuLng_Click(Index As Integer)
    
    Dim i As Long
    
    If m_nLngNum = 0 Then Exit Sub
    
    For i = 0 To m_nLngNum - 1
        mnuLng(i).Checked = False
    Next
    
    mnuLng(Index).Checked = True
    SaveSetting App.Title, "Settings", "Language", mnuLng(Index).Caption
    
    ChangeLanguage (mnuLng(Index).Caption)
    
End Sub

Private Sub mnuOopCode_Click()
    mnuOopCode.Checked = Not mnuOopCode.Checked
    SaveSetting App.Title, "Settings", "OopCode", IIf(mnuOopCode.Checked, "1", "0")
    mnuSaveMainOnly.Visible = Not mnuOopCode.Checked
    mnuSaveUiOnly.Visible = mnuOopCode.Checked
    mnuCopyToClipMainOnly.Visible = Not mnuOopCode.Checked
    mnuCopyToClipUiOnly.Visible = mnuOopCode.Checked
End Sub

Private Sub mnuSaveAll_Click()
    
    Dim sF As String
    sF = FileDialog(Me, True, L("l_fdSave", "将Python文件保存到："), "*.py", m_prevsf)
    
    If Len(sF) Then
        If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  '如果文件名没有扩展名，自动添加.py扩展名
        Utf8File_Write_VB sF, TxtCode.Text
    End If
    
    m_prevsf = sF
    
End Sub

'仅输出main()函数，用于之前已经建好框架，并且也写了一些代码，现在修改空间布局，不用影响其他代码
Private Sub mnuSaveMainOnly_Click()
    
    Dim sF As String, s As String, nm As Long, nf As Long
    
    '分析代码，仅提取main(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nm = InStr(1, s, "def main(argv):")
    nf = InStr(1, s, "if __name__")
    If nm > 0 And nf > 0 Then
        sF = FileDialog(Me, True, L("l_fdSave", "将Python文件保存到："), "*.py", m_prevsf)
        If Len(sF) Then
            If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  '如果文件名没有扩展名，自动添加.py扩展名
            Utf8File_Write_VB sF, Mid(s, nm, nf - nm)
        End If
    Else
        MsgBox L("l_msgNoMain", "代码中找不到main()函数！"), vbInformation
    End If
    
    m_prevsf = sF
    
End Sub

'仅输出界面生成类，用于之前已经建好框架，并且也写了一些代码，现在修改空间布局，不用影响其他代码
Private Sub mnuSaveUiOnly_Click()
    
    Dim sF As String, s As String, nui As Long, napp As Long
    
    '分析代码，仅提取main(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nui = InStr(1, s, "class Application_ui(Frame):")
    napp = InStr(1, s, "class Application(Application_ui):")
    If nui > 0 And napp > 0 Then
        sF = FileDialog(Me, True, "将Python文件保存到：", "*.py", m_prevsf)
        If Len(sF) Then
            If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  '如果文件名没有扩展名，自动添加.py扩展名
            Utf8File_Write_VB sF, Mid(s, nui, napp - nui)
        End If
    Else
        MsgBox L("l_msgNoClsUi", "代码中找不到Application_ui类！"), vbInformation
    End If
    
    m_prevsf = sF
    
End Sub

Private Sub mnuPreview_Click()
    
    Dim bExeExisted As Boolean, sTmpFile As String
    Dim nIdxIcon1 As Long, nIdxIcon2 As Long, s As String, sCode As String, sFrmFile As String
    
    '首先判断PYTHONEXE是否存在
    If Len(g_PythonExe) = 0 Then
        mnuPythonExe_Click
    ElseIf Dir(g_PythonExe) = "" Then
        g_PythonExe = ""
        mnuPythonExe_Click  '如果没有设置python.exe位置，则先打开设置界面进行设置
    Else
        bExeExisted = True
    End If
    
    '前面设置后再确认是否正确，不正确则退出
    If Not bExeExisted And Len(g_PythonExe) = 0 Then
        Exit Sub
    End If
    
    '创建一个临时文件
    sTmpFile = CreateTempFile("vt")
    If Len(sTmpFile) Then
        '暂存临时文件名，在ADDIN退出时全部删除
        ReDim Preserve m_saTmpFile(UBound(m_saTmpFile) + 1) As String
        m_saTmpFile(UBound(m_saTmpFile)) = sTmpFile
        
        sCode = TxtCode.Text
        
        '为了体验更好，如果设置了图标的话，则将图标也拷贝到临时文件夹[假定图标在窗体文件同一个目录]
        nIdxIcon1 = InStr(1, sCode, ".iconbitmap(default=r'")
        If nIdxIcon1 > 0 Then
            nIdxIcon1 = nIdxIcon1 + Len(".iconbitmap(default=r'")
            nIdxIcon2 = InStr(nIdxIcon1 + 1, sCode, "'")
            If nIdxIcon2 > 0 And (nIdxIcon2 - nIdxIcon1 < 256) Then
                s = Mid$(sCode, nIdxIcon1, nIdxIcon2 - nIdxIcon1)
                On Error Resume Next
                sFrmFile = m_curFrm.FileNames(1)
                On Error GoTo 0
                If Len(sFrmFile) Then
                    On Error Resume Next
                    FileCopy AddSlash(PathName(sFrmFile)) & s, AddSlash(PathName(sTmpFile)) & s
                    If Err.Number = 0 Then
                        ReDim Preserve m_saTmpFile(UBound(m_saTmpFile) + 1) As String
                        m_saTmpFile(UBound(m_saTmpFile)) = AddSlash(PathName(sTmpFile)) & s
                        sCode = Left$(sCode, nIdxIcon1 - 1) & m_saTmpFile(UBound(m_saTmpFile)) & Mid$(sCode, nIdxIcon2)
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
        
        Utf8File_Write_VB sTmpFile, sCode
        Shell Chr(34) & g_PythonExe & """ """ & sTmpFile & Chr(34)
    Else
        MsgBox L("l_msgCreateTempFileFailed", "创建临时文件失败！"), vbInformation
    End If
    
End Sub

Private Sub mnuPythonExe_Click()
    Dim sExe As String, sExes() As String, i As Long
    
    Load frmOption
    
    '将系统中所有已经安装的Python都添加到组合框中提供选择
    frmOption.cmbPythonExe.Clear
    sExes = GetAllInstalledPython()
    If UBound(sExes) >= 0 Then
        For i = 0 To UBound(sExes)
            frmOption.cmbPythonExe.AddItem sExes(i)
        Next
    End If
    
    '从注册表中读取保存的配置
    sExe = GetSetting(App.Title, "Settings", "PythonExe", "")
    If Len(sExe) > 0 Then
        If Dir(sExe) = "" Then
            sExe = ""
        End If
    End If
    
    If Len(sExe) Then
        frmOption.cmbPythonExe.Text = sExe
    ElseIf frmOption.cmbPythonExe.ListCount > 0 Then
        frmOption.cmbPythonExe.ListIndex = 0
    End If
    
    frmOption.Show vbModal  '在frmOption窗体中会设置g_PythonExe变量并保存到注册表
    
End Sub

Private Sub mnuQuit_Click()
    mConnect.Hide
End Sub

Private Sub mnuRefreshForms_Click()
    CmdRefsFormsList_Click
End Sub

Private Sub mnuRelPos_Click()
    
    Dim o As Object
    
    mnuRelPos.Checked = Not mnuRelPos.Checked
    
    '绝对坐标
'    If Not m_curFrm Is Nothing Then
'        If Not mnuRelPos.Checked And m_curFrm.Properties("ScaleMode") <> vbTwips Then
'            '如果使用绝对坐标，不支持Frame控件
'            For Each o In m_curFrm.Designer.VBControls
'                If o.ClassName = "Frame" Then
'                    MsgBox L("l_msgFrameNotSupportInAbs", "绝对坐标布局不支持Frame控件，请改用相对坐标或去掉Frame控件。"), vbInformation
'                    mnuRelPos.Checked = True
'                    Exit For
'                End If
'            Next
'        End If
'    End If
'
    SaveSetting App.Title, "Settings", "RelPos", IIf(mnuRelPos.Checked, "1", "0")
End Sub

'在文件中恢复配置
Private Sub mnuRestoreConfig_Click()
    
    Dim cSerial As New clsSerialization
    Dim sIn As String, i As Long, s As String, sF As String
    Dim re As RegExp, Matches As MatchCollection, Mth As Match
    Dim csa() As String
    
    If Len(cmbFrms.Text) = 0 Or LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Then
        MsgBox L("l_msgChooseAForm", "请先选择一个窗体！"), vbInformation
        Exit Sub
    End If
    
    '判断是否安装了正则表达式组件VBScript.dll
    On Error Resume Next
    Set re = New RegExp
    If Err.Number <> 0 Then
        MsgBox L("l_msgNoRegExp", "系统没有注册VBScript.dll正则表达式组件，无法执行此功能！"), vbInformation
        mnuRestoreConfig.Enabled = False
        mnuSaveConfig.Enabled = False
        Exit Sub
    End If
    Err.Clear
    
    sF = m_curFrm.FileNames(1) & ".save"
    
    On Error GoTo 0
    
    If sF = "" Or sF = ".save" Then
        MsgBox L("l_msgFrmNoSaved", "设计窗体尚未保存，请先保存设计窗体。"), vbInformation
        Exit Sub
    End If
    
    If Len(Dir(sF)) = 0 Then
        MsgBox L_F("l_msgFileNotExist", "{0} 文件不存在！", sF), vbInformation
        Exit Sub
    End If
    
    sIn = Utf8File_Read_VB(sF)
    
    re.MultiLine = True
    re.Global = True
    
    'On Error Resume Next
    're.Pattern = REGX_PATTERN_FRM
    'Set Matches = re.Execute(sIn)
    'Set Mth = Matches(0)
    'cSerial.SerialString = Mth.SubMatches(0)
    'cSerial.Deserializer m_curFrm.Caption, m_curFrm.ScaleWidth, m_curFrm.ScaleHeight
    
    re.Pattern = REGX_PATTERN_CTL
    Set Matches = re.Execute(sIn)
    For Each Mth In Matches
        csa = Split(Mth, SEP_NAME_FROM_CONTENT)
        csa(0) = Replace(csa(0), REGX_INC_CTL_S, "")
        csa(1) = Replace(csa(1), REGX_INC_CTL_E, "")
        For i = 0 To UBound(g_Comps)
            If g_Comps(i).Name = csa(0) Then
                cSerial.SerialString = csa(1)
                cSerial.Deserializer g_Comps(i)
                Exit For
            End If
        Next
    Next
    
    '避免覆盖，先获取第一个控件数据到表格
    FetchCfgFromCls 0
    m_PrevCompIdx = 0
    LstComps.ListIndex = 0
    
    Set re = Nothing
    
    MsgBox L_F("l_msgRestoreCfgSuccesed", "已经成功从文件\n{0}\n恢复自定义配置！", sF), vbInformation
    
End Sub

'保存当前配置到窗体同名文件
Private Sub mnuSaveConfig_Click()
    
    Dim sOut As New cStrBuilder, i As Long, s As String, sF As String
    Dim cSerial As New clsSerialization
    
    On Error Resume Next
    sF = m_curFrm.FileNames(1) & ".save"
    On Error GoTo 0
    
    If sF = "" Or sF = ".save" Then
        MsgBox L("l_msgFrmNoSaved", "设计窗体尚未保存，请先保存设计窗体。"), vbInformation
        Exit Sub
    End If
    
    If Len(cmbFrms.Text) = 0 Or LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Then
        MsgBox L("l_msgHasNoCfgToSave", "当前没有可以保存的配置！"), vbInformation
        Exit Sub
    End If
    
    '先保存主窗体配置
    'sOut.Append REGX_INC_FRM_S
    'cSerial.Serializer m_curFrm.Caption, m_curFrm.ScaleWidth, m_curFrm.ScaleHeight
    'sOut.Append cSerial.SerialString()
    'sOut.Append REGX_INC_FRM_E
    
    For i = 0 To UBound(g_Comps)
        sOut.Append REGX_INC_CTL_S
        sOut.Append g_Comps(i).Name
        sOut.Append SEP_NAME_FROM_CONTENT
        
        cSerial.Reset
        cSerial.Serializer g_Comps(i)
        
        sOut.Append cSerial.SerialString()
        sOut.Append REGX_INC_CTL_E
    Next
    
    '保存到文件
    Utf8File_Write_VB sF, sOut.toString()
    MsgBox L_F("l_msgCfgSaved", "配置已经保存到：\n{0}", sF), vbInformation
    
End Sub

Private Sub mnuUnicodePrefixU_Click()
    
    If MsgBox(L("l_msgChangePrefixU", "改变UNICODE字符串的前缀选项需要全部刷新控件状态。\n继续吗？"), vbQuestion + vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    mnuUnicodePrefixU.Checked = Not mnuUnicodePrefixU.Checked
    g_bUnicodePrefixU = mnuUnicodePrefixU.Checked
    SaveSetting App.Title, "Settings", "UnicodePrefix", IIf(mnuUnicodePrefixU.Checked, "1", "0")
    
    CmdRefsFormsList_Click
    
End Sub

Private Sub mnuUseTtk_Click()
    Dim i As Long, s As String
    
    If LstComps.ListCount > 0 And LstComps.ListIndex >= 0 Then
        If InStr(1, LstComps.List(LstComps.ListIndex), "ComboBox") Then
            LstComps_Click                                                      '先保存配置，避免万一组合框切换后配置不对
        End If
    End If
    
    mnuUseTtk.Checked = Not mnuUseTtk.Checked
    
    '判断是否有TTK都有的控件，如果有，则不允许取消TTK选项
    If Not mnuUseTtk.Checked Then
        For i = 0 To LstComps.ListCount - 1
            s = Mid(LstComps.List(i), InStr(1, LstComps.List(i), "(") + 1)
            s = Left(s, Len(s) - 1)
            If InStr(1, " ProgressBar, TreeView, TabStrip, Line, ", " " & s & ",") > 0 Then
                MsgBox L("l_msgCantCancelTTK", "窗体中有部分控件仅在TTK库中存在，不能取消TTK选项。"), vbInformation
                mnuUseTtk.Checked = True
                Exit For
            End If
        Next
    End If
    
    '切换组合框适配器的TTK属性
    If LstComps.ListCount > 0 Then
        For i = 0 To UBound(g_Comps)
            If TypeName(g_Comps(i)) = "clsComboboxAdapter" Then
                g_Comps(i).TTK = mnuUseTtk.Checked
            End If
        Next
        
        If LstComps.ListIndex >= 0 Then
            If InStr(1, LstComps.List(LstComps.ListIndex), "ComboBox") Then
                FetchCfgFromCls LstComps.ListIndex                              '重新获取组合框信息
            End If
        End If
        LstComps_Click
    End If
    
    SaveSetting App.Title, "Settings", "UseTtk", IIf(mnuUseTtk.Checked, "1", "0")
    
End Sub

Private Sub mnuV2andV3Code_Click()
    mnuV2andV3Code.Checked = Not mnuV2andV3Code.Checked
    SaveSetting App.Title, "Settings", "V2andV3Code", IIf(mnuV2andV3Code.Checked, "1", "0")
End Sub

'自动填充编辑用的组合框内容
Private Sub FillcmbEdit(Row As Long, Col As Long)
    
    Dim sa() As String, i As Long, nEditType As Long, fn As String
    Static s_NoFirstcmbEditList As Boolean, s_NoFirstcmbEditCombo As Boolean
    
    If LstComps.ListCount = 0 Or LstComps.ListIndex < 0 Then Exit Sub
    
    '0表示内置文本编辑框，1表示仅限下拉列表，2表示下拉列表加文本输入
    nEditType = g_Comps(LstComps.ListIndex).GetAttrValueList(LstCfg.CellText(Row, 0), sa)
    
    If nEditType = 1 Then
        LstCfg.BindControl 1, cmbEditList
        cmbEditList.Clear
        cmbEditList.AddItem "" '在第一行放一个空字符串，这样就可以不设置对应的参数。
        For i = 0 To UBound(sa)
            cmbEditList.AddItem sa(i)
        Next
        For i = 0 To cmbEditList.ListCount - 1
            If cmbEditList.List(i) = LstCfg.CellText(Row, Col) Then
                cmbEditList.ListIndex = i
                Exit For
            End If
        Next
        cmbEditList.Refresh
    ElseIf nEditType = 2 Then
        LstCfg.BindControl 1, cmbEditCombo
        cmbEditCombo.Clear
        For i = 0 To UBound(sa)
            cmbEditCombo.AddItem sa(i)
        Next
        cmbEditCombo.Text = LstCfg.CellText(Row, Col)
        cmbEditCombo.Refresh
    Else
        LstCfg.BindControl 1, Nothing  '使用内置文本编辑框
    End If
    
End Sub

Private Sub stabar_DblClick()
    MsgBox L("l_msgCtlsSupported", "支持控件列表：") & vbCrLf & "Menu, Label, TextBox, PictureBox, Frame, CommandButton, CheckBox, OptionButton, ComboBox," & vbCrLf & _
            "ListBox, HScrollBar, VScrollBar, Slider, ProgressBar, TreeView, StatusBar, CommonDialog, Line" & vbCrLf & vbCrLf
End Sub

Private Sub stabar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetText "https://github.com/cdhigh/Visual-Tkinter-for-Python"
    End If
End Sub

Private Sub TxtCode_Change()
    mnuPreview.Enabled = (Len(TxtCode.Text) > 0)
End Sub

Private Sub TxtCode_DblClick()
    Static s_l As Single, s_t As Single, s_w As Single, s_h As Single
    Static s_txt As String
    
    If m_TxtCodeExpanded Then
        TxtCode.Move s_l, s_t, s_w, s_h
        m_TxtCodeExpanded = False
    Else
        s_l = TxtCode.Left
        s_t = TxtCode.Top
        s_w = TxtCode.Width
        s_h = TxtCode.Height
        TxtCode.ZOrder 0
        TxtCode.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        m_TxtCodeExpanded = True
    End If
End Sub

Private Sub TxtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And TxtCode.Width = Me.ScaleWidth Then
        TxtCode_DblClick
    End If
End Sub

'查看鼠标和按键消息详情
Private Sub TxtTips_DblClick()
    Static s_l As Single, s_t As Single, s_w As Single, s_h As Single
    Static s_txt As String
    
    Dim s As String
    s = TxtTips.Text
    If Len(s) Then
        If Left(s, Len("bindcommand")) = "bindcommand" Then
            s_l = TxtTips.Left
            s_t = TxtTips.Top
            s_w = TxtTips.Width
            s_h = TxtTips.Height
            TxtTips.ZOrder 0
            TxtTips.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            s_txt = TxtTips.Text
            m_TxtTipsExpanded = True
            TxtTips.Text = "<再次双击返回>" & vbCrLf & _
            "bindcommand" & vbCrLf & _
            "使用bind()绑定的事件处理事件列表，需要绑定多个则使用逗号分隔，如果不需要则留空。" & vbCrLf & _
            "所有事件列表如下：" & vbCrLf & _
            "<ButtonPress-n> : 鼠标按钮n按下，n:1(左键);2(中键);3(右键)" & vbCrLf & _
            "<Button-n>,<n> : 都是<ButtonPress-n>的简化形式" & vbCrLf & _
            "<ButtonRelease-n> : 鼠标按钮n被松开" & vbCrLf & _
            "<Bn-Motion> : 在按住按钮n的同时，鼠标发生移动" & vbCrLf & _
            "<Double-Button-n> : 鼠标按钮n双击" & vbCrLf & _
            "<Triple-Button-n> : 鼠标按钮n三击" & vbCrLf & _
            "<Enter> : 鼠标指针进入组件" & vbCrLf & _
            "<Leave> : 鼠标指针离开组件" & vbCrLf & _
            "<FocusIn> / <FocusOut> : 组件获得或失去焦点" & vbCrLf & _
            "<KeyPress> : 按下任意键" & vbCrLf & _
            "<KeyRelease> : 松开任意键" & vbCrLf & _
            "<KeyPress-key> : 按下key，比如<KeyPress-H>表示按下H键，可以简化为使用双引号代替尖括号将字符括起来，比如：""H""。" & vbCrLf & _
            "<KeyRelease-key> : 松开key" & vbCrLf & _
            "<Key> : 按下任意键。" & vbCrLf & _
            "<Key-key> : <KeyPress-key>的简化形式，比如<Key-H>。" & vbCrLf & _
            "<key> : 使用后附的特殊键定义替换key，表示按下特定键。" & vbCrLf & _
            "<Prefix-key> : 在按住Prefix的同时，按下key，可以使用Alt,Shift,Control的单个或组合比如<Control-Alt-key>" & vbCrLf
            
            TxtTips.Text = TxtTips.Text & "<Configure> : 控件大小改变后触发。" & vbCrLf & _
            "附全部特殊键定义：" & vbCrLf & _
            "Cancel,Break,BackSpace,Tab,Return," & vbCrLf & _
            "Sift_L , Shift_R, Control_L, Control_R, Alt_L, Alt_R, Pause" & vbCrLf & _
            "Caps_Loack,Escape,Prior(PageUp),Next(PageDown),End,Home,Left,Up,Right,Down,Print," & vbCrLf & _
            "Insert,Delete,F1-12,Num_Lock,Scroll_Lock,space,less"
            TxtTips.SelStart = 1
            TxtTips.SelLength = 0
        ElseIf Left(s, Len("<再次双击返回>")) = "<再次双击返回>" Then
            TxtTips.Move s_l, s_t, s_w, s_h
            TxtTips.Text = s_txt
            m_TxtTipsExpanded = False
        End If
    End If
End Sub

Private Sub TxtTips_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
    End If
End Sub

Private Sub TxtTips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staTips", "属性解析窗口，在有些属性状态下可以双击变大。")
End Sub

Private Sub LstComps_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staComps", "控件列表窗口，列出对应窗体上所有控件名和控件类型。")
End Sub

Private Sub cmbFrms_GotFocus()
    stabar.SimpleText = L("l_staFrms", "窗体列表，程序中支持多个设计窗口。")
End Sub

Private Sub CmdCopyToClipboard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCopyCode", "拷贝代码到剪贴板，可以选择拷贝全部还是仅拷贝界面生成部分。")
End Sub

Private Sub CmdGenCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCmdGenCode", "全部的控件属性都配置完成后，使用这个按钮生成Python代码。")
End Sub

Private Sub CmdQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staQuit", "直接退出！")
End Sub

Private Sub CmdRefsFormsList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staRefsFrms", "刷新窗体和控件，如果插件运行后修改了窗体和控件，请刷新后再重新生成代码。")
End Sub

Private Sub CmdSaveToFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCmdSaveFile", "如果必要，可以选择代码保存到文件(UTF-8带BOM格式)。")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = ""
End Sub

Private Sub LstCfg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staLstCfg", "属性列表窗口，双击属性值或按F2键可以编辑，程序只生成对应前面打钩的属性的代码。")
End Sub

Private Sub TxtCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staTxtCode", "代码预览窗口，双击可以放大。如果需要，也可以直接在这里修改代码。")
End Sub

Private Sub stabar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = "https://github.com/cdhigh/Visual-Tkinter-for-Python  ['Ctrl+Click' copy url to clipboard]"
End Sub

