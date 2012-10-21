VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Visual Tkinter of Python - cdhigh@sohu.com"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12960
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   864
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar stabar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   8400
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "退出(&Q)"
      Height          =   615
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdRefsFormsList 
      Caption         =   "刷新窗体列表(&R)"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox cmbFrms 
      Height          =   300
      ItemData        =   "MainForm.frx":0CCA
      Left            =   120
      List            =   "MainForm.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton CmdSaveFile 
      Caption         =   "保存到文件(&F)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7860
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdAddUsrProperty 
      Caption         =   "增加一个自定义属性(&P)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   6015
   End
   Begin VB.TextBox TxtTips 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.ListBox LstComps 
      Height          =   3660
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VisualTkinter.GridOcx LstCfg 
      Height          =   6855
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12091
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdClip 
      Caption         =   "拷贝到剪贴板(&C)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.5
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
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton CmdOutput 
      Caption         =   "生成代码(&G)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2700
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblCurPrjName 
      Caption         =   "当前工程："
      Height          =   345
      Left            =   8760
      TabIndex        =   13
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblWP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   10560
      TabIndex        =   11
      Top             =   960
      Width           =   2280
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuRefreshForms 
         Caption         =   "刷新窗体列表(&R)"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenCode 
         Caption         =   "生成代码(&G)"
      End
      Begin VB.Menu mnuSeparator2 
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
      Begin VB.Menu mnuSeparator3 
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
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "语言(&L)"
      Begin VB.Menu mnuLng 
         Caption         =   "简体中文(&C)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuChooseOut 
      Caption         =   "保存文件选项"
      Visible         =   0   'False
      Begin VB.Menu mnuOutAll 
         Caption         =   "输出全部内容"
      End
      Begin VB.Menu mnuOutMainOnly 
         Caption         =   "仅输出main()函数"
      End
      Begin VB.Menu mnuOutUiOnly 
         Caption         =   "仅输出界面生成类"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuChosseClip 
      Caption         =   "剪贴板输出选项"
      Visible         =   0   'False
      Begin VB.Menu mnuClipOutAll 
         Caption         =   "拷贝全部内容"
      End
      Begin VB.Menu mnuClipOutMainOnly 
         Caption         =   "仅拷贝main()函数"
      End
      Begin VB.Menu mnuClipOutUiOnly 
         Caption         =   "仅拷贝界面生成类"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBE As VBIDE.VBE
Public Connect As Connect

Private m_Comps() As Object                                                     '和LstComps行数一样多，对应各组件生成的实例
Private m_MainMenu As clsMenu                                                   '菜单对象
Private m_PrevCompIdx As Long
Private m_curFrm As Object
Private m_prevsf As String
Private m_nLngNum As Long                                                       ' 语言种类
Private m_HasCommonDialog As Boolean

Const NAME_TOPWINDOW = "top"

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
    
    '多语种支持
    InitMultiLanguage
    
    LstCfg.Redraw = False
    LstCfg.Editable = True
    LstCfg.EditType = EnterKey Or MouseDblClick Or F2Key
    LstCfg.CheckBoxes = True
    LstCfg.AddColumn "Property", 2260, lgAlignCenterCenter
    LstCfg.AddColumn "Value", 3400, lgAlignCenterCenter
    LstCfg.ColAlignment(0) = lgAlignLeftCenter
    LstCfg.ColAlignment(1) = lgAlignLeftCenter
    LstCfg.SelectBackColor = &HFCC597 'vbHighlight
    LstCfg.Redraw = True
    
    CmdRefsFormsList_Click
    
    Me.Caption = "Visual Tkinter of Python - cdhigh@sohu.com - v" & App.Major & "." & App.Minor
    
    mnuOopCode.Checked = GetSetting(App.Title, "Settings", "OopCode", "1") = "1"
    mnuV2andV3Code.Checked = GetSetting(App.Title, "Settings", "V2andV3Code", "1") = "1"
    mnuUseTtk.Checked = GetSetting(App.Title, "Settings", "UseTtk", "1") = "1"
    mnuRelPos.Checked = GetSetting(App.Title, "Settings", "RelPos", "1") = "1"
    
    m_HasCommonDialog = False
    
End Sub

'多语种支持初始化
Private Sub InitMultiLanguage()
    
    Dim i As Long, s As String, sa() As String
    
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
    Connect.Hide
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
        m_PrevCompIdx = -1
    Else
        LstComps.ListIndex = 0
        LstComps_Click
    End If
    
    If LstComps.ListCount > 0 Then
        CmdOutput.Enabled = True
        CmdClip.Enabled = True
        CmdSaveFile.Enabled = True
        CmdAddUsrProperty.Enabled = True
        mnuSaveConfig.Enabled = True
        mnuRestoreConfig.Enabled = True
        mnuGenCode.Enabled = True
    Else
        CmdOutput.Enabled = False
        CmdClip.Enabled = False
        CmdSaveFile.Enabled = False
        CmdAddUsrProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub


'增加自定义配置
Private Sub CmdAddUsrProperty_Click()
    
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

Private Sub CmdClip_Click()
    
    mnuClipOutMainOnly.Visible = Not mnuOopCode.Checked
    mnuClipOutUiOnly.Visible = mnuOopCode.Checked
    
    Me.PopupMenu mnuChosseClip
End Sub

'更新各个列表，创建对应的控件类实例, 返回false表示初始化失败
Private Function ResetLstComps(frm As Object) As Boolean
    
    Dim obj As Object, ObjClsModule As Object, i As Long, s As String, j As Long, nScaleWidth As Long, nScaleHeight As Long
    
    ResetLstComps = False
    If frm Is Nothing Then Exit Function
    
    LstComps.Clear
    Erase m_Comps
    Set m_MainMenu = Nothing
    
    '创建窗体实例做为列表第一项
    ReDim m_Comps(0) As Object
    Set m_Comps(0) = New clsForm
    
    '因为ScaleX/ScaleY为窗体类独有方法，只能先在这里转换窗体大小为像素单位
    nScaleWidth = Round(ScaleX(frm.Properties("ScaleWidth"), frm.Properties("ScaleMode"), vbPixels))
    nScaleHeight = Round(ScaleY(frm.Properties("ScaleHeight"), frm.Properties("ScaleMode"), vbPixels))
    m_Comps(0).InitConfig frm, nScaleWidth, nScaleHeight
    m_Comps(0).Name = NAME_TOPWINDOW
    LstComps.AddItem m_Comps(0).Name & " (Form)"
    i = 1
    
    m_HasCommonDialog = False
    
    '将控件添加到列表中
    For Each obj In frm.Designer.VBControls
        
        CreateObj obj, ObjClsModule                                             '生成对应类模块实例
        
        If Not ObjClsModule Is Nothing Then
            
            '用于自动单位转换，需要在InitConfig之前设置这个值
            ObjClsModule.ScaleMode = frm.Properties("ScaleMode")
            
            '如果窗体存在菜单控件，则创建主菜单对象，主菜单控件将管理所有的菜单项
            If obj.ClassName = "Menu" And m_MainMenu Is Nothing Then
                ReDim Preserve m_Comps(i) As Object
                Set m_MainMenu = New clsMenu
                Set m_Comps(i) = m_MainMenu
                LstComps.AddItem m_MainMenu.Name & " (MainMenu)"
                m_MainMenu.InitConfig
                i = i + 1
            End If
            
            '添加控件到控件列表
            ReDim Preserve m_Comps(i) As Object
            Set m_Comps(i) = ObjClsModule
            LstComps.AddItem obj.Properties("Name") & " (" & obj.ClassName & ")"
            
            '初始化各控件对应的类模块对象
            If obj.Container Is frm.Designer Then
                m_Comps(i).InitConfig obj, frm.Properties("ScaleWidth"), frm.Properties("ScaleHeight")
                m_Comps(i).Parent = IIf(obj.ClassName = "Menu", "MainMenu", NAME_TOPWINDOW)
            ElseIf obj.Container.ClassName = "Menu" Then  '子菜单
                m_Comps(i).InitConfig obj, 0, 0
                m_Comps(i).Parent = obj.Container.Properties("Name")
            Else
                On Error Resume Next
                nScaleWidth = obj.Container.Properties("ScaleWidth")
                nScaleHeight = obj.Container.Properties("ScaleHeight")
                If Err.Number Then                                     'Frame和个别其他容器不支持ScaleWidth属性，则使用Width代替
                    nScaleWidth = obj.Container.Properties("Width")
                    nScaleHeight = obj.Container.Properties("Height")
                End If
                Err.Clear
                On Error GoTo 0
                m_Comps(i).InitConfig obj, nScaleWidth, nScaleHeight
                m_Comps(i).Parent = obj.Container.Properties("Name")
            End If
            
            i = i + 1
            ResetLstComps = True
        ElseIf obj.ClassName = "CommonDialog" Then
            m_HasCommonDialog = True
        Else
            MsgBox L_F("l_msgCtlNotSupport", "当前暂不支持'{0}'控件\n\n程序将不生成此控件的代码。", obj.ClassName), vbInformation, App.Title
        End If
    Next
    
    '生成菜单的树形层次关系，为生成代码建立基础
    '子类储存父类的名字，父类储存所有子类的引用
    If Not m_MainMenu Is Nothing Then
        For i = 0 To UBound(m_Comps)
            If TypeName(m_Comps(i)) = "clsMenu" Then
                '将所有的顶层菜单做为clsMenu的子控件
                For j = 0 To UBound(m_Comps)
                    If TypeName(m_Comps(j)) = "clsMenuItem" And m_Comps(j).Parent = "MainMenu" Then
                        m_Comps(i).AddChild m_Comps(j)
                    End If
                Next
            ElseIf TypeName(m_Comps(i)) = "clsMenuItem" Then
                '子菜单有可能还有子菜单
                For j = 0 To UBound(m_Comps)
                    If TypeName(m_Comps(j)) = "clsMenuItem" And m_Comps(j).Parent = m_Comps(i).Name Then
                        m_Comps(i).AddChild m_Comps(j)
                    End If
                Next
            End If
        Next
    End If
    
End Function

'生成一个控件字符实例对象:输入ctlobj:控件对象，clsobj:对应的字符串对象
Private Function CreateObj(ByRef ctlobj As Object, ByRef clsobj As Object) As Object
    
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
        Case "ListBox"
            Set clsobj = New clsListbox
        Case "HScrollBar", "VScrollBar"
            Set clsobj = New clsScrollbar
        Case "Slider"
            Set clsobj = New clsScale
        Case "Frame"
            Set clsobj = New clsLabelFrame
        Case "PictureBox"
            Set clsobj = New clsCanvas
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
        Case "StatusBar"
            Set clsobj = New clsStatusbar
        Case Else:
            Set clsobj = Nothing
    End Select
    
    Set CreateObj = clsobj
    
End Function

Private Sub CmdOutput_Click()
    
    Dim i As Long, o As Object
    Dim strHead As New cStrBuilder, strOut As New cStrBuilder, strCmd As New cStrBuilder, s As String
    Dim OutOnlyV3 As Boolean, OutOOP As Boolean, OutRelPos As Boolean, usettk As Boolean
    
    If LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Then Exit Sub
    
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
    If Not OutRelPos Then
        '如果使用绝对坐标，不支持Frame控件
        For Each o In m_curFrm.Designer.VBControls
            If o.ClassName = "Frame" Then
                MsgBox L("l_msgFrameNotSupportInAbs", "绝对坐标布局不支持Frame控件，请改用相对坐标或去掉Frame控件。"), vbInformation
                Exit Sub
            End If
        Next
    End If
    
    '在输出代码前先更新一下当前显示的数据
    UpdateCfgtoCls LstComps.ListIndex
    
    strHead.Append "#!#!/usr/bin/python"
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
        If m_HasCommonDialog Then strHead.Append "import tkinter.colorchooser as tkColorChooser  #askcolor()"
        strHead.Append vbCrLf
    Else
        strHead.Append "try:"
        strHead.Append "    from tkinter import *"
        strHead.Append "except ImportError:  #Python 2.x"
        strHead.Append "    PythonVersion = 2"
        strHead.Append "    from Tkinter import *"
        strHead.Append "    from tkFont import Font"
        If usettk Then strHead.Append "    from ttk import *"
        strHead.Append "    #Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel"
        strHead.Append "    from tkMessageBox import *"
        strHead.Append "    #Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkSimpleDialog"
        If m_HasCommonDialog Then strHead.Append "    import tkColorChooser  #askcolor()"
        strHead.Append "else:  #Python 3.x"
        strHead.Append "    PythonVersion = 3"
        strHead.Append "    from tkinter.font import Font"
        If usettk Then strHead.Append "    from tkinter.ttk import *"
        strHead.Append "    from tkinter.messagebox import *"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.filedialog as tkFileDialog"
        strHead.Append "    " & IIf(m_HasCommonDialog, "", "#") & "import tkinter.simpledialog as tkSimpleDialog    #askstring()"
        If m_HasCommonDialog Then strHead.Append "    import tkinter.colorchooser as tkColorChooser  #askcolor()"
        strHead.Append vbCrLf
    End If
    
    '如果存在状态栏控件，则先输出状态栏控件的类定义
    For i = 1 To UBound(m_Comps)  '0固定为窗体，不用判断
        If TypeName(m_Comps(i)) = "clsStatusbar" Then
            strHead.Append m_Comps(i).WidgetCode()
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
        strOut.Append m_Comps(0).toString(strCmd, OutRelPos, OutOOP, usettk)  'm_Comps(0)固定是Form
        strOut.Append "        self.createWidgets()" & vbCrLf
        strOut.Append "    def createWidgets(self):"
        strOut.Append "        self." & NAME_TOPWINDOW & " = self.winfo_toplevel()" & vbCrLf
        If usettk Then strOut.Append "        self.style = Style()" & vbCrLf
    Else
        strHead.Append L("l_cmtgComps", "#所有控件和控件绑定变量引用字典，使用这个字典是为了方便在其他函数中引用所有控件。")
        strHead.Append "gComps = {}"
        strHead.Append vbCrLf & vbCrLf
        
        strOut.Append vbCrLf
        strOut.Append "def main(argv):"
        strOut.Append m_Comps(0).toString(strCmd, OutRelPos, OutOOP, usettk)  'm_Comps(0)固定是Form
        If usettk Then
            strOut.Append "    style = Style()"
            strOut.Append "    gComps['style'] = style" & vbCrLf
        End If
    End If
    
    '遍历各控件，由各控件自己输出自己的界面生成代码
    '为了保证容器控件先于其内部的控件生成，将容器不是顶层窗口的控件集中放到最后生成
    For i = 1 To UBound(m_Comps)
        If m_Comps(i).Parent = NAME_TOPWINDOW And TypeName(m_Comps(i)) <> "clsMenuItem" Then ' clsMenuItem由clsMenu处理
            strOut.Append m_Comps(i).toString(strCmd, OutRelPos, OutOOP, usettk) & vbCrLf
        End If
    Next
    For i = 1 To UBound(m_Comps)
        If m_Comps(i).Parent <> NAME_TOPWINDOW And TypeName(m_Comps(i)) <> "clsMenuItem" Then
            strOut.Append m_Comps(i).toString(strCmd, OutRelPos, OutOOP, usettk) & vbCrLf
        End If
    Next
    
    '输出到文本框
    If OutOOP Then
        strCmd.Append "if __name__ == ""__main__"":"
        strCmd.Append "    " & NAME_TOPWINDOW & " = Tk()"
        strCmd.Append "    Application(" & NAME_TOPWINDOW & ").mainloop()"
        strCmd.Append "    try: " & NAME_TOPWINDOW & ".destroy()"
        strCmd.Append "    except: pass" & vbCrLf
        TxtCode.Text = strHead.toString(vbCrLf) & strOut.toString(vbCrLf) & strCmd.toString(vbCrLf)
    Else
        strOut.Append "    " & NAME_TOPWINDOW & ".mainloop()"
        strOut.Append "    try: " & NAME_TOPWINDOW & ".destroy()"
        strOut.Append "    except: pass"
        strOut.Append vbCrLf & vbCrLf
        strOut.Append "if __name__ == ""__main__"":"
        strOut.Append "    main(sys.argv)" & vbCrLf
        TxtCode.Text = strHead.toString(vbCrLf) & strCmd.toString(vbCrLf) & strOut.toString(vbCrLf)
    End If
    
    strOut.Reset
    strHead.Reset
    strCmd.Reset
    
End Sub

Private Sub CmdRefsFormsList_Click()
    
    Dim frm As Object
    
    cmbFrms.Clear
    LstComps.Clear
    LstCfg.Clear
    
    If VBE.ActiveVBProject Is Nothing Then
        CmdOutput.Enabled = False
        CmdClip.Enabled = False
        CmdSaveFile.Enabled = False
        CmdAddUsrProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
        lblWP.Caption = ""
        Exit Sub
    End If
    
    lblWP.Caption = VBE.ActiveVBProject.Name
    
    '查找工程中所有的窗体,全部添加到组合框供选择输出
    For Each frm In VBE.ActiveVBProject.VBComponents
        If frm.Type = vbext_ct_VBForm Then
            If frm.Properties("ScaleMode") <> vbTwips And frm.Properties("ScaleMode") <> vbPoints And frm.Properties("ScaleMode") <> vbPixels Then
                MsgBox L_F("l_msgFailedScaleMode", "查找到窗体'{0}'，但是ScaleMode={1}，程序仅支持模式1/2/3。", _
                         frm.Properties("Name"), frm.Properties("ScaleMode")), vbInformation
            Else
                cmbFrms.AddItem frm.Name
            End If
        End If
    Next
    
    If cmbFrms.ListCount >= 1 Then
        cmbFrms.ListIndex = 0      '触发cmbFrms_Click
    Else
        CmdOutput.Enabled = False
        CmdClip.Enabled = False
        CmdSaveFile.Enabled = False
        CmdAddUsrProperty.Enabled = False
        mnuSaveConfig.Enabled = False
        mnuRestoreConfig.Enabled = False
        mnuGenCode.Enabled = False
    End If
    
End Sub

Private Sub CmdSaveFile_Click()
    
    mnuOutMainOnly.Visible = Not mnuOopCode.Checked
    mnuOutUiOnly.Visible = mnuOopCode.Checked
    
    Me.PopupMenu mnuChooseOut
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TxtCode.Width = Me.ScaleWidth Then
        TxtCode_DblClick
        Cancel = True
    ElseIf TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
        Cancel = True
    End If
End Sub

Private Sub lblWP_Click()
    MsgBox L("l_msgCtlsSupported", "支持控件列表：") & vbCrLf & "Menu, Label, TextBox, PictureBox, Frame, CommandButton, CheckBox, OptionButton, " & vbCrLf & _
    "ComboBox, ListBox, HScrollBar, VScrollBar, Slider, ProgressBar, TreeView, StatusBar, CommonDialog"
End Sub

Private Sub LstCfg_ItemChecked(Row As Long)
    If InStr(1, " x, y, relx, rely, width, height, relwidth, relheight,", " " & LstCfg.CellText(Row, 0) & ",") Then
        LstCfg.ItemChecked(Row) = True
    End If
    
    '更新列表中的数值到实例对象和数组
    UpdateCfgtoCls m_PrevCompIdx
    
End Sub

Private Sub LstCfg_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or InStr(1, " x, y, relx, rely, width, height, relwidth, relheight, ", " " & LstCfg.CellText(Row, 0) & ",") Then Cancel = True
End Sub

Private Sub LstCfg_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    If LstCfg.CellText(Row, Col) <> "" And NewValue = "" And LstCfg.ItemChecked(Row) Then
        LstCfg.ItemChecked(Row) = False
    ElseIf NewValue <> "" Then
        LstCfg.ItemChecked(Row) = True
    End If
End Sub

Private Sub LstCfg_RowColChanged()
    If LstComps.ListIndex >= 0 Then
        TxtTips.Text = m_Comps(LstComps.ListIndex).Tips(LstCfg.CellText(LstCfg.Row, 0))
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
    TxtTips.Text = m_Comps(LstComps.ListIndex).Description
    
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
    
    If idx < 0 Or idx > UBound(m_Comps) Then Exit Sub
    
    LstCfg.Redraw = False
    LstCfg.Clear
    
    Set cItms = m_Comps(idx).Allitems()
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
    
    If idx < 0 Or idx > UBound(m_Comps) Then Exit Sub
    
    s = ""
    For i = 0 To LstCfg.ItemCount - 1
        If LstCfg.ItemChecked(i) Then
            s = s & IIf(i > 0, "|", "") & LstCfg.CellText(i, 0) & "|" & LstCfg.CellText(i, 1)
        End If
    Next
    
    If Len(s) Then m_Comps(idx).SetConfig s
    
End Sub

Private Sub mnuClipOutAll_Click()
    Clipboard.Clear
    Clipboard.SetText TxtCode.Text
End Sub

Private Sub mnuClipOutMainOnly_Click()
    
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

Private Sub mnuClipOutUiOnly_Click()
    
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
    mnuSaveConfig.Enabled = LstComps.ListCount > 0
    mnuRestoreConfig.Enabled = LstComps.ListCount > 0
    mnuGenCode.Enabled = LstComps.ListCount > 0
End Sub

Private Sub mnuGenCode_Click()
    CmdOutput_Click
End Sub

Private Sub mnuLng_Click(Index As Integer)
    
    Dim i As Long
    
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
End Sub

Private Sub mnuOutAll_Click()
    
    Dim sF As String
    sF = FileDialog(Me, True, L("l_fdSave", "将Python文件保存到："), "*.py", m_prevsf)
    
    If Len(sF) Then Utf8File_Write_VB sF, TxtCode.Text
    
    m_prevsf = sF
    
End Sub

'仅输出main()函数，用于之前已经建好框架，并且也写了一些代码，现在修改空间布局，不用影响其他代码
Private Sub mnuOutMainOnly_Click()
    
    Dim sF As String, s As String, nm As Long, nf As Long
    
    '分析代码，仅提取main(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nm = InStr(1, s, "def main(argv):")
    nf = InStr(1, s, "if __name__")
    If nm > 0 And nf > 0 Then
        sF = FileDialog(Me, True, L("l_fdSave", "将Python文件保存到："), "*.py", m_prevsf)
        If Len(sF) Then
            Utf8File_Write_VB sF, Mid(s, nm, nf - nm)
        End If
    Else
        MsgBox L("l_msgNoMain", "代码中找不到main()函数！"), vbInformation
    End If
    
    m_prevsf = sF
    
End Sub

'仅输出界面生成类，用于之前已经建好框架，并且也写了一些代码，现在修改空间布局，不用影响其他代码
Private Sub mnuOutUiOnly_Click()
    
    Dim sF As String, s As String, nui As Long, napp As Long
    
    '分析代码，仅提取main(),使用正则表达式也可以，但是这里使用简单字符串分析
    s = TxtCode.Text
    nui = InStr(1, s, "class Application_ui(Frame):")
    napp = InStr(1, s, "class Application(Application_ui):")
    If nui > 0 And napp > 0 Then
        sF = FileDialog(Me, True, "将Python文件保存到：", "*.py", m_prevsf)
        If Len(sF) Then
            Utf8File_Write_VB sF, Mid(s, nui, napp - nui)
        End If
    Else
        MsgBox L("l_msgNoClsUi", "代码中找不到Application_ui类！"), vbInformation
    End If
    
    m_prevsf = sF
    
End Sub

Private Sub mnuQuit_Click()
    Connect.Hide
End Sub

Private Sub mnuRefreshForms_Click()
    CmdRefsFormsList_Click
End Sub

Private Sub mnuRelPos_Click()
    
    Dim o As Object
    
    mnuRelPos.Checked = Not mnuRelPos.Checked
    
    '绝对坐标
    If Not mnuRelPos.Checked And Not m_curFrm Is Nothing Then
        '如果使用绝对坐标，不支持Frame控件
        For Each o In m_curFrm.Designer.VBControls
            If o.ClassName = "Frame" Then
                MsgBox L("l_msgFrameNotSupportInAbs", "绝对坐标布局不支持Frame控件，请改用相对坐标或去掉Frame控件。"), vbInformation
                mnuRelPos.Checked = True
                Exit For
            End If
        Next
    End If
    
    SaveSetting App.Title, "Settings", "RelPos", IIf(mnuRelPos.Checked, "1", "0")
End Sub

'在文件中恢复配置
Private Sub mnuRestoreConfig_Click()
    
    Dim cSerial As New clsSerialization
    Dim sIn As String, i As Long, s As String, sF As String
    Dim re As New RegExp, Matches As MatchCollection, Mth As Match
    Dim csa() As String
    
    If Len(cmbFrms.Text) = 0 Or LstComps.ListCount = 0 Or LstCfg.ItemCount = 0 Then
        MsgBox L("l_msgChooseAForm", "请先选择一个窗体！"), vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
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
        For i = 0 To UBound(m_Comps)
            If m_Comps(i).Name = csa(0) Then
                cSerial.SerialString = csa(1)
                cSerial.Deserializer m_Comps(i)
                Exit For
            End If
        Next
    Next
    
    '避免覆盖，先获取第一个控件数据到表格
    FetchCfgFromCls 0
    m_PrevCompIdx = 0
    LstComps.ListIndex = 0
    
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
    
    For i = 0 To UBound(m_Comps)
        sOut.Append REGX_INC_CTL_S
        sOut.Append m_Comps(i).Name
        sOut.Append SEP_NAME_FROM_CONTENT
        
        cSerial.Reset
        cSerial.Serializer m_Comps(i)
        
        sOut.Append cSerial.SerialString()
        sOut.Append REGX_INC_CTL_E
    Next
    
    '保存到文件
    Utf8File_Write_VB sF, sOut.toString()
    MsgBox L_F("l_msgCfgSaved", "配置已经保存到：\n{0}", sF), vbInformation
    
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
            If InStr(1, " ProgressBar, TreeView, TabStrip, ", " " & s & ",") > 0 Then
                MsgBox L("l_msgCantCancelTTK", "窗体中有部分控件仅在TTK库中存在，不能取消TTK选项。"), vbInformation
                mnuUseTtk.Checked = True
                Exit For
            End If
        Next
    End If
    
    '切换组合框适配器的TTK属性
    If LstComps.ListCount > 0 Then
        For i = 0 To UBound(m_Comps)
            If TypeName(m_Comps(i)) = "clsComboboxAdapter" Then
                m_Comps(i).TTK = mnuUseTtk.Checked
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

Private Sub TxtCode_DblClick()
    Static s_l As Single, s_t As Single, s_w As Single, s_h As Single
    Static s_txt As String, s_Expand As Boolean
    
    If s_Expand Then
        TxtCode.Move s_l, s_t, s_w, s_h
        s_Expand = False
    Else
        s_l = TxtCode.Left
        s_t = TxtCode.Top
        s_w = TxtCode.Width
        s_h = TxtCode.Height
        TxtCode.ZOrder 0
        TxtCode.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        s_Expand = True
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
        ElseIf Left(s, Len("<再次双击返回>")) = "<再次双击返回>" Then
            TxtTips.Move s_l, s_t, s_w, s_h
            TxtTips.Text = s_txt
        End If
    End If
End Sub

'不管当前的坐标单位是什么，转换为像素
Private Sub Convert2Pixel(ByRef nWidth As Long, ByRef nHeight As Long)
    nWidth = ScaleX(nWidth, Me.ScaleMode, vbPixels)
    nHeight = ScaleY(nHeight, Me.ScaleMode, vbPixels)
End Sub

Private Sub TxtTips_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And TxtTips.Width = Me.ScaleWidth Then
        TxtTips_DblClick
    End If
End Sub

Private Sub TxtTips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staTips", "属性解析窗口，在有些属性状态下可以双击变大。")
End Sub

Private Sub CmdSerializer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staSrl", "注意：仅保存那些被选择（选择框被打勾）的配置项。")
End Sub

Private Sub LstComps_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staComps", "控件列表窗口，列出对应窗体上所有控件名和控件类型。")
End Sub

Private Sub cmbFrms_GotFocus()
    stabar.SimpleText = L("l_staFrms", "窗体列表，程序中支持多个设计窗口。")
End Sub

Private Sub CmdAddUsrProperty_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staAddAttr", "目前列表中的选项并不完整，仅包括常用属性，如果需要，可以根据手册手工添加其他属性。")
End Sub

Private Sub CmdClip_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCopyCode", "拷贝代码到剪贴板，可以选择拷贝全部还是仅拷贝界面生成部分。")
End Sub

Private Sub CmdDeserializer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staRestoreCfg", "仅恢复那些同名保存的控件配置项，新增的控件不受影响。")
End Sub

Private Sub CmdOutput_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staCmdGenCode", "全部的控件属性都配置完成后，使用这个按钮生成Python代码。")
End Sub

Private Sub CmdQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staQuit", "直接退出！")
End Sub

Private Sub CmdRefsFormsList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    stabar.SimpleText = L("l_staRefsFrms", "刷新窗体和控件，如果插件运行后修改了窗体和控件，请刷新后再重新生成代码。")
End Sub

Private Sub CmdSaveFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

