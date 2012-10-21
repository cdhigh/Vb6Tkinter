VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9585
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10425
   _ExtentX        =   18389
   _ExtentY        =   16907
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "VisualTkinter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VisualTkinter"
Option Explicit

Public FormDisplayed          As Boolean
Public VBE                    As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mcbToolBoxCommandBar      As Office.CommandBarControl
Dim mfrmAddIn                 As New FrmMain
Public WithEvents MenuHandler As CommandBarEvents                               '菜单栏事件句柄
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents ToolBoxHandler As CommandBarEvents
Attribute ToolBoxHandler.VB_VarHelpID = -1

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    Unload mfrmAddIn
    
End Sub

Sub Show()
    
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New FrmMain
    End If
    
    Set mfrmAddIn.VBE = VBE
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
    
End Sub

'------------------------------------------------------
'这个方法添加外接程序到 VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    Dim cbStandard As CommandBar
    
    On Error Resume Next
    Err.Clear
    Set cbStandard = Application.CommandBars("标准")
    If Err.Number Then Set cbStandard = Application.CommandBars("Standard")
    Err.Clear
    
    On Error GoTo error_handler
    
    '保存 vb 实例
    Set VBE = Application
    
    '这里是一个设置断点以及测试不同外接程序
    '对象,属性及方法的适当位置
    'Debug.Print VBE.FullName
    
    If ConnectMode = ext_cm_External Then
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(App.Title & "(&T)")
        '吸取事件
        Set Me.MenuHandler = VBE.Events.CommandBarEvents(mcbMenuCommandBar)
        
        If Not cbStandard Is Nothing Then
            Set mcbToolBoxCommandBar = cbStandard.Controls.Add(msoControlButton, , , cbStandard.Controls.Count)
            mcbToolBoxCommandBar.BeginGroup = True
            mcbToolBoxCommandBar.Caption = App.Title
            Clipboard.SetData LoadResPicture(101, vbResBitmap)
            mcbToolBoxCommandBar.PasteFace
            Set Me.ToolBoxHandler = VBE.Events.CommandBarEvents(mcbToolBoxCommandBar)
        End If
    End If
    
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            Me.Show
        End If
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'这个方法从 VB 中删除外接程序
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    '删除命令栏条目
    If Not mcbMenuCommandBar Is Nothing Then mcbMenuCommandBar.Delete
    If Not mcbToolBoxCommandBar Is Nothing Then mcbToolBoxCommandBar.Delete
    Set mcbMenuCommandBar = Nothing
    Set mcbToolBoxCommandBar = Nothing
    
    '关闭外接程序
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        '设置这个到连接显示的窗体
        Me.Show
    End If
End Sub

'当 IDE 中的菜单被单击时,这个事件被激活
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl                            '命令栏对象
    Dim cbMenu As Object
    
    On Error GoTo AddToAddInCommandBarErr
    
    '察看能否找到外接程序菜单
    Set cbMenu = VBE.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    '添加它到命令栏
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.BeginGroup = True
    
    '设置标题
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
    
End Function

Private Sub ToolBoxHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub
