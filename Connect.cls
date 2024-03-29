VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Vb6Tkinter"
Option Explicit

Private mcbMenuItem         As Office.CommandBarControl
Private mcbToolBoxItem      As Office.CommandBarControl
Private mfrmAddIn           As New FrmMain
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents ToolBoxHandler As CommandBarEvents
Attribute ToolBoxHandler.VB_VarHelpID = -1

'ADDIN必须要实现的接口，使用这个接口而不是直接使用VB6提供的Designer是为了支持VB6绿色精简版
Implements IDTExtensibility

Private Sub Class_Terminate()
    Set IDTExtensibility = Nothing
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    '预防编译器删除函数，因为这个接口必须要实现
    Dim i As Long
    i = 1
End Sub
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    '预防编译器删除函数，因为这个接口必须要实现
    Dim i As Long
    i = 1
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    Set VbeInst = VBInst
    
    If ConnectMode = vbext_cm_Startup Or ConnectMode = vbext_cm_AfterStartup Then
        AddToMenu (App.Title & "(&T)")
        AddToToolBox (App.Title)
    ElseIf ConnectMode = vbext_cm_External Then
        Show
    End If

End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    'If RemoveMode = vbext_dm_UserClosed Then
    'End If
    
    On Error Resume Next
    
    '删除命令栏条目
    If Not (mcbMenuItem Is Nothing) Then mcbMenuItem.Delete
    If Not (mcbToolBoxItem Is Nothing) Then mcbToolBoxItem.Delete
    Set mcbMenuItem = Nothing
    Set mcbToolBoxItem = Nothing
    
    If Not (mfrmAddIn Is Nothing) Then Unload mfrmAddIn
    Set mfrmAddIn = Nothing
End Sub

Public Sub Hide()
    If Not (mfrmAddIn Is Nothing) Then
        Unload mfrmAddIn
        Set mfrmAddIn = Nothing
    End If
End Sub

Public Sub Show()
    If mfrmAddIn Is Nothing Then Set mfrmAddIn = New FrmMain
    Set mfrmAddIn.mConnect = Me
    mfrmAddIn.Show
End Sub

'在外接程序菜单下增加一个菜单项
Private Sub AddToMenu(sCaption As String)
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Variant
    
    On Error Resume Next
    
    '察看能否找到外接程序菜单
    Set cbMenu = VbeInst.CommandBars("外接程序")
    If cbMenu Is Nothing Then Set cbMenu = VbeInst.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then Exit Sub

    '添加它到命令栏
    Set cbMenuCommandBar = cbMenu.Controls.Add(msoControlButton)
    If cbMenuCommandBar Is Nothing Then Exit Sub
    
    cbMenuCommandBar.BeginGroup = True
    
    '设置标题
    cbMenuCommandBar.Caption = sCaption
    
    'DoEvents
    'Clipboard.SetData LoadResPicture(101, vbResBitmap)
    'cbMenuCommandBar.PasteFace
    'DoEvents
    
    Set mcbMenuItem = cbMenuCommandBar
    Set MenuHandler = VbeInst.Events.CommandBarEvents(mcbMenuItem)
End Sub

'在工具栏增加一个图标
Private Sub AddToToolBox(sCaption As String)
    Dim cbToolboxCommandBar As Office.CommandBarControl
    Dim cbStandard As CommandBar
    
    '察看能否找到标准工具栏
    On Error Resume Next
    Set cbStandard = VbeInst.CommandBars("标准")
    If cbStandard Is Nothing Then Set cbStandard = VbeInst.CommandBars("Standard")
    If cbStandard Is Nothing Then Set cbStandard = VbeInst.CommandBars(2)
    If cbStandard Is Nothing Then Exit Sub
    
    Err.Clear
    'On Error GoTo AddToAddInToolboxErr
    On Error Resume Next
    
    '添加它到工具栏
    Set cbToolboxCommandBar = cbStandard.Controls.Add(msoControlButton, , , cbStandard.Controls.Count)
    cbToolboxCommandBar.BeginGroup = True
    cbToolboxCommandBar.Caption = sCaption
    Set mcbToolBoxItem = cbToolboxCommandBar
    DoEvents
    Clipboard.SetData LoadResPicture(101, vbResBitmap)
    cbToolboxCommandBar.PasteFace
    DoEvents
    
    Set ToolBoxHandler = VbeInst.Events.CommandBarEvents(mcbToolBoxItem)
'AddToAddInToolboxErr:
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub

Private Sub ToolBoxHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub
