VERSION 5.00
Begin VB.Form frmEncodeAFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "编码文件为Base64字符串"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13230
   Icon            =   "frmEncodeAFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   13230
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtCharsPerLine 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "80"
      Top             =   720
      Width           =   975
   End
   Begin TkinterDesigner.xpcmdbutton cmdCancelEncode 
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   9000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "退出(&Q)"
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
   Begin TkinterDesigner.xpcmdbutton cmdSaveBase64Result 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   9000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "保存(&S)"
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
   Begin TkinterDesigner.xpcmdbutton cmdBase64It 
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   9000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "编码(&E)"
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
   Begin VB.TextBox txtBase64Result 
      Height          =   7455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1200
      Width           =   12975
   End
   Begin TkinterDesigner.xpcmdbutton cmdChooseSourceToEncode 
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "..."
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
   Begin VB.TextBox txtSourceToEncode 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   10695
   End
   Begin VB.Label lblCharsPerLine 
      Alignment       =   1  'Right Justify
      Caption         =   "每行字符数"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblSourceToEncode 
      Alignment       =   1  'Right Justify
      Caption         =   "源文件"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmEncodeAFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'实际编码一个文件
Private Sub cmdBase64It_Click()
    Dim sFileName As String, sResult As String, abContent() As Byte, charsPerLine As Integer
    Dim sF As String
    
    sFileName = Trim$(txtSourceToEncode.Text)
    If Len(sFileName) <= 0 Then
        MsgBox L("l_msgFileFieldNull", "文件不能为空！"), vbInformation
        Exit Sub
    End If
    
    On Error GoTo DirErr
    
    charsPerLine = CInt(txtCharsPerLine.Text)
    
    If Dir(sFileName) = "" Then
        MsgBox L_F("l_msgFileNotExist", "文件{0}不存在，请重新选择文件。", sFileName), vbInformation
        Exit Sub
    ElseIf FileLen(sFileName) > 500000 Then
        MsgBox L("l_msgFileTooBig", "文件太大，速度会很慢，暂时不支持！"), vbInformation
        Exit Sub
    End If
    
    '用二进制方式读取内容
    If ReadFileBinaryContent(sFileName, abContent) = 0 Then
        MsgBox L_F("l_msgReadFileError", "读取文件{0}出错。", sFileName), vbInformation
        Exit Sub
    End If
    
    Base64Encode abContent, sResult, "", charsPerLine
    
    If Len(sResult) >= 65530 Then
        MsgBox L("l_msgEncodeResultTooLong", "转换后的编码字符串太长，文本框装不下，请选择一个文件直接用于保存结果！"), vbInformation
        txtBase64Result.Text = ""
        
        sF = FileDialog(Me, True, L("l_fdSave", "将文件保存到："), "All Files (*.*)|*.*")
        If Len(sF) > 0 Then
            SaveStringToFile sF, sResult
        End If
    Else
        txtBase64Result.Text = sResult
    End If
    
    Exit Sub
DirErr:
    MsgBox L_F("l_msgFileNotExist", "文件{0}不存在，请重新选择文件。", sFileName), vbInformation
    
End Sub

Private Sub cmdCancelEncode_Click()
    Unload Me
End Sub

'打开文件浏览框，选择一个文件进行编码
Private Sub cmdChooseSourceToEncode_Click()
    Dim sF As String
    sF = FileDialog(Me, False, L("l_fdOpen", "请选择文件"), "All Files (*.*)|*.*", txtSourceToEncode.Text)
    If Len(sF) Then
        txtSourceToEncode.Text = sF
    End If
End Sub

'将文本框的内容保存到磁盘文本文件
Private Sub cmdSaveBase64Result_Click()
    Dim sF As String, s As String, nm As Long, nf As Long
    
    s = txtBase64Result.Text
    If Len(s) > 2 Then
        sF = FileDialog(Me, True, L("l_fdSave", "将文件保存到："), "Python Files (*.py)|*.py|Text Files (*.txt)|*.txt|All Files (*.*)|*.*")
        If Len(sF) Then
            If Len(FileExt(sF)) = 0 Then sF = sF & ".py"  '如果文件名没有扩展名，自动添加.py扩展名
            SaveStringToFile sF, s
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim ctl As Control
    
    '多语种支持
    Me.Caption = L(Me.Name, Me.Caption)
    For Each ctl In Me.Controls
        If TypeName(ctl) = "xpcmdbutton" Or TypeName(ctl) = "Label" Then
            ctl.Caption = L(ctl.Name, ctl.Caption)
        End If
    Next
    
End Sub

Private Sub SaveStringToFile(ByRef sFileName As String, ByRef s As String)
    Dim fileNum As Integer
    On Error GoTo errHandler
    fileNum = FreeFile()
    Open sFileName For Output As fileNum
    Print #fileNum, s
    Close fileNum
    Exit Sub
errHandler:
    MsgBox L_F("l_msgWriteFileError", "写文件{0}出错。", sFileName), vbInformation
End Sub

'添加Ctrl+A快捷键
Private Sub txtBase64Result_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtBase64Result.SelStart = 0
        txtBase64Result.SelLength = Len(txtBase64Result.Text) + 1
    End If
End Sub
