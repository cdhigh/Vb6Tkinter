VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   1275
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7710
   Icon            =   "frmOption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox cmdOptionOK 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   0
      Width           =   1000
   End
   Begin VB.PictureBox cmdPythonExe 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   0
      Width           =   1000
   End
   Begin VB.ComboBox cmbPythonExe 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "示例 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "示例 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "示例 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox cmdOptionApply 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   0
      Width           =   1000
   End
   Begin VB.PictureBox cmdOptionCancel 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblPythonExe 
      Caption         =   "Python主文件"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOptionApply_Click()
    ApplySetting
End Sub

Private Sub cmdOptionOK_Click()
    If ApplySetting() Then
        Unload Me
    End If
End Sub

Private Sub cmdOptionCancel_Click()
    Unload Me
End Sub

'应用配置，返回true表示成功
Private Function ApplySetting() As Boolean
    
    Dim sExe As String
    sExe = Trim$(cmbPythonExe.Text)
    If Len(sExe) Then
        If Dir(sExe) = "" Then
            MsgBox L_F("l_msgFileNotExist", "{0} 文件不存在！", sExe), vbInformation
            Exit Function
        End If
    Else
        MsgBox L("l_msgFileFieldNull", "文件不能为空！"), vbInformation
        Exit Function
    End If
    
    g_PythonExe = sExe
    SaveSetting App.Title, "Settings", "PythonExe", sExe
    ApplySetting = True
    
End Function

Private Sub cmdPythonExe_Click()
    Dim sF As String
    sF = FileDialog(Me, False, L("l_fdOpen", "请选择文件"), "python(w).exe|python*.exe", cmbPythonExe.Text)
    If Len(sF) Then
        cmbPythonExe.Text = sF
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
