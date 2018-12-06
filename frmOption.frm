VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Config"
   ClientHeight    =   1305
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TkinterDesigner.xpcmdbutton cmdOptionCancel 
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Cancel(&C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TkinterDesigner.xpcmdbutton cmdOptionOK 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Ok(&O)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TkinterDesigner.xpcmdbutton cmdOptionApply 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Apply(&A)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TkinterDesigner.xpcmdbutton cmdPythonExe 
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbPythonExe 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblPythonExe 
      Caption         =   "Python EXE"
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
            MsgBox L_F("l_msgFileNotExist", "File '{0}' not exist!", sExe), vbInformation
            Exit Function
        End If
    Else
        MsgBox L("l_msgFileFieldNull", "File can't be null."), vbInformation
        Exit Function
    End If
    
    g_PythonExe = sExe
    SaveSetting App.Title, "Settings", "PythonExe", sExe
    ApplySetting = True
    
End Function

Private Sub cmdPythonExe_Click()
    Dim sF As String
    sF = FileDialog(Me, False, L("l_fdOpen", "Please Choose file:"), "python(w).exe|python*.exe", cmbPythonExe.Text)
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
