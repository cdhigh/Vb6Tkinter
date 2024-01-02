VERSION 5.00
Begin VB.Form frmNewVer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New version found"
   ClientHeight    =   2595
   ClientLeft      =   5955
   ClientTop       =   4425
   ClientWidth     =   6900
   Icon            =   "frmNewVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6900
   Begin Vb6Tkinter.xpcmdbutton cmdCancelVer 
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Cancel"
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
   Begin Vb6Tkinter.xpcmdbutton cmdDownload 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "Download"
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
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "frmNewVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelVer_Click()
    Unload Me
End Sub

Private Sub cmdDownload_Click()
    Call ShellExecute(FrmMain.hWnd, "open", OFFICIAL_RELEASES, vbNullString, vbNullString, &H0)
End Sub

Private Sub Form_Load()
    Dim ctl As Control
    
    '多语种支持
    Me.Caption = L(Me.Name, Me.Caption)
    For Each ctl In Me.Controls
        If TypeName(ctl) = "xpcmdbutton" Then
            ctl.Caption = L(ctl.Name, ctl.Caption)
        End If
    Next
End Sub
