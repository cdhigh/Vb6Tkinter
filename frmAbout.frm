VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4230
   ClientLeft      =   4755
   ClientTop       =   3660
   ClientWidth     =   10830
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10830
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   240
      Picture         =   "frmAbout.frx":058A
      ScaleHeight     =   3735
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblAbout2 
      Caption         =   "Visit <https://github.com/cdhigh/Vb6Tkinter> for more detailed information"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label lblAbout1 
      Caption         =   "With Vb6Tkinter, you can create your Tkinter UI in the most intuitive way"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label lblAboutAuthor 
      Caption         =   "Made by cdhigh"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label lblAboutVer 
      Caption         =   "Vb6Tkinter v1.7.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblAboutVer.Caption = "Vb6Tkinter v" & g_AppVerString
    
    '多语种支持
    Me.Caption = L(Me.Name, Me.Caption)
    lblAboutAuthor.Caption = L("lblAboutAuthor", "Made by cdhigh")
    lblAbout1.Caption = L("lblAbout1", "With Vb6Tkinter, you can create your Tkinter UI in the most intuitive way")
    lblAbout2.Caption = L("lblAbout2", "Visit <https://github.com/cdhigh/Vb6Tkinter> for more detailed information")
End Sub

Private Sub lblAbout2_Click()
    Call ShellExecute(FrmMain.hWnd, "open", OFFICIAL_SITE, vbNullString, vbNullString, &H0)
    Unload Me
End Sub

Private Sub picAbout_Click()
    Unload Me
End Sub
