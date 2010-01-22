VERSION 5.00
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   8940
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   6360
      Top             =   120
   End
   Begin VB.Label lblOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Checking..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label picNews 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Receiving News..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Server Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label picIpConfig 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label picDeleteAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label picCredits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   1500
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim FileName As String

    Call GUI_PictureLoad(frmMainMenu, "GUI\MainMenu")

    FileName = ReadINI("CONFIG", "MenuMusic", App.Path & "\Config.ini")

    If FileExists("Music\" & FileName) Then
        Call MusicPlay(FileName)
    End If

    Call MainMenuInit
End Sub

Private Sub Form_GotFocus()
    If frmMirage.Socket.State = 0 Then
        frmMirage.Socket.Connect
    End If
End Sub

Private Sub picIpConfig_Click()
    Me.Visible = False
    frmIpconfig.Visible = True
End Sub

Private Sub picNewAccount_Click()
    Me.Visible = False
    frmNewAccount.Visible = True
End Sub

Private Sub picDeleteAccount_Click()
    frmDeleteAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picLogin_Click()
    If LenB(frmLogin.txtPassword.Text) <> 0 Then
        frmLogin.Check1.Value = vbChecked
    Else
        frmLogin.Check1.Value = vbUnchecked
    End If

    Me.Visible = False
    frmLogin.Visible = True
End Sub

Private Sub picCredits_Click()
    Me.Visible = False
    frmCredits.Visible = True
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub Status_Timer()
    If ConnectToServer = True Then
        If Not AllDataReceived Then
            Call SendData(POut.GiveMeTheMax & END_CHAR)
        Else
            lblOnline.Caption = "Online"
            lblOnline.ForeColor = vbGreen
        End If
    Else
        picNews.Caption = "Could not connect. The server may be down."

        lblOnline.Caption = "Offline"
        lblOnline.ForeColor = vbRed
    End If
End Sub
