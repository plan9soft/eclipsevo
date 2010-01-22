VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure Server IP"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmIpconfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label PicCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label PicConfirm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   2775
   End
End
Attribute VB_Name = "frmIpconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call GUI_PictureLoad(frmIpconfig, "GUI\IPConfig")

    txtIP.Text = ReadINI("IPCONFIG", "IP", App.Path & "\Config.ini")
    txtPort.Text = ReadINI("IPCONFIG", "PORT", App.Path & "\Config.ini")
End Sub

Private Sub picCancel_Click()
    frmIpconfig.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub picConfirm_Click()
    Dim IPAddr As String
    Dim Port As Long

    ' Set the IP address to a local variable.
    IPAddr = CStr(txtIP.Text)

    ' Set the port to a local variable.
    Port = CLng(txtPort.Text)

    ' Check if the IP address is valid.
    If LenB(IPAddr) = 0 Then
        Call MsgBox("Error: Please enter a valid IP address!")
        Exit Sub
    End If

    ' Check if the port is valid.
    If Port < 1 Or Port > 65535 Then
        Call MsgBox("Error: Please enter a valid port range!")
        Exit Sub
    End If

    ' Write the configuration to the server list file.
    Call WriteINI("IPCONFIG", "IP", IPAddr, App.Path & "\Config.ini")
    Call WriteINI("IPCONFIG", "PORT", CStr(Port), App.Path & "\Config.ini")

    ' Restart the TCP control.
    Call TcpDestroy

    ' Set the new configuration settings.
    frmMirage.Socket.RemoteHost = IPAddr
    frmMirage.Socket.RemotePort = Port

    ' Revert back to the old window.
    frmIpconfig.Visible = False
    frmMainMenu.Visible = True
End Sub
