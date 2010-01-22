VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9c.ocx"
Begin VB.Form frmFlash 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flash Event"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPlayFlash 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   5520
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   5280
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5730
      _cx             =   10107
      _cy             =   9313
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label lblSkipFlash 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   5535
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrPlayFlash_Timer()
    ' Check if we're playing a flash video.
    If Flash.CurrentFrame > 0 Then

        ' Check if the flash video is over.
        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then

            ' Reset the flash movie.
            Flash.FrameNum = 0
            Flash.Stop

            ' Disable this timer.
            tmrPlayFlash.Enabled = False

            ' Start playing the map music.
            Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))

            ' Unload the flash window.
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Call GUI_PictureLoad(frmFlash, "GUI\FlashTheatre")
End Sub

Private Sub lblSkipFlash_Click()
    ' Reset the flash video.
    Flash.FrameNum = 0
    Flash.Stop

    ' Disable the timer.
    tmrPlayFlash.Enabled = False

    ' Start playing the map music.
    Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))

    ' Unload this window.
    Unload Me
End Sub
