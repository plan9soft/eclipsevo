VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Server"
   ClientHeight    =   4890
   ClientLeft      =   420
   ClientTop       =   840
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScriptedTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8149
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   370
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "frmServer.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLogTime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraChatOpt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSTab2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "picCMsg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tmrChatLogs"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TPO"
      Tab(1).Control(1)=   "lvUsers"
      Tab(1).Control(2)=   "Command66"
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(4)=   "cmdKickPlayerReason"
      Tab(1).Control(5)=   "cmdBanPlayerReason"
      Tab(1).Control(6)=   "cmdJailPlayer"
      Tab(1).Control(7)=   "cmdViewInfo"
      Tab(1).Control(8)=   "cmdMsgPlayer"
      Tab(1).Control(9)=   "cmdMutePlayer"
      Tab(1).Control(10)=   "cmdUnmutePlayer"
      Tab(1).Control(11)=   "cmdKillPlayer"
      Tab(1).Control(12)=   "cmdHealPlayer"
      Tab(1).Control(13)=   "picBan"
      Tab(1).Control(14)=   "picStats"
      Tab(1).Control(15)=   "picJail"
      Tab(1).Control(16)=   "cmdWarpPlayer"
      Tab(1).Control(17)=   "cmdGiveAccess"
      Tab(1).Control(18)=   "picWarp"
      Tab(1).Control(19)=   "picKick"
      Tab(1).Control(20)=   "picMessage"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Control Panel"
      TabPicture(2)   =   "frmServer.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "News"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Socket(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "PlayerTimer"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "tmrShutdown"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "tmrGameAI"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "tmrSpawnMapItems"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "tmrPlayerSave"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame6"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "picWarpAll"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Frame9"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "picExp"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "picMap"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "picWeather"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Script"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Time"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmServer.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblTopic"
      Tab(3).Control(1)=   "lblWebsite"
      Tab(3).Control(2)=   "lblContent"
      Tab(3).Control(3)=   "lblForum"
      Tab(3).Control(4)=   "lstTopics"
      Tab(3).Control(5)=   "txtTopic"
      Tab(3).ControlCount=   6
      Begin VB.PictureBox picMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -70320
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   192
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtPlayerMsg 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   195
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton cmdMsgCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   194
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdServMsg 
            Caption         =   "Send Message"
            Height          =   255
            Left            =   120
            TabIndex        =   193
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblMessage 
            Caption         =   "Message:"
            Height          =   240
            Left            =   120
            TabIndex        =   196
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.PictureBox picKick 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -70320
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   187
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdServKick 
            Caption         =   "Kick Player"
            Height          =   255
            Left            =   120
            TabIndex        =   191
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdKickCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   190
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtKickReason 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   189
            Top             =   360
            Width           =   3075
         End
         Begin VB.CheckBox chkKickReason 
            Caption         =   "With Reason"
            Height          =   240
            Left            =   120
            TabIndex        =   188
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.PictureBox picWarp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   -70320
         ScaleHeight     =   199
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   175
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdServWarp 
            Caption         =   "Warp Player"
            Height          =   255
            Left            =   120
            TabIndex        =   182
            Top             =   2640
            Width           =   1575
         End
         Begin VB.HScrollBar scrlWarpMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   181
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   180
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   179
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CommandButton cmdWarpCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   178
            Top             =   2640
            Width           =   1575
         End
         Begin VB.CheckBox chkWarpReason 
            Caption         =   "With Reason"
            Height          =   240
            Left            =   120
            TabIndex        =   177
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtWarpReason 
            Height          =   285
            Left            =   120
            TabIndex        =   176
            Top             =   2280
            Width           =   3135
         End
         Begin VB.Label lblWarpMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   185
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblWarpX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   184
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblWarpY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   183
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.Frame Time 
         Caption         =   "Time"
         Height          =   735
         Left            =   -74760
         TabIndex        =   163
         Top             =   3720
         Width           =   9855
         Begin VB.CommandButton cmdSetTime 
            Caption         =   "Set"
            Height          =   285
            Left            =   1920
            TabIndex        =   172
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtTimeS 
            Height          =   285
            Left            =   1320
            TabIndex        =   171
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtTimeM 
            Height          =   285
            Left            =   720
            TabIndex        =   170
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtTimeH 
            Height          =   285
            Left            =   120
            TabIndex        =   169
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox GameTimeSpeed 
            Height          =   285
            Left            =   4200
            TabIndex        =   166
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command68 
            Caption         =   "Change Speed"
            Height          =   285
            Left            =   4800
            TabIndex        =   165
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command69 
            Caption         =   "Disable Time"
            Height          =   285
            Left            =   6480
            TabIndex        =   164
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   7680
            TabIndex        =   168
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Game Speed:"
            Height          =   255
            Left            =   3120
            TabIndex        =   167
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.PictureBox Script 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -74880
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   153
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command72 
            Caption         =   "Run"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton Command71 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1770
            TabIndex        =   154
            Top             =   1920
            Width           =   1455
         End
         Begin CodeSenseCtl.CodeSense ServerScript 
            Height          =   1455
            Left            =   120
            OleObjectBlob   =   "frmServer.frx":5C82
            TabIndex        =   156
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Script:"
            Height          =   195
            Left            =   120
            TabIndex        =   157
            Top             =   120
            Width           =   465
         End
      End
      Begin VB.TextBox txtTopic 
         Height          =   3570
         Left            =   -72480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   150
         Top             =   600
         Width           =   7575
      End
      Begin VB.CommandButton cmdGiveAccess 
         Caption         =   "Give Access"
         Height          =   255
         Left            =   -66600
         TabIndex        =   149
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Timer tmrDayNight 
         Interval        =   1000
         Left            =   -68040
         Top             =   0
      End
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   9840
         Top             =   0
      End
      Begin VB.PictureBox picWeather 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   -74880
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   129
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command65 
            Caption         =   "Snow"
            Height          =   255
            Left            =   1680
            TabIndex        =   137
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Rain"
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Thunder"
            Height          =   255
            Left            =   1680
            TabIndex        =   135
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command62 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   1080
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRainIntensity 
            Height          =   255
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   132
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   130
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Weather: None"
            Height          =   195
            Left            =   120
            TabIndex        =   133
            Top             =   720
            Width           =   1710
         End
         Begin VB.Label lblRainIntensity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensity: 25"
            Height          =   195
            Left            =   120
            TabIndex        =   131
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.PictureBox picCMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   1320
         ScaleHeight     =   1905
         ScaleWidth      =   3345
         TabIndex        =   28
         Top             =   5640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtMsg 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   3075
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            MaxLength       =   13
            TabIndex        =   33
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   360
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3015
         Left            =   240
         TabIndex        =   118
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         TabsPerRow      =   8
         TabHeight       =   353
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Main"
         TabPicture(0)   =   "frmServer.frx":5DE8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtText(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtChat"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Broadcast"
         TabPicture(1)   =   "frmServer.frx":5E04
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Global"
         TabPicture(2)   =   "frmServer.frx":5E20
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtText(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Map"
         TabPicture(3)   =   "frmServer.frx":5E3C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Private"
         TabPicture(4)   =   "frmServer.frx":5E58
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Admin"
         TabPicture(5)   =   "frmServer.frx":5E74
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtText(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Emote"
         TabPicture(6)   =   "frmServer.frx":5E90
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtText(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Errors"
         TabPicture(7)   =   "frmServer.frx":5EAC
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "txtText(7)"
         Tab(7).ControlCount=   1
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   7
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   198
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   6
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   126
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   5
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   125
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   4
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   124
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   3
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   123
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   2
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   122
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   1
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   121
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtChat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   2640
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2250
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   119
            Top             =   360
            Width           =   9375
         End
      End
      Begin VB.CommandButton cmdWarpPlayer 
         Caption         =   "Warp Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   115
         Top             =   2640
         Width           =   1575
      End
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   231
         TabIndex        =   98
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ListBox lstNPC 
            Height          =   2400
            Left            =   1680
            TabIndex        =   113
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command41 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   99
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPCs"
            Height          =   195
            Index           =   13
            Left            =   1680
            TabIndex        =   114
            Top             =   120
            Width           =   375
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indoors:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   112
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shop:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   111
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootY:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   110
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootX:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   109
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootMap:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   108
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Music:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   107
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   106
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   105
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   104
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Up:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   255
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moral:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   600
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revision:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   660
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   100
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.PictureBox picExp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -74880
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   92
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.HScrollBar scrlExp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   197
            Top             =   360
            Value           =   1
            Width           =   3015
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   95
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Execute"
            Height          =   255
            Left            =   1560
            TabIndex        =   93
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblMassExp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experience: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Map List"
         Height          =   1695
         Left            =   -69000
         TabIndex        =   96
         Top             =   360
         Width           =   4095
         Begin VB.ListBox MapList 
            Height          =   1035
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.PictureBox picWarpAll 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -74880
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command38 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   91
            Top             =   2160
            Width           =   1575
         End
         Begin VB.HScrollBar scrlMY 
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMX 
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMM 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   83
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Warp"
            Height          =   255
            Left            =   1680
            TabIndex        =   82
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblMY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblMX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblMM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Commands"
         Height          =   3255
         Left            =   -72840
         TabIndex        =   74
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton Command34 
            Caption         =   "Mass Level"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Mass Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Mass Warp"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Mass Heal"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Mass Kill"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Mass Kick"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ListBox lstTopics 
         Height          =   3570
         Left            =   -74760
         TabIndex        =   71
         Top             =   600
         Width           =   2175
      End
      Begin VB.PictureBox picJail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   -70320
         ScaleHeight     =   199
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtJailReason 
            Height          =   285
            Left            =   120
            TabIndex        =   174
            Top             =   2280
            Width           =   3135
         End
         Begin VB.CheckBox chkJailReason 
            Caption         =   "With Reason"
            Height          =   240
            Left            =   120
            TabIndex        =   173
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CommandButton cmdJailCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   90
            Top             =   2640
            Width           =   1575
         End
         Begin VB.HScrollBar scrlJailY 
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlJailX 
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlJailMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   65
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton cmdServJail 
            Caption         =   "Jail Player"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblJailY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblJailX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblJailMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.PictureBox picStats 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -71640
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Command8 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index:"
            Height          =   195
            Index           =   20
            Left            =   2400
            TabIndex        =   62
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   61
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magi:"
            Height          =   195
            Index           =   18
            Left            =   2400
            TabIndex        =   60
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            Height          =   195
            Index           =   17
            Left            =   2400
            TabIndex        =   59
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Def:"
            Height          =   195
            Index           =   16
            Left            =   2400
            TabIndex        =   58
            Top             =   840
            Width           =   315
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            Height          =   195
            Index           =   15
            Left            =   2400
            TabIndex        =   57
            Top             =   600
            Width           =   270
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Access:"
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   56
            Top             =   360
            Width           =   945
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild:"
            Height          =   195
            Index           =   13
            Left            =   2400
            TabIndex        =   55
            Top             =   120
            Width           =   405
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   54
            Top             =   3000
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   52
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   51
            Top             =   2280
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PK:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   50
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   49
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXP: /"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SP: /"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP: /"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP: /"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Character:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   780
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   645
         End
      End
      Begin VB.PictureBox picBan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -70320
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkBanReason 
            Caption         =   "With Reason"
            Height          =   240
            Left            =   120
            TabIndex        =   186
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdBanCancel 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   89
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdServBan 
            Caption         =   "Ban Player"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtBanReason 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   3075
         End
      End
      Begin VB.Frame fraChatOpt 
         Caption         =   "Chat Options"
         Height          =   855
         Left            =   240
         TabIndex        =   22
         Top             =   3480
         Width           =   7215
         Begin VB.CommandButton cmdSaveLogs 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   5520
            TabIndex        =   127
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chkLogAdmin 
            Caption         =   "Admin"
            Height          =   255
            Left            =   4680
            TabIndex        =   36
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkLogGlobal 
            Caption         =   "Global"
            Height          =   255
            Left            =   3840
            TabIndex        =   35
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkLogPM 
            Caption         =   "Private"
            Height          =   255
            Left            =   2880
            TabIndex        =   26
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkLogMap 
            Caption         =   "Map"
            Height          =   255
            Left            =   2160
            TabIndex        =   25
            Top             =   360
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkLogEmote 
            Caption         =   "Emote"
            Height          =   255
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkLogBC 
            Caption         =   "Broadcast"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdHealPlayer 
         Caption         =   "Heal Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   21
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Timer tmrPlayerSave 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   -67560
         Top             =   0
      End
      Begin VB.Timer tmrSpawnMapItems 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -65640
         Top             =   0
      End
      Begin VB.Timer tmrGameAI 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   -66120
         Top             =   0
      End
      Begin VB.Timer tmrShutdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -66600
         Top             =   0
      End
      Begin VB.Timer PlayerTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -67080
         Top             =   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Classes"
         Height          =   1095
         Left            =   -70920
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
         Begin VB.CommandButton Command30 
            Caption         =   "Edit"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Reload"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Server"
         Height          =   1455
         Left            =   -69000
         TabIndex        =   17
         Top             =   2160
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "Shutdown"
            Height          =   255
            Left            =   2280
            TabIndex        =   162
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command59 
            Caption         =   "Weather"
            Height          =   255
            Left            =   2280
            TabIndex        =   161
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Map List"
            Height          =   255
            Left            =   2280
            TabIndex        =   160
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command36 
            Caption         =   "Map Info"
            Height          =   255
            Left            =   2280
            TabIndex        =   159
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkChat 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   810
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox mnuServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   1050
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Closed 
            Caption         =   "Server Closed"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   570
            Width           =   1335
         End
         Begin VB.CheckBox GMOnly 
            Caption         =   "Admin Only"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scripts"
         Height          =   2055
         Left            =   -70920
         TabIndex        =   12
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton Command70 
            Caption         =   "Run script"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Edit Script"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Turn Off"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Turn On"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Reload"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblScriptOn 
            Alignment       =   2  'Center
            Caption         =   "Scripts: OFF"
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdKillPlayer 
         Caption         =   "Kill Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnmutePlayer 
         Caption         =   "UnMute Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   10
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdMutePlayer 
         Caption         =   "Mute Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdMsgPlayer 
         Caption         =   "Message (PM)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdViewInfo 
         Caption         =   "View Info"
         Height          =   255
         Left            =   -66600
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdJailPlayer 
         Caption         =   "Jail Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdBanPlayerReason 
         Caption         =   "Ban Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdKickPlayerReason 
         Caption         =   "Kick Player"
         Height          =   255
         Left            =   -66600
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gridlines"
         Height          =   255
         Left            =   -67920
         TabIndex        =   2
         Top             =   3840
         Value           =   1  'Checked
         Width           =   975
      End
      Begin MSWinsockLib.Winsock Socket 
         Index           =   0
         Left            =   -65160
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -69600
         TabIndex        =   139
         Top             =   3840
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Character"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Level"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sprite"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Access"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Frame News 
         Caption         =   "News"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   141
         Top             =   2040
         Width           =   1815
         Begin VB.CommandButton btnEventSend 
            Caption         =   "Send Events"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton btnEventEdit 
            Caption         =   "Edit Events"
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command73 
            Caption         =   "Send News"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Engine Info"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   144
         Top             =   360
         Width           =   1815
         Begin VB.Label lblGetIP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check IP Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   158
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label lblVer 
            Caption         =   "Build: (...)"
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblEngine 
            Caption         =   "Eclipse Evolution"
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label lblForum 
         Caption         =   "Official Support Forums"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73320
         TabIndex        =   152
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label lblContent 
         Caption         =   "Contents:"
         Height          =   255
         Left            =   -72480
         TabIndex        =   151
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblLogTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Log Save In"
         Height          =   195
         Left            =   7680
         TabIndex        =   138
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Label lblWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Official Website"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   4320
         Width           =   1125
      End
      Begin VB.Label lblTopic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   72
         Top             =   360
         Width           =   510
      End
      Begin VB.Label TPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Players Online:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   3
         Top             =   3840
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGiveAccess_Click()
    Dim Access As String
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the player is in-game.
    If IsPlaying(Index) Then

        ' Prompt the user for the new access level of the selected player.
        Access = InputBox("Give player what access?" & vbNewLine & vbNewLine & "0 - Player" & vbNewLine & "1 - Moderator" & vbNewLine & "2 - Mapper" & vbNewLine & "3 - Developer" & vbNewLine & "4 - Admin" & vbNewLine & "5 - Owner" & vbNewLine, "Give Access", CStr(Player(Index).Char(Player(Index).CharNum).Access))

        ' Check if the data entered is numeric (numbers only).
        If IsNumeric(Access) Then

            ' Check for a valid access level.
            If CLng(Access) < 0 Or CLng(Access) > 5 Then
                Call MsgBox("Please enter any value between 0 and 5.")
                Exit Sub
            End If

            ' Set the new access level server-side.
            Call SetPlayerAccess(Index, CLng(Access))

            ' Set the new access level client-side.
            Call SendPlayerData(Index)

            ' Check if the person has been given administrative powers.
            If GetPlayerAccess(Index) > 0 Then
                Call PlayerMsg(Index, "You have been given administrative status.", AdminColor)
            End If

            ' Update the player list with the new access level.
            Call ShowPLR(Index)
        End If
    End If
End Sub

Private Sub cmdJailCancel_Click()
    ' Hide the jail picture box.
    picJail.Visible = False
End Sub

Private Sub cmdBanCancel_Click()
    ' Hide the ban picture box.
    picBan.Visible = False
End Sub

Private Sub cmdKickCancel_Click()
    ' Hide the kick picture box.
    picKick.Visible = False
End Sub

Private Sub cmdMsgCancel_Click()
    ' Hide the message picture box.
    picMessage.Visible = False
End Sub

Private Sub cmdServBan_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is currently playing.
    If Not IsPlaying(Index) Then Exit Sub

    ' Check if the administrator wants to send a ban message.
    If chkBanReason.Value = vbChecked Then

        ' Check if the ban message exists.
        If LenB(txtBanReason.Text) = 0 Then
            Call MsgBox("Please input a reason to ban this player!")
            Exit Sub
        End If

        ' Ban the player with the reason.
        Call BanByServer(Index, txtBanReason.Text)
    Else
        ' Ban the player without the reason.
        Call BanByServer(Index, vbNullString)
    End If

    ' Hide the ban picture box.
    picBan.Visible = False
End Sub

Private Sub cmdServKick_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is currently playing.
    If Not IsPlaying(Index) Then Exit Sub

    ' Check if the administrator wants to send a kick message.
    If chkKickReason.Value = vbChecked Then
    
        ' Check if the kick message exists.
        If LenB(txtKickReason.Text) = 0 Then
            Call MsgBox("Please input a reason to kick this player!")
            Exit Sub
        End If

        ' Kick the player with the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been kicked by the server! Reason(" & txtWarpReason.Text & ")", WHITE)
        Call AlertMsg(Index, "You have been kicked by the server! Reason(" & txtKickReason.Text & ")")
    Else
        ' Kick the player without the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been kicked by the server!", WHITE)
        Call AlertMsg(Index, "You have been kicked by the server!")
    End If

    ' Hide the kick picture box.
    picKick.Visible = False
End Sub

Private Sub cmdServMsg_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is currently playing.
    If Not IsPlaying(Index) Then Exit Sub

    ' Send a personal message through the server.
    Call PlayerMsg(Index, "Server: " & txtPlayerMsg.Text, BRIGHTGREEN)

    ' Hide the message picture box.
    picMessage.Visible = False
End Sub

Private Sub cmdServWarp_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is currently playing.
    If Not IsPlaying(Index) Then Exit Sub

    ' Check if the administrator wants to send a warp message.
    If chkWarpReason.Value = vbChecked Then

        ' Check if the warp message exists.
        If LenB(txtWarpReason.Text) = 0 Then
            Call MsgBox("Please input a reason to warp this player!")
            Exit Sub
        End If

        ' Warp the player with the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been warped by the server! Reason(" & txtWarpReason.Text & ")", WHITE)
        Call PlayerWarp(Index, scrlWarpMap.Value, scrlWarpX.Value, scrlWarpY.Value)
    Else
        ' Warp the player without the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been warped by the server!", WHITE)
        Call PlayerWarp(Index, scrlWarpMap.Value, scrlWarpX.Value, scrlWarpY.Value)
    End If

    ' Hide the warp picture box.
    picWarp.Visible = False
End Sub

Private Sub cmdServJail_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is currently playing.
    If Not IsPlaying(Index) Then Exit Sub

    ' Check if the administrator wants to send a jail message.
    If chkJailReason.Value = vbChecked Then

        ' Check if the jail message exists.
        If LenB(txtJailReason.Text) = 0 Then
            Call MsgBox("Please input a reason to jail this player!")
            Exit Sub
        End If

        ' Jail the player with the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been jailed by the server! Reason(" & txtJailReason.Text & ")", WHITE)
        Call PlayerWarp(Index, scrlJailMap.Value, scrlJailX.Value, scrlJailY.Value)
    Else
        ' Jail the player without the reason.
        Call GlobalMsg(GetPlayerName(Index) & " has been jailed by the server!", WHITE)
        Call PlayerWarp(Index, scrlJailMap.Value, scrlJailX.Value, scrlJailY.Value)
    End If

    ' Hide the jail picture box.
    picJail.Visible = False
End Sub

Private Sub cmdSetTime_Click()
    Dim TimeH As Integer
    Dim TimeM As Integer
    Dim TimeS As Integer

    ' Get the time from the controls.
    TimeH = CLng(txtTimeH.Text)
    TimeM = CLng(txtTimeM.Text)
    TimeS = CLng(txtTimeS.Text)

    ' Check for a valid hour.
    If TimeH < 1 Or TimeH > 24 Then Exit Sub

    ' Check for a valid minute.
    If TimeM < 0 Or TimeM > 59 Then Exit Sub

    ' Check for a valid second.
    If TimeS < 0 Or TimeS > 59 Then Exit Sub

    ' Check for a valid time.
    If TimeH = 24 Then
        If TimeM > 0 Or TimeS > 0 Then
            Exit Sub
        End If
    End If

    ' Set the new time server-side.
    Hours = TimeH
    Minutes = TimeM
    Seconds = TimeS

    ' Set the new time client-side.
    Call SendGameClockToAll
End Sub

Private Sub cmdWarpCancel_Click()
    ' Hide the warp picture box.
    picWarp.Visible = False
End Sub

Private Sub Command68_Click()
    Dim TempSpeed As Long

    ' Get the game speed from the control.
    TempSpeed = CLng(GameTimeSpeed.Text)

    ' Check for valid input from the user.
    If TempSpeed < 0 Or TempSpeed > 59 Then
        Call MsgBox("Please enter a positive number less than 60.")
        Exit Sub
    End If

    ' Set the new game speed server-side.
    Gamespeed = TempSpeed

    ' Set the new game speed client-side.
    Call SendGameClockToAll
End Sub


Private Sub Command69_Click()
    If Not TimeDisable Then
        Gamespeed = 0
        GameTimeSpeed.Text = 0
        TimeDisable = True
        tmrDayNight.Enabled = False
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        GameTimeSpeed.Text = 1
        TimeDisable = False
        tmrDayNight.Enabled = True
        frmServer.Command69.Caption = "Disable Time"
    End If

    Call DisabledTime

    If Not TimeDisable Then
        SendGameClockToAll
    End If
End Sub

Private Sub Command70_Click()
    Script.Visible = True
End Sub

Private Sub Command71_Click()
    Script.Visible = False
End Sub

Private Sub Command72_Click()
    On Local Error Resume Next

    If SCRIPTING = 1 Then
        Call MyScript.SControl.ExecuteStatement(ServerScript.Text)
        ServerScript.Text = vbNullString
    Else
        Call MsgBox("Scripting is disabled. This action cannot be completed.")
    End If

    Script.Visible = False
End Sub

Private Sub Command73_Click()
    Dim I As Integer

    For I = 1 To MAX_PLAYERS
        If IsConnected(I) Then
            Call SendNewsTo(I)
        End If
    Next I
End Sub

Private Sub Form_Load()
    Hours = Rand(1, 24)
    Minutes = Rand(0, 59)
    Seconds = Rand(0, 59)

    Gamespeed = 1

    lblVer.Caption = "Build: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        lvUsers.GridLines = True
    Else
        lvUsers.GridLines = False
    End If
End Sub

Private Sub Command1_Click()
    If Not tmrShutdown.Enabled Then
        tmrShutdown.Enabled = True
    End If
    
    Command1.Enabled = False
End Sub

Private Sub Command12_Click()
    Dim Index As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerHP(Index) < GetPlayerMaxHP(Index) Then
                Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            End If
        End If
    Next Index

    Call GlobalMsg("The server has healed the wounded!", BRIGHTGREEN)
End Sub

Private Sub cmdKickPlayerReason_Click()
    If picKick.Visible Then
        picKick.Visible = False
    Else
        picKick.Visible = True
    End If
End Sub

Private Sub cmdBanPlayerReason_Click()
    If picBan.Visible Then
        picBan.Visible = False
    Else
        picBan.Visible = True
    End If
End Sub

Private Sub cmdJailPlayer_Click()
    If picJail.Visible Then
        picJail.Visible = False
    Else
        scrlJailMap.Max = MAX_MAPS
        scrlJailX.Max = MAX_MAPX
        scrlJailY.Max = MAX_MAPY

        picJail.Visible = True
    End If
End Sub

Private Sub cmdViewInfo_Click()
    Dim Index As Long

    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    If IsPlaying(Index) Then
        CharInfo(0).Caption = "Account: " & GetPlayerLogin(Index)
        CharInfo(1).Caption = "Character: " & GetPlayerName(Index)
        CharInfo(2).Caption = "Level: " & GetPlayerLevel(Index)
        CharInfo(3).Caption = "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index)
        CharInfo(4).Caption = "MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index)
        CharInfo(5).Caption = "SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index)
        CharInfo(6).Caption = "EXP: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index)
        CharInfo(7).Caption = "Access: " & GetPlayerAccess(Index)
        CharInfo(8).Caption = "PK: " & GetPlayerPK(Index)
        CharInfo(9).Caption = "Class: " & ClassData(GetPlayerClass(Index)).Name
        CharInfo(10).Caption = "Sprite: " & GetPlayerSprite(Index)
        CharInfo(11).Caption = "Sex: " & CStr(Player(Index).Char(Player(Index).CharNum).Sex)
        CharInfo(12).Caption = "Map: " & GetPlayerMap(Index)
        CharInfo(13).Caption = "Guild: " & GetPlayerGuild(Index)
        CharInfo(14).Caption = "Guild Access: " & GetPlayerGuildAccess(Index)
        CharInfo(15).Caption = "STR: " & GetPlayerSTR(Index)
        CharInfo(16).Caption = "DEF: " & GetPlayerDEF(Index)
        CharInfo(17).Caption = "Speed: " & GetPlayerSPEED(Index)
        CharInfo(18).Caption = "Magi: " & GetPlayerMAGI(Index)
        CharInfo(19).Caption = "Points: " & GetPlayerPOINTS(Index)
        CharInfo(20).Caption = "Index: " & Index
        picStats.Visible = True
    End If
End Sub

Private Sub cmdMsgPlayer_Click()
    If picMessage.Visible Then
        picMessage.Visible = False
    Else
        picMessage.Visible = True
    End If
End Sub

Private Sub cmdMutePlayer_Click()
    Dim Index As Long

    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    If IsPlaying(Index) Then
        Call PlayerMsg(Index, "You have been muted!", WHITE)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has been muted!", True)
        Player(Index).Mute = True
    End If
End Sub

Private Sub cmdUnmutePlayer_Click()
    Dim Index As Long

    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    If IsPlaying(Index) Then
        Call PlayerMsg(Index, "You have been unmuted!", WHITE)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has been unmuted!", True)
        Player(Index).Mute = False
    End If
End Sub

Private Sub cmdKillPlayer_Click()
    Dim Index As Long
    Dim ClassID As Long

    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    If IsPlaying(Index) Then
        Call SetPlayerHP(Index, 0)

        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
        Else
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).BootMap, Map(GetPlayerMap(Index)).BootX, Map(GetPlayerMap(Index)).BootY)
            Else
                If CLASSES = 0 Then
                    Call PlayerWarp(Index, START_MAP, START_X, START_Y)
                Else
                    ClassID = GetPlayerClass(Index)

                    Call PlayerWarp(Index, GetClassSpawnMap(ClassID), GetClassSpawnX(ClassID), GetClassSpawnY(ClassID))
                End If
            End If
        End If

        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))

        Call PlayerMsg(Index, "You have been killed by the server.", BRIGHTRED)
    End If
End Sub

Private Sub Command25_Click()
    If SCRIPTING = 1 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing

        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands

        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

        MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"

        Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
        Call AdminMsg("The scripts were reloaded by the server.", 15)
    End If
End Sub

Private Sub Command26_Click()
    If SCRIPTING = 0 Then
        ' Check for Main.txt
        If Not FileExists("\Scripts\Main.txt") Then
            Call MsgBox("The file 'Scripts\Main.txt' could not be found!", vbExclamation)
            Exit Sub
        End If

        SCRIPTING = 1

        PutVar FILE_DATAINI, "CONFIG", "Scripting", 1

        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands

        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

        MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"
        
        lblScriptOn.Caption = "Scripts: ON"
    End If
End Sub

Private Sub Command27_Click()
    If SCRIPTING = 1 Then
        SCRIPTING = 0

        Call PutVar(FILE_DATAINI, "CONFIG", "Scripting", 0)

        Set MyScript = Nothing
        Set clsScriptCommands = Nothing

        lblScriptOn.Caption = "Scripts: OFF"
    End If
End Sub

Private Sub Command28_Click()
    If FileExists("Editor.exe") Then
        Call Shell(App.Path & "\Editor.exe Scripts\Main.txt", vbNormalNoFocus)
    Else
        Call MsgBox("The eclipse editor cannot be found!", vbOKOnly, "Error")
    End If
End Sub

Private Sub Command29_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
End Sub

Private Sub cmdHealPlayer_Click()
    Dim Index As Long

    ' Get the selected index from the player list.
    Index = CLng(lvUsers.ListItems(lvUsers.SelectedItem.Index).Text)

    ' Check if the user is playing.
    If Not IsPlaying(Index) Then Exit Sub
    
    ' Check if the user has below their max HP.
    If Not GetPlayerHP(Index) = GetPlayerMaxHP(Index) Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))

        Call PlayerMsg(Index, "You have been healed by the server.", BRIGHTGREEN)
    End If
End Sub

Private Sub Command30_Click()
    Dim ClassNum As Long

    If FileExists("Editor.exe") Then
        ' We use Val() so we can catch empty results and strings, converting them to 0.
        ClassNum = Val(InputBox("Which class file would you like to open?", "Choose File", "0"))

        If FileExists("Classes\Class" & ClassNum & ".ini") Then
            Call Shell(App.Path & "\Editor.exe Classes\Class" & ClassNum & ".ini", vbNormalNoFocus)
        Else
            Call MsgBox("The file 'Classes\Class" & ClassNum & ".ini' could not be found.")
        End If
    Else
        Call MsgBox("The eclipse editor cannot be found!", vbOKOnly, "Error")
    End If
End Sub

Private Sub Command31_Click()
    Dim Index As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerAccess(Index) = 0 Then
                Call SetPlayerHP(Index, 0)
                Call PlayerMsg(Index, "You have been killed by the server!", BRIGHTRED)

                ' Warp player away
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
                Else
                    If Map(GetPlayerMap(Index)).BootMap > 0 Then
                        Call PlayerWarp(Index, Map(GetPlayerMap(Index)).BootMap, Map(GetPlayerMap(Index)).BootX, Map(GetPlayerMap(Index)).BootY)
                    Else
                        If CLASSES = 0 Then
                            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
                        Else
                            Call PlayerWarp(Index, GetClassSpawnMap(GetPlayerClass(Index)), GetClassSpawnX(GetPlayerClass(Index)), GetClassSpawnY(GetPlayerClass(Index)))
                        End If
                    End If
                End If

                Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
            End If
        End If
    Next Index
End Sub

Private Sub Command32_Click()
    scrlMM.Max = MAX_MAPS
    scrlMX.Max = MAX_MAPX
    scrlMY.Max = MAX_MAPY
    picWarpAll.Visible = True
End Sub

Private Sub Command33_Click()
    picExp.Visible = True
End Sub

Private Sub Command34_Click()
    Dim Index As Long
    Dim I As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerLevel(Index) >= MAX_LEVEL Then
                Call SetPlayerExp(Index, Experience(MAX_LEVEL))
            Else
                Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)

                I = Int(GetPlayerSPEED(Index) / 10)

                If I < 1 Then I = 1
                If I > 3 Then I = 3

                Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + I)

                If GetPlayerLevel(Index) >= MAX_LEVEL Then
                    Call SetPlayerExp(Index, Experience(MAX_LEVEL))
                End If
            End If

            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendPTS(Index)
        End If
    Next Index

    Call GlobalMsg("The server gave everyone a free level!", BRIGHTGREEN)
End Sub

Private Sub Command35_Click()
    Dim I As Long

    MapList.Clear

    For I = 1 To MAX_MAPS
        MapList.AddItem I & ": " & Map(I).Name
    Next I

    frmServer.MapList.Selected(0) = True
End Sub

Private Sub Command36_Click()
    Dim MapNum As Long
    Dim I As Long

    MapNum = MapList.ListIndex + 1

    MapInfo(0).Caption = "Map " & MapNum & " - " & Map(MapNum).Name
    MapInfo(1).Caption = "Revision: " & Map(MapNum).Revision
    MapInfo(2).Caption = "Moral: " & Map(MapNum).Moral
    MapInfo(3).Caption = "Up: " & Map(MapNum).Up
    MapInfo(4).Caption = "Down: " & Map(MapNum).Down
    MapInfo(5).Caption = "Left: " & Map(MapNum).Left
    MapInfo(6).Caption = "Right: " & Map(MapNum).Right
    MapInfo(7).Caption = "Music: " & Map(MapNum).music
    MapInfo(8).Caption = "BootMap: " & Map(MapNum).BootMap
    MapInfo(9).Caption = "BootX: " & Map(MapNum).BootX
    MapInfo(10).Caption = "BootY: " & Map(MapNum).BootY
    MapInfo(11).Caption = "Shop: " & Map(MapNum).Shop
    MapInfo(12).Caption = "Indoors: " & Map(MapNum).Indoors

    lstNPC.Clear

    For I = 1 To MAX_MAP_NPCS
        lstNPC.AddItem I & ": " & NPC(Map(MapNum).NPC(I)).Name
    Next I

    picMap.Visible = True
End Sub

Private Sub Command37_Click()
    Dim Index As Long
    Dim MapNum As Long
    Dim MapX As Long
    Dim MapY As Long

    MapNum = Int(scrlMM.Value)
    MapX = Int(scrlMX.Value)
    MapY = Int(scrlMY.Value)

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerAccess(Index) = 0 Then
                Call PlayerWarp(Index, MapNum, MapX, MapY)
            End If
        End If
    Next Index

    Call GlobalMsg("The server has warped everyone to map " & MapNum & ".", YELLOW)

    picWarpAll.Visible = False
End Sub

Private Sub Command38_Click()
    picWarpAll.Visible = False
End Sub

Private Sub Command39_Click()
    picExp.Visible = False
End Sub

Private Sub Command40_Click()
    Dim Index As Long
    Dim TotalExp As Long

    TotalExp = CLng(scrlExp.Value)

    If TotalExp > 0 Then
        For Index = 1 To MAX_PLAYERS
            If IsPlaying(Index) Then
                Call SetPlayerExp(Index, GetPlayerExp(Index) + TotalExp)
                Call CheckPlayerLevelUp(Index)
            End If
        Next Index

        Call GlobalMsg("The server gave everyone " & TotalExp & " experience!", BRIGHTGREEN)
    End If

    picExp.Visible = False
End Sub

Private Sub Command41_Click()
    picMap.Visible = False
End Sub

Private Sub cmdWarpPlayer_Click()
    If picWarp.Visible Then
        picWarp.Visible = False
    Else
        scrlWarpMap.Max = MAX_MAPS
        scrlWarpX.Max = MAX_MAPX
        scrlWarpY.Max = MAX_MAPY

        picWarp.Visible = True
    End If
End Sub

Private Sub Command5_Click()
    picCMsg.Visible = False
End Sub

Private Sub Command59_Click()
    picWeather.Visible = True
End Sub

Private Sub cmdSaveLogs_Click()
    Call SaveLogs
End Sub

Private Sub Command61_Click()
    picWeather.Visible = False
End Sub

Private Sub Command62_Click()
    WeatherType = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command63_Click()
    WeatherType = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Private Sub Command64_Click()
    WeatherType = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Private Sub Command65_Click()
    WeatherType = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Private Sub Command66_Click()
    Dim I As Long

    lvUsers.ListItems.Clear

    For I = 1 To MAX_PLAYERS
        Call ShowPLR(I)
    Next I
End Sub

Private Sub Command8_Click()
    picStats.Visible = False
End Sub

Private Sub Command9_Click()
    Dim Index As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerAccess(Index) = 0 Then
                Call GlobalMsg(GetPlayerName(Index) & " has been kicked by the server!", WHITE)
                Call AlertMsg(Index, "You have been kicked by the server!")
            End If
        End If
    Next Index
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide

        With NID
            .cbSize = Len(NID)
            .hWnd = Me.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon
            .szTip = Chr$(0)
            .uTimeout = 3000
            .dwState = NIS_SHAREDICON
            .dwInfoFlags = vbInformation
        End With
        
        Call Shell_NotifyIcon(NIM_ADD, NID)
    Else
        Call Shell_NotifyIcon(NIM_DELETE, NID)
    End If
End Sub

Private Sub Form_Terminate()
    Call SaveAllPlayersOnline
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveAllPlayersOnline
    Call DestroyServer
End Sub

Private Sub lblForum_Click()
    Call Shell("explorer http://www.touchofdeathforums.com/smf/index.php", vbNormalFocus)
End Sub

Private Sub lblGetIP_Click()
    Call Shell("explorer http://www.ipchicken.com", vbNormalFocus)
End Sub

Private Sub lblWebsite_Click()
    Call Shell("explorer http://freemmorpgmaker.com", vbNormalFocus)
End Sub

Private Sub lstTopics_Click()
    Dim FileName As String
    Dim hfile As Long

    txtTopic.Text = vbNullString

    FileName = lstTopics.ListIndex + 1 & ".txt"

    If FileExists("Guides\" & FileName) = True Then
        hfile = FreeFile

        Open App.Path & "\Guides\" & FileName For Input As #hfile
            txtTopic.Text = Input$(LOF(hfile), hfile)
        Close #hfile
    End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Value = vbChecked Then
        ServerLog = False
    Else
        ServerLog = True
    End If
End Sub

Private Sub PlayerTimer_Timer()
    If PlayerI <= MAX_PLAYERS Then
        If IsPlaying(PlayerI) Then
            Call SavePlayer(PlayerI)
        End If

        PlayerI = PlayerI + 1
    End If

    If PlayerI >= MAX_PLAYERS Then
        PlayerI = 1
        PlayerTimer.Enabled = False
        tmrPlayerSave.Enabled = True
    End If
End Sub

Private Sub scrlExp_Change()
    lblMassExp.Caption = "Experience: " & scrlExp.Value
End Sub

Private Sub scrlJailMap_Change()
    lblJailMap.Caption = "Map: " & scrlJailMap.Value
End Sub

Private Sub scrlMM_Change()
    lblMM.Caption = "Map: " & scrlMM.Value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.Value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.Value
End Sub

Private Sub scrlRainIntensity_Change()
    lblRainIntensity.Caption = "Intensity: " & scrlRainIntensity.Value
    WeatherLevel = scrlRainIntensity.Value

    Call SendWeatherToAll
End Sub

Private Sub scrlJailX_Change()
    lblJailX.Caption = "X: " & scrlJailX.Value
End Sub

Private Sub scrlJailY_Change()
    lblJailY.Caption = "Y: " & scrlJailY.Value
End Sub

Private Sub Timer1_Timer()
    Dim AMorPM As String
    Dim TempSeconds As Integer
    Dim PrintSeconds As String
    Dim PrintMinutes As String
    Dim PrintHours As Integer

    Seconds = Seconds + Gamespeed

    If Seconds > 59 Then
        Minutes = Minutes + 1
        Seconds = 0
    End If

    If Minutes > 59 Then
        Hours = Hours + 1
        Minutes = 0
    End If

    If Hours > 24 Then
        Hours = 1
    End If

    If Hours > 12 Then
        AMorPM = "PM"
        PrintHours = Hours - 12
    Else
        AMorPM = "AM"
        PrintHours = Hours
    End If

    If Hours = 24 Then
        AMorPM = "AM"
    End If

    TempSeconds = Seconds

    If Seconds > 9 Then
        PrintSeconds = TempSeconds
    Else
        PrintSeconds = "0" & CStr(Seconds)
    End If

    If Minutes > 9 Then
        PrintMinutes = Minutes
    Else
        PrintMinutes = "0" & CStr(Minutes)
    End If

    Label8.Caption = "Time: " & PrintHours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM

    If Hours < 7 Or Hours > 20 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
            Call SendTimeToAll
        End If
    Else
        If GameTime = TIME_NIGHT Then
            GameTime = TIME_DAY
            Call SendTimeToAll
        End If
    End If
 
    If Hours > 11 Then
        GameClock = Hours - 12 & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
    Else
        GameClock = Hours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
    End If

    ' Sync game clock every 10 minutes
    If Minutes Mod 10 = 0 Then
        Call SendGameClockToAll
    End If

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "TimedEvent " & Hours & "," & Minutes & "," & Seconds
    End If
End Sub

Private Sub tmrChatLogs_Timer()
    If frmServer.chkChat.Value = vbUnchecked Then
        CHATLOG_TIMER = 3600
        lblLogTime.Caption = "Chat Log Save Disabled!"
        Exit Sub
    End If

    If CHATLOG_TIMER < 1 Then
        CHATLOG_TIMER = 3600
    End If

    If CHATLOG_TIMER > 60 Then
        lblLogTime.Caption = "Chat Log Save In " & Int(CHATLOG_TIMER / 60) & " Minute(s)"
    Else
        lblLogTime.Caption = "Chat Log Save In " & Int(CHATLOG_TIMER) & " Second(s)"
    End If

    CHATLOG_TIMER = CHATLOG_TIMER - 1

    If CHATLOG_TIMER <= 0 Then
        Call TextAdd(txtText(0), "The chat logs were successfully saved!", True)
        Call SaveLogs
    End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrScriptedTimer_Timer()
    Call ScriptedTimer
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) <> 0 Then
            Call GlobalMsg(txtChat.Text, WHITE)
            Call TextAdd(frmServer.txtText(0), "Server: " & txtChat.Text, True)
            txtChat.Text = vbNullString
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub tmrShutdown_Timer()
    If SHUTDOWN_TIMER < 1 Then
        SHUTDOWN_TIMER = 30
    End If

    If SHUTDOWN_TIMER Mod 5 = 0 Or SHUTDOWN_TIMER <= 10 Then
        Call GlobalMsg("Server is shutting down in " & SHUTDOWN_TIMER & " second(s).", BRIGHTBLUE)
        Call TextAdd(frmServer.txtText(0), "Automated server shutdown in " & SHUTDOWN_TIMER & " second(s).", True)
    End If
    
    SHUTDOWN_TIMER = SHUTDOWN_TIMER - 1
    
    If SHUTDOWN_TIMER < 1 Then
        Call GlobalMsg("Server has been shutdown.", BRIGHTRED)
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub txtText_GotFocus(Index As Integer)
    txtChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
