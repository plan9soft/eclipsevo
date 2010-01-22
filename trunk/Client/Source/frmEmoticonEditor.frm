VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "frmEmoticonEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2265
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3995
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Emoction"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmoticon"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlEmoticon"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCommand"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCommand 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "/"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   120
         Max             =   1000
         TabIndex        =   3
         Top             =   360
         Value           =   1
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2400
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   1
         Top             =   720
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   4
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picEmoticons 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   6
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Command :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblEmoticon 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "Emoticon :"
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
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Dim i As Long
    Dim Cmd As String

    ' Get the string from the control.
    Cmd = Trim$(txtCommand.Text)

    ' Check if the command name is blank.
    If Cmd = "/" Then
        Call MsgBox("Please enter a unique command name.")
        Exit Sub
    End If

    ' Loop for any possible errors.
    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) = Cmd Then
            If i <> EditorIndex - 1 Then
                Call MsgBox("There already is a '" & Cmd & "' command!")
                Exit Sub
            End If
        End If
    Next i

    ' Send the emoticon data to the server.
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' Define the source rectangle.
    sRECT.Top = scrlEmoticon.Value * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + PIC_X

    ' Define the destination rectangle.
    dRECT.Top = 0
    dRECT.Bottom = dRECT.Top + PIC_Y
    dRECT.Left = 0
    dRECT.Right = dRECT.Left + PIC_X

    ' Draw the emoticon to the picture.
    Call DD_EmoticonSurf.BltToDC(picEmoticons.hDC, sRECT, dRECT)

    ' Refresh the picture.
    picEmoticons.Refresh
End Sub

Private Sub scrlEmoticon_Change()
    Dim sRECT As DxVBLib.RECT
    Dim dRECT As DxVBLib.RECT

    ' Update the label.
    lblEmoticon.Caption = CStr(scrlEmoticon.Value)

    ' Define the source rectangle.
    sRECT.Top = scrlEmoticon.Value * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + PIC_X

    ' Define the destination rectangle.
    dRECT.Top = 0
    dRECT.Bottom = dRECT.Top + PIC_Y
    dRECT.Left = 0
    dRECT.Right = dRECT.Left + PIC_X

    ' Draw the emoticon to the picture.
    Call DD_EmoticonSurf.BltToDC(picEmoticons.hDC, sRECT, dRECT)

    ' Refresh the picture.
    picEmoticons.Refresh
End Sub

Private Sub txtCommand_Change()
    Dim Cmd As String
    
    ' Get the text from the control.
    Cmd = txtCommand.Text

    ' Check if the far left character is a slash.
    If Left$(Cmd, 1) <> "/" Then
        txtCommand.Text = "/" & Cmd
    End If

    ' Set the marker at the end.
    txtCommand.SelStart = Len(txtCommand.Text)
End Sub
