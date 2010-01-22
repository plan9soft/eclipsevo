VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Character"
   ClientHeight    =   5985
   ClientLeft      =   135
   ClientTop       =   315
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   0  'User
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   4440
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   36
         Top             =   720
         Width           =   480
         Begin VB.PictureBox picBody 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   2
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   37
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   34
         Top             =   360
         Width           =   480
         Begin VB.PictureBox picBody 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   35
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   0
         Width           =   480
         Begin VB.PictureBox picBody 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   33
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.HScrollBar hsLegs 
      Height          =   255
      Left            =   3960
      Max             =   200
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar hsBody 
      Height          =   255
      Left            =   3960
      Max             =   200
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar hsHead 
      Height          =   255
      Left            =   3960
      Max             =   200
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   18
      Top             =   3540
      Width           =   2040
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   17
      Top             =   3240
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4440
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   7
      Top             =   1800
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   8
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblClassDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Class Description.>"
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
      Height          =   975
      Left            =   3360
      TabIndex        =   38
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Legs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   30
      Top             =   3720
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Head:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   24
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   22
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   20
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   13
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   12
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   2805
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   5565
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastCharUpdate As Long
Private LastAnimation As Byte

Private Sub cmbClass_Click()
    Dim ClassNum As Byte
    
    For ClassNum = 0 To Max_Classes
        If Trim(Class(ClassNum).Name) = cmbClass.List(cmbClass.ListIndex) Then
            Exit For
        End If
    Next

    lblHP.Caption = CStr(Class(ClassNum).HP)
    lblMP.Caption = CStr(Class(ClassNum).MP)
    lblSP.Caption = CStr(Class(ClassNum).SP)

    lblSTR.Caption = CStr(Class(ClassNum).STR)
    lblDEF.Caption = CStr(Class(ClassNum).DEF)
    lblSPEED.Caption = CStr(Class(ClassNum).Speed)
    lblMAGI.Caption = CStr(Class(ClassNum).MAGI)

    lblClassDesc.Caption = Class(ClassNum).desc
End Sub

Private Sub Form_Activate()
    Dim sRECT As RECT
    Dim dRECT As RECT

    If CUSTOM_PLAYERS = 0 Then
        NewChar_FormLoop
    Else
        frmNewChar.Picture4.Visible = False
        frmNewChar.hsHead.Visible = True
        frmNewChar.hsBody.Visible = True
        frmNewChar.hsLegs.Visible = True
        frmNewChar.Label14.Visible = True
        frmNewChar.Label11.Visible = True
        frmNewChar.Label12.Visible = True
        frmNewChar.picBackEnd.Visible = True
        
        dRECT = CreateRECT(0, 32, 0, 32)

        If SpriteSize = 1 Then
            sRECT = CreateRECT(hsHead.Value * 64 + 15, 32, 32 * 3, 32)
            Call DD_PlayerHead.BltToDC(picBody(0).hDC, sRECT, dRECT)
        Else
            sRECT = CreateRECT(hsHead.Value * 32, 32, 32 * 3, 32)
            Call DD_PlayerHead.BltToDC(picBody(0).hDC, sRECT, dRECT)
        End If

        If SpriteSize = 1 Then
            sRECT = CreateRECT(hsBody.Value * 64 + 35, 32, 32 * 3, 32)
            Call DD_PlayerBody.BltToDC(picBody(1).hDC, sRECT, dRECT)
        Else
            sRECT = CreateRECT(hsBody.Value * 32, 32, 32 * 3, 32)
            Call DD_PlayerBody.BltToDC(picBody(1).hDC, sRECT, dRECT)
        End If

        If SpriteSize = 1 Then
            sRECT = CreateRECT(hsLegs.Value * 64 + 35, 32, 32 * 3, 32)
            Call DD_PlayerLegs.BltToDC(picBody(2).hDC, sRECT, dRECT)
        Else
            sRECT = CreateRECT(hsLegs.Value * 32, 32, 32 * 3, 32)
            Call DD_PlayerLegs.BltToDC(picBody(2).hDC, sRECT, dRECT)
        End If
        
        picBody(0).Refresh
        picBody(1).Refresh
        picBody(2).Refresh
    End If
End Sub

Private Sub hsHead_Change()
    Dim sRECT As RECT
    Dim dRECT As RECT

    dRECT = CreateRECT(0, 32, 0, 32)
    
    If SpriteSize = 1 Then
        sRECT = CreateRECT(hsHead.Value * 64 + 15, 32, 32 * 3, 32)
        Call DD_PlayerHead.BltToDC(picBody(0).hDC, sRECT, dRECT)
    Else
        sRECT = CreateRECT(hsHead.Value * 32, 32, 32 * 3, 32)
        Call DD_PlayerHead.BltToDC(picBody(0).hDC, sRECT, dRECT)
    End If

    picBody(0).Refresh
End Sub

Private Sub hsBody_Change()
    Dim sRECT As RECT
    Dim dRECT As RECT

    dRECT = CreateRECT(0, 32, 0, 32)

    If SpriteSize = 1 Then
        sRECT = CreateRECT(hsBody.Value * 64 + 35, 32, 32 * 3, 32)
        Call DD_PlayerBody.BltToDC(picBody(1).hDC, sRECT, dRECT)
    Else
        sRECT = CreateRECT(hsBody.Value * 32, 32, 32 * 3, 32)
        Call DD_PlayerBody.BltToDC(picBody(1).hDC, sRECT, dRECT)
    End If

    picBody(1).Refresh
End Sub

Private Sub hsLegs_Change()
    Dim sRECT As RECT
    Dim dRECT As RECT

    dRECT = CreateRECT(0, 32, 0, 32)

    If SpriteSize = 1 Then
        sRECT = CreateRECT(hsLegs.Value * 64 + 35, 32, 32 * 3, 32)
        Call DD_PlayerLegs.BltToDC(picBody(2).hDC, sRECT, dRECT)
    Else
        sRECT = CreateRECT(hsLegs.Value * 32, 32, 32 * 3, 32)
        Call DD_PlayerLegs.BltToDC(picBody(2).hDC, sRECT, dRECT)
    End If

    picBody(2).Refresh
End Sub

Private Sub picAddChar_Click()
    Dim Msg As String
    Dim i As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)

        If Len(Trim$(txtName.Text)) < 3 Then
            MsgBox "Character name must be at least three characters in length."
            Exit Sub
        End If

        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 255 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = vbNullString
                Exit Sub
            End If
        Next i

        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call GUI_PictureLoad(frmNewChar, "GUI\NewCharacter")
End Sub

Private Sub NewChar_FormLoop()
    While Me.Visible
        If GetTickCount() > LastCharUpdate + 250 Then
            Call NewChar_DrawSprite
            Call NewChar_UpdateLastAnimation
            LastCharUpdate = GetTickCount()
        End If
        DoEvents
    Wend
End Sub

Private Sub NewChar_UpdateLastAnimation()
    LastAnimation = LastAnimation + 1
    
    If LastAnimation > 11 Then LastAnimation = 0
End Sub

Private Sub NewChar_DrawSprite()
    Dim sRECT As RECT
    Dim dRECT As RECT

    ' Check to make sure a class is selected.
    If cmbClass.ListIndex = -1 Then Exit Sub

    ' Check if custom player graphics are enabled.
    If CUSTOM_PLAYERS = 0 Then
    
        ' Check for the sprite size (either 32x32 or 64x32).
        If SpriteSize = 1 Then
        
            ' Check if we're dealing with a male or female.
            If optMale.Value Then
                sRECT = CreateRECT(Class(cmbClass.ListIndex).MaleSprite * 64, 64, LastAnimation * PIC_X, 32)
                dRECT = CreateRECT(0, 64, 0, 32)

                Call DD_SpriteSurf.BltToDC(Picpic.hDC, sRECT, dRECT)
            Else
                sRECT = CreateRECT(Class(cmbClass.ListIndex).FemaleSprite * 64, 64, LastAnimation * PIC_X, 32)
                dRECT = CreateRECT(0, 64, 0, 32)

                Call DD_SpriteSurf.BltToDC(Picpic.hDC, sRECT, dRECT)
            End If
        Else
            ' Check if we're dealing with a male or female.
            If optMale.Value Then
                sRECT = CreateRECT(Class(cmbClass.ListIndex).MaleSprite * 32, 32, LastAnimation * PIC_X, 32)
                dRECT = CreateRECT(0, 32, 0, 32)

                Call DD_SpriteSurf.BltToDC(Picpic.hDC, sRECT, dRECT)
            Else
                sRECT = CreateRECT(Class(cmbClass.ListIndex).FemaleSprite * 32, 32, LastAnimation * PIC_X, 32)
                dRECT = CreateRECT(0, 32, 0, 32)

                Call DD_SpriteSurf.BltToDC(Picpic.hDC, sRECT, dRECT)
            End If
        End If

        ' Refresh the picture box.
        Picpic.Refresh
    End If
End Sub
