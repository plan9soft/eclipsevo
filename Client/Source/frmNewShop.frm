VERSION 5.00
Begin VB.Form frmNewShop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shop"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picItemInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2775
      Left            =   90
      ScaleHeight     =   2745
      ScaleWidth      =   3450
      TabIndex        =   12
      Top             =   2520
      Width           =   3480
      Begin VB.Label lblNamePrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   73
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblMagicReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   1560
         Width           =   1700
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
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
         Height          =   735
         Left            =   30
         TabIndex        =   21
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblSpdBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "SpdBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   2400
         Width           =   1700
      End
      Begin VB.Label lblMagiBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "MagiBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   19
         Top             =   2400
         Width           =   1700
      End
      Begin VB.Label lblDefBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "DefBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   2160
         Width           =   1700
      End
      Begin VB.Label lblAddStr 
         BackStyle       =   0  'Transparent
         Caption         =   "StrBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   17
         Top             =   2160
         Width           =   1700
      End
      Begin VB.Label lblVital 
         BackStyle       =   0  'Transparent
         Caption         =   "Vital Mod:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblSpdReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   15
         Top             =   1560
         Width           =   1700
      End
      Begin VB.Label lblDefReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   1700
      End
      Begin VB.Label lblStrReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   1320
         Width           =   1700
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   19
      Left            =   3000
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   70
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   19
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   71
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   19
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   72
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   18
      Left            =   2280
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   67
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   18
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   68
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   18
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   69
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   17
      Left            =   1560
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   64
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   17
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   65
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   17
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   66
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   16
      Left            =   840
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   61
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   16
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   62
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   16
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   63
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   15
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   58
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   15
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   59
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   15
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   60
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   14
      Left            =   3000
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   55
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   14
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   56
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   14
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   57
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   13
      Left            =   2280
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   13
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   53
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   13
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   54
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   12
      Left            =   1560
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   49
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   12
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   50
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   12
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   51
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   11
      Left            =   840
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   46
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   11
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   47
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   11
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   48
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   10
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   10
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   44
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   10
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   45
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   9
      Left            =   3000
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   9
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   41
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   9
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   42
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   8
      Left            =   2280
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   8
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   38
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   8
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   39
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   7
      Left            =   1560
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   7
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   7
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   36
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   6
      Left            =   840
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   6
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   6
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
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   840
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   29
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
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
            TabIndex        =   30
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   5
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   5
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   5
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   27
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   4
      Left            =   3000
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   4
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   4
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   10
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   2280
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   3
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   3
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   7
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   1560
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
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
            TabIndex        =   4
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
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
            TabIndex        =   11
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Label lblSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sell Items"
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
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fix Items"
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
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmNewShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public shopNum As Integer
Public fixItems As Boolean 'Is the shop fixing items?
Public SellItems As Boolean 'Is the shop selling items?

' Loads shop data into the form for the first time.
Public Sub loadShop(ByVal sNum As Integer)
    Dim I As Byte
    Dim shopCurrency As String
    Dim curItem As ShopItemRec
    Dim amtTxt As String
    On Error GoTo loadShop_Error
    
    Me.Visible = False
    shopNum = sNum

    Me.Caption = Shop(sNum).Name

    ' Check if this shop fixes items
    If Shop(sNum).FixesItems = YES Then
        lblFix.Visible = True
    Else
        lblFix.Visible = False
    End If

    ' Check if this shop buys back items
    If Shop(sNum).BuysItems = YES Then
        lblSell.Visible = True
    Else
        lblSell.Visible = False
    End If

    ' Set it not to fix item mode by default
    fixItems = False
       
    shopCurrency = Trim$(Item(Shop(shopNum).currencyItem).Name)

    For I = 1 To MAX_SHOP_ITEMS
        curItem = Shop(shopNum).ShopItem(I)
        If curItem.ItemNum = 0 Then
            Exit For
        Else
            imgBox(I - 1).Visible = True
            Me.iconn(I - 1).Cls

            Call BltIcon(I - 1)
        End If

    Next I

    I = (I - 2) \ 5

    picItemInfo.Top = 168 - 40 * (3 - I)
    lblFix.Top = 360 - 40 * (3 - I)
    lblSell.Top = 360 - 40 * (3 - I)
    Me.Height = 6060 - 600 * (3 - I)
    
    If Not fixItems And Not SellItems Then
        Me.Height = Me.Height - 300
    End If
    
    Me.Visible = True
    
    On Error GoTo 0
    Exit Sub

loadShop_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadShop of Form frmNewShop"
    If MsgBox("Could not load shop.", vbRetryCancel) = vbRetry Then
        loadShop sNum
    Else
        frmNewShop.Visible = False
    End If
    Exit Sub
End Sub

Public Sub BuyItem(shopItemIndex As Integer)
    Dim BuyItem As Long

    ' Prompt the user to make sure it's the right item.
    BuyItem = MsgBox("Are you SURE you want to buy this item?", vbYesNo)
    Call HideItemInfo
    
    ' Send the buy packet if they choose yes.
    If BuyItem = vbYes Then
        Call SendData(POut.BuyItem & SEP_CHAR & shopNum & SEP_CHAR & shopItemIndex & END_CHAR)
    End If
End Sub

Public Sub FixItem(ByVal Item As Integer)
    Call SendData(POut.FixItem & SEP_CHAR & snumber & SEP_CHAR & Item & END_CHAR)
End Sub

Public Sub Buyback(ByVal Item As Integer, ByVal slot As Integer, Optional ByVal AMT As Integer = 1)
    Call SendData(POut.SellItem & SEP_CHAR & shopNum & SEP_CHAR & Item & SEP_CHAR & slot & SEP_CHAR & AMT & END_CHAR)
End Sub

Private Sub BltIcon(ByVal IconNum As Integer)
    Dim sRECT As DXVBLib.RECT
    Dim dRECT As DXVBLib.RECT
    Dim ItemNum As Long
    ItemNum = Shop(shopNum).ShopItem(IconNum + 1).ItemNum
    
    sRECT.Top = Int(Item(ItemNum).Pic / 6) * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = (Item(ItemNum).Pic - Int(Item(ItemNum).Pic / 6) * 6) * PIC_X
    sRECT.Right = sRECT.Left + PIC_X

    dRECT.Top = 0
    dRECT.Bottom = dRECT.Top + PIC_Y
    dRECT.Left = 0
    dRECT.Right = dRECT.Left + PIC_X

    Call DD_ItemSurf.BltToDC(iconn(IconNum).hDC, sRECT, dRECT)
    
    iconn(IconNum).Refresh
End Sub

Private Sub ShowItemInfo(CurShopItem As ShopItemRec)
    Dim ItemNum As Integer
    
    ItemNum = CurShopItem.ItemNum
    
    'Update name, price, and description
    lblDesc.Caption = Item(ItemNum).desc
    lblNamePrice.Caption = Trim(Item(ItemNum).Name) & " - " & CStr(CurShopItem.Price) & " " & Trim(CStr(Item(Shop(shopNum).currencyItem).Name))
    
    ' Update the strength requirement.
    If Item(ItemNum).StrReq > 0 Then
        lblStrReq.Caption = STAT1 & " Req: " & Item(ItemNum).StrReq
    Else
        lblStrReq.Caption = STAT1 & " Req: None"
    End If

    ' Update the defense requirement.
    If Item(ItemNum).DefReq > 0 Then
        lblDefReq.Caption = STAT2 & " Req: " & Item(ItemNum).DefReq
    Else
        lblDefReq.Caption = STAT2 & " Req: None"
    End If

    ' Update the magic requirement.
    If Item(ItemNum).MagicReq > 0 Then
        lblMagicReq.Caption = STAT3 & " Req: " & Item(ItemNum).MagicReq
    Else
        lblMagicReq.Caption = STAT3 & " Req: None"
    End If

    ' Update the speed requirement.
    If Item(ItemNum).SpeedReq > 0 Then
        lblSpdReq.Caption = STAT4 & " Req: " & Item(ItemNum).SpeedReq
    Else
        lblSpdReq.Caption = STAT4 & " Req: None"
    End If

    ' Update the attack, defense, or value.
    If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
        lblVital.Caption = "Attack: " & Item(ItemNum).Data2
    ElseIf Item(ItemNum).Type >= ITEM_TYPE_ARMOR And Item(ItemNum).Type <= ITEM_TYPE_LEGS Then
        lblVital.Caption = "Defense: " & Item(ItemNum).Data2
    ElseIf Item(ItemNum).Type >= ITEM_TYPE_POTIONADDHP And Item(ItemNum).Type <= ITEM_TYPE_POTIONSUBSP Then
        lblVital.Caption = "Value: " & Item(ItemNum).Data2
    Else
        lblVital.Caption = vbNullString
    End If

    ' Update the amount of strength this adds.
    If Item(ItemNum).AddSTR > 0 Then
        lblAddStr.Caption = STAT1 & " Bonus: " & Item(ItemNum).AddSTR
    Else
        lblAddStr.Caption = STAT1 & " Bonus: None"
    End If

    ' Update the amount of defense this adds.
    If Item(ItemNum).AddDEF > 0 Then
        lblDefBonus.Caption = STAT2 & " Bonus: " & Item(ItemNum).AddDEF
    Else
        lblDefBonus.Caption = STAT2 & " Bonus: None"
    End If

    ' Update the amount of magic this adds.
    If Item(ItemNum).AddMAGI > 0 Then
        lblMagiBonus.Caption = STAT3 & " Bonus: " & Item(ItemNum).AddMAGI
    Else
        lblMagiBonus.Caption = STAT3 & " Bonus: None"
    End If

    ' Update the amount of speed this adds.
    If Item(ItemNum).AddSpeed > 0 Then
        lblSpdBonus.Caption = STAT4 & " Bonus: " & Item(ItemNum).AddSpeed
    Else
        lblSpdBonus.Caption = STAT4 & " Bonus: None"
    End If

    ' Display the item information.
    picItemInfo.Visible = True
End Sub

Private Sub HideItemInfo()
    picItemInfo.Visible = False
End Sub

Private Sub lblFix_Click()
    frmFixItem.Show vbModal
End Sub

Private Sub lblSell_Click()
    frmSellItem.Visible = True
End Sub

Private Sub Form_Load()
    Call GUI_PictureLoad(frmNewShop, "GUI\Shop")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call HideItemInfo
End Sub

' Buy item
Private Sub imgBox_Click(Index As Integer)
    Call BuyItem(Index + 1)
End Sub

' Buy item
Private Sub iconn_Click(Index As Integer)
    Call BuyItem(Index + 1)
End Sub

Private Sub iconn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    ShowItemInfo Shop(shopNum).ShopItem(Index + 1)
End Sub
