VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RGB Color Editor"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin MSComctlLib.Slider sldrBlue 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider sldrGreen 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider sldrRed 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin VB.Shape shpRGB 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ' Unload the form from memory.
    Unload Me
End Sub

Private Sub cmdOkay_Click()
    ' Set the global color variable since we use two forms.
    NEWS_RED = CByte(sldrRed.Value)
    NEWS_GREEN = CByte(sldrGreen.Value)
    NEWS_BLUE = CByte(sldrBlue.Value)

    ' Display the new colors as the fore color.
    frmNews.txtTitle.ForeColor = RGB(NEWS_RED, NEWS_GREEN, NEWS_BLUE)
    frmNews.txtBody.ForeColor = RGB(NEWS_RED, NEWS_GREEN, NEWS_BLUE)

    ' Unload the form from memory.
    Unload Me
End Sub

Private Sub Form_Load()
    ' Set the sliders to the global variables.
    sldrRed.Value = CInt(NEWS_RED)
    sldrGreen.Value = CInt(NEWS_GREEN)
    sldrBlue.Value = CInt(NEWS_BLUE)
End Sub

Private Sub sldrRed_Click()
    ' Update the shape with the new red blend.
    shpRGB.BackColor = RGB(sldrRed.Value, sldrGreen.Value, sldrBlue.Value)
End Sub

Private Sub sldrGreen_Click()
    ' Update the shape with the new green blend.
    shpRGB.BackColor = RGB(sldrRed.Value, sldrGreen.Value, sldrBlue.Value)
End Sub

Private Sub sldrBlue_Click()
    ' Update the shape with the new blue blend.
    shpRGB.BackColor = RGB(sldrRed.Value, sldrGreen.Value, sldrBlue.Value)
End Sub
