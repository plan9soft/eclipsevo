VERSION 5.00
Begin VB.Form frmNews 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "News Editor"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdColor 
      Caption         =   "Change Color"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtBody 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblNews 
      Caption         =   "News Content:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCaption 
      Caption         =   "News Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    ' Write the news title and body to the file.
    Call PutVar(App.Path & "\News.ini", "DATA", "NewsTitle", txtTitle.Text)
    Call PutVar(App.Path & "\News.ini", "DATA", "NewsBody", txtBody.Text)

    ' Write the RGB color values to the file.
    Call PutVar(App.Path & "\News.ini", "COLOR", "Red", CStr(NEWS_RED))
    Call PutVar(App.Path & "\News.ini", "COLOR", "Green", CStr(NEWS_GREEN))
    Call PutVar(App.Path & "\News.ini", "COLOR", "Blue", CStr(NEWS_BLUE))

    ' Unload this form from memory.
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ' Unload this form from memory.
    Unload Me
End Sub

Private Sub cmdColor_Click()
    ' Display the RGB color form.
    frmColor.Visible = True
End Sub

Private Sub Form_Load()
    ' Read the news title and body from the file.
    txtTitle.Text = GetVar(App.Path & "\News.ini", "DATA", "NewsTitle")
    txtBody.Text = GetVar(App.Path & "\News.ini", "DATA", "NewsBody")

    ' Read the news RGB color values from the file.
    NEWS_RED = CByte(GetVar(App.Path & "\News.ini", "COLOR", "Red"))
    NEWS_GREEN = CByte(GetVar(App.Path & "\News.ini", "COLOR", "Green"))
    NEWS_BLUE = CByte(GetVar(App.Path & "\News.ini", "COLOR", "Blue"))

    ' Set the background for the news title and body.
    txtTitle.ForeColor = RGB(NEWS_RED, NEWS_GREEN, NEWS_BLUE)
    txtBody.ForeColor = RGB(NEWS_RED, NEWS_GREEN, NEWS_BLUE)
End Sub
