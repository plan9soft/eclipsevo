VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Server"
   ClientHeight    =   1200
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6375
   ClipControls    =   0   'False
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   ScaleHeight     =   1200
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing..."
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
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Display the loader form.
    Me.Show

    ' Initiate the server and start it.
    Call InitServer
End Sub
