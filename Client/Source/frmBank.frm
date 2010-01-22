VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   0  'User
   ScaleWidth      =   399
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4320
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4320
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblMsg 
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
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   5520
   End
   Begin VB.Label lblBankWithdraw 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3090
      TabIndex        =   4
      Top             =   4800
      Width           =   2685
   End
   Begin VB.Label lblBankClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label lblBankDeposit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   315
      TabIndex        =   2
      Top             =   4800
      Width           =   2580
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblBankDeposit_Click()
    Dim InvNum As Long
    Dim StackAmt As String

    On Local Error Resume Next

    ' Cache the inventory number from the list box control.
    InvNum = lstInventory.ListIndex + 1

    ' Check if it's a valid item number.
    If GetPlayerInvItemNum(MyIndex, InvNum) < 0 Or GetPlayerInvItemNum(MyIndex, InvNum) > MAX_ITEMS Then Exit Sub

    ' Check if the item is stackable.
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
        
        ' Ask for the amount the player wants to deposit.
        StackAmt = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmBank.Left, frmBank.Top)

        ' Check if the string is numeric characters.
        If IsNumeric(StackAmt) Then
            Call SendData(POut.BankDeposit & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & StackAmt & END_CHAR)
        End If
    Else
        Call SendData(POut.BankDeposit & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & 0 & END_CHAR)
    End If
End Sub

Private Sub lblBankWithdraw_Click()
    Dim BankNum As Long
    Dim StackAmt As String

    On Local Error Resume Next

    ' Cache the inventory number from the list box control.
    BankNum = lstBank.ListIndex + 1

    ' Check if it's a valid item number.
    If GetPlayerInvItemNum(MyIndex, BankNum) < 0 Or GetPlayerInvItemNum(MyIndex, BankNum) > MAX_ITEMS Then Exit Sub

    ' Check if the item is stackable.
    If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Stackable = 1 Then

        ' Ask for the amount the player wants to withdraw.
        StackAmt = InputBox("How much " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to withdraw?", "Withdraw " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name), 0, frmBank.Left, frmBank.Top)

        ' Check if the string is numeric characters.
        If IsNumeric(StackAmt) Then
            Call SendData(POut.BankWithdraw & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & StackAmt & END_CHAR)
        End If
    Else
        Call SendData(POut.BankWithdraw & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & 0 & END_CHAR)
    End If
End Sub

Private Sub lblBankClose_Click()
    ' Unload the bank form.
    Unload Me
End Sub

Private Sub Form_Load()
    ' Load the bank form GUI.
    Call GUI_PictureLoad(frmBank, "GUI\Bank")
End Sub
