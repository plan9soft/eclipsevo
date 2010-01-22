VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmPlayerTrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Items2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   3120
      TabIndex        =   2
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ListBox Items1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ListBox PlayerInv1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   3120
      TabIndex        =   9
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblAddItem 
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
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblRemoveItem 
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
      TabIndex        =   6
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label lblQuitTrade 
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
      Top             =   5520
      Width           =   5565
   End
   Begin VB.Label Command2 
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
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblAcceptTrade 
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
      Left            =   3120
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Load the trading form GUI.
    Call GUI_PictureLoad(frmPlayerTrade, "GUI\Trade")
End Sub

Private Sub lblAddItem_Click()
    Dim InvNum As Long
    Dim TradeIndex As Long
    Dim TradeAmt As String

    ' Get the index of the item they want to trade.
    InvNum = PlayerInv1.ListIndex + 1

    ' Check for a valid item bounds.
    If GetPlayerInvItemNum(MyIndex, InvNum) < 0 Or GetPlayerInvItemNum(MyIndex, InvNum) > MAX_ITEMS Then Exit Sub

    ' Check if the item is being worn.
    If ItemIsEquipped(MyIndex, GetPlayerInvItemNum(MyIndex, InvNum)) Then
        Call AddText("You cannot trade worn items!", BRIGHTRED)
        Exit Sub
    End If

    ' Check if the item is already being traded.
    If TradeItemExists(InvNum) Then
        Call AddText("You cannot trade an item already being traded!", BRIGHTRED)
        Exit Sub
    End If

    ' Get an available trade slot.
    TradeIndex = FindOpenTradeSlot()

    ' Check if there's an available trade slot.
    If TradeIndex = -1 Then
        Call AddText("The trade window is full.", WHITE)
        Exit Sub
    End If

    ' Determine if we want to trade stackables or not.
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
        ' Get the amount of the items the player want to trade.
        TradeAmt = InputBox("How many " & Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name & "'s would you like to trade?", "Trade")

        ' Check if the string is numeric.
        If Not IsNumeric(TradeAmt) Then Exit Sub

        ' Update the player trade inventory that the item has been added
        ' to the trade window successfully.
        PlayerInv1.List(InvNum - 1) = PlayerInv1.Text & " **"

        ' Update the trade window to display the item and amount being traded.
        Items1.List(TradeIndex - 1) = CStr(TradeIndex) & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name & " [" & TradeAmt & "]")

        ' Add the item to the trade buffer.
        Trading(TradeIndex).InvNum = InvNum
        Trading(TradeIndex).InvName = Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
        Trading(TradeIndex).InvAmt = CLng(TradeAmt)
    Else
        ' Update the player trade inventory that the item has been added
        ' to the trade window successfully.
        PlayerInv1.List(InvNum - 1) = PlayerInv1.Text & " **"

        ' Update the trade window to display the item and amount being traded.
        Items1.List(TradeIndex - 1) = CStr(TradeIndex) & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)

        ' Add the item to the trade buffer.
        Trading(TradeIndex).InvNum = InvNum
        Trading(TradeIndex).InvName = Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
        Trading(TradeIndex).InvAmt = 0
    End If

    ' Update the server and trade client with the item update.
    Call SendData(POut.UpdateTradeInventory & SEP_CHAR & TradeIndex & SEP_CHAR & Trading(TradeIndex).InvNum & SEP_CHAR & Trading(TradeIndex).InvName & SEP_CHAR & Trading(TradeIndex).InvAmt & END_CHAR)
End Sub

Private Sub lblRemoveItem_Click()
    Dim TradeIndex As Long

    ' Get the index of the item they want to remove.
    TradeIndex = Items1.ListIndex + 1

    ' Check if an item exists in that trade slot.
    If Trading(TradeIndex).InvNum = 0 Then Exit Sub

    ' Update the trade window to display the item and amount being traded.
    PlayerInv1.List(Trading(TradeIndex).InvNum - 1) = Mid$(Trim$(PlayerInv1.List(Trading(TradeIndex).InvNum - 1)), 1, Len(PlayerInv1.List(Trading(TradeIndex).InvNum - 1)) - 3)

    ' Update the player trade inventory that the item has been removed
    ' from the trade window successfully.
    Items1.List(TradeIndex - 1) = CStr(TradeIndex) & ": <Nothing>"

    ' Remove the item to the trade buffer.
    Trading(TradeIndex).InvNum = 0
    Trading(TradeIndex).InvName = vbNullString
    Trading(TradeIndex).InvAmt = 0

    ' Update the server and trade client with the item update.
    Call SendData(POut.UpdateTradeInventory & SEP_CHAR & TradeIndex & SEP_CHAR & 0 & SEP_CHAR & vbNullString & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub lblAcceptTrade_Click()
    ' Swap the items and complete the trade.
    Call SendData(POut.SwapItems & END_CHAR)
End Sub

Private Sub lblQuitTrade_Click()
    ' Stop the trade and close the window.
    Call SendData(POut.QuitTrade & END_CHAR)
End Sub

Private Function FindOpenTradeSlot() As Long
    Dim i As Long

    ' Loop through all of the trade slots.
    For i = 1 To MAX_PLAYER_TRADES
        ' Check if an item slot exists.
        If Trading(i).InvNum = 0 Then
            FindOpenTradeSlot = i
            Exit Function
        End If
    Next i

    FindOpenTradeSlot = -1
End Function

Private Function TradeItemExists(ByVal InvNum As Long) As Boolean
    Dim i As Long

    ' Loop through all of the trade slots.
    For i = 1 To MAX_PLAYER_TRADES
        ' Check if the iventory number matches the argument.
        If Trading(i).InvNum = InvNum Then
            TradeItemExists = True
            Exit Function
        End If
    Next i
End Function
