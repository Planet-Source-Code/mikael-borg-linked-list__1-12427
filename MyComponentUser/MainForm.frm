VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Remove items"
      Height          =   975
      Left            =   2640
      TabIndex        =   21
      Top             =   1320
      Width           =   2295
      Begin VB.CommandButton Command6 
         Caption         =   "Remove oldest"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7800
      Top             =   2400
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Animate!"
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   120
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   15
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Send message"
      Height          =   2175
      Left            =   5040
      TabIndex        =   7
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "Send message"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Msg 2 = Animate and draw"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Msg 1 = Draw"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Param 1"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Msg"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "ID"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Count items"
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command3 
         Caption         =   "Count'em"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add items"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "20"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add items of type 2"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add circles"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Items to add"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      Caption         =   $"MainForm.frx":0000
      Height          =   855
      Left            =   5160
      TabIndex        =   20
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   $"MainForm.frx":0097
      Height          =   855
      Left            =   5160
      TabIndex        =   17
      Top             =   3000
      Width           =   3855
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AddSequence
    FromID As Long
    ToID As Long
End Type

Private arrAddList() As AddSequence
Private lngAddListCount As Long
Private lngAddListToRemove As Long
Private objMyList As LinkedList.List

Private Function AddToAddList(lngFrom As Long, lngTo As Long) As Long
    If lngAddListCount = 0 Then
        ReDim arrAddList(0)
    Else
        ReDim Preserve arrAddList(lngAddListCount)
    End If
    arrAddList(lngAddListCount).FromID = lngFrom
    arrAddList(lngAddListCount).ToID = lngTo
    lngAddListCount = lngAddListCount + 1
End Function

Private Sub Command6_Click()
    Dim i As Long
    If lngAddListCount > 0 And lngAddListToRemove < lngAddListCount - 1 Then
        For i = arrAddList(lngAddListToRemove).FromID To arrAddList(lngAddListToRemove).ToID
            objMyList.Remove i
        Next i
        lngAddListToRemove = lngAddListToRemove + 1
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Long
    Dim objMyListItem1 As MyComponent.MyListItem1
    Dim X As Long
    Dim Y As Long
    Dim dX As Long
    Dim dY As Long
    Dim R As Long
    Dim AddFrom As Long
    Dim AddTo As Long
    Dim Added As Long
    
    If IsNumeric(Text1.Text) Then
        For i = 1 To CLng(Text1.Text)
            Set objMyListItem1 = New MyComponent.MyListItem1
            R = Rnd * 40 + 20
            dX = Int((4 - (-4) + 1) * Rnd + (-4))
            dY = Int((4 - (-4) + 1) * Rnd + (-4))
            'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
            X = Int((Picture1.ScaleWidth - (R / 2) - (R / 2) + 1) * Rnd + (R / 2))
            Y = Int((Picture1.ScaleHeight - (R / 2) - (R / 2) + 1) * Rnd + (R / 2))
            objMyListItem1.SetCircleProperties X, Y, dX, dY, R _
                                             , Picture1.ScaleLeft _
                                             , Picture1.ScaleWidth _
                                             , Picture1.ScaleTop _
                                             , Picture1.ScaleHeight
            Added = objMyList.AddItem(objMyListItem1)
            If i = 1 Then AddFrom = Added
            If i = CLng(Text1.Text) Then AddTo = Added
            Set objMyListItem1 = Nothing
        Next i
        AddToAddList AddFrom, AddTo
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Dim objMyItem As LinkedList.Item
    If IsNumeric(Text1.Text) Then
        For i = 1 To CLng(Text1.Text)
            Set objMyItem = New MyComponent.MyListItem2
            objMyList.AddItem objMyItem
            Set objMyItem = Nothing
        Next i
    End If
End Sub

Private Sub Command3_Click()
    MsgBox objMyList.Count
End Sub

Private Sub Command4_Click()
    Picture1.Cls
    If IsNumeric(Text2.Text) And IsNumeric(Text3.Text) Then
        objMyList.SendMessage CLng(Text2.Text), CLng(Text3.Text), CStr(Text4.Text), Picture1.hDC
    End If
End Sub


Private Sub Command5_Click()
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled Then
        Command5.Caption = "Stop animation!"
    Else
        Command5.Caption = "Animate!"
    End If
End Sub


Private Sub Form_Load()
    Set objMyList = New LinkedList.List
    MainForm.Caption = App.Title
    Randomize
End Sub


Private Sub Timer1_Timer()
    Picture1.Cls
    objMyList.SendMessage 0, 2, 0, Picture1.hDC
End Sub
