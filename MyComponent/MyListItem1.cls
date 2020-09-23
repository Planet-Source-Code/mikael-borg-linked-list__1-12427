VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyListItem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class implements the list class and is an example of what you can do with
'the LinkedList component.
'This specific class is implemented as circles to be drawn on a picture.

Option Explicit
Implements LinkedList.Item

Private mlngMyID As Long

Private mX As Double
Private mY As Double
Private mdX As Double
Private mdY As Double
Private mddY As Double
Private mRadius As Long
Private mMaxX As Long
Private mMaxY As Long
Private mMinX As Long
Private mMinY As Long

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Property Get Item_ListID() As Long
    Item_ListID = mlngMyID
End Property

Private Property Let Item_ListID(ByVal RHS As Long)
    mlngMyID = RHS
End Property

Private Function Item_ListProc(lngID As Long, lngMsg As Long, varParam1 As Variant, varParam2 As Variant) As Long
    Select Case lngMsg
    Case 1  'Draw
        Call Ellipse(CLng(varParam2), mX - mRadius / 2, mY - mRadius / 2, mX + mRadius / 2, mY + mRadius / 2)
    Case 2  'Animate
        mY = mY + mdY
        mdY = mdY + mddY
        If (mY - mRadius / 2) <= mMinY Or (mY + mRadius / 2) >= mMaxY Then
            mdY = -mdY
            mY = mY + mdY
            mdY = mdY * 0.4
        End If
        mX = mX + mdX
        If (mX - mRadius / 2) <= mMinX Or (mX + mRadius / 2) >= mMaxX Then
            mdX = -mdX
            mX = mX + mdX
            mdX = mdX * 0.4
        End If
        Call Ellipse(CLng(varParam2), mX - mRadius / 2, mY - mRadius / 2, mX + mRadius / 2, mY + mRadius / 2)
    End Select
End Function

Public Sub SetCircleProperties(lngX As Long _
                             , lngY As Long _
                             , lngdX As Long _
                             , lngdY As Long _
                             , lngRaduis As Long _
                             , lngMinX As Long _
                             , lngMaxX As Long _
                             , lngMinY As Long _
                             , lngMaxY As Long)
    mX = lngX
    mY = lngY
    mdX = lngdX
    mdY = lngdY
    mRadius = lngRaduis
    mMinX = lngMinX
    mMaxX = lngMaxX
    mMinY = lngMinY
    mMaxY = lngMaxY
    mddY = 0.01
End Sub

