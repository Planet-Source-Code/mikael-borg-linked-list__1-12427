VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyListItem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class implements the list class and is an example of what you can do with
'the LinkedList component.
'Please experiment.

Option Explicit

Implements LinkedList.Item

Private mlngMyID As Long

Private Property Let Item_ListID(ByVal RHS As Long)
    mlngMyID = RHS
End Property

Private Property Get Item_ListID() As Long
    Item_ListID = mlngMyID
End Property

Private Function Item_ListProc(lngID As Long, lngMsg As Long, varParam1 As Variant, varParam2 As Variant) As Long
    'This object does not do anything
End Function
