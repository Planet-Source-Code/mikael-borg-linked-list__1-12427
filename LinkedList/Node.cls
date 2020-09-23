VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class defines a node in the linked list.

Option Explicit

Public PreviousNode As LinkedList.Node  'The previous node in the list.

Public NextNode As LinkedList.Node      'The next node in the list.

Public MyItem As LinkedList.Item        'The list item attached to this node.

