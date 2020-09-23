VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'This class is the class to implement for any list item class you want to
'add to a list.
'See MyComponent.MyListItem1 for an example. Please experiment with the other class.

Public ListID As Long   'The identity of the item. This public variable is
                        'converted into Property Get and Property Let statements
                        'in the implementing class.

Public Function ListProc(lngID As Long _
                       , lngMsg As Long _
                       , varParam1 As Variant _
                       , varParam2 As Variant) As Long
    'This method defines the function for the list class to call for each item.
    'The method is to be implemented by the list item class.
    'lngID      The item id for the called item.
    'lngMsg     The message.
    'varParam1  A parameter of any kind.
    'varParam2  A parameter of any kind.
    'The method shall return a value less than zero if an error occurs.
End Function


