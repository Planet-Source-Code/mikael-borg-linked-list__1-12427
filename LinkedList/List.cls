VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This class implements the actual list. This is the only class to createable in
'this component.

Option Explicit

Private mobjFirstNode As LinkedList.Node    'First node of the list.

Private mlngCurrentID As Long               'The current ID of any item to add.

Private Const LM_BROADCAST As Long = &H0    'Constant for broadcasting of messages.

Public Function SendMessage(lngID As Long _
                          , lngMsg As Long _
                          , varParam1 As Variant _
                          , varParam2 As Variant) As Long
    'Sends a message to one or all items in the list.
    'lngID      The item id for the item to send the message to.
    '           If parameter is set to LM_BROADCAST (0) then the message is
    '           sent to all items in the list.
    'lngMsg     The message.
    'varParam1  A parameter of any kind.
    'varParam2  A parameter of any kind.
    'Returns the error code from the item class that received the message. If lngID
    'was set to LM_BROADCAST and a ListProc method of a link item returned a value
    'less than zero, the sending of messages is stopped at that item.
                                    
    On Error GoTo SysErrorExit
    Dim objCurrentNode As LinkedList.Node
    Dim lngRet As Long
    'Get the first node.
    Set objCurrentNode = mobjFirstNode
    'While the node is not empty...
    While Not objCurrentNode Is Nothing
        '.. check the item ID or check if all items shall receive the message.
        If lngID = LM_BROADCAST Or objCurrentNode.MyItem.ListID = lngID Then
            'Sent the message.
            lngRet = objCurrentNode.MyItem.ListProc(lngID, lngMsg, varParam1, varParam2)
            If lngRet < 0 Then
                SendMessage = lngRet
                GoTo ExitFunction
            End If
        End If
        'Get the next node.
        Set objCurrentNode = objCurrentNode.NextNode
    Wend
ExitFunction:
    Exit Function
SysErrorExit:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo ExitFunction
End Function

Public Function AddItem(objItem As LinkedList.Item) As Long
    'Adds an item to the list.
    'objItem    The item to add.
    'Returns the id of the added item on success. If an error occurs the
    'method returns a value less than zero.
                
    On Error GoTo SysErrorExit
    Dim objNewNode As LinkedList.Node
    Dim objSecondNode As LinkedList.Node
    'Prepare the new node.
    
    'Create a new ID and set the ID of the item.
    mlngCurrentID = mlngCurrentID + 1
    AddItem = mlngCurrentID
    objItem.ListID = mlngCurrentID
    'Create a new node and attach the item.
    Set objNewNode = New LinkedList.Node
    Set objNewNode.MyItem = objItem
    
    'Attach the new node to the list.
    
    'Get the old first node
    Set objSecondNode = mobjFirstNode
    'Attach the new node as the first node
    Set mobjFirstNode = objNewNode
    Set mobjFirstNode.PreviousNode = Nothing
    Set mobjFirstNode.NextNode = objSecondNode
    
    'Link back from the second node.
    If Not objSecondNode Is Nothing Then
        Set objSecondNode.PreviousNode = mobjFirstNode
    End If
    
ExitFunction:
    Exit Function
SysErrorExit:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo ExitFunction
End Function

Public Function Count() As Long
    'Returns the number of items in the list.

    On Error GoTo SysErrorExit
    Dim objCurrentNode As LinkedList.Node
    Dim lngCount As Long
    
    'Set the counter to zero to be sure (!).
    lngCount = 0
    'Get the first node.
    Set objCurrentNode = mobjFirstNode
    'While the node is not empty....
    While Not objCurrentNode Is Nothing
        '... increase the count..
        lngCount = lngCount + 1
        '... and get the next node.
        Set objCurrentNode = objCurrentNode.NextNode
    Wend
ExitFunction:
    'Return the node count.
    Count = lngCount
    Exit Function
SysErrorExit:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo ExitFunction
End Function

Public Function Remove(lngItemID As Long) As LinkedList.Item
    'Removes the item from the list.
    'Returns the removed item in case the item shall not be destroyed.

    'On Error GoTo SysErrorExit
    Dim objCurrentNode As LinkedList.Node
    Dim objNodeToRemove As LinkedList.Node
    Dim blnItemToRemoveFound As Boolean
    
    'Get the first node.
    Set objCurrentNode = mobjFirstNode
    While (Not objCurrentNode Is Nothing) Or Not blnItemToRemoveFound
        If objCurrentNode.MyItem.ListID = lngItemID Then
            Set objNodeToRemove = objCurrentNode
            blnItemToRemoveFound = True
        End If
        Set objCurrentNode = objCurrentNode.NextNode
    Wend
    If blnItemToRemoveFound Then
        If Not objNodeToRemove.PreviousNode Is Nothing Then
            Set objNodeToRemove.PreviousNode.NextNode = objNodeToRemove.NextNode
        End If
        If Not objNodeToRemove.NextNode Is Nothing Then
            Set objNodeToRemove.NextNode.PreviousNode = objNodeToRemove.PreviousNode
        End If
        Set objNodeToRemove.NextNode = Nothing
        Set objNodeToRemove.PreviousNode = Nothing
        Set Remove = objNodeToRemove.MyItem
        Set objNodeToRemove.MyItem = Nothing
        Set objNodeToRemove = Nothing
    End If
ExitFunction:
    Exit Function
SysErrorExit:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo ExitFunction
End Function
