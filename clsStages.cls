Option Explicit

Private mcolStages As Collection

Private Sub Class_Initialize()
'creates the collection in memory for usage.
    Set mcolStages = New Collection
End Sub

Private Sub Class_Terminate()
'remove the collection from memory.
    Set mcolStages = Nothing
End Sub

Private Function InCollection(col As Collection, key As Long) As Boolean
'check to see if the object is already in the collection by searching for the collection object key.
    Dim var As Variant
    Dim errNumber As Long               'error number storage.

    InCollection = False                'ensure that the boolean value is set to false before anything happens.
    Set var = Nothing

    Err.Clear                           'remove any previous errors.
    On Error Resume Next                'tell the VBA interperator to carry on during errors after this point.
        var = col.Item(key)             'try to pull the item using the Stage key, any errors here will create an error code.
        errNumber = CLng(Err.Number)    'store the error number for later usage (5 = not in collection, 0 or 438 = in collection).
    On Error GoTo 0                     'tell the VBA interperator to halt on any errors after this point.

    '5 is not in, 0 and 438 represent incollection
    If errNumber = 5 Then               'it is 5 if not in collection.
        InCollection = False            'return as false
    Else
        InCollection = True             'return as true
    End If
End Function

Public Sub Add(cStage As clsStage)
'add the Stage class to the collection, will update an existing object in the collection by removing the one stored and then adding the changes to the collection as a new object.
    Dim StageID As Long
    
    StageID = cStage.stageNo    'get the collection id from the Stage itself.
    If InCollection(mcolStages, StageID) Then
        Debug.Print "Stage found in Stages collection"
        mcolStages.Remove StageID:          Debug.Print "Removing the Stage from the collection"            'remove from collection
        mcolStages.Add cStage, StageID:     Debug.Print "Adding the Stage back to the Stages collection"    'mcolStages.Add cStage, StageID
    Else
        mcolStages.Add cStage, StageID:     Debug.Print "Adding the new Stage to the Stages collection"     'add the Stage to the collecion.
    End If
End Sub

Public Sub Remove(StageID As Long)
'remove an object from the collection.
    mcolStages.Remove (StageID)
End Sub

Public Property Get Count() As Long
'return the number of Stages entered into the collection.
    Count = mcolStages.Count
End Property

Public Property Get Item(NameOrNumber As Variant) As clsStage
'return a particular Stage via ID.
    Set Item = mcolStages(NameOrNumber)
End Property

Public Property Get Items() As Collection
'return the objects in a collection, not sure what they are returned as.
    Set Items = mcolStages
End Property

