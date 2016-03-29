Option Explicit

Private mcolPhases As Collection

Private Sub Class_Initialize()
'creates the collection in memory for usage.
    Set mcolPhases = New Collection
End Sub

Private Sub Class_Terminate()
'remove the collection from memory.
    Set mcolPhases = Nothing
End Sub

Private Function InCollection(col As Collection, key As Long) As Boolean
'check to see if the object is already in the collection by searching for the collection object key.
    Dim var As Variant
    Dim errNumber As Long               'error number storage.

    InCollection = False                'ensure that the boolean value is set to false before anything happens.
    Set var = Nothing

    Err.Clear                           'remove any previous errors.
    On Error Resume Next                'tell the VBA interperator to carry on during errors after this point.
        var = col.Item(key)             'try to pull the item using the phase key, any errors here will create an error code.
        errNumber = CLng(Err.Number)    'store the error number for later usage (5 = not in collection, 0 or 438 = in collection).
    On Error GoTo 0                     'tell the VBA interperator to halt on any errors after this point.

    '5 is not in, 0 and 438 represent incollection
    If errNumber = 5 Then               'it is 5 if not in collection.
        InCollection = False            'return as false
    Else
        InCollection = True             'return as true
    End If
End Function

Public Sub Add(cPhase As clsPhase)
'add the phase class to the collection, will update an existing object in the collection by removing the one stored and then adding the changes to the collection as a new object.
    Dim phaseID As Long
    
    phaseID = cPhase.phaseNo    'get the collection id from the phase itself.
    If InCollection(mcolPhases, phaseID) Then
        Debug.Print "Phase found in phases collection"
        mcolPhases.Remove phaseID:          Debug.Print "Removing the phase from the collection"            'remove from collection
        mcolPhases.Add cPhase, phaseID:     Debug.Print "Adding the phase back to the phases collection"    'mcolPhases.Add cPhase, phaseID
    Else
        mcolPhases.Add cPhase, phaseID:     Debug.Print "Adding the new phase to the phases collection"     'add the phase to the collecion.
    End If
End Sub

Public Sub Remove(phaseID As Long)
'remove an object from the collection.
    mcolPhases.Remove (phaseID)
End Sub

Public Property Get Count() As Long
'return the number of phases entered into the collection.
    Count = mcolPhases.Count
End Property

Public Property Get Item(NameOrNumber As Variant) As clsPhase
'return a particular phase via ID.
    Set Item = mcolPhases(NameOrNumber)
End Property

Public Property Get Items() As Collection
'return the objects in a collection, not sure what they are returned as.
    Set Items = mcolPhases
End Property
