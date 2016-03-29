Option Explicit

Private mcolProAltIgnores As Collection

Private Sub Class_Initialize()
'creates the collection in memory for usage.
    Set mcolProAltIgnores = New Collection
End Sub

Private Sub Class_Terminate()
'remove the collection from memory.
    Set mcolProAltIgnores = Nothing
End Sub

Private Function InCollection(col As Collection, key As Long) As Boolean
'check to see if the object is already in the collection by searching for the collection object key.
    Dim var As Variant
    Dim errNumber As Long               'error number storage.

    InCollection = False                'ensure that the boolean value is set to false before anything happens.
    Set var = Nothing

    Err.Clear                           'remove any previous errors.
    On Error Resume Next                'tell the VBA interperator to carry on during errors after this point.
        var = col.Item(key)             'try to pull the item using the ProAltIgnore key, any errors here will create an error code.
        errNumber = CLng(Err.Number)    'store the error number for later usage (5 = not in collection, 0 or 438 = in collection).
    On Error GoTo 0                     'tell the VBA interperator to halt on any errors after this point.

    '5 is not in, 0 and 438 represent incollection
    If errNumber = 5 Then               'it is 5 if not in collection.
        InCollection = False            'return as false
    Else
        InCollection = True             'return as true
    End If
End Function

Public Sub Add(cProAltIgnore As clsProAltIgnore)
'add the ProAltIgnore class to the collection, will update an existing object in the collection by removing the one stored and then adding the changes to the collection as a new object.
    Dim ProAltIgnoreID As Long
    
    ProAltIgnoreID = cProAltIgnore.ProAltIgnoreNo    'get the collection id from the ProAltIgnore itself.
    If InCollection(mcolProAltIgnores, ProAltIgnoreID) Then
        Debug.Print "ProAltIgnore found in ProAltIgnores collection"
        mcolProAltIgnores.Remove ProAltIgnoreID:          Debug.Print "Removing the ProAltIgnore from the collection"                           'remove from collection
        mcolProAltIgnores.Add cProAltIgnore, ProAltIgnoreID:     Debug.Print "Adding the ProAltIgnore back to the ProAltIgnores collection"     'mcolProAltIgnores.Add cProAltIgnore, ProAltIgnoreID
    Else
        mcolProAltIgnores.Add cProAltIgnore, ProAltIgnoreID:     Debug.Print "Adding the new ProAltIgnore to the ProAltIgnores collection"      'add the ProAltIgnore to the collecion.
    End If
End Sub

Public Sub Remove(ProAltIgnoreID As Long)
'remove an object from the collection.
    mcolProAltIgnores.Remove (ProAltIgnoreID)
End Sub

Public Property Get Count() As Long
'return the number of ProAltIgnores entered into the collection.
    Count = mcolProAltIgnores.Count
End Property

Public Property Get Item(NameOrNumber As Variant) As clsProAltIgnore
'return a particular ProAltIgnore via ID.
    Set Item = mcolProAltIgnores(NameOrNumber)
End Property

Public Property Get Items() As Collection
'return the objects in a collection, not sure what they are returned as.
    Set Items = mcolProAltIgnores
End Property

