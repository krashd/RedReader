Private pGroupNo As Long             'DFM group number.
Private pActiveTimes As Long         'Active DFM times.          'Stored as a 1 dimensional array !!!! needs looking into.       stored as long, 4 entries.
Private pInactiveTimes As Long       'Inactive DFM times.        'Stored as a 1 dimensional array !!!! needs looking into.       stored as long, 4 entries.

Public Property Get groupNo() As Long
    groupNo = pGroupNo
End Property

Public Property Let groupNo(Value As Long)
    pGroupNo = groupNo
End Property

Public Property Get activeTimes() As Long
    activeTimes = pActiveTimes
End Property

Public Property Let activeTimes(Value As Long)
    pActiveTimes = activeTimes
End Property

Public Property Get inactiveTimes() As Long
    inactiveTimes = pInactiveTimes
End Property

Public Property Let inactiveTimes(Value As Long)
    pInactiveTimes = inactiveTimes
End Property
