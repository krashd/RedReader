Private pSequenceID As Long                          'Lamp sequence ID
Private pTraffic As Long                             '?
Private pPedestrian As Long                          '?
Private pGreenArrow As Long                          '?
Private pPelVehicle As Long                          '?
Private pPelPedestrian As Long                       '?
Private pAmberIsWait As Long                         '?
Private pFlexible As Long                            '?
Private pLampStatesToBeMonitoredSelection As Long    '?

Public Property Get sequenceID() As Long
    sequenceID = pSequenceID
End Property

Public Property Let sequenceID(Value As Long)
    pSequenceID = sequenceID
End Property

Public Property Get traffic() As Long
    traffic = pTraffic
End Property

Public Property Let traffic(Value As Long)
    pTraffic = traffic
End Property

Public Property Get greenArrow() As Long
    greenArrow = pGreenArrow
End Property

Public Property Let greenArrow(Value As Long)
    pGreenArrow = greenArrow
End Property

Public Property Get pelVehicle() As Long
    pelVehicle = pPelVehicle
End Property

Public Property Let pelVehicle(Value As Long)
    pPelVehicle = pelVehicle
End Property

Public Property Get pelPedestrian() As Long
    pelPedestrian = pPelPedestrian
End Property

Public Property Let pelPedestrian(Value As Long)
    pPelPedestrian = pelPedestrian
End Property

Public Property Get amberIsWait() As Long
    amberIsWait = pAmberIsWait
End Property

Public Property Let amberIsWait(Value As Long)
    pAmberIsWait = amberIsWait
End Property

Public Property Get flexible() As Long
    flexible = pFlexible
End Property

Public Property Let flexible(Value As Long)
    pFlexible = flexible
End Property

Public Property Get lampStatesToBeMonitoredSelection() As Long
    lampStatesToBeMonitoredSelection = pLampStatesToBeMonitoredSelection
End Property

Public Property Let lampStatesToBeMonitoredSelection(Value As Long)
    pLampStatesToBeMonitoredSelection = lampStatesToBeMonitoredSelection
End Property
