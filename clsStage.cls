Private pStageNo As Long                         'Stage ID number
Private pWindowTime As Long                      '?
Private pUnlDemToStartMaxs As Long               'Unlatched Demands to start maxs.
Private pDemandedOnStartUp As Boolean            'Is the stage demanded on system start-up?
Private pDemandedOnLeavingMAN_FT As Boolean      'Is the stage demanded on leaving Manual or Fixed-Time mode?
Private pExceptionalStage As Boolean             '?
Private pPhasesInStage As Variant                'Stored as a 1 dimensional array !!!! needs looking into, dynamic array where the size is dependant on the number of phases, each entry is stored as a boolean value.
Private pFixedPhases As Variant                  'Stored as a 1 dimensional array !!!! needs looking into, dynamic array where the size is dependant on the number of phases, each entry is stored as a boolean value.

Public Property Get stageNo() As Long
    stageNo = pStageNo
End Property

Public Property Let stageNo(Value As Long)
    pStageNo = stageNo
End Property

Public Property Get windowTime() As Long
    windowTime = pWindowTime
End Property

Public Property Let windowTime(Value As Long)
    pWindowTime = windowTime
End Property

Public Property Get unlDemToStartMaxs() As Long
    unlDemToStartMaxs = pUnlDemToStartMaxs
End Property

Public Property Let unlDemToStartMaxs(Value As Long)
    pUnlDemToStartMaxs = unlDemToStartMaxs
End Property

Public Property Get demandedOnStartUp() As Boolean
    demandedOnStartUp = pDemandedOnStartUp
End Property

Public Property Let demandedOnStartUp(Value As Boolean)
    pDemandedOnStartUp = demandedOnStartUp
End Property

Public Property Get demandedOnLeavingMAN_FT() As Boolean
    demandedOnLeavingMAN_FT = pDemandedOnLeavingMAN_FT
End Property

Public Property Let demandedOnLeavingMAN_FT(Value As Boolean)
    pDemandedOnLeavingMAN_FT = demandedOnLeavingMAN_FT
End Property

Public Property Get exceptionalStage() As Boolean
    exceptionalStage = pExceptionalStage
End Property

Public Property Let exceptionalStage(Value As Boolean)
    pExceptionalStage = exceptionalStage
End Property

Public Property Get phasesInStage() As Variant
    phasesInStage = pPhasesInStage
End Property

Public Property Let phasesInStage(Value As Variant)
    pPhasesInStage = phasesInStage
End Property

Public Property Get fixedPhases() As Variant
    fixedPhases = pFixedPhases
End Property

Public Property Let fixedPhases(Value As Variant)
    pFixedPhases = fixedPhases
End Property
