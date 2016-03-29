Private pPhaseNo As Long                             'Phase Number.
Private pRealPhase As Boolean                        'Whether real or dummy phase.
Private pPhaseRef As String                          'Phase Letter.
Private pApproach As Variant                         '?
Private pMinGreen As Long                            'Minimum Green Time in seconds.
Private pBlackoutTime As Long                        'Length of blackout time in seconds.
Private pRATime As Long                              '?
Private pExtTimeInt As Long                          'Some intergreen time in seconds.
Private pExtTimeDec As Long                          '?
Private pTypeDesc As String                          'Description or name of the phase.
Private pTypeNo As Long                              'Phase Type Number !!!! Need to find out which type is which number.
Private pDirection As String                         '?
Private pAppearance As Long                          'Appearance Type.
Private pTermination As Long                         'Termination Type   0-2
Private pUnlDemStartMaxs As Boolean                  'Unlatched Demands starts max timers.
Private pDemandedOnStartUp As Boolean                'Is the phase demanded on start-up?
Private pDemandedOnLeavingFT_MAN As Boolean          'Is the phase demanded on leaving Fixed time or Manula mode?
Private pMinLimitMinGrn As Long                      'Minimum allowed min green Limit.
Private pMaxLimitMinGrn As Long                      'Maximum allowed min green limit.
Private pRDL As Long                                 '?
Private pADL As Long                                 '?
Private pGDL As Long                                 '?
Private pFailFlash As Long                           '?  defaulted to 255
Private pAssociatedPhase As Variant                  'Stored as a 1 dimensional array !!!! needs looking into.       stored as boolean.
Private pRevertiveDemand As Variant                  'Stored as a 1 dimensional array !!!! needs looking into.       stored as boolean.
Private pMaxGreens As Variant                        'Stored as a 1 dimensional array !!!! needs looking into.       stored as long.
Private pLeavingTimes As Variant                     'Stored as a 1 dimensional array !!!! needs looking into.       stored as long.
Private pLeavingTimesFractions As Variant            'Stored as a 1 dimensional array !!!! needs looking into.       stored as long.
Private pDemandByVA As Variant                       'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pExtendedByVA As Variant                     'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pDemandByUTC As Variant                      'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pExtendedByUTC As Variant                    'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pPushbuttonKerbsideAssociations As Variant   'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pPhaseAssociationsDefaulted As Boolean       '### Total guesswork on this one, looks booleanish, should really change it to a long.
Private pOldDemands As Variant                       'Stored as a 1 dimensional array !!!! needs looking into.       unknown on what is entered into this field, needs more research.
Private pPDDTime As Long                             'PDD Time in seconds.
Private pPDXTimeInt As Long                          'PDX Time in seconds.
Private pPDXTimeDec As Long                          'PDX Time in seconds.
Private pCMXTime As Long                             'CMX Time in seconds.
Private pCDY0Time As Long                            'CDY0 Time in seconds.
Private pCDY1Time As Long                            'CDY1 Time in seconds.
Private pCRDTime As Long                             'CDR Time in seconds.
Private pPreTimedMax As Long                         'Pre-Timed Max in seconds.
Private pOnCrossingDetection As Variant              'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.
Private pKerbsideDetection As Variant                'Stored as a 1 dimensional array !!!! needs looking into.       stored as string, detector names, maximum 4 entries.

Public Property Get phaseNo() As Long
    phaseNo = pPhaseNo
End Property

Public Property Let phaseNo(Value As Long)
    pPhaseNo = phaseNo
End Property

Public Property Get realPhase() As Boolean
    realPhase = pRealPhase
End Property

Public Property Let realPhase(Value As Boolean)
    pRealPhase = realPhase
End Property

Public Property Get phaseRef() As String
    phaseRef = pPhaseRef
End Property

Public Property Let phaseRef(Value As String)
    pPhaseRef = phaseRef
End Property

Public Property Get approach() As Variant
    approach = pApproach
End Property

Public Property Let approach(Value As Variant)
    pApproach = approach
End Property

Public Property Get minGreen() As Long
    minGreen = pMinGreen
End Property

Public Property Let minGreen(Value As Long)
    pMinGreen = minGreen
End Property

Public Property Get blackoutTime() As Long
    blackoutTime = pBlackoutTime
End Property

Public Property Let blackoutTime(Value As Long)
    pBlackoutTime = blackoutTime
End Property

Public Property Get RATime() As Long
    RATime = pRATime
End Property

Public Property Let RATime(Value As Long)
    pRATime = RATime
End Property

Public Property Get extTimeInt() As Long
    extTimeInt = pExtTimeInt
End Property

Public Property Let extTimeInt(Value As Long)
    pExtTimeInt = extTimeInt
End Property

Public Property Get extTimeDec() As Long
    extTimeDec = pExtTimeDec
End Property

Public Property Let extTimeDec(Value As Long)
    pExtTimeDec = extTimeDec
End Property

Public Property Get typeDesc() As String
    typeDesc = pTypeDesc
End Property

Public Property Let typeDesc(Value As String)
    pTypeDesc = typeDesc
End Property

Public Property Get typeNo() As Long
    typeNo = pTypeNo
End Property

Public Property Let typeNo(Value As Long)
    pTypeNo = typeNo
End Property

Public Property Get direction() As String
    direction = pDirection
End Property

Public Property Let direction(Value As String)
    pDirection = direction
End Property

Public Property Get appearance() As Long
    direction = pDirection
End Property

Public Property Let direction(Value As Long)
    pDirection = direction
End Property

Public Property Get termination() As Long
    termination = pTermination
End Property

Public Property Let termination(Value As Long)
    pTermination = termination
End Property

Public Property Get unlDemStartMaxs() As Boolean
    unlDemStartMaxs = pUnlDemStartMaxs
End Property

Public Property Let unlDemStartMaxs(Value As Boolean)
    pUnlDemStartMaxs = unlDemStartMaxs
End Property

Public Property Get demandedOnStartUp() As Boolean
    demandedOnStartUp = pDemandedOnStartUp
End Property

Public Property Let demandedOnStartUp(Value As Boolean)
    pDemandedOnStartUp = demandedOnStartUp
End Property

Public Property Get demandedOnLeavingFT_MAN() As Boolean
    demandedOnLeavingFT_MAN = pDemandedOnLeavingFT_MAN
End Property

Public Property Let demandedOnLeavingFT_MAN(Value As Boolean)
    pDemandedOnLeavingFT_MAN = demandedOnLeavingFT_MAN
End Property

Public Property Get minLimitMinGrn() As Long
    minLimitMinGrn = pMinLimitMinGrn
End Property

Public Property Let minLimitMinGrn(Value As Long)
    pMinLimitMinGrn = minLimitMinGrn
End Property

Public Property Get maxLimitMinGrn() As Long
    maxLimitMinGrn = pMaxLimitMinGrn
End Property

Public Property Let maxLimitMinGrn(Value As Long)
    pMaxLimitMinGrn = maxLimitMinGrn
End Property

Public Property Get RDL() As Long
    RDL = pRDL
End Property

Public Property Let RDL(Value As Long)
    pRDL = RDL
End Property

Public Property Get ADL() As Long
    ADL = pADL
End Property

Public Property Let ADL(Value As Long)
    pADL = ADL
End Property

Public Property Get GDL() As Long
    GDL = pGDL
End Property

Public Property Let GDL(Value As Long)
    pGDL = GDL
End Property

Public Property Get failFlash() As Long
    failFlash = pFailFlash
End Property

Public Property Let failFlash(Value As Long)
    pFailFlash = failFlash
End Property

