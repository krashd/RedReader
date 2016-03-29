Private pSetNumber As Long                       'ProAltIgnoreIDNumber
Private pAppliesToVA As Long
Private pAppliesToCLF As Long
Private pAppliesToMan As Long
Private pAppliesToHC As Long
Private pAppliesToUTC As Long
Private pAppliesToEV As Long
Private pAppliesToBP As Long
Private pAppliesToPT As Long
Private pAppliesToFT As Long
Private pAppliesToMova As Long
Private pAppliesToLRT As Long

'Each of these modules has an entry with is built up of a 2d array with each row designated for each phase.
'Research on how to attain this is needed to fully implement this method.

Public Property Get setNumber() As Long
    setNumber = pSetNumber
End Property

Public Property Let setNumber(Value As Long)
    pSetNumber = setNumber
End Property

Public Property Get appliesToVA() As Long
    appliesToVA = pAppliesToVA
End Property

Public Property Let appliesToVA(Value As Long)
    pAppliesToVA = appliesToVA
End Property

Public Property Get appliesToCLF() As Long
    appliesToCLF = pAppliesToCLF
End Property

Public Property Let appliesToCLF(Value As Long)
    pAppliesToCLF = appliesToCLF
End Property

Public Property Get appliesToMAN() As Long
    appliesToMAN = pAppliesToMan
End Property

Public Property Let appliesToMAN(Value As Long)
    pAppliesToMan = appliesToMAN
End Property

Public Property Get appliesToHC() As Long
    appliesToHC = pAppliesToHC
End Property

Public Property Let appliesToHC(Value As Long)
    pAppliesToHC = appliesToHC
End Property

Public Property Get appliesToUTC() As Long
    appliesToUTC = pAppliesToUTC
End Property

Public Property Let appliesToUTC(Value As Long)
    pAppliesToUTC = appliesToUTC
End Property

Public Property Get appliesToEV() As Long
    appliesToEV = pAppliesToEV
End Property

Public Property Let appliesToEV(Value As Long)
    pAppliesToEV = appliesToEV
End Property

Public Property Get appliesToBP() As Long
    appliesToBP = pAppliesToBP
End Property

Public Property Let appliesToBP(Value As Long)
    pAppliesToBP = appliesToBP
End Property

Public Property Get appliesToPT() As Long
    appliesToPT = pAppliesToPT
End Property

Public Property Let appliesToPT(Value As Long)
    pAppliesToPT = appliesToPT
End Property

Public Property Get appliesToFT() As Long
    appliesToFT = pAppliesToFT
End Property

Public Property Let appliesToFT(Value As Long)
    pAppliesToFT = appliesToFT
End Property

Public Property Get appliesToMOVA() As Long
    appliesToMOVA = pAppliesToMova
End Property

Public Property Let appliesToMOVA(Value As Long)
    pAppliesToMova = appliesToMOVA
End Property

Public Property Get appliesToLRT() As Long
    appliesToLRT = pAppliesToLRT
End Property

Public Property Let appliesToLRT(Value As Long)
    pAppliesToLRT = appliesToLRT
End Property
