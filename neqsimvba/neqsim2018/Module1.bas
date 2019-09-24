Attribute VB_Name = "Module1"
Public fluidSystem As New NeqSim_NET.NeqSimNETService
Public phaseType As String
Public compositionRange As Range
Public numberOfComponents As Integer
Public compNames() As String
Public molFractions() As Double

Sub InitializeThermo()
    fluidSystem.readFluidFromGQIT (Range("B1").Value)
    numberOfComponents = fluidSystem.getNumberOfComponents
    
    
    ReDim compNames(numberOfComponents)
    ReDim molFractions(numberOfComponents)
    
    For i = 0 To numberOfComponents - 1
     compNames(i) = fluidSystem.getThermoSystem().getPhase(0).getComponent(i).getComponentName()
     molFractions(i) = fluidSystem.getThermoSystem().getPhase(0).getComponent(i).getz()
     
     Range("A4").Offset(i, 0).Value = compNames(i)
     Range("B4").Offset(i, 0).Value = molFractions(i)
    Next
    
    Dim StartCell As Range
    Set StartCell = Range("B4")
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, StartCell.Column).End(xlUp).Row
    Set compositionRange = Range(StartCell, Cells(lastRow, 2))
    compositionRange2
End Sub

Function compositionRange2() As Range
    Set compositionRange2 = compositionRange
End Function

Sub Button1_Click()
    Range("A4:B100").Clear
    InitializeThermo
End Sub

Sub setTPFraction(T As Double, P As Double, x As Range)
    For i = 0 To x.Count - 1
     molFractions(i) = x(i + 1).Value2
    Next
    fluidSystem.setTPFraction T + 273.15, P, molFractions, 1
End Sub

Sub TPflash(T As Double, P As Double, Optional ByVal flashType As Integer = -1, Optional ByVal initType As Integer = 3)
    'flashType = -2=MultiPhaseFlash, -1=TwoPhaseFlash,0=SingelPhaseGas,=1=SinglePhaseLiquid
    Dim multphase As Boolean
    
    Select Case flashType
        Case -2
            multphase = True
            Call fluidSystem.getThermoSystem().setMultiPhaseCheck(multphase)
        Case -1
            multphase = False
            Call fluidSystem.getThermoSystem().setMultiPhaseCheck(multphase)
    End Select
    
    Select Case flashType
        Case 0
            phaseType = "Vapor"
            Call fluidSystem.init(phaseType, initType)
        Case 1
            phaseType = "Liquid"
            Call fluidSystem.init(phaseType, initType)
        Case Is <= -1
            fluidSystem.TPflash
            fluidSystem.getThermoSystem().init (initType)
    End Select
    
End Sub



Function enthalpy(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            enthalpy = fluidSystem.getThermoSystem().getEnthalpy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            enthalpy = fluidSystem.getThermoSystem().getPhase(phaseNumber).getEnthalpy() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        enthalpy = fluidSystem.getThermoSystem().getEnthalpy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function speedOfSound(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            speedOfSound = fluidSystem.getThermoSystem().getSoundSpeed()
        Else
            speedOfSound = fluidSystem.getThermoSystem().getPhase(phaseNumber).getSoundSpeed()
        End If
    Else
        speedOfSound = fluidSystem.getThermoSystem().getSoundSpeed()
    End If
End Function

Function JouleThomsonCoefficient(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            JouleThomsonCoefficient = fluidSystem.getThermoSystem().getJouleThomsonCoefficient()
        Else
            JouleThomsonCoefficient = fluidSystem.getThermoSystem().getPhase(phaseNumber).getJouleThomsonCoefficient()
        End If
    Else
        JouleThomsonCoefficient = fluidSystem.getThermoSystem().getJouleThomsonCoefficient()
    End If
End Function

Function internalEnergy(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            internalEnergy = fluidSystem.getThermoSystem().getInternalEnergy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            internalEnergy = fluidSystem.getThermoSystem().getPhase(phaseNumber).getInternalEnergy() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        internalEnergy = fluidSystem.getThermoSystem().getInternalEnergy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function gibbsEnergy(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            gibbsEnergy = fluidSystem.getThermoSystem().getGibbsEnergy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            gibbsEnergy = fluidSystem.getThermoSystem().getPhase(phaseNumber).getGibbsEnergy() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        gibbsEnergy = fluidSystem.getThermoSystem().getGibbsEnergy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function entropy(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            entropy = fluidSystem.getThermoSystem().getEntropy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            entropy = fluidSystem.getThermoSystem().getPhase(phaseNumber).getEntropy() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        entropy = fluidSystem.getThermoSystem().getEntropy() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function Cp(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            Cp = fluidSystem.getThermoSystem().getCp() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            Cp = fluidSystem.getThermoSystem().getPhase(phaseNumber).getCp() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        Cp = fluidSystem.getThermoSystem().getCp() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function Cv(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 2)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            Cv = fluidSystem.getThermoSystem().getCv() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
        Else
            Cv = fluidSystem.getThermoSystem().getPhase(phaseNumber).getCv() / (fluidSystem.getThermoSystem().getPhase(phaseNumber).getNumberOfMolesInPhase() * fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#)
        End If
    Else
        Cv = fluidSystem.getThermoSystem().getCv() / (fluidSystem.getThermoSystem().getNumberOfMoles() * fluidSystem.getThermoSystem().getMolarMass() * 1000#)
    End If
End Function

Function density(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            density = fluidSystem.getThermoSystem().getDensity()
        Else
            density = fluidSystem.getThermoSystem().getPhase(phaseNumber).getPhysicalProperties().getDensity()
        End If
    Else
        density = fluidSystem.getThermoSystem().getDensity()
    End If
    
End Function

Function compressibilityFactor(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            compressibilityFactor = fluidSystem.getThermoSystem().getz()
        Else
            compressibilityFactor = fluidSystem.getThermoSystem().getPhase(phaseNumber).getz()
        End If
    Else
        compressibilityFactor = fluidSystem.getThermoSystem().getz()
    End If
    
End Function

Function molecularWeight(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    If flashType < 0 Then
        If phaseNumber = -1 Then
            molecularWeight = fluidSystem.getThermoSystem().getMolarMass() * 1000#
        Else
            molecularWeight = fluidSystem.getThermoSystem().getPhase(phaseNumber).getMolarMass() * 1000#
        End If
    Else
        molecularWeight = fluidSystem.getThermoSystem().getMolarMass() * 1000#
    End If
    
End Function

Function molFraction(T As Double, P As Double, x As Range, flashType As Integer, phase As Integer, compNo As Integer) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    If phase < 0 Then
        phase = 0
    End If
        
    molFraction = fluidSystem.getThermoSystem().getPhase(phase).getComponent(compNo).getx()
End Function

Function phaseFraction(T As Double, P As Double, x As Range, flashType As Integer, phase As Integer) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
    phaseFraction = fluidSystem.getThermoSystem().getPhase(phase).getBeta()
    
End Function

Function viscosity(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
     If flashType < 0 Then
     If phaseNumber = -1 Then
            viscosity = fluidSystem.getThermoSystem().getViscosity()
        Else
            viscosity = fluidSystem.getThermoSystem().getPhase(phaseNumber).getPhysicalProperties().getViscosity()
        End If
    Else
        viscosity = fluidSystem.getThermoSystem().getViscosity()
    End If
End Function

Function conductivity(T As Double, P As Double, x As Range, Optional ByVal flashType As Integer = -1, Optional ByVal phaseNumber As Integer = -1) As Double
    Call setTPFraction(T, P, x)
    Call TPflash(T, P, flashType, 1)
     If flashType < 0 Then
        If phaseNumber = -1 Then
            conductivity = fluidSystem.getThermoSystem().getConductivity()
        Else
            conductivity = fluidSystem.getThermoSystem().getPhase(phaseNumber).getPhysicalProperties().getConductivity()
        End If
    Else
        conductivity = fluidSystem.getThermoSystem().getConductivity()
    End If
End Function


