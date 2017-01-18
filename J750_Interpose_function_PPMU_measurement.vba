'J750 interpose function calling convention need to be like this:
'(argc As Long, argv() As String)
''argc' is numbers of parameters
''argv' is parameters lists, separated by comma
'Parameter for this function is: LowLimit, HighLimit, Pins, ForceCurrent

Public Function InterposePPMUMeasureSingle(argc As Long, argv() As String) As Long

    If (argc <> 4) Then
        GoTo errHandler
    End If
    
    Dim LowLimit As Double
    Dim highLimit As Double
    Dim pins As String
    Dim ForceCurrent As Double
    
    LowLimit = CDbl(Trim(argv(0)))
    highLimit = CDbl(Trim(argv(1)))
    pins = Trim(argv(2))
    ForceCurrent = CDbl(Trim(argv(3)))
    
    'Clamp value should be taking care automatically by spec sheets
    'Set limits for the tests
    With thehdw.PPMU.pins(pins)
        .TestLimitLow = LowLimit
        .TestLimitHigh = highLimit
        '.TestLimitValid = pmuBothLimitsValid 'Should not be necessary
        .ForceCurrent(ppmuSmartRange) = ForceCurrent
        .Connect
    End With
    
    'Decalre a PinListData variable to store measured result
    'Use TestLimit method to datalogging the result
    Dim ResultPLD As New PinListData
    thehdw.PPMU.pins(pins).MeasureVoltages ResultPLD
    TheExec.Flow.TestLimit resultVal:=ResultPLD, LowLimit:=LowLimit, highLimit:=highLimit, _
        ForceValue:=ForceCurrent, forceUnit:=unitVolt, scaleValue:=scaleNone
    
    'Cold switching PPMU relay of the pins
    thehdw.PPMU.pins(pins).ForceCurrent = 0.000000001
    thehdw.PPMU.pins(pins).Disconnect
    
errHandler:

    TheExec.datalog.WriteComment "Error encountered within the InterposePPMUMeasure interpose function"

End Function
