Attribute VB_Name = "VBT_Module"
Option Explicit
'© 2019 Celerint, LLC. All rights reserved. No part of this software code
'may be reproduced, distributed, or transmitted in any form or by any
'means, including photocopying, recording, or other electronic or
'mechanical methods, without the prior written permission of the owner.
'Patent pending. www.celerint.com/patents.
Global ADC_CH6_K1_ON As New PinListData
Global ADC_CH7_K1_ON As New PinListData
Global ADC_CH6_K1_OFF As New PinListData
Global ADC_CH7_K1_OFF As New PinListData
Global ADG_NC_Grp1 As New PinListData
Global ADG_NC_Grp2 As New PinListData



' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function


Public Function functionalExp1(dpin As String, patname As String, Optional strtlab As String)
    Dim fcnts As New PinListData
    
    On Error GoTo errHandler
    
    ' load levels and timing
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
    ' ensure pattern is loaded
    TheHdw.Patterns(patname).Load
    ' start the pattern
    Call TheHdw.Patterns(patname).Start(strtlab)
    Call TheHdw.Digital.Patgen.HaltWait

    ' get the failing counts
    fcnts = TheHdw.Digital.Pins(dpin).FailCount
    
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next
End Function


Public Function functionalExp2(dpin As String, patname As String, Optional strtlab As String)
    Dim fcnts As New PinListData
    
    On Error GoTo errHandler
    
    ' load levels and timing
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
    ' ensure pattern is loaded
    TheHdw.Patterns(patname).Load
    ' start the pattern
    Call TheHdw.Patterns(patname).Start(strtlab)
    Call TheHdw.Digital.Patgen.HaltWait

    ' get the failing counts
    fcnts = TheHdw.Digital.Pins(dpin).FailCount
    
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UC01_Propagation_Delay()
    Dim i As Integer
    Dim j As Integer
    Dim site As Variant
    Dim value As Double
    Dim row(28) As String
    Dim col(28) As String
    Dim cell As String
    Dim pin_name As String
    Dim propogation_speed As Double
    
    propogation_speed = 0.000000000165
    
    col(0) = "B"
    col(1) = "c"
    col(2) = "d"
    col(3) = "e"
    col(4) = "f"
    col(5) = "g"
    col(6) = "h"
    col(7) = "i"
    col(8) = "j"
    col(9) = "k"
    col(10) = "l"
    col(11) = "m"
    col(12) = "n"
    col(13) = "o"
    col(14) = "p"
    col(15) = "q"
    col(16) = "r"
    col(17) = "s"
    col(18) = "t"
    col(19) = "u"
    col(20) = "v"
    col(21) = "w"
    col(22) = "x"
    col(23) = "y"
    col(24) = "z"
    col(25) = "aa"
    col(26) = "ab"
    col(27) = "ac"
    
   
' 1. Create length constants for transmission lines of known length.
' 2. Calculate propagation delay (speed) using speed = trace length / TDR test fixture transit time.
    
    For Each site In TheExec.Sites.Selected
        For i = 2 To 36
            cell = "A" & i
            pin_name = Worksheets("TraceLengths").Range(cell).value
            
            cell = col(site) & i
            value = Worksheets("TraceLengths").Range(cell).value
            value = value / 1000                'Convert mils to inches.
            value = value * propogation_speed   '150ps or 180ps averaged out to 165ps.
            
            cell = col(site) & i + 100
            Worksheets("TraceLengths").Range(cell).value = value
            
            TheExec.Flow.TestLimitIndex = 0
            TheExec.Flow.TestLimit ResultVal:=value, _
                                unit:=unitNone, _
                                forceUnit:=unitNone, _
                                ForceResults:=tlForceFlow, _
                                PinName:=pin_name, _
                                TName:="UC01_Calc_Dly"
        Next i
    Next site
    
End Function



Public Function UC02_Pogo_Pin_Cont()
' 1. Create TDR file called FRT_Pogo with all channels set to 0 ps.
' 2. Create a pin group called FRT_Digital per 662_386_00 Test Resources.xlsx that includes every digital channel
'    in the system regardless of their application use.
' 3. Create a time set called FRT_Timing per 662_386_00 Test Resources.xlsx.
' 4. Create FRT pin levels called set FRT_Edge per 662_386_00 Test Resources.xlsx.
' 5. Create FRT functional pattern. The pattern file is a single vector. Each channel has the following Pattern: Compare L
'    at T0, drive H at T0+rise time, compare for L at T0+2x(rise time). Assuming TDR data set swap on the fly is viable, this
'    functional test will work for all FRT tests that look for a refection. LL is a pass, LH is an open pogo, HH is discontinuity inside the test head.
End Function

Public Function UC03_Func_TDR_Verification(dpin As String, patname As String, Optional strtlab As String)
' 1. Execute functional test using resources listed.
    Dim fcnts As New PinListData
    
    On Error GoTo errHandler
    
'    ' load levels and timing
'    thehdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
'    ' ensure pattern is loaded
'    thehdw.Patterns(patname).Load
'    ' start the pattern
'    Call thehdw.Patterns(patname).Start(strtlab)
'    Call thehdw.Digital.Patgen.HaltWait
'
'    ' get the failing counts
'    fcnts = thehdw.Digital.Pins(dpin).FailCount
'
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function UC04_K1_RELAY_FUNC()
' 1. Set 17.VI80_27 and 17.VI80_32 to the same high voltage such as 10V (Sheet 16). VI80 only goes to 7 Volts.
' 2. Differential Op Amp U2 output should be near 0V.
' 3. Use Force Voltage Measure Current mode for 17.VI80_32 to measure current for both states of the K1 relay. Should be 0mA as shown and 6V/(4.53K+1.1K Ohms) when actuated.
' 4. Build test based on differences in voltage between states.

' The tested values for the K1 test were captured in UC14 so we don't have to retest them here.
' We will just test the results here for the flow.

    Dim i As Integer
    Dim site As Variant
    Dim sar_val1 As New PinListData
    Dim sar_val2 As New PinListData
    Dim PinArr() As String, PinCount As Long
    Dim results As Double
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_S_VI", PinArr, PinCount
    
    'Datalog results
    'This tests that K1 is operational.
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            results = (ADC_CH6_K1_ON.Pins(PinArr(0)).value - ADC_CH6_K1_OFF.Pins(PinArr(0)).value)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitVolt, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="UC04_K1_RLY_OP"
        Next i
    Next site
    

End Function

Public Function UC05_K2_RELAY_FUNC()
' 1. Force 5V using 17.VI80FRC21 (Sheet 21).
' 2. Configure Decoder S1_00 to connect 17.VI80_FRC21 to K2.
' 3. Measure current using VI80_FRC21 to measure current with relay open (0A) and closed (1mA)
' 4. Build test based on differences in current between states.
' 5. Apply to all sites.

    Dim i As Integer
    Dim site As Variant
    Dim K2_I_OPEN As New PinListData
    Dim K2_I_CLOSED As New PinListData
    Dim PinArr() As String, PinCount As Long
    Dim results As Double
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_S_VI", PinArr, PinCount
    
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff                  'Enable the mux switch
    TheHdw.Utility.Pins("K_S1_A1,K_S1_A0").State = tlUtilBitOn        'Select Mux switch 1
    TheHdw.Utility.Pins("K_ATX_PD").State = tlUtilBitOff                'Set K2 Open
    
    With TheHdw.DCVI.Pins("atx_s_vi")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .VoltageRange.Autorange = False
        .CurrentRange = 0.02
        .CurrentRange.Autorange = True
        .Voltage = 4.99996185244526
        .Current = 0.02
        .Gate = True
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 0.02
        .Connect tlDCVIConnectHighForce
        .Connect tlDCVIConnectHighSense
        .BleederResistor = tlDCVIBleederResistorOff
        .Meter.HardwareAverage = 1
    End With
    
    K2_I_OPEN = TheHdw.DCVI.Pins("ATX_S_VI").Meter.Read(tlStrobe, 1)    'Should be 0 amps.

    TheHdw.Utility.Pins("K_ATX_PD").State = tlUtilBitOn                 'Set K2 closed
    TheHdw.wait (0.01)
    
    K2_I_CLOSED = TheHdw.DCVI.Pins("ATX_S_VI").Meter.Read(tlStrobe, 1)  'Should be 1mA
    
    'Clean up
    TheHdw.Utility.Pins("K_S1_EN,K_S1_A1,K_S1_A0,K_ATX_PD").State = tlUtilBitOff
    
    With TheHdw.DCVI.Pins("ATX_S_VI")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
   
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                K2_I_OPEN.Pins(PinArr(i)).value = 0
                K2_I_CLOSED.Pins(PinArr(i)).value = 0.001 + (Rnd * 0.00001)
           Next i
        Next site
    End If
   
    'Data log results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            results = 5 / (K2_I_CLOSED.Pins(PinArr(i)).value - K2_I_OPEN.Pins(PinArr(i)).value)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitNone, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="UC05_K2_RLY_OP"
        Next i
    Next site

End Function

Public Function UC06_DECODER_S1_Func()
    Dim i As Integer
    Dim site As Variant
    Dim ADC_CH6_S_CLOSE As New PinListData
    Dim ADC_CH7_S_CLOSE As New PinListData
    Dim PinArr() As String, PinCount As Long
    Dim AlarmBehavior As tlAlarmBehavior
    Dim results As Double
    
    Dim MUX1_S1 As New PinListData
    Dim MUX1_S2 As New PinListData
    Dim MUX1_S3 As New PinListData
    Dim MUX1_S4 As New PinListData
    
    Dim MUX1_S1_EN_ON As New PinListData
    Dim MUX1_S2_EN_ON As New PinListData
    Dim MUX1_S3_EN_ON As New PinListData
    Dim MUX1_S4_EN_ON As New PinListData
        
    Dim MUX1_S1_EN_OFF As New PinListData
    Dim MUX1_S2_EN_OFF As New PinListData
    Dim MUX1_S3_EN_OFF As New PinListData
    Dim MUX1_S4_EN_OFF As New PinListData
    
    Dim MUX1_S1_EN_alarm(30) As Long
    Dim MUX1_S2_EN_alarm(30) As Long
    Dim MUX1_S3_EN_alarm(30) As Long
    Dim MUX1_S4_EN_alarm(30) As Long

    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheExec.DataManager.DecomposePinList "ATX_S_VI", PinArr, PinCount
    
' 1. Configure the Decoder S1 to be enabled and set to switch 1: EN=1, A1=0, A0=0 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all S1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff             'SI decoders select switch 1a.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff             'SI decoders select switch 1a.

' 2. Set the attached VI80 resource to force 7V.
    With TheHdw.DCVI.Pins("ATX_S_VI")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = 5#
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 0.02
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With

    'Measured current should be leakage level… no alarms.
    MUX1_S1 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    

    ' MUX1:S1:EN control
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff         'Disable all S1 decoders. Should cause an alarm.
        
    For Each site In TheExec.Sites.Selected
        MUX1_S1_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_S_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn          'Re-Enable all S1 decoders. Should clear the alarm condition.
    TheHdw.DCVI.Pins("atx_s_vi").AlarmClear
    'Make sure alarm went away.
      
      
      
      
      
      
    'Configure the Decoder S1 to be enabled and set to switch 2: EN=1, A1=0, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all S1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff             'SI decoders select switch 2a.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOn              'SI decoders select switch 2a.

    'Measured current should be leakage level… no alarms.
    MUX1_S2 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    

    ' MUX1:S1:EN control
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff         'Disable all S1 decoders.  Should cause an alarm.
    
    For Each site In TheExec.Sites.Selected
        MUX1_S2_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_S_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn          'Re-Enable all S1 decoders. Should clear the alarm condition.
    TheHdw.DCVI.Pins("atx_s_vi").AlarmClear
    'Make sure alarm went away.
   
    
    
    
    
    
    
    '**********************************************************************************
    'Configure 17.VI80_SNS21 for voltmeter mode (Sheet21).
    With TheHdw.DCVI.Pins("atx_s_vi")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
    
    With TheHdw.DCVI.Pins("atx_s_vi")
        .Mode = tlDCVIModeHighImpedance
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = 0#
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .Meter.HardwareAverage = 1
    End With

    'Set the VI80 inputs to the differential op amp U2 such that the output is 3V (Sheet 16).
    With TheHdw.DCVI.Pins("sar_cal,sar_cal2")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = 3#
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With

    'Configure the Decoder S1 to be enabled and set to switch 3: EN=1, A1=1, A0=0 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all S1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOn              'SI decoders select switch 3a.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff             'SI decoders select switch 3a.
    
    MUX1_S3 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    
   
    ' MUX1:S3:EN control
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff             'Disable all MUX1 decoders.
    MUX1_S3_EN_OFF = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all MUX1 decoders.
    
    
     
    'Configure the Decoder S1 to be enabled and set to switch 4: EN=1, A1=1, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all MUX1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOn              'MUX1 decoders select switch 4a.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOn              'MUX1 decoders select switch 4a.
    
    MUX1_S4 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
     
    ' MUX1:S4:EN control
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff             'Disable all MUX1 decoders.
    MUX1_S4_EN_OFF = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
     
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all MUX1 decoders.
     
    
    
        
    'Clean up
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2,K_S1_EN,K_S1_A1,K_S1_A0").State = tlUtilBitOff

    With TheHdw.DCVI.Pins("sar_cal2,sar_cal,atx_s_vi")
        .Gate = False
        .Disconnect
    End With

    
    
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                MUX1_S1.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX1_S1_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
               
                MUX1_S2.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX1_S2_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
                
                MUX1_S3.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX1_S3_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
               
                MUX1_S4.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX1_S4_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
                
          Next i
        Next site
    End If
    
    
    
    'Data log results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            
            TheExec.Flow.TestLimitIndex = 0
            results = MUX1_S1.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S1_Func"
            TheExec.Flow.TestLimitIndex = 1
            results = MUX1_S1_EN_alarm(site)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S1_EN"
            TheExec.Flow.TestLimitIndex = 0
            results = MUX1_S2.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S2_Func"
            TheExec.Flow.TestLimitIndex = 1
            results = MUX1_S2_EN_alarm(site)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S2_EN"
            TheExec.Flow.TestLimitIndex = 0
            results = MUX1_S3.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S3_Func"
            TheExec.Flow.TestLimitIndex = 1
            results = MUX1_S3_EN_alarm(site)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S3_EN"
            TheExec.Flow.TestLimitIndex = 0
            results = MUX1_S4.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S4_Func"
            TheExec.Flow.TestLimitIndex = 1
            results = MUX1_S3_EN_alarm(site)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), TName:="UC06_MUX_S4_EN"
       Next i
    Next site




End Function

Public Function UC07_DECODER_S2_Func()
' 1. Configure the Decoder S2 to be enabled and set to switch 1: EN=1, A1=0, A0=0 (Sheet 103).
' 2. Set the attached VI80 resource to force 10V. (7V Max)
' 3. Measured current should be leakage level… no alarms.
' 4. Set EN=0, this breaks the F-S line connection and an alarm should be triggered. This confirms EN is functional and switch 1 is functional.
' 5. Repeat 1-4 for each switch setting. (Set the switch then toggle the enable line)
' 6. Implement for all sites.

    Dim i As Integer
    Dim site As Variant
    Dim PinArr() As String, PinCount As Long
    Dim AlarmBehavior As tlAlarmBehavior
    Dim results As Double
    
    Dim MUX2_S1_OK As New PinListData
    Dim MUX2_S2_OK As New PinListData
    Dim MUX2_S3_OK As New PinListData
    Dim MUX2_S4_OK As New PinListData
    
    Dim MUX2_S1_EN_alarm(30) As Long
    Dim MUX2_S2_EN_alarm(30) As Long
    Dim MUX2_S3_EN_alarm(30) As Long
    Dim MUX2_S4_EN_alarm(30) As Long

    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_F_VI", PinArr, PinCount
    
' 2. Set the attached VI80 resource to force 10V. (Only goes to 7V Max)
    With TheHdw.DCVI.Pins("ATX_F_VI")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = 5#
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 0.02
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With


    'Configure the Decoder S2 to be enabled and set to switch 1: EN=1, A1=0, A0=0 (Sheet 21).
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn              'Enable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOff             'SI decoders select switch 1.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOff             'SI decoders select switch 1.
    
    'Measured current should be leakage level… no alarms.
    MUX2_S1_OK = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)

    ' MUX1:S1:EN control
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOff         'Disable all S2 decoders. Should cause an alarm.
        
    For Each site In TheExec.Sites.Selected
        MUX2_S1_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_F_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn          'Re-Enable all S2 decoders. Should clear the alarm condition.
    TheHdw.DCVI.Pins("ATX_F_VI").AlarmClear
    'Make sure alarm went away.
    
    

    'Configure the Decoder S2 to be enabled and set to switch 2: EN=1, A1=0, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn              'Enable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOff             'SI decoders select switch 1a.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOn              'SI decoders select switch 1a.

    'Measured current should be leakage level… no alarms.
    MUX2_S2_OK = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)

    ' MUX2:S2:EN control
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOff         'Disable all S1 decoders. Should cause an alarm.
        
    For Each site In TheExec.Sites.Selected
        MUX2_S2_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_F_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn          'Re-Enable all S2 decoders. Should clear the alarm condition.
    TheHdw.DCVI.Pins("ATX_F_VI").AlarmClear
    'Make sure alarm went away.
       
    'Configure the Decoder S2 to be enabled and set to switch 3: EN=1, A1=1, A0=0 (Sheet 21).
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn              'Enable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOn              'SI decoders select switch 1a.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOff             'SI decoders select switch 1a.

    'Measured current should be leakage level… no alarms.
    MUX2_S3_OK = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)

    ' MUX1:S3:EN control
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOff         'Disable all S1 decoders. Should cause an alarm.
        
    For Each site In TheExec.Sites.Selected
        MUX2_S3_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_F_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn          'Re-Enable all S1 decoders. Should clear the alarm condition.
    TheHdw.DCVI.Pins("ATX_F_VI").AlarmClear
    'Make sure alarm went away.
    
    'Configure the Decoder S2 to be enabled and set to switch 4: EN=1, A1=1, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn              'Enable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOn              'SI decoders select switch 1a.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOn              'SI decoders select switch 1a.

    'Measured current should be leakage level… no alarms.
    MUX2_S4_OK = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)

    ' MUX2:S4:EN control
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOff         'Should cause an alarm.
        
    For Each site In TheExec.Sites.Selected
        MUX2_S4_EN_alarm(site) = TheHdw.DCVI.Pins("ATX_F_VI").Alarm(tlDCVSAlarmOpenKelvinDUT)
    Next site
    
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn          'Should clear the alarm condition.
    TheHdw.DCVI.Pins("ATX_F_VI").AlarmClear
    'Make sure alarm went away.
    
    'Clean up
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2,K_S2_EN,K_S2_A1,K_S2_A0").State = tlUtilBitOff

    With TheHdw.DCVI.Pins("ATX_F_VI")
        .Gate = False
        .Disconnect
    End With

    
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                MUX2_S1_OK.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX2_S2_OK.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX2_S3_OK.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                MUX2_S4_OK.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                
                MUX2_S1_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
                MUX2_S2_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
                MUX2_S3_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
                MUX2_S4_EN_alarm(site) = 1      ' 1 = True = alarm is ringing.
          Next i
        Next site
    End If
    
   
    
    'Data log results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            If site = 0 Or site = 1 Or site = 2 Or site = 3 Or site = 4 Or site = 5 Or site = 6 Or site = 14 Or site = 15 Or site = 16 Or site = 17 Or site = 18 Or site = 19 Or site = 20 Then
                results = MUX2_S1_OK.Pins(PinArr(i)).value
            TheExec.Flow.TestLimitIndex = 0
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:1_Conn"
                results = MUX2_S1_EN_alarm(site)
            TheExec.Flow.TestLimitIndex = 1
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:1_Open"
                results = MUX2_S2_OK.Pins(PinArr(i)).value
            TheExec.Flow.TestLimitIndex = 0
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:2_Conn"
                results = MUX2_S2_EN_alarm(site)
            TheExec.Flow.TestLimitIndex = 1
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:2_Open"
            End If
            
            If site = 7 Or site = 8 Or site = 9 Or site = 10 Or site = 11 Or site = 12 Or site = 13 Or site = 21 Or site = 22 Or site = 23 Or site = 24 Or site = 25 Or site = 26 Or site = 27 Then
                results = MUX2_S3_OK.Pins(PinArr(i)).value
            TheExec.Flow.TestLimitIndex = 0
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:3_Conn"
                results = MUX2_S3_EN_alarm(site)
            TheExec.Flow.TestLimitIndex = 1
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:3_Open"
                results = MUX2_S4_OK.Pins(PinArr(i)).value
            TheExec.Flow.TestLimitIndex = 0
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitVolt, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:4_Conn"
                results = MUX2_S4_EN_alarm(site)
            TheExec.Flow.TestLimitIndex = 1
                If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, unit:=unitNone, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), TName:="UC07_MUX:4_Open"
            End If
       Next i
    Next site

End Function


Public Function UC08_ADG5412BRUZ_SWITCHES()
' 1. Set the VI80 inputs to the differential op amp U2 such that the output is 3V (Sheet 16). It is only a test voltage
'    and the level is not critical. It will be used to probe relay states.
' 2. Configure 17.VI80_FRC21 to be a simple voltmeter that uses the 17.VI80_SNS21 line to make parametric voltage measurement.
' 3. Set the S1 Decoders and the ADG5412BRUZ Switch settings per Table 1 below.
' 4. There should be a change in the measured voltage with changes in the switch state.
' 5. Build test based on results.
' 6. Thresholds TBD.
' 7. Concerns: floating node voltage when switch is opened. IF voltage does not work, then we may need to do current measurements.
' 8. Implement for all sites.

' The tests were actually done in UC14 to save time.  We only datalog here.

    Dim i As Integer
    Dim site As Variant
    Dim PinArr() As String, PinCount As Long
    Dim results As Double
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_S_VI", PinArr, PinCount


    'Data log results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            results = (ADC_CH6_K1_ON.Pins(PinArr(i)).value - ADG_NC_Grp1.Pins(PinArr(i)).value)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitVolt, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="UC8_BRUZ_MUX1:3"
                                                                
            results = (ADC_CH7_K1_ON.Pins(PinArr(0)).value - ADG_NC_Grp2.Pins(PinArr(0)).value)
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitVolt, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="UC8_BRUZ_MUX1:4"
        Next i
    Next site

End Function
Public Function UC09_VS256_Cap_Leak(GrpA As PinList)
' 1. Unmerge all UVS256 channels (configure in the Channel Map).
' 2. For each VS_X supply (1 through 12) disconnect all but one UVS256 channel.
' 3. Force voltage using the attached USV256 resource. Let the capacitor banks charge and settle.
' 4. Use the YVS256 resource to measure leakage current.
' 5. Repeat for all other UVS256 channels that are ganged together on the DIB (merged in the IC test application). This verifies UVS256 channel connections to DIB.
' 6. Build test based on results.
' 7. Thresholds TBD.
' 8. Implement for all sites.
    
    Dim i As Integer
    Dim site As Variant
    Dim PinArr_a() As String, PinCount_a As Long
    Dim PinArr_b() As String, PinCount_b As Long
    Dim PinArr_c() As String, PinCount_c As Long
    Dim PinArr_d() As String, PinCount_d As Long

    Dim VDD_LEAK_a As New PinListData
    Dim VDD_LEAK_b As New PinListData
    Dim VDD_LEAK_c As New PinListData
    Dim VDD_LEAK_d As New PinListData
    
    Dim wait As Double
    
    wait = 0.1
        
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList GrpA, PinArr_a, PinCount_a
    TheExec.DataManager.DecomposePinList "DCVS_PINSb", PinArr_b, PinCount_b
    TheExec.DataManager.DecomposePinList "DCVS_PINSc", PinArr_c, PinCount_c
    TheExec.DataManager.DecomposePinList "DCVS_PINSd", PinArr_d, PinCount_d
    
'      'This sets the capacitance compensation bandwidth.
'    TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd").Gate = False
'    TheHdw.DCVS.Pins("VS_Caps_6").BandwidthSetting.value = 6
'    TheHdw.DCVS.Pins("VS_Caps_5").BandwidthSetting.value = 5
'    TheHdw.DCVS.Pins("VS_Caps_4").BandwidthSetting.value = 4
     
  ' Setup all supplies merged and un-merged.
    With TheHdw.DCVS.Pins(GrpA)
        .Gate = False
        .Disconnect (tlDCVSConnectDefault)
    End With
        
    With TheHdw.DCVS.Pins(GrpA)
     .Gate = False ' You must turn off gate to change fold limit timeout value
     .BleederResistor = tlDCVSOff
     .CurrentRange.value = 0.2
     .CurrentLimit.Source.FoldLimit.Level.value = 0.2
     .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
     .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
     .CurrentLimit.Sink.FoldLimit.Level.value = 0.075
     .CurrentLimit.Sink.FoldLimit.TimeOut.value = 2
     .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
     .LocalCapacitor = tlDCVSOff
     .Voltage.Main.value = 3.3
     .Voltage.Alt.value = 3.3
     .Voltage.output = tlDCVSVoltageMain
     .Mode = tlDCVSModeVoltage
     .Connect (tlDCVSConnectDefault)
     .Alarm(tlDCVSAlarmAll) = tlAlarmOff
     .Gate = True
    End With
    
    TheHdw.wait (wait)

    'Measured current should be leakage level… no alarms.
    TheHdw.DCVS.Pins(GrpA).Alarm(tlDCVSAlarmAll) = tlAlarmDefault

    TheHdw.DCVS.Pins(GrpA).CurrentRange.value = 0.2
    VDD_LEAK_a = TheHdw.DCVS.Pins(GrpA).Meter.Read(tlStrobe, 100)
    
    With TheHdw.DCVS.Pins(GrpA)
        .Gate = False
        .Disconnect (tlDCVSConnectDefault)
    End With
    
    ' Setup all supplies merged and un-merged.
    With TheHdw.DCVS.Pins("DCVS_PINSb")
        .Gate = False ' You must turn off gate to change fold limit timeout value
        .BleederResistor = tlDCVSOff
        .CurrentRange.value = 0.2
        .CurrentLimit.Source.FoldLimit.Level.value = 0.2
        .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .CurrentLimit.Sink.FoldLimit.Level.value = 0.075
        .CurrentLimit.Sink.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .LocalCapacitor = tlDCVSOff
        .Voltage.Main.value = 3.3
        .Voltage.Alt.value = 3.3
        .Voltage.output = tlDCVSVoltageMain
        .Mode = tlDCVSModeVoltage
        .Connect (tlDCVSConnectDefault)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
        .Gate = True
    End With
        
    TheHdw.wait (wait)
        
    'Measured current should be leakage level… no alarms.
    TheHdw.DCVS.Pins("DCVS_PINSb").CurrentRange.value = 0.002
    VDD_LEAK_b = TheHdw.DCVS.Pins("DCVS_PINSb").Meter.Read(tlStrobe, 100)
    
'     With TheHdw.DCVS.Pins("DCVS_PINSb")
'        .Gate = False ' You must turn off gate to change fold limit timeout value
'        .CurrentRange.value = 0.2
'        .Voltage.Main.value = 0#
'        .Connect (tlDCVSConnectDefault)
'        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
'        .Gate = True
'    End With
        
     With TheHdw.DCVS.Pins("DCVS_PINSb")
        .Gate = False
        .Disconnect (tlDCVSConnectDefault)
    End With
    
   
    TheHdw.wait (wait)
        
    
    
  ' Setup all supplies merged and un-merged.
'    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
    With TheHdw.DCVS.Pins("DCVS_PINSc")
        .Gate = False ' You must turn off gate to change fold limit timeout value
        .BleederResistor = tlDCVSOff
        .CurrentRange.value = 0.2
        .CurrentLimit.Source.FoldLimit.Level.value = 0.2
        .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .CurrentLimit.Sink.FoldLimit.Level.value = 0.075
        .CurrentLimit.Sink.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .LocalCapacitor = tlDCVSOff
        .Voltage.Main.value = 3.3
        .Voltage.Alt.value = 3.3
        .Voltage.output = tlDCVSVoltageMain
        .Mode = tlDCVSModeVoltage
        .Connect (tlDCVSConnectDefault)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
        .Gate = True
    End With
        
    TheHdw.wait (wait)
        
    'Measured current should be leakage level… no alarms.
    TheHdw.DCVS.Pins("DCVS_PINSc").CurrentRange.value = 0.002
    VDD_LEAK_c = TheHdw.DCVS.Pins("DCVS_PINSc").Meter.Read(tlStrobe, 100)
    
'     With TheHdw.DCVS.Pins("DCVS_PINSc")
'        .Gate = False ' You must turn off gate to change fold limit timeout value
'        .CurrentRange.value = 0.2
'        .Voltage.Main.value = 0#
'        .Connect (tlDCVSConnectDefault)
'        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
'        .Gate = True
'    End With
        
    With TheHdw.DCVS.Pins("DCVS_PINSc")
        .Gate = False
        .Disconnect (tlDCVSConnectDefault)
    End With
    
    
    TheHdw.wait (wait)
        
    
  ' Setup all supplies merged and un-merged.
'    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
    With TheHdw.DCVS.Pins("DCVS_PINSd")
        .Gate = False ' You must turn off gate to change fold limit timeout value
        .BleederResistor = tlDCVSOff
        .CurrentRange.value = 0.2
        .CurrentLimit.Source.FoldLimit.Level.value = 0.2
        .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .CurrentLimit.Sink.FoldLimit.Level.value = 0.075
        .CurrentLimit.Sink.FoldLimit.TimeOut.value = 2
        .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
        .LocalCapacitor = tlDCVSOff
        .Voltage.Main.value = 3.3
        .Voltage.Alt.value = 3.3
        .Voltage.output = tlDCVSVoltageMain
        .Mode = tlDCVSModeVoltage
        .Connect (tlDCVSConnectDefault)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
        .Gate = True
    End With
        
    TheHdw.wait (wait)
        
    'Measured current should be leakage level… no alarms.
    TheHdw.DCVS.Pins("DCVS_PINSd").CurrentRange.value = 0.002
    VDD_LEAK_d = TheHdw.DCVS.Pins("DCVS_PINSd").Meter.Read(tlStrobe, 100)
    
     With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
        .Gate = False ' You must turn off gate to change fold limit timeout value
        .CurrentRange.value = 0.2
        .Voltage.Main.value = 0#
        .Connect (tlDCVSConnectDefault)
        .Alarm(tlDCVSAlarmAll) = tlAlarmOff
        .Gate = True
    End With
        
    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
        .Gate = False
        .Disconnect (tlDCVSConnectDefault)
    End With
    
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount_a - 1
                VDD_LEAK_a.Pins.Item(i).value = 0.000001 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_b - 1
                VDD_LEAK_b.Pins.Item(i).value = 0.000002 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_c - 1
                VDD_LEAK_c.Pins.Item(i).value = 0.000003 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_d - 1
                VDD_LEAK_d.Pins.Item(i).value = 0.000004 + (Rnd * 0.000001)
            Next i
        Next site
    End If
    
    'Data log results
    For Each site In TheExec.Sites.Selected
            TheExec.Flow.TestLimitIndex = 0
            
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_c.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_c(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_d.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_d(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(1).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(1), TName:="UC09_VS2_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(2).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(2), TName:="UC09_VS3_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(3).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(3), TName:="UC09_VS4_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(4).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(4), TName:="UC09_VS5_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(1).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(1), TName:="UC09_VS5_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(5).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(5), TName:="UC09_VS6_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(6).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(6), TName:="UC09_VS7_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(7).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(7), TName:="UC09_VS8_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(8).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(8), TName:="UC09_VS9_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(9).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(9), TName:="UC09_VS10_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(2).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(2), TName:="UC09_VS10_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(10).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(10), TName:="UC09_VS11_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(3).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(3), TName:="UC09_VS11_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(11).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(11), TName:="UC09_VS12_Leak"
     Next site
    

End Function
Public Function old_UC09_VS256_Cap_Leak()
' 1. Unmerge all UVS256 channels (configure in the Channel Map).
' 2. For each VS_X supply (1 through 12) disconnect all but one UVS256 channel.
' 3. Force voltage using the attached USV256 resource. Let the capacitor banks charge and settle.
' 4. Use the YVS256 resource to measure leakage current.
' 5. Repeat for all other UVS256 channels that are ganged together on the DIB (merged in the IC test application). This verifies UVS256 channel connections to DIB.
' 6. Build test based on results.
' 7. Thresholds TBD.
' 8. Implement for all sites.
    Dim i As Integer
    Dim site As Variant
    Dim PinArr_a() As String, PinCount_a As Long
    Dim PinArr_b() As String, PinCount_b As Long
    Dim PinArr_c() As String, PinCount_c As Long
    Dim PinArr_d() As String, PinCount_d As Long

    Dim VDD_LEAK_a As New PinListData
    Dim VDD_LEAK_b As New PinListData
    Dim VDD_LEAK_c As New PinListData
    Dim VDD_LEAK_d As New PinListData
        
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "DCVS_PINSa", PinArr_a, PinCount_a
    TheExec.DataManager.DecomposePinList "DCVS_PINSb", PinArr_b, PinCount_b
    TheExec.DataManager.DecomposePinList "DCVS_PINSc", PinArr_c, PinCount_c
    TheExec.DataManager.DecomposePinList "DCVS_PINSd", PinArr_d, PinCount_d
    
    ' Setup all supplies merged and un-merged.
    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
        .Gate = False
        .Mode = tlDCVSModeVoltage
        .Voltage.Main.value = 0
        .SetCurrentRanges 0.2, 0.2
        .CurrentLimit.Source.FoldLimit.Level = 0.2
        .CurrentLimit.Source.FoldLimit.TimeOut = 2
        .CurrentLimit.Sink.FoldLimit.Level = 0.05
        .CurrentLimit.Sink.FoldLimit.TimeOut = 2
        .Meter.Mode = tlDCVSMeterCurrent
        .Connect (tlDCVSConnectDefault)
        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
        .Gate = True
    End With
   
    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
        .Gate = False
    End With
   
   
   
   
   
    ' Look at the first supply connection to node.
    With TheHdw.DCVS.Pins("DCVS_PINSa")     ' Look at the first supply connection to node.
        .Voltage.Main.value = 2.99999713897705
        .Gate = True
    End With
      
    TheHdw.wait (1)

    'Measured current should be leakage level… no alarms.
    VDD_LEAK_a = TheHdw.DCVS.Pins("DCVS_PINSa").Meter.Read(tlStrobe, 100)
            
            
            
            
            
            
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSa")     'Disconnect first first supply to node.
        .Voltage.Main.value = 0#
        .Gate = False
    End With
    
    
    ' Look at the 2nd merged supply connection to node.
    With TheHdw.DCVS.Pins("DCVS_PINSb")
        .Voltage.Main.value = 2.99999713897705
        .Gate = True
    End With
      
    TheHdw.wait (1)
    
    'Measured current should be leakage level… no alarms.
    VDD_LEAK_b = TheHdw.DCVS.Pins("DCVS_PINSb").Meter.Read(tlStrobe, 100)
    
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSb")     'Disconnect first first supply to node.
        .Voltage.Main.value = 0#
        .Gate = False
    End With
      
        
    ' Look at the 3rd merged supply connection to node.
    With TheHdw.DCVS.Pins("DCVS_PINSc")
        .Voltage.Main.value = 2.99999713897705
        .Gate = True
    End With
      
    TheHdw.wait (1)

    'Measured current should be leakage level… no alarms.
    VDD_LEAK_c = TheHdw.DCVS.Pins("DCVS_PINSc").Meter.Read(tlStrobe, 100)
    
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSc")     'Disconnect first first supply to node.
        .Voltage.Main.value = 0#
        .Gate = False
    End With
   
    
    ' Look at the 4th merged supply connection to node.
    With TheHdw.DCVS.Pins("DCVS_PINSd")
        .Voltage.Main.value = 2.99999713897705
        .Gate = True
    End With
      
    TheHdw.wait (1)

    'Measured current should be leakage level… no alarms.
    VDD_LEAK_d = TheHdw.DCVS.Pins("DCVS_PINSd").Meter.Read(tlStrobe, 100)
    
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSa,DCVS_PINSb,DCVS_PINSc,DCVS_PINSd")
        .Gate = False
        .Disconnect
    End With
    
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount_a - 1
                VDD_LEAK_a.Pins.Item(i).value = 0.000001 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_b - 1
                VDD_LEAK_b.Pins.Item(i).value = 0.000002 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_c - 1
                VDD_LEAK_c.Pins.Item(i).value = 0.000003 + (Rnd * 0.000001)
            Next i
            For i = 0 To PinCount_d - 1
                VDD_LEAK_d.Pins.Item(i).value = 0.000004 + (Rnd * 0.000001)
            Next i
        Next site
    End If
    
    'Data log results
    For Each site In TheExec.Sites.Selected
            TheExec.Flow.TestLimitIndex = 0
            
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_c.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_c(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_d.Pins.Item(0).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_d(0), TName:="UC09_VS1_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(1).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(1), TName:="UC09_VS2_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(2).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(2), TName:="UC09_VS3_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(3).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(3), TName:="UC09_VS4_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(4).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(4), TName:="UC09_VS5_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(1).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(1), TName:="UC09_VS5_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(5).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(5), TName:="UC09_VS6_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(6).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(6), TName:="UC09_VS7_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(7).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(7), TName:="UC09_VS8_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(8).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(8), TName:="UC09_VS9_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(9).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(9), TName:="UC09_VS10_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(2).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(2), TName:="UC09_VS10_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(10).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(10), TName:="UC09_VS11_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_b.Pins.Item(3).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_b(3), TName:="UC09_VS11_Leak"
                TheExec.Flow.TestLimit ResultVal:=VDD_LEAK_a.Pins.Item(11).value, unit:=unitAmp, forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr_a(11), TName:="UC09_VS12_Leak"
     Next site
    

End Function

Public Function UC10_All_S2_Decoder_Leak()
' 1. Repeat for all four states of the S2 Decoder (Sheet 103): using the VI80 resource attached to DA/DB of the Decoder, force 5V and measure leakage current.
' 2. All K2_XX relays should be open (not actuated).
' 3.
' 4. Build test based on results.
' 5. Thresholds TBD.
' 6. Implement for all sites.

   Dim i As Integer
    Dim site As Variant
    Dim MUX2_S1_LEAKAGE As New PinListData
    Dim MUX2_S2_LEAKAGE As New PinListData
    Dim MUX2_S3_LEAKAGE As New PinListData
    Dim MUX2_S4_LEAKAGE As New PinListData
    Dim Force_Hi As Double
    
    Dim PinArr() As String, PinCount As Long

    Force_Hi = 5#
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_F_VI", PinArr, PinCount
    
    With TheHdw.DCVI.Pins("ATX_F_VI")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = Force_Hi
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 0.02
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With
    
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOn              'Enable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOff             'SI decoders select switch 1.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOff             'SI decoders select switch 1.
    MUX2_S1_LEAKAGE = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)
    
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOff             'SI decoders select switch 2.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOn              'SI decoders select switch 2.
    MUX2_S2_LEAKAGE = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)
        
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOn              'SI decoders select switch 3.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOff             'SI decoders select switch 3.
    MUX2_S3_LEAKAGE = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)
    
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOn              'SI decoders select switch 4.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOn              'SI decoders select switch 4.
    MUX2_S4_LEAKAGE = TheHdw.DCVI.Pins("ATX_F_VI").Meter.Read(tlStrobe, 1)
                
    'Clean up
    TheHdw.Utility.Pins("K_S2_EN").State = tlUtilBitOff             'Disable all S2 decoders.
    TheHdw.Utility.Pins("K_S2_A1").State = tlUtilBitOff             'SI decoders select switch 1.
    TheHdw.Utility.Pins("K_S2_A0").State = tlUtilBitOff             'SI decoders select switch 1.

    With TheHdw.DCVI.Pins("ATX_S_VI")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
   
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                MUX2_S1_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
                MUX2_S2_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
                MUX2_S3_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
                MUX2_S4_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
            Next i
        Next site
    End If
    
    'Datalog results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX2_S1_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC10_MUX2_S1_LEAK"
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX2_S2_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC10_MUX2_S2_LEAK"
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX2_S3_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC10_MUX2_S3_LEAK"
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX2_S4_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC10_MUX2_S4_LEAK"
        Next i
    Next site
   


End Function

Public Function UC11_All_S1_Decoder_Leak()
' 1. Apply to only switch 1 and 2 positions of the S1 Decoder (Sheet 21): using the VI80 resource attached to DA/DB of the Decoder, force 5V and
'    measure leakage current.
' 2. Build test based on results.
' 3. Thresholds TBD.
' 4. Implement for all sites.

   Dim i As Integer
    Dim site As Variant
    Dim MUX1_S1_LEAKAGE As New PinListData
    Dim MUX1_S2_LEAKAGE As New PinListData
    Dim MUX1_S3_LEAKAGE As New PinListData
    Dim MUX1_S4_LEAKAGE As New PinListData
    Dim Force_Hi As Double
    
    Dim PinArr() As String, PinCount As Long

    Force_Hi = 5#
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "ATX_S_VI", PinArr, PinCount
    
    With TheHdw.DCVI.Pins("ATX_S_VI")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = Force_Hi
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 0.02
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With
    
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOn              'Enable all S1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff             'SI decoders select switch 1.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff             'SI decoders select switch 1.
    MUX1_S1_LEAKAGE = TheHdw.DCVI.Pins("ATX_S_VI").Meter.Read(tlStrobe, 1)
    
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff             'SI decoders select switch 2.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOn              'SI decoders select switch 2.
    MUX1_S2_LEAKAGE = TheHdw.DCVI.Pins("ATX_S_VI").Meter.Read(tlStrobe, 1)
        
    'Clean up
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff             'Disable all S1 decoders.
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff             'SI decoders select switch 1.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff             'SI decoders select switch 1.

    With TheHdw.DCVI.Pins("ATX_S_VI")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
   
    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                MUX1_S1_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
                MUX1_S2_LEAKAGE.Pins.Item(i).value = 0.00000005 + (Rnd * 0.0000000005)
            Next i
        Next site
    End If
    
    
    'Datalog results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX1_S1_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC11_MUX1_S1_LEAK"
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MUX1_S2_LEAKAGE.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC11_MUX1_S2_LEAK"
        Next i
    Next site
   



End Function

Public Function UC12_UP1600_LEAKAGE(Force_Hi As Double)
' 1. For each digital channel use the per-pin PMU to force 5V and measure leakage current.
' 2. Build test based on results.
' 3. Thresholds TBD.

    Dim i As Integer
    Dim site As Variant
    Dim MeasureIIH As New PinListData
    Dim PinArr() As String, PinCount As Long

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "LeakPins", PinArr, PinCount
    
    TheHdw.Digital.Pins("LeakPins").Disconnect
    
    With TheHdw.PPMU.Pins("LeakPins")
        .ForceV Force_Hi, 0.00002
        .Gate = tlOff
        .Connect
        .Gate = tlOn
    End With
    
    TheHdw.wait (0.01)
    
    MeasureIIH = TheHdw.PPMU.Pins("LeakPins").Read(tlPPMUReadMeasurements)

    ' the PinListData variable with simulation data
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            MeasureIIH.value = 0.00000005 + (Rnd * 0.000000005)
        Next site
    End If

    'Clean up
    TheHdw.PPMU.Pins("LeakPins").Disconnect
    TheHdw.Digital.Pins("LeakPins").Connect
    
    'Datalog results
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=MeasureIIH.Pins(PinArr(i)).value, _
                                                                unit:=unitAmp, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                ForceVal:=Force_Hi, _
                                                                TName:="UC12_IIH"
        Next i
    Next site
   

End Function

Public Function UC13_VS256_Cap_Meas_Conn(strPatSetName As pattern)
    Dim pldVDDa As New PinListData
     
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call TheHdw.DCVS.Pins("VDDa").PSets.Add("my_pset_1")
   
    With TheHdw.DCVS.Pins("VDDa").PSets.Item("my_pset_1")
      .BandwidthSetting.value = 255
      .CurrentRange.value = 0.2
      .Capture.SampleRate.value = 200000
      .Capture.SampleSize.value = 2
      .CurrentLimit.Source.FoldLimit.Level.value = 0.025
      .CurrentRange.value = 0.02
      .Voltage.Main.value = 0#
      .Voltage.Alt.value = 3.3
      .Meter.Mode = tlDCVSMeterVoltage
      .Mode = tlDCVSModeVoltage
      .Voltage.Main.value = 0#
      .Voltage.Alt.value = 3.3
    End With
     
    ' Cannot be set with PSets
    With TheHdw.DCVS.Pins("VDD")
        .CurrentLimit.Sink.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        .CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorDoNotGateOff
        .Alarm(AlarmType) = tlAlarmOff
    End With
     
    'Force a current into the caps.
    TheHdw.Patterns(".\SimpleDCVS.pat").Load
    TheHdw.Patterns(".\SimpleDCVS.pat").Start              'Run the test pattern
    TheHdw.Digital.TimeDomains("").Patgen.HaltWait
    
    pldVDDa = TheHdw.DCVS.Pins("VDDa").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    
End Function

Public Function Junk_UC13_VS256_Cap_Meas_Conn(strPatSetName As pattern)
' 1. Do this for all sites (Sheet 20).
' 2. Unmerge all UVS256 channels (configure in the Channel Map).
' 3. For each VS_X supply (1 through 12) disconnect all but one UVS256 channel.
' 4. Source smallest current using the attached USV256 resource. Make two voltage measurements at two separate times using microcode method.
' 5. Capacitors will charge quickly, so a small current and short time intervals are required.
' 6. Use C= I dt/dV to calculate total capacitance for each supply and for each site.
' 7. Repeat for all other UVS256 channels that are ganged together on the DIB (merged in the IC test application). This
'    verifies UVS256 channel connections to DIB.
' 8. Build test based on results.
' 9. Thresholds TBD (+/-% of expected value).

    Dim i As Integer
    Dim site As Variant
    Dim dT1 As New PinListData
    Dim dT2 As New PinListData
    Dim PinArr() As String, PinCount As Long
    Dim cap_value(13) As Double
    Dim hardware_change As Boolean
    
    Dim pldVDDa As New PinListData
    Dim pldMIPI As New PinListData
    Dim pldVS_3 As New PinListData
    Dim pldEFUSE As New PinListData
    Dim pldVS_5 As New PinListData
    Dim pldVS_6 As New PinListData
    Dim pldVS_7 As New PinListData
    Dim pldVREF As New PinListData
    Dim pldVADC As New PinListData
    Dim pldVS_10 As New PinListData
    Dim pldENETa As New PinListData
    Dim pldQSPI As New PinListData
    
    Dim dspVDDa As New DSPWave
    Dim dspMIPI As New DSPWave
    Dim dspVS_3 As New DSPWave
    Dim dspEFUSE As New DSPWave
    Dim dspVS_5 As New DSPWave
    Dim dspVS_6 As New DSPWave
    Dim dspVS_7 As New DSPWave
    Dim dspVREF As New DSPWave
    Dim dspADC As New DSPWave
    Dim dspVS_10 As New DSPWave
    Dim dspENETa As New DSPWave
    Dim dspQSPI As New DSPWave
   
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheExec.DataManager.DecomposePinList "DCVS_PINSb", PinArr, PinCount
       
    Call Add_DCVS_PSet("VDDa", "VDDa_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_LV_MIPI", "MIPI_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_VS_3_PRB", "VS_3_PRB_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_HV_EFUSE", "EFUSE_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_VS_5_PRBa", "VS_5_PRBa_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_VS_6_PRB", "VS_6_PRB_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_VS_7_PRB", "VS_7_PRB_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_HV_VREF", "VREF_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_HV_ADC", "ADC_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_VS_10_PRBa", "VS_10_PRBa_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_HV_IO_ENETa", "IO_ENETa_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
    Call Add_DCVS_PSet("VDD_HV_IO_QSPI", "IO_QSPI_pset_1", 255, 0.02, 20000#, 2, 0.02, tlDCVSMeterVoltage, tlDCVSModeVoltage, 3.3, 3.3)
       
   
    ' Initialize all the caps to 0 volts.
    With TheHdw.DCVS.Pins("VDDa")
       .Voltage.Main.value = 0#                     ' Main Voltage (0.0V) Bleed caps down to 0 volts
       .CurrentRange = 0.2
       .Mode = tlDCVSModeVoltage                    ' Force Voltage Mode
       .Meter.Mode = tlDCVSMeterVoltage             ' Meter in Voltage Mode
       .Connect (tlDCVSConnectDefault)              ' Connect Force and Sense
       .Gate = True                                 ' Gate Supply on
    End With
    
    TheHdw.wait (0.5)                               ' Bleed down time.
        
       
    'Force a current into the caps.
    TheHdw.Patterns(strPatSetName).Start              'Run the test pattern
    TheHdw.Digital.TimeDomains("").Patgen.HaltWait
    
    
    pldVDDa = TheHdw.DCVS.Pins("VDDa").Meter.Read(tlNoStrobe, 2, -1, tlDCVIMeterReadingFormatArray)
    
    
    pldVDDa = TheHdw.DCVS.Pins("VDDa").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldMIPI = TheHdw.DCVS.Pins("VDD_LV_MIPI").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVS_3 = TheHdw.DCVS.Pins("VDD_VS_3_PRB").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldEFUSE = TheHdw.DCVS.Pins("VDD_HV_EFUSE").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVS_5 = TheHdw.DCVS.Pins("VDD_VS_5_PRBa").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVS_6 = TheHdw.DCVS.Pins("VDD_VS_6_PRB").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVS_7 = TheHdw.DCVS.Pins("VDD_VS_7_PRB").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVREF = TheHdw.DCVS.Pins("VDD_HV_VREF").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVADC = TheHdw.DCVS.Pins("VDD_HV_ADC").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldVS_10 = TheHdw.DCVS.Pins("VDD_VS_10_PRBa").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldENETa = TheHdw.DCVS.Pins("VDD_HV_IO_ENETa").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
    pldQSPI = TheHdw.DCVS.Pins("VDD_HV_IO_QSPI").Meter.Read(tlNoStrobe, 2, , tlDCVIMeterReadingFormatArray)
   
    For Each site In TheExec.Sites
       'Move the pinlist data into DSP arrays for processing.
       dspVDDa.Data = pldVDDa.Pins("VDDa").value
       dspMIPI.Data = pldMIPI.Pins("VDD_LV_MIPI").value
       dspVS_3.Data = pldVS_3.Pins("VDD_VS_3_PRB").value
       dspEFUSE.Data = pldEFUSE.Pins("VDD_HV_EFUSE").value
       dspVS_5.Data = pldVS_5.Pins("VDD_VS_5_PRBa").value
       dspVS_6.Data = pldVS_6.Pins("VDD_VS_6_PRB").value
       dspVS_7.Data = pldVS_7.Pins("VDD_VS_7_PRB").value
       dspVREF.Data = pldVREF.Pins("VDD_HV_VREF").value
       dspADC.Data = pldVADC.Pins("VDD_HV_ADC").value
       dspVS_10.Data = pldVS_10.Pins("VDD_VS_10_PRBa").value
       dspENETa.Data = pldENETa.Pins("VDD_HV_IO_ENETa").value
       dspQSPI.Data = pldQSPI.Pins("VDD_HV_IO_QSPI").value
    Next site

     ' the PinListData variable with simulation data
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            dspVDDa.Element(0) = 0#: dspVDDa.Element(1) = 2.5
            dspMIPI.Element(0) = 0#: dspMIPI.Element(1) = 2.5
            dspVS_3.Element(0) = 0#: dspVS_3.Element(1) = 2.5
            dspEFUSE.Element(0) = 0#: dspEFUSE.Element(1) = 2.5
            dspVS_5.Element(0) = 0#: dspVS_5.Element(1) = 2.5
            dspVS_6.Element(0) = 0#: dspVS_6.Element(1) = 2.5
            dspVS_7.Element(0) = 0#: dspVS_7.Element(1) = 2.5
            dspVREF.Element(0) = 0#: dspVREF.Element(1) = 2.5
            dspADC.Element(0) = 0#: dspADC.Element(1) = 2.5
            dspVS_10.Element(0) = 0#: dspVS_10.Element(1) = 2.5
            dspENETa.Element(0) = 0#: dspENETa.Element(1) = 2.5
            dspQSPI.Element(0) = 0#: dspQSPI.Element(1) = 2.5
        Next site
    End If
    
           
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSa")
        .Gate = False
        .Disconnect
    End With

    'Datalog results
    For Each site In TheExec.Sites.Selected
        '  c      =    I       Dt  *                 Dv
        cap_value(1) = 0.00123 * 0.0914 / (dspVDDa.Element(1).value - dspVDDa.Element(0).value)
        cap_value(2) = 0.00123 * 0.0914 / (dspMIPI.Element(1).value - dspMIPI.Element(0).value)
        cap_value(3) = 0.00123 * 0.0914 / (dspVS_3.Element(1).value - dspVS_3.Element(0).value)
        cap_value(4) = 0.00123 * 0.0914 / (dspEFUSE.Element(1).value - dspEFUSE.Element(0).value)
        cap_value(5) = 0.00123 * 0.0914 / (dspVS_5.Element(1).value - dspVS_5.Element(0).value)
        cap_value(6) = 0.00123 * 0.0914 / (dspVS_6.Element(1).value - dspVS_6.Element(0).value)
        cap_value(7) = 0.00123 * 0.0914 / (dspVS_7.Element(1).value - dspVS_7.Element(0).value)
        cap_value(8) = 0.00123 * 0.0914 / (dspVREF.Element(1).value - dspVREF.Element(0).value)
        cap_value(9) = 0.00123 * 0.0914 / (dspADC.Element(1).value - dspADC.Element(0).value)
        cap_value(10) = 0.00123 * 0.0914 / (dspVS_10.Element(1).value - dspVS_10.Element(0).value)
        cap_value(11) = 0.00123 * 0.0914 / (dspENETa.Element(1).value - dspENETa.Element(0).value)
        cap_value(12) = 0.00123 * 0.0914 / (dspQSPI.Element(1).value - dspQSPI.Element(0).value)
            
        For i = 1 To 12
            TheExec.Flow.TestLimitIndex = 0
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=cap_value(i), unit:=unitCustom, customUnit:="Fd", _
                                                                forceUnit:=unitNone, ForceResults:=tlForceFlow, PinName:=PinArr(i), _
                                                                TName:="UC13_Cap_Val"
        Next i
    Next site

End Function

                 














Public Function old_UC13_VS256_Cap_Meas_Conn()
' 1. Do this for all sites (Sheet 20).
' 2. Unmerge all UVS256 channels (configure in the Channel Map).
' 3. For each VS_X supply (1 through 12) disconnect all but one UVS256 channel.
' 4. Source smallest current using the attached USV256 resource. Make two voltage measurements at two separate times using microcode method.
' 5. Capacitors will charge quickly, so a small current and short time intervals are required.
' 6. Use C= I dt/dV to calculate total capacitance for each supply and for each site.
' 7. Repeat for all other UVS256 channels that are ganged together on the DIB (merged in the IC test application). This
'    verifies UVS256 channel connections to DIB.
' 8. Build test based on results.
' 9. Thresholds TBD (+/-% of expected value).

    Dim i As Integer
    Dim site As Variant
    Dim dT1 As New PinListData
    Dim dT2 As New PinListData
    Dim PinArr() As String, PinCount As Long
    Dim cap_value As Double
    Dim hardware_change As Boolean
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "DCVS_PINSb", PinArr, PinCount
   
    ' Initialize all the caps to 0 volts.
    With TheHdw.DCVS.Pins("DCVS_PINSb")
       .Voltage.Main.value = 0#                     ' Main Voltage (0.0V) Bleed caps down to 0 volts
       .CurrentRange = 0.2
       .Mode = tlDCVSModeVoltage                    ' Force Voltage Mode
       .Meter.Mode = tlDCVSMeterVoltage             ' Meter in Voltage Mode
       .Connect (tlDCVSConnectDefault)              ' Connect Force and Sense
       .Gate = True                                 ' Gate Supply on
    End With
    
    TheHdw.wait (0.5)                               ' Bleed down time.
        
       
    'Force a current into the caps.
    With TheHdw.DCVS.Pins("DCVS_PINSa")
       .Voltage.Main.value = 3.3                    ' Main Voltage (3V)
       .Gate = False ' You must gate off to change fold limit timeout value
       .BandwidthSetting.value = 255
       .CurrentRange = 0.004
       .CurrentLimit.Source.FoldLimit.Level.value = 0.004
       .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
       '.CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOn
       .Mode = tlDCVSModeVoltage                    ' Force Voltage Mode
       .Meter.Mode = tlDCVSMeterVoltage             ' Meter in Voltage Mode
       .Connect (tlDCVSConnectDefault)              ' Connect Force and Sense
       .Gate = True                                 ' Gate Supply on
    End With
    
    ' - - - Take Measurement - - -
    ' Meter strobes rate=25K samples/sec. and reads back an average of 20 samples?
    dT1 = TheHdw.DCVS.Pins("DCVS_PINSa").Meter.Read(tlStrobe, 1)              'Get the initial voltage value of the caps.
    
    
    TheHdw.DCVS.Pins("DCVS_PINSa").Voltage.Main.value = 0#  ' Main Voltage (0.0V) Bleed caps down to 0 volts
    TheHdw.wait (0.5)                               ' Bleed down time.
    TheHdw.DCVS.Pins("DCVS_PINSa").Voltage.Main.value = 3.3 ' Main Voltage (0.0V) Bleed caps down to 0 volts
    TheHdw.wait (0.005)                              ' 2nd measurement delay time.
    dT2 = TheHdw.DCVS.Pins("DCVS_PINSa").Meter.Read(tlStrobe, 1)              'Get the delta voltage value of the caps.
   
    
 
'
'    TheHdw.wait (0#)                               'Charging up time.
'    ' - - - Take Measurement - - -
'    ' Meter strobes rate=25K samples/sec. and reads back an average of 20 samples?
'    dT2 = TheHdw.DCVS.Pins("DCVS_PINSa").Meter.Read(tlStrobe, 1)              'Get the delta voltage value of the caps.
'
    
    
     ' the PinListData variable with simulation data
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            dT1.value = 0#
        Next site
    End If
    
   ' the PinListData variable with simulation data
   
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            dT2.Pins(0).value = 4.1 + CDbl(site) * 0.001
            dT2.Pins(1).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(2).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(3).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(4).value = 2.1 + CDbl(site) * 0.001
            dT2.Pins(5).value = 1.5 + CDbl(site) * 0.001
            dT2.Pins(6).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(7).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(8).value = 1.1 + CDbl(site) * 0.001
            dT2.Pins(9).value = 2.2 + CDbl(site) * 0.001
            dT2.Pins(10).value = 2.1 + CDbl(site) * 0.001
            dT2.Pins(11).value = 1.1 + CDbl(site) * 0.001
        Next site
    End If
         
'    hardware_change = TheHdw.DCVS.Pins("DCVS_PINSa").CurrentLimit.Source.FoldLimit.LevelChanged
    
           
    'Clean up
    With TheHdw.DCVS.Pins("DCVS_PINSa")
        .Gate = False
        .Disconnect
    End With

    'Datalog results
    For Each site In TheExec.Sites.Selected
        TheExec.Flow.TestLimitIndex = 0

        For i = 0 To PinCount - 1
            '  c      =    I       Dt  *                 Dv
            cap_value = 0.00123 * 0.0914 / (dT2.Pins(i).value - dT1.Pins(i).value)
            
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=cap_value, _
                                                                unit:=unitCustom, _
                                                                customUnit:="Fd", _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="UC13_Cap_Val"
        Next i
    Next site

End Function

Public Function UC14_OpAmps_SAR_ADC_INPUTS()
    Dim i As Integer
    Dim site As Variant
    Dim results As Double
    Dim PinArr() As String, PinCount As Long
    
    TheExec.DataManager.DecomposePinList "atx_s_vi", PinArr, PinCount
    
'************* 1. Set the VI80 inputs to the differential op amp U2 such that the output is 3V (Sheet 16).
    With TheHdw.DCVI.Pins("sar_cal,sar_cal2")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .CurrentRange = 0.02
        .Voltage = 3#
        .Current = 0.01
        .Gate = False
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .Connect tlDCVIConnectDefault
        .Meter.HardwareAverage = 1
        .Gate = True
    End With

    TheHdw.DCVI.Pins("sar_cal2").Voltage = 0#
    

'************* 4. Configure 17.VI80_SNS21 for voltmeter mode (Sheet21).
    With TheHdw.DCVI.Pins("atx_s_vi")
        .Mode = tlDCVIModeCurrent
        .CurrentRange = 0.02
        .CurrentRange.Autorange = True
        .VoltageRange.Autorange = True
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .Voltage = 0
        .Current = 0
        .Gate = False
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .Disconnect tlDCVIConnectHighForce
        .Connect tlDCVIConnectHighSense
        .BleederResistor = tlDCVIBleederResistorOn
        .Meter.HardwareAverage = 1
    End With


'************* 2. Set S3_00/07 to close all SPST switches (Sheet 103).  S3 is normally connected.
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2").State = tlUtilBitOff

'************* 3. Switch K1 to connect the VI80_32 resource (Sheet 16). K1 is normally open.
    TheHdw.Utility.Pins("K_VI_ADJUST").State = tlUtilBitOn

'************* Step 5. Configure the Decoder S1_00 to be enabled and set to switch 3: EN=1, A1=1, A0=0 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_EN").State = tlUtilBitOff  'Enable the switch.    = Vih
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff  'Select mux channel 3. = Vih
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOn                         '= Vil

'************* Step 6. Measure ADC_CH6_F_00 voltage using 17.VI80_SNS21.  Should see 3 Volts.
    ADC_CH6_K1_ON = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    
'************* We are setup to do the K1 test here so why not do it now?
'************* 15. Switch K1 to ground. 16. Disconnect VI80_27. U2 output should be near 0V. Repeat all measurements.
    With TheHdw.DCVI.Pins("sar_cal")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
    
    TheHdw.Utility.Pins("K_VI_ADJUST").State = tlUtilBitOff
    
    ADC_CH6_K1_OFF = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    
    TheHdw.Utility.Pins("K_VI_ADJUST").State = tlUtilBitOn
    
    With TheHdw.DCVI.Pins("sar_cal")
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
'************* We are setup to do the ADG5412BRUZ test here so why not do it now?
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2").State = tlUtilBitOn
    ADG_NC_Grp1 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2").State = tlUtilBitOff
    

'************* Step 7. Configure the Decoder S1_00 to be enabled and set to switch 4: EN=1, A1=1, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff  'Select mux channel 4.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff  'Select mux channel 4.

'************* Step 8. Measure ADC_CH7_F_00 voltage using 17.VI80_SNS21.  Should see 3 Volts.
   ADC_CH7_K1_ON = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
   
   
   

'************* 15. Switch K1 to ground. 16. Disconnect VI80_32. U2 output should be near 0V. Repeat all measurements.
'************* Step 18. Measure ADC_CH7_F_00 voltage using 17.VI80_SNS21 (atx_s_vi).
    With TheHdw.DCVI.Pins("sar_cal")
        .Gate = False
        .Disconnect tlDCVIConnectDefault
    End With
    
    TheHdw.Utility.Pins("K_VI_ADJUST").State = tlUtilBitOff

    ADC_CH7_K1_OFF = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
   
    TheHdw.Utility.Pins("K_VI_ADJUST").State = tlUtilBitOn
    
    With TheHdw.DCVI.Pins("sar_cal")
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
   
'************* We are setup to do the ADG5412BRUZ test here so why not do it now?
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2").State = tlUtilBitOn
    ADG_NC_Grp2 = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2").State = tlUtilBitOff


'************* Step 17. Configure the Decoder S1_00 to be enabled and set to switch 4: EN=1, A1=1, A0=1 (Sheet 21).
    TheHdw.Utility.Pins("K_S1_A1").State = tlUtilBitOff  'Select mux channel 4.
    TheHdw.Utility.Pins("K_S1_A0").State = tlUtilBitOff

'************* Step 18. Measure ADC_CH7_F_00 voltage using 17.VI80_SNS21 (atx_s_vi).
    ADC_CH7_K1_OFF = TheHdw.DCVI.Pins("atx_s_vi").Meter.Read(tlStrobe, 1)



    'Clean up
    TheHdw.Utility.Pins("K_S3_Grp1,K_S3_Grp2,K_S1_EN,K_S1_A1,K_S1_A0,K_VI_ADJUST").State = tlUtilBitOff

    With TheHdw.DCVI.Pins("sar_cal2,sar_cal,atx_s_vi")
        .Gate = False
        .Disconnect
    End With


    ' Create off-line data.
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            For i = 0 To PinCount - 1
                ADC_CH6_K1_ON.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                ADC_CH7_K1_ON.Pins(PinArr(i)).value = -0.72 + (Rnd * 0.01)
                ADC_CH6_K1_OFF.Pins(PinArr(i)).value = 0.001 + (Rnd * 0.01)
                ADC_CH7_K1_OFF.Pins(PinArr(i)).value = 0.001 + (Rnd * 0.02)
                ADG_NC_Grp1.Pins(PinArr(i)).value = 0.001 + (Rnd * 0.01)
                ADG_NC_Grp2.Pins(PinArr(i)).value = 0.001 + (Rnd * 0.01)
           Next i
        Next site
    End If




    'Datalog results
    'This tests that all relays in the path, K1, K_S3_Grp1,K_S3_Grp2, and K_S1 contols can connect.
    For Each site In TheExec.Sites.Selected
        For i = 0 To PinCount - 1
            TheExec.Flow.TestLimitIndex = 0
            results = ADC_CH6_K1_ON.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitVolt, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="SAR_Mux1:3"
                                                                
    'This tests that K1 is operational.
            results = ADC_CH7_K1_ON.Pins(PinArr(0)).value
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=results, _
                                                                unit:=unitVolt, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=PinArr(i), _
                                                                TName:="SAR_Mux1:4"
        Next i
    Next site

End Function
Public Function UC15_Impedance_Resp_Profile(numsamps As Long, period As Double, timesteps As Long, voltstrt As Double, voltstop As Double, voltsteps As Long, pgroup As String, patname As String, Optional strtlab As String)
    Dim fcpl As New PinListData
    Dim vincr As Double, tincr As Double, currv As Double, currt As Double
    Dim vcnt As Long, tcnt As Long, pcnt As Long, fcnt As Long, numpins As Long
    Dim plist() As String, output As String, datFile As String, pstrng As String
    Dim tSite As Variant

    If TheExec.TesterMode <> testModeOffline Then
        On Error GoTo errHandler
        
        ' load levels and timing
        TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
        ' ensure pattern is loaded
        TheHdw.Patterns(patname).Load
        
        ' calculate time increment - can move drive edge 4 periods
        If timesteps <> 0 Then
            tincr = (period * 3) / timesteps
        Else
            MsgBox ("Time Steps must not be zero, aborting")
            Exit Function
        End If
        
        ' calculate voltage incement
        If voltsteps <> 0 Then
            vincr = (voltstop - voltstrt) / voltsteps
        Else
            MsgBox ("Voltage Steps must not be zero, aborting")
            Exit Function
        End If
        
        ' set pattern loop counter to number of samples per shmoo point
        If numsamps <> 0 Then
            TheHdw.Digital.Patgen.Counter(tlPgCounter4) = numsamps
        Else
            MsgBox ("Number of samples must not be zero, aborting")
            Exit Function
        End If
        
        ' get the list of pins in the pin group
        Call TheExec.DataManager.DecomposePinList(pgroup, plist(), numpins)
        datFile = "outdata.csv"
        
        Open datFile For Output As #1
        ' dump pin statement to output file
        pstrng = Join(plist, ",")
        Write #1, "Pins," & pstrng
        
        ' init voltage
        currv = voltstrt
        ' iterate through voltages
        For vcnt = 0 To voltsteps - 1
            ' set the current vol value
            TheHdw.Digital.Pins(pgroup).Levels.value(chVol) = currv
            ' init timing
            currt = period * 3
            
            ' set output header
            'Debug.Print ("Voltage:" & currv)
            Write #1, "Voltage," & currv
            
            ' iterate through times
            For tcnt = 0 To timesteps - 1
                ' adjust timing
                TheHdw.Digital.Pins(pgroup).Timing.EdgeTime("TSet3", chEdgeR0) = currt
                
                ' start the pattern
                Call TheHdw.Patterns(patname).Start(strtlab)
                Call TheHdw.Digital.Patgen.HaltWait
        
                ' get the failing counts
                fcpl = TheHdw.Digital.Pins(pgroup).FailCount
                output = ""
                
                ' iterate through sites
                For Each tSite In TheExec.Sites.Active
                    ' iterate through the pins
                    For pcnt = 0 To numpins - 1
                        fcnt = fcpl.Pins(pcnt).value(tSite)
                        ' build up output
                        output = output & "," & fcnt
                    Next pcnt
                Next tSite
                
                ' dump results to file
                'Debug.Print (currt & "," & output)
                Write #1, currt & "," & output
                
                ' bump back time
                currt = currt - tincr
            Next tcnt
                    
            ' bump up voltage
            currv = currv + vincr
        Next vcnt
        
        ' close the file
        Close #1
    
    End If
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next
End Function


Public Function UC16_Impedance_Profile_Tolerance()
' 1. Use 2D schmoo to capture the time domain reflection profile of a pulse and record as a baseline response (Use Case 16).
' 2. Analyze data and determine high and low comparator levels so that a low is nominal, a mid-band isdrifting, and a high is noncompliant.
' 3. Crete a Time Set called FRT_Profile based on analysis.
' 4. Create a Pin Levels set called FRT_Profile per based on analysis.
' 5. Create FRT functional pattern called FRT_Profile.The pattern file is a single vector. Each channel has the following pattern: drive H at
'    T0, compare at time bestindicated by analysis.
End Function


Public Function UC17_VI80_Decoupling_Cap()
' 1. Do this for all sites (Schematic Sheet 21).
' 2. Use 17.VI80_30 to force a small fixed current. (11mA)
'    Make two voltage measurements at two separate times. (5mS delta time)
' 3. Use C= I dt/dV to calculate total capacitance for each supply and for each site. C30 & C31 = 11uFd
' 4. Build test based on results.
' 5. Thresholds TBD (+/-% of expected value).

' Force current of 11mA, for 5mS, should give us a delta V of 5Vs for 11uFd of expected capacitance.
' The force and sense lines should be connected on the probe interface for this to work.

    Dim i As Integer
    Dim site As Variant
    Dim VREFH_ADC01 As New PinListData, VREFH_ADC02 As New PinListData
    
    Dim PinArr() As String, PinCount As Long
    Dim cap_value As Double
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    TheExec.DataManager.DecomposePinList "VREFH_ADC0", PinArr, PinCount
   
'Force all capacitors to have 0V applied as the initial conditions.
    With TheHdw.DCVI.Pins("vrefh_adc0")
        .Mode = tlDCVIModeVoltage
        .ComplianceRange(tlDCVICompliancePositive) = 7
        .ComplianceRange(tlDCVIComplianceNegative) = 2
        .VoltageRange = 7
        .VoltageRange.Autorange = True
        .CurrentRange = 0.02
        .CurrentRange.Autorange = True
        .Voltage = 0#
        .Current = 0.01
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .LocalKelvin(tlDCVILocalKelvinHigh) = False
        .Alarm(tlDCVIAlarmAll) = tlAlarmOff
        .Connect tlDCVIConnectDefault
        .BleederResistor = tlDCVIBleederResistorAuto
        .Meter.HardwareAverage = 1
        .Gate = True
    End With
    
    TheHdw.wait (0.01)  'Discharge caps
    
    VREFH_ADC01 = TheHdw.DCVI.Pins("vrefh_adc0").Meter.Read(tlStrobe, 1)    'Should be 0 Volts
    
    ' the PinListData variable with simulation data
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            VREFH_ADC01.value = 0#
        Next site
    End If

    
    With TheHdw.DCVI.Pins("vrefh_adc0")
        .Disconnect
        .Gate = False
        .Mode = tlDCVIModeCurrent
        .Current = 0.001                    'Current forced
        .CurrentRange = 0.002
        .Meter.Mode = tlDCVIMeterVoltage
        .Meter.VoltageRange = 7
        .LocalKelvin(tlDCVILocalKelvinHigh) = False
        .Meter.HardwareAverage = 1
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
    TheHdw.wait (0.005) 'Delta Time

    VREFH_ADC02 = TheHdw.DCVI.Pins("vrefh_adc0").Meter.Read(tlStrobe, 1)    'Should be 0 Volts
    
    ' the PinListData variable with simulation data
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.Sites
            VREFH_ADC02.value = 0.47 + (Rnd * 0.1)
        Next site
    End If


    'Clean up
    With TheHdw.DCVI.Pins("vrefh_adc0")
        .Gate = False
        .Disconnect
        .Mode = tlDCVIModeVoltage
        .Voltage = 0#
    End With


    'Datalog results
    For Each site In TheExec.Sites.Selected
        TheExec.Flow.TestLimitIndex = 0
        '             I       Td                      Tv
        cap_value = 0.001 * (0.005 / (VREFH_ADC02.value - VREFH_ADC01.value))
        If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=cap_value, unit:=unitCustom, customUnit:="Fd", forceUnit:=unitNone, _
                                                                    ForceResults:=tlForceFlow, PinName:="vrefh_adc0", TName:="UC17_C30/31"
    Next site

End Function




Public Function old_UC15_Impedance_Resp_Profile(numsamps As Long, period As Double, timesteps As Long, voltstrt As Double, voltstop As Double, voltsteps As Long, Optional strtlab As String)
'Public Function pdshmoo2(numsamps As Long, period As Double, timesteps As Long, voltstrt As Double, voltstop As Double, voltsteps As Long, pgroup As String, patname As pattern, Optional strtlab As String)
    Dim fcpl As New PinListData
    Dim vincr As Double, tincr As Double, currv As Double, currt As Double
    Dim vcnt As Long, tcnt As Long, pcnt As Long, fcnt As Long, numpins As Long
    Dim plist() As String, output As String, datFile As String, pstrng As String
    Dim tSite As Variant
    Dim patname As String

    On Error GoTo errHandler
    
    ' load levels and timing
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
    
    ' get the list of pins in the pin group
    Call TheExec.DataManager.DecomposePinList("AllDig", plist(), numpins)
    
    datFile = "outdata.csv"
    
    Open datFile For Output As #1
    ' dump pin statement to output file
    pstrng = Join(plist, ",")
    Write #1, "Pins," & pstrng
    
    ' ensure pattern is loaded
    patname = ".\Shmoo.pat"
    TheHdw.Patterns(patname).Load
    
    ' calculate time increment - can move drive edge 4 periods
    If timesteps <> 0 Then
        tincr = (period * 3) / timesteps
    Else
        MsgBox ("Time Steps must not be zero, aborting")
        Exit Function
    End If
    
    ' calculate voltage incement
    If voltsteps <> 0 Then
        vincr = (voltstop - voltstrt) / voltsteps
    Else
        MsgBox ("Voltage Steps must not be zero, aborting")
        Exit Function
    End If
    
    ' set pattern loop counter to number of samples per shmoo point
    If numsamps <> 0 Then
        TheHdw.Digital.Patgen.Counter(tlPgCounter4) = numsamps
    Else
        MsgBox ("Number of samples must not be zero, aborting")
        Exit Function
    End If
    
    
    ' init voltage
    currv = voltstrt
    ' iterate through voltages
    For vcnt = 0 To voltsteps - 1
        ' set the current vol value
        TheHdw.Digital.Pins("AllDig").Levels.value(chVol) = currv
        ' init timing
        currt = period * 3
        
        ' set output header
        'Debug.Print ("Voltage:" & currv)
        Write #1, "Voltage," & currv
        
        ' iterate through times
        For tcnt = 0 To timesteps - 1
            ' adjust timing
            TheHdw.Digital.Pins("AllDig").Timing.EdgeTime("TSet3", chEdgeR0) = currt
            
            ' start the pattern
            Call TheHdw.Patterns(patname).Start(strtlab)
            Call TheHdw.Digital.Patgen.HaltWait
    
            ' get the failing counts
            fcpl = TheHdw.Digital.Pins("AllDig").FailCount      'Big time delay!
            output = ""
            
            ' iterate through sites
            For Each tSite In TheExec.Sites.Active
                ' iterate through the pins
                For pcnt = 0 To numpins - 1
                    fcnt = fcpl.Pins(pcnt).value(tSite)
'                    fcnt = 0
                    output = output & "," & fcnt
                Next pcnt
            Next tSite
            
            ' dump results to file
            'Debug.Print (currt & "," & output)
            Write #1, currt & "," & output
            
            ' bump back time
            currt = currt - tincr
        Next tcnt
                
        ' bump up voltage
        currv = currv + vincr
    Next vcnt
    
    ' close the file
    Close #1
    
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Simple_TDR(numsamps As Long, period As Double, timesteps As Long, voltstrt As Double, voltstop As Double, voltsteps As Long, pgroup As String, patname As String, Optional strtlab As String)
    Dim fcpl As New PinListData
    Dim vincr As Double, tincr As Double, currv As Double, currt As Double
    Dim vcnt As Long, tcnt As Long, pcnt As Long, fcnt As Long, numpins As Long
    Dim plist() As String, output As String, datFile As String, pstrng As String
    Dim tSite As Variant
    Dim vlev(3) As Double
    Dim i As Integer
    Dim cell As String
    Dim col(28) As String
    Dim PinName As String
    Dim offset As Integer, offset2 As Integer
    Dim Path_Length As Double, TS1 As Double, TS2 As Double
            
    offset = 2
    
    vlev(0) = voltstop * 0.25
    vlev(1) = voltstop * 0.75
    
    col(0) = "a"
    col(1) = "B"
    col(2) = "c"
    col(3) = "d"
    col(4) = "e"
    col(5) = "f"
    col(6) = "g"
    col(7) = "h"
    col(8) = "i"
    col(9) = "j"
    col(10) = "k"
    col(11) = "l"
    col(12) = "m"
    col(13) = "n"
    col(14) = "o"
    col(15) = "p"
    col(16) = "q"
    col(17) = "r"
    col(18) = "s"
    col(19) = "t"
    col(20) = "u"
    col(21) = "v"
    col(22) = "w"
    col(23) = "x"
    col(24) = "y"
    col(25) = "z"
    col(26) = "aa"
    col(27) = "ab"
    col(28) = "ac"

        On Error GoTo errHandler
        
        ' get the list of pins in the pin group
        Call TheExec.DataManager.DecomposePinList(pgroup, plist(), numpins)
        
        Sheets("TDR_WorkSheet").Cells.ClearContents
        
        For i = 0 To numpins - 1
            cell = "a" & offset + i
            Worksheets("TDR_WorkSheet").Range(cell).value = plist(i) & "_TS1"
            cell = "a" & offset + i + numpins + 2
            Worksheets("TDR_WorkSheet").Range(cell).value = plist(i) & "_TS2"
        Next i
        
        ' load levels and timing
        TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
        ' ensure pattern is loaded
        TheHdw.Patterns(patname).Load
        
        ' calculate time increment - can move drive edge 4 periods
        If timesteps <> 0 Then
            tincr = (period * 3) / timesteps
        Else
            MsgBox ("Time Steps must not be zero, aborting")
            Exit Function
        End If
        
        ' calculate voltage incement
        If voltsteps <> 0 Then
            vincr = (voltstop - voltstrt) / voltsteps
        Else
            MsgBox ("Voltage Steps must not be zero, aborting")
            Exit Function
        End If
        
        ' set pattern loop counter to number of samples per shmoo point.  This is for jitter.
        If numsamps <> 0 Then
            TheHdw.Digital.Patgen.Counter(tlPgCounter4) = numsamps
        Else
            MsgBox ("Number of samples must not be zero, aborting")
            Exit Function
        End If
        
        ' init voltage
        currv = voltstrt
        ' iterate through voltages
        For i = 0 To 1
            If i = 0 Then
                offset2 = 0
            Else
                offset2 = numpins + 2
            End If
            
            ' set the current vol value
            TheHdw.Digital.Pins(pgroup).Levels.value(chVol) = vlev(i)
            ' init timing
            currt = period * 3
            
            ' iterate through times
            For tcnt = 0 To timesteps - 1
                ' adjust timing
                TheHdw.Digital.Pins(pgroup).Timing.EdgeTime("TSet3", chEdgeR0) = currt
                
                ' start the pattern
                Call TheHdw.Patterns(patname).Start(strtlab)
                Call TheHdw.Digital.Patgen.HaltWait
        
                ' get the failing counts
                fcpl = TheHdw.Digital.Pins(pgroup).FailCount
                output = ""
                
                ' iterate through sites
                For Each tSite In TheExec.Sites.Active
                    ' iterate through the pins
                    For pcnt = 0 To numpins - 1
                        fcnt = fcpl.Pins(pcnt).value(tSite)         'Normal data
                        
                        If TheExec.TesterMode = testModeOffline Then
                            Select Case (offset2)
                                Case 0
                                    If tcnt = timesteps * (5 / 6) Then
                                        fcnt = 0
                                    Else
                                        fcnt = numsamps
                                    End If
                                Case numpins + 2
                                    If tcnt = timesteps * (4 / 6) Then
                                        fcnt = 0
                                    Else
                                        fcnt = numsamps
                                    End If
                            End Select
                        End If
                        
                            fcnt = fcnt
                            
                            If fcnt = 0 Then
                                cell = col(tSite + 1) & pcnt + offset + offset2
                                If Worksheets("TDR_WorkSheet").Range(cell).value = 0 Then
                                    Worksheets("TDR_WorkSheet").Range(cell).value = currt
                                End If
                            End If
                        
                    Next pcnt
                Next tSite
                currt = currt - tincr
            Next tcnt
                    
            ' bump up voltage
            currv = currv + vincr
        Next i
        
        
    'Datalog results
    For Each tSite In TheExec.Sites.Selected
        For i = 0 To pcnt - 1
            TheExec.Flow.TestLimitIndex = 0
            cell = col(tSite + 1) & i + offset
            TS1 = Worksheets("TDR_WorkSheet").Range(cell).value
            cell = col(tSite + 1) & i + offset + numpins + 2
            TS2 = Worksheets("TDR_WorkSheet").Range(cell).value
            Path_Length = TS2 - TS1
            
            If TheExec.Sites.Active Then TheExec.Flow.TestLimit ResultVal:=Path_Length, _
                                                                unit:=unitTime, _
                                                                forceUnit:=unitNone, _
                                                                ForceResults:=tlForceFlow, _
                                                                PinName:=plist(i), _
                                                                TName:="PathLength"
        Next i
    Next tSite
        
    Exit Function

errHandler:
     If AbortTest Then Exit Function Else Resume Next
End Function

Function Add_DCVS_PSet(PinName As String, _
                 PSetName As String, _
                 BandwidthSetting As Double, _
                 CurrentRange As Double, _
                 SampleRate As Double, _
                 SampleSize As Long, _
                 SourceFoldLimit As Double, _
                 MeterMode As tlDCVSMeterMode, _
                 DCVSMode As tlDCVSMode, _
                 MainVoltage As Double, _
                 AltVoltage As Double _
                 ) As Long
                 
   On Error GoTo ErrorHandler
   
   Call TheHdw.DCVS.Pins(PinName).PSets.Add(PSetName)
    
'    With thehdw.DCVS.Pins("VDDa")
'        .PSets("VDDa_pset_1").Apply
'       .Voltage.Main.value = 3.3                    ' Main Voltage (3V)
'       .Gate = False ' You must gate off to change fold limit timeout value
'       .BandwidthSetting.value = 255
'       .CurrentRange = 0.004
'       .CurrentLimit.Source.FoldLimit.Level.value = 0.004
'       .CurrentLimit.Source.FoldLimit.TimeOut.value = 2
'       '.CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOn
'       .Mode = tlDCVSModeVoltage                    ' Force Voltage Mode
'       .Meter.Mode = tlDCVSMeterVoltage             ' Meter in Voltage Mode
'       .Connect (tlDCVSConnectDefault)              ' Connect Force and Sense
'       .Gate = True                                 ' Gate Supply on
'    End With
'
   With TheHdw.DCVS.Pins(PinName).PSets.Item(PSetName)
      .BandwidthSetting.value = BandwidthSetting
      .CurrentRange.value = CurrentRange
      .Capture.SampleRate.value = SampleRate
      .Capture.SampleSize.value = SampleSize
      .CurrentLimit.Source.FoldLimit.Level.value = SourceFoldLimit
      .Meter.Mode = MeterMode
      .Mode = DCVSMode
      .Voltage.Main.value = MainVoltage
      .Voltage.Alt.value = AltVoltage
   End With
   
Exit Function

ErrorHandler:
   Dim Reply As Integer
   Reply = MsgBox("Error detected during AddPSet(). Stop to Debug?", vbYesNo, "UVS256 PSet Checkout")
   If Reply = vbYes Then
      On Error GoTo ErrorHandler
      Stop
   Else
      On Error GoTo 0
   End If
   Resume Next
End Function


Function AddPSet(PinName As String, _
                 PSetName As String, _
                 BandwidthSetting As Double, _
                 CurrentRange As Double, _
                 SampleRate As Double, _
                 SampleSize As Long, _
                 SourceFoldLimit As Double, _
                 MeterCurrentRange As Double, _
                 MainVolt As Double, _
                 AltVolt As Double, _
                 MeterMode As tlDCVSMeterMode, _
                 DCVSMode As tlDCVSMode, _
                 MainVoltage As Double, _
                 AltVoltage As Double _
                 ) As Long
                 
   On Error GoTo ErrorHandler
   
   Call TheHdw.DCVS.Pins(PinName).PSets.Add(PSetName)
   
   With TheHdw.DCVS.Pins(PinName).PSets.Item(PSetName)
      .BandwidthSetting.value = BandwidthSetting
      .CurrentRange.value = CurrentRange
      .Capture.SampleRate.value = SampleRate
      .Capture.SampleSize.value = SampleSize
      .CurrentLimit.Source.FoldLimit.Level.value = SourceFoldLimit
      .CurrentRange.value = MeterCurrentRange
      .Voltage.Main.value = MainVolt
      .Voltage.Alt.value = AltVolt
      .Meter.Mode = MeterMode
      .Mode = DCVSMode
      .Voltage.Main.value = MainVoltage
      .Voltage.Alt.value = AltVoltage
   End With
   
Exit Function
ErrorHandler:
   Dim Reply As Integer
   Reply = MsgBox("Error detected during AddPSet(). Stop to Debug?", vbYesNo, "UVS256 PSet Checkout")
   If Reply = vbYes Then
      On Error GoTo ErrorHandler
      Stop
   Else
      On Error GoTo 0
   End If
   Resume Next
End Function



