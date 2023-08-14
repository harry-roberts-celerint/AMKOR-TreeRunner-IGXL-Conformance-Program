Attribute VB_Name = "RunVBT"
' This ALWAYS GENERATED file contains wrappers for VBT tests.
' Do not edit.

Private Sub HandleUntrappedError()
    ' Sanity clause
    If TheExec Is Nothing Then
        MsgBox "IG-XL is not running!  VBT tests cannot execute unless IG-XL is running."
        Exit Sub
    End If
    ' If the last site has failed out, let's ignore the error
    If TheExec.Sites.Active.Count = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 6.10 aborts gracefully.)
    Do While TheExec.Sites.InSiteLoop
        Call TheExec.Sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Select all active sites in case a subset of sites was selected when error occurred.
    TheExec.Sites.Selected = TheExec.Sites.Active
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    TheExec.ErrorLogMessage "Test " + TheExec.DataManager.instanceName + ": VBT error #" + Trim(Str(Err.Number)) + " '" + Err.Description + "'"
End Sub

Public Function DCVIPowerSupply_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As New pattern
    p8.value = v(7)
    Dim p9 As New PinList
    p9.value = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(17)
    Dim p14 As New PinList
    p14.value = v(18)
    Dim p15 As tlPSSource
    p15 = v(19)
    Dim p16 As tlRelayMode
    p16 = v(34)
    Dim p17 As New PinList
    p17.value = v(35)
    Dim p18 As New PinList
    p18.value = v(36)
    Dim p19 As tlPSTestControl
    p19 = v(37)
    Dim p20 As New InterposeName
    p20.value = v(39)
    Dim p21 As tlWaitVal
    p21 = v(41)
    Dim p22 As tlWaitVal
    p22 = v(42)
    Dim p23 As tlWaitVal
    p23 = v(43)
    Dim p24 As tlWaitVal
    p24 = v(44)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVIPowerSupply_T__ = Template.VBT_DCVIPowerSupply_T.DCVIPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, CDbl(v(12)), CLng(v(13)), CStr(v(14)), CDbl(v(15)), CDbl(v(16)), p13, p14, p15, CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CStr(v(32)), CBool(v(33)), p16, p17, p18, p19, CBool(v(38)), p20, CStr(v(40)), p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(46)), , CStr(v(47)), CBool(v(48)), CBool(v(49)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVSPowerSupply_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As New pattern
    p8.value = v(7)
    Dim p9 As New PinList
    p9.value = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(16)
    Dim p15 As tlPSSource
    p15 = v(17)
    Dim p16 As tlRelayMode
    p16 = v(31)
    Dim p17 As New PinList
    p17.value = v(32)
    Dim p18 As New PinList
    p18.value = v(33)
    Dim p19 As tlPSTestControl
    p19 = v(34)
    Dim p20 As tlWaitVal
    p20 = v(35)
    Dim p21 As tlWaitVal
    p21 = v(36)
    Dim p22 As tlWaitVal
    p22 = v(37)
    Dim p23 As tlWaitVal
    p23 = v(38)
    Dim p24 As New FormulaArg
    p24.value = v(40)
    Dim p25 As New FormulaArg
    p25.value = v(41)
    Dim p26 As New FormulaArg
    p26.value = v(42)
    Dim p27 As New FormulaArg
    p27.value = v(43)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    DCVSPowerSupply_T__ = Template.VBT_DCVSPowerSupply_T.DCVSPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CDbl(v(13)), CLng(v(14)), CStr(v(15)), p14, p15, CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(30)), p16, p17, p18, p19, p20, p21, p22, p23, CBool(v(UBound(v))), p24, p25, p26, p27, , CStr(v(44)), CBool(v(45)), CBool(v(46)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Empty_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New InterposeName
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New PinList
    p7.value = v(12)
    Dim p8 As New PinList
    p8.value = v(13)
    Dim p9 As New PinList
    p9.value = v(14)
    Dim p10 As New PinList
    p10.value = v(15)
    Dim p11 As New PinList
    p11.value = v(16)
    Dim p12 As New PinList
    p12.value = v(17)
    Dim p13 As New PinList
    p13.value = v(18)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Empty_T__ = Template.VBT_Empty_T.Empty_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, p13, pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(13)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New PinList
    p16.value = v(21)
    Dim p17 As New InterposeName
    p17.value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Functional_T__ = Template.VBT_Functional_T.Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PinPMU_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New InterposeName
    p1.value = v(1)
    Dim p2 As New InterposeName
    p2.value = v(2)
    Dim p3 As New InterposeName
    p3.value = v(3)
    Dim p4 As New InterposeName
    p4.value = v(4)
    Dim p5 As New InterposeName
    p5.value = v(5)
    Dim p6 As New InterposeName
    p6.value = v(6)
    Dim p7 As New pattern
    p7.value = v(7)
    Dim p8 As New pattern
    p8.value = v(8)
    Dim p9 As New PinList
    p9.value = v(10)
    Dim p10 As New PinList
    p10.value = v(11)
    Dim p11 As New PinList
    p11.value = v(12)
    Dim p12 As New PinList
    p12.value = v(13)
    Dim p13 As New PinList
    p13.value = v(14)
    Dim p14 As New PinList
    p14.value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.value = v(18)
    Dim p17 As New FormulaArg
    p17.value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New PinList
    p19.value = v(36)
    Dim p20 As New PinList
    p20.value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.value = v(52)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim p27 As New PinList
    p27.value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.value = v(55)
    PinPMU_T__ = Template.VBT_PinPmu_T.PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MtoMemory_T__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As New PinList
    p9.value = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(19)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New InterposeName
    p16.value = v(21)
    Dim p17 As tlRelayMode
    p17 = v(24)
    Dim pStep As SubType
    pStep = TheExec.Flow.StepType
    Dim ExtraArgs(0 To 49) As Variant
    Dim i As Integer
    For i = 0 To 49
        ExtraArgs(i) = v(51 + i)
    Next i
    MtoMemory_T__ = Template.VBT_MTOMemory_T.MtoMemory_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(23)), p17, CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CLng(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CLng(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CStr(v(38)), CLng(v(39)), CLng(v(40)), CBool(v(UBound(v))), pStep, ExtraArgs, CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function
Public Function functionalExp1__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    functionalExp1__ = VBAProject.VBT_Module.functionalExp1(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function functionalExp2__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    functionalExp2__ = VBAProject.VBT_Module.functionalExp2(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC01_Propagation_Delay__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC01_Propagation_Delay__ = VBAProject.VBT_Module.UC01_Propagation_Delay()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC02_Pogo_Pin_Cont__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC02_Pogo_Pin_Cont__ = VBAProject.VBT_Module.UC02_Pogo_Pin_Cont()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC03_Func_TDR_Verification__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC03_Func_TDR_Verification__ = VBAProject.VBT_Module.UC03_Func_TDR_Verification(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC04_K1_RELAY_FUNC__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC04_K1_RELAY_FUNC__ = VBAProject.VBT_Module.UC04_K1_RELAY_FUNC()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC05_K2_RELAY_FUNC__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC05_K2_RELAY_FUNC__ = VBAProject.VBT_Module.UC05_K2_RELAY_FUNC()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC06_DECODER_S1_Func__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC06_DECODER_S1_Func__ = VBAProject.VBT_Module.UC06_DECODER_S1_Func()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC07_DECODER_S2_Func__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC07_DECODER_S2_Func__ = VBAProject.VBT_Module.UC07_DECODER_S2_Func()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC08_ADG5412BRUZ_SWITCHES__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC08_ADG5412BRUZ_SWITCHES__ = VBAProject.VBT_Module.UC08_ADG5412BRUZ_SWITCHES()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC09_VS256_Cap_Leak__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New PinList
    p1.value = v(0)
    UC09_VS256_Cap_Leak__ = VBAProject.VBT_Module.UC09_VS256_Cap_Leak(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function old_UC09_VS256_Cap_Leak__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    old_UC09_VS256_Cap_Leak__ = VBAProject.VBT_Module.old_UC09_VS256_Cap_Leak()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC10_All_S2_Decoder_Leak__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC10_All_S2_Decoder_Leak__ = VBAProject.VBT_Module.UC10_All_S2_Decoder_Leak()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC11_All_S1_Decoder_Leak__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC11_All_S1_Decoder_Leak__ = VBAProject.VBT_Module.UC11_All_S1_Decoder_Leak()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC12_UP1600_LEAKAGE__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC12_UP1600_LEAKAGE__ = VBAProject.VBT_Module.UC12_UP1600_LEAKAGE(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC13_VS256_Cap_Meas_Conn__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    UC13_VS256_Cap_Meas_Conn__ = VBAProject.VBT_Module.UC13_VS256_Cap_Meas_Conn(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Junk_UC13_VS256_Cap_Meas_Conn__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As New pattern
    p1.value = v(0)
    Junk_UC13_VS256_Cap_Meas_Conn__ = VBAProject.VBT_Module.Junk_UC13_VS256_Cap_Meas_Conn(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function old_UC13_VS256_Cap_Meas_Conn__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    old_UC13_VS256_Cap_Meas_Conn__ = VBAProject.VBT_Module.old_UC13_VS256_Cap_Meas_Conn()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC14_OpAmps_SAR_ADC_INPUTS__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC14_OpAmps_SAR_ADC_INPUTS__ = VBAProject.VBT_Module.UC14_OpAmps_SAR_ADC_INPUTS()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC15_Impedance_Resp_Profile__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC15_Impedance_Resp_Profile__ = VBAProject.VBT_Module.UC15_Impedance_Resp_Profile(CLng(v(0)), CDbl(v(1)), CLng(v(2)), CDbl(v(3)), CDbl(v(4)), CLng(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC16_Impedance_Profile_Tolerance__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC16_Impedance_Profile_Tolerance__ = VBAProject.VBT_Module.UC16_Impedance_Profile_Tolerance()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UC17_VI80_Decoupling_Cap__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    UC17_VI80_Decoupling_Cap__ = VBAProject.VBT_Module.UC17_VI80_Decoupling_Cap()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function old_UC15_Impedance_Resp_Profile__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    old_UC15_Impedance_Resp_Profile__ = VBAProject.VBT_Module.old_UC15_Impedance_Resp_Profile(CLng(v(0)), CDbl(v(1)), CLng(v(2)), CDbl(v(3)), CDbl(v(4)), CLng(v(5)), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Simple_TDR__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Simple_TDR__ = VBAProject.VBT_Module.Simple_TDR(CLng(v(0)), CDbl(v(1)), CLng(v(2)), CDbl(v(3)), CDbl(v(4)), CLng(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Add_DCVS_PSet__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As tlDCVSMeterMode
    p1 = v(7)
    Dim p2 As tlDCVSMode
    p2 = v(8)
    Add_DCVS_PSet__ = VBAProject.VBT_Module.Add_DCVS_PSet(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CLng(v(5)), CDbl(v(6)), p1, p2, CDbl(v(9)), CDbl(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AddPSet__(v As Variant) As Long
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    Dim p1 As tlDCVSMeterMode
    p1 = v(10)
    Dim p2 As tlDCVSMode
    p2 = v(11)
    AddPSet__ = VBAProject.VBT_Module.AddPSet(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CLng(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), p1, p2, CDbl(v(12)), CDbl(v(13)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































