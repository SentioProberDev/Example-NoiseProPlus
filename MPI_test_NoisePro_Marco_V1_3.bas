Function ProbeStationInitialize(GPIBAddress As String, GPIBTimeout As String) As Boolean

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    If OpenGPIBProbe(Val(GPIBAddress), Val(GPIBTimeout)) <> 0 Then GoTo ProbeStationInitialize_err:

	'------- Load Bin table-----------'
    'Cmd = "map:bins:load C:\\ProgramData\\MPI Corporation\\Sentio\\config\\defaults\\default_bins.xbt"
    'If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo ProbeStationInitialize_err
    'If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo ProbeStationInitialize_err
	'  -- Take Die Color from bin
    'Cmd = "map:set_color_scheme 1"
    'If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo ProbeStationInitialize_err
    'If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo ProbeStationInitialize_err
	'  -- clear the binning
    'Cmd = "map:bins:set_all -1"
    'If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo ProbeStationInitialize_err
    'If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo ProbeStationInitialize_err
	'  -- Take Die Color from bin
    'Cmd = "map:set_color_scheme 2"
    'If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo ProbeStationInitialize_err
    'If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo ProbeStationInitialize_err

	'  -- Move to chuch home die
	Cmd = "move_chuck_home"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo ProbeStationInitialize_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo ProbeStationInitialize_err:


    ProbeStationInitialize = True
    Exit Function
ProbeStationInitialize_err:
    ProbeStationInitialize = False
End Function

Function ReleaseProbeGPIB() As Boolean
  CloseGPIBProbe()
End Function

Function StepFirstDie(r1 As String, r2 As String, r3 As String, r4 As String) As Boolean
'************************************************************************************************
'Step to First Die.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "map:step_first_die"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo StepFirstDie_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo StepFirstDie_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
    If CStr(r1) <> "0" Then
    	GoTo StepFirstDie_err:
    Else
    	r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
	End If

StepFirstDie = True
    Exit Function
StepFirstDie_err:
    MsgBox (r1)
    StepFirstDie_err = False
End Function

Function StepNextDieWithBin(bin As Integer, r1 As String, r2 As String, r3 As String, r4 As String) As Boolean
'************************************************************************************************
'Description:    Assigns a bin to the current die and steps to the next die of the wafer according to the selected routing algorithm.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "map:bin_step_next_die" & " " & CStr(bin)
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo StepNextDieWithBin_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo StepNextDieWithBin_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

	If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
	ElseIf r1 = "1024" Then
		GoTo StepNextDie_last:
	Else
		GoTo StepNextDieWithBin_err:
	End If


StepNextDieWithBin = True
    Exit Function
StepNextDie_last:
'	MsgBox ("Last Die")
	StepNextDieWithBin = True
	Exit Function
StepNextDieWithBin_err:
	r3 = (Trim(tmp_str(2)))
    MsgBox (r3)
    StepNextDieWithBin = False
End Function

Function StepNextDie(r1 As String, r2 As String, r3 As String, r4 As String) As Boolean
'************************************************************************************************
'Step to the next die of the wafer according to the selected routing algorithm. Only executable if currently active die is part of the route.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
	
    Cmd = "map:step_next_die"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo StepNextDie_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo StepNextDie_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
				
	If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
	ElseIf r1 = "1024" Then
		GoTo StepNextDie_last:
	Else
		GoTo StepNextDie_err:
	End If

StepNextDie = True
    Exit Function
StepNextDie_last:
'	MsgBox ("Last Die")
	StepNextDie = True
	Exit Function
StepNextDie_err:
	r3 = (Trim(tmp_str(2)))
    MsgBox (r3)
    StepNextDie = False
End Function

Function StepNextSubsiteWithBin(bin As Integer, r1 As String, r2 As String, r3 As String, r4 As String, r5 As String) As Boolean
'************************************************************************************************
'Sets a bin code to the current subsite and steps to the next subsite of the die. If the current subsite is the last one of the die it steps to the first subsite of the next die.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "map:subsite:bin_step_next" & " " &CStr(bin)
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo StepNextSubsiteWithBin_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo StepNextSubsiteWithBin_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

	If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
		r5 = (Trim(tmp_str(4)))
	ElseIf r1 = "2048" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
		r5 = (Trim(tmp_str(4)))
		GoTo StepNextSubsite_last:
	Else
		GoTo StepNextSubsiteWithBin_err:
	End If


StepNextSubsiteWithBin = True
    Exit Function
StepNextSubsite_last:
	MsgBox ("Last subsite")
	StepNextSubsiteWithBin = True
	Exit Function
StepNextSubsiteWithBin_err:
	r3 = (Trim(tmp_str(2)))
    MsgBox (r3)
    StepNextSubsiteWithBin = False
End Function

Function StepNextSubsite(r1 As String, r2 As String, r3 As String, r4 As String , r5 As String) As Boolean
'************************************************************************************************
'Steps to the next subsite of the current die. If the current subsite is the last one of the die and currently active die is part of the route it steps to the first subsite of the next die.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
	Dim test_result As Integer
	
    Cmd = "map:subsite:step_next"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo StepNextSubsite_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo StepNextSubsite_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

	test_result = Val(r1) Mod 1024
	
	If test_result <> 0 Then
        GoTo StepNextSubsite_err:
    End If
	
    If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
		r5 = (Trim(tmp_str(4)))
	ElseIf r1 = "2048" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
		r5 = (Trim(tmp_str(4)))
		GoTo StepNextSubsite_last:
	Else
		GoTo StepNextSubsite_err:
	End If


StepNextSubsite = True
    Exit Function
StepNextSubsite_last:
	MsgBox ("Last subsite")
	StepNextSubsite = True
	Exit Function
StepNextSubsite_err:
	r3 = (Trim(tmp_str(2)))
    MsgBox (r3)
    StepNextSubsite = False
End Function

Function MoveChuckContact(r1 As String, r2 As String, r3 As String) As Boolean
'************************************************************************************************
'Moves chuck to contact height. If overtravel is enabled chuck moves to overtravel height. If contact height is not set the command is not carried out.


    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "move_chuck_contact"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo MoveChuckContact_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo MoveChuckContact_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
    r3 = (Trim(tmp_str(2)))
    If CStr(r1) <> "0" Then GoTo MoveChuckContact_err:


MoveChuckContact = True
    Exit Function
MoveChuckContact_err:
    MsgBox (r3)
    MoveChuckContact_err = False
End Function

Function MoveChuckSeparation(r1 As String, r2 As String, r3 As String) As Boolean
'************************************************************************************************
'Moves chuck to separation height. If contact height is not set the command is not carried out.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "move_chuck_separation"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo MoveChuckSeparation_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo MoveChuckSeparation_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
    r3 = (Trim(tmp_str(2)))
    If r1 <> "0" Then GoTo MoveChuckSeparation_err:


MoveChuckSeparation = True
    Exit Function
MoveChuckSeparation_err:
    MsgBox (r3)
    MoveChuckSeparation = False
End Function

Function MoveChuckHome(r1 As String, r2 As String, r3 As String, r4 As String) As Boolean
'************************************************************************************************
'Moves chuck xy to home position of the current site.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "move_chuck_home"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo MoveChuckHome_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo MoveChuckHome_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

    If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
	Else
		GoTo MoveChuckHome_err:
	End If


MoveChuckHome = True
    Exit Function
MoveChuckHome_err:
    MsgBox (r3)
    MoveChuckHome = False
End Function

Function MoveChuckCenter(r1 As String, r2 As String, r3 As String, r4 As String) As Boolean
'************************************************************************************************
'Moves chuck xy to home position of the current site.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "move_chuck_center"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo MoveChuckCenter_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo MoveChuckCenter_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

    If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
		r4 = (Trim(tmp_str(3)))
	Else
		GoTo MoveChuckCenter_err:
	End If


MoveChuckCenter = True
    Exit Function
MoveChuckCenter_err:
    MsgBox (r3)
    MoveChuckCenter = False
End Function

Function GetTestDiesNumber(r1 As String, r2 As String, r3 As String) As Boolean
'************************************************************************************************
'Get the Die number of wafer map.

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "map:get_num_dies Selected"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo GetDiesNumber_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo GetDiesNumber_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))

    If r1 = "0" Then
		r3 = (Trim(tmp_str(2)))
	Else
		GoTo GetDiesNumber_err:
	End If


GetTestDiesNumber = True
    Exit Function
GetDiesNumber_err:
    MsgBox (r1)
    GetTestDiesNumber = False
End Function

Function LightONOFF(Camera As String, Status As String) As Boolean
'************************************************************************************************
'Switch Light ON/OFF
'2 Input: Camera Switch, Status Switch

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "vis:switch_light" & " " & Camera & "," & Status
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo LightONOFF_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo LightONOFF_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
    r3 = (Trim(tmp_str(2)))
    If CStr(r1) <> "0" Then GoTo LightONOFF_err:


LightONOFF = True
    Exit Function
LightONOFF_err:
    MsgBox (r3)
    LightONOFF_err = False
End Function

Function SetChuckTemperature(TargetTemp As Double) As Boolean
'************************************************************************************************
'Set Chuck Temperautre 
'1 Input: TargetTemperature

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "status:set_chuck_temp" & " " &CStr(TargetTemp)
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo SetTemp_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo SetTemp_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
    r3 = (Trim(tmp_str(2)))
    If CStr(r1) <> "0" Then GoTo SetTemp_err:


SetChuckTemperature = True
    Exit Function
SetTemp_err:
    MsgBox (r3)
    SetTemp_err = False
End Function
'************************************************************************************************

Function GetChuckTemperature(CurrTemp As String) As Boolean
'************************************************************************************************
'Get Chuck Temperautre 
'Output: Current Temperature

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "status:get_chuck_temp"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo GetTemp_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo GetTemp_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	
	If r1 = "0" Then
		CurrTemp = (Trim(tmp_str(2)))
	Else
		GoTo GetTemp_err:
	End If
		
GetChuckTemperature = True
    Exit Function
GetTemp_err:
    MsgBox (CurrTemp)
    GetTemp_err = False
End Function
'************************************************************************************************

Function GetChuckSetTemperature(SetTemp As String) As Boolean
'************************************************************************************************
'Get Chuck Setpoint Temperautre 
'Output: Setpoint Temperature

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "status:get_chuck_temp_setpoint"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo GetSetTemp_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo GetSetTemp_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	If r1 = "0" Then
		SetTemp = (Trim(tmp_str(2)))
	Else
		GoTo GetSetTemp_err:
	End If
		
GetChuckSetTemperature = True
    Exit Function
GetSetTemp_err:
    MsgBox (SetTemp)
    GetSetTemp_err = False
End Function
'************************************************************************************************

Function SetChuckHoldMode(Status As String) As Boolean
'************************************************************************************************
'Set Chuck Hold mode 
'1 Input: ON/OFF

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "status:set_chuck_thermo_hold_mode" & " " & Status
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo SetHoldMode_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo SetHoldMode_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))	
	If CStr(r1) <> "0" Then GoTo SetHoldMode_err:
		
SetChuckHoldMode = True
    Exit Function
SetHoldMode_err:
    MsgBox (SetTemp)
    SetHoldMode_err = False
End Function
'************************************************************************************************

Function GetChuckHoldMode(Status As String) As Boolean
'************************************************************************************************
'Get Chuck Hold Mode
'Output: Hold Status

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String

    Cmd = "status:get_chuck_thermo_hold_mode"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo GetHoldMode_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo GetHoldMode_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	
	If r1 = "0" Then
		Status = (Trim(tmp_str(2)))
	Else
		GoTo GetHoldMode_err:
	End If
		
GetChuckHoldMode = True
    Exit Function
GetHoldMode_err:
    MsgBox (SetTemp)
    GetHoldMode_err = False
End Function
'**************************************************************************************************


Function ScanStation(Station As String, ret As String, wafer_number As Integer) As Boolean
'************************************************************************************************
'Scanning wafer in station 
'Output: wafer existing & number of wafer

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String, tmp2_str() As String
    Dim r3 As String
    wafer_number = 0
	
    Cmd = "loader:scan_station" & " " &CStr(Station)
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo scan_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo scan_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
	'tmp2_str = Split(r3, " ")

	If r1 = "0" Then
		ret = r3
	Else
		GoTo scan_err:
	End If

	
	For i = 1 to Len(ret)
		If	Mid(ret, i, 1) = "1" Then
			wafer_number = wafer_number + 1
		End If
	Next
	
ScanStation = True
    Exit Function
scan_err:
    MsgBox (Station)
    scan_err = False
End Function
'**************************************************************************************************

Function LoadWafer(Station As String, TragetSlot As String, AlignAngle As String) As Boolean
'************************************************************************************************
'Load wafer to chuck
'Output: OK

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
    Dim r1 As String, r2 As String, r3 As String

    Cmd = "loader:load_wafer" & " " & Station & "," & TragetSlot & "," & AlignAngle
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo load_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo load_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
		
	If r1 <> "0" Then
		GoTo load_err:
	End If
		

LoadWafer = True
    Exit Function
load_err:
    MsgBox (r3)
    LoadWafer = False
End Function
'**************************************************************************************************


Function WaferAlign() As Boolean
'************************************************************************************************
'Wafer Align
'Output: OK

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
    Dim r1 As String, r2 As String, r3 As String

    Cmd = "vis:align_wafer"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo align_wafer_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo align_wafer_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
		
	If r1 <> "0" Then
		GoTo align_wafer_err:
	End If
		

WaferAlign = True
    Exit Function
align_wafer_err:
    MsgBox (r3)
    WaferAlign = False
End Function
'**************************************************************************************************

Function AutoFocus() As Boolean
'************************************************************************************************
'auto focus
'Output: OK

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
    Dim r1 As String, r2 As String, r3 As String

    Cmd = "vis:auto_focus"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo auto_focus_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo auto_focus_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
		
	If r1 <> "0" Then
		GoTo auto_focus_err:
	End If
		

AutoFocus = True
    Exit Function
auto_focus_err:
    MsgBox (r3)
    AutoFocus = False
End Function
'**************************************************************************************************

Function FindHome() As Boolean
'************************************************************************************************
'auto focus
'Output: OK

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
    Dim r1 As String, r2 As String, r3 As String

    Cmd = "vis:find_home"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo auto_focus_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo auto_focus_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
		
	If r1 <> "0" Then
		GoTo auto_focus_err:
	End If
		

FindHome = True
    Exit Function
auto_focus_err:
    MsgBox (r3)
    FindHome = False
End Function
'**************************************************************************************************

Function WaferUnload() As Boolean
'************************************************************************************************
'wafer unload
'Output: OK

    Dim i As Integer, response As String
    Dim Cmd As String, tmp_str() As String
    Dim r1 As String, r2 As String, r3 As String

    Cmd = "loader:unload_wafer"
    If WriteGPIBCommandToProbe(Cmd) <> 0 Then GoTo wafer_unload_err:
    If ReadGPIBFromProbe(response, 256) <> 0 Then GoTo wafer_unload_err:
    response = Trim(response)
    tmp_str = Split(response, ",")
    r1 = (Trim(tmp_str(0)))
    r2 = (Trim(tmp_str(1)))
	r3 = (Trim(tmp_str(2)))
		
	If r1 <> "0" Then
		GoTo wafer_unload_err:
	End If
		

WaferUnload = True
    Exit Function
wafer_unload_err:
    MsgBox (r3)
    WaferUnload = False
End Function
'**************************************************************************************************

Sub Main

'Dim Cmd As String
    Dim MPICommandResult1 As String, MPICommandResult2 As String, MPICommandResult3 As String, MPICommandResult4 As String, MPICommandResult5 As String
    Dim DieRow As String, DieColumn As String
    Dim TotalDieNumber As String, Status As String,SetTemp As String,CurrTemp As String
	Dim WaferExistingStatus As String
'Dim i As Integer
	Dim RandomBinNumber As Integer
	Dim TotalSubsNum As Integer,RetInt As Integer
    Dim TotalDieNum As Integer
    Dim TotalWaferNum As Integer
	Dim SourceStation As String
	Dim AlignAngle As String
	Dim SlotNumber As Integer

	Dim ProbeGPIBaddress As String
	Dim probeGPIBtimeout As String
	Dim currentInfoset As String
	Dim Cmd As String, tmp_str() As String, Msg_str As String


If ShowMeasureStatus("Start API measurement ....","")<>0 Then GoTo Measurement_err
	infoStr = ""

	Dim probgpibformshowstr As String
	Dim probegpibtimeoutformshowstr As String


	probgpibformshowstr = "Probe GPIB Address"
	probegpibtimeoutformshowstr = "Probe GPIB TimeOut"

	currentInfoset = "Probe GPIB Address,text,true,true@Probe GPIB TimeOut: 17 for 1000 sec. 16 for 300 sec.,text,true,true"

SetupFile_again:
	RetInt = LoadMessage (currentInfoset)
	currentInfoset = ""

	ProbeGPIBaddress = GetMessage(probgpibformshowstr)
	If RetInt = -2 Or  Len(Trim(ProbeGPIBaddress)) <= 0  Or Val(Trim(ProbeGPIBaddress)) <= 0 Then
		RetInt = MsgBox ("GPIB address set error...." &vbCrLf & "Do you want To try again?",vbYesNo)
		If RetInt = vbYes Then
			GoTo SetupFile_again
		Else
			GoTo Measurement_err
		End If
	End If


	probeGPIBtimeout = GetMessage(probegpibtimeoutformshowstr)
	If RetInt = -2 Or  Len(Trim(probeGPIBtimeout)) <= 0 Or Val(Trim(probeGPIBtimeout)) <= 0 Then
		RetInt = MsgBox ("Time out set error...." &vbCrLf & "Do you want To try again?",vbYesNo)
		If RetInt = vbYes Then
			GoTo SetupFile_again
		Else
			GoTo Measurement_err
		End If
	End If

'--------Prepare the station----------'

	If ProbeStationInitialize(ProbeGPIBaddress, probeGPIBtimeout) = False Then
        GoTo Measurement_err:
    End If


'--------Thermal control testing---------------------'

'	If  SetChuckTemperature(28)= False Then
'        GoTo Measurement_err:
'    End If

'	If  GetChuckSetTemperature(SetTemp)= False Then
'        GoTo Measurement_err:
'    End If

'	If  GetChuckTemperature(CurrTemp)= False Then
'        GoTo Measurement_err:
'    End If

'	If  SetChuckHoldMode("on")= False Then		' Hold Mode: ON'
'        GoTo Measurement_err:
'    End If

'    If  GetChuckHoldMode(Status)= False Then
'        GoTo Measurement_err:
'    End If

'	WaitSeconds(10)

'	If  SetChuckHoldMode("off")= False Then		' Hold Mode: Off'
'        GoTo Measurement_err:
'    End If

'    If  GetChuckHoldMode(Status)= False Then
'        GoTo Measurement_err:
'    End If

'---------------------------------------------------'

	SourceStation = "Cassette1"
	SlotNumber = 25
	AlignAngle = "0"
	'---Parameter can be: Cassette1, Cassette2, waferwellet ---'
	If ScanStation(SourceStation, WaferExistingStatus, TotalWaferNum)= False Then
	    GoTo Measurement_err:
	End If

	For j = 1 To SlotNumber
		
		'---Transfer Wafer---'	
		If	Mid(WaferExistingStatus, j, 1) = "1" Then
			If  LoadWafer(SourceStation, SlotNumber - j + 1, AlignAngle)= False Then
				GoTo Measurement_err:
			End If

		
			'---Move Separation---'	
			If  MoveChuckSeparation(MPICommandResult1,MPICommandResult2,MPICommandResult3)= False Then
				GoTo Measurement_err:
			End If
		
			'---AutoFocus---'	
			If  AutoFocus()= False Then
				GoTo Measurement_err:
			End If
		
			'---Align Wafer---'
			If WaferAlign() = False Then
				GoTo Measurement_err:
			End If
		
			'---Find Home---'
			If FindHome() = False Then
				GoTo Measurement_err:
			End If

			'---------Stepping and measurement---------'
			If  StepFirstDie(MPICommandResult1,MPICommandResult2,DieColumn,DieRow)= False Then
				GoTo Measurement_err:
			End If

			'    If  LightONOFF("Scope", "ON")= False Then
			'        GoTo Measurement_err:
			'    End If

			If  GetTestDiesNumber(MPICommandResult1,MPICommandResult2,TotalDieNumber)= False Then
				GoTo Measurement_err:
			End If

			TotalDieNum= Val(TotalDieNumber)

			If OpenMeasureNoiseWindow() <> 0 Then GoTo Measurement_err

			For i = 1 To TotalDieNum
				WaitSeconds(0.5)
				RandomBinNumber = ((6 - 1 + 1) * Rnd + 1)

			'	If doing the subsite stepping, please using command below

			'		If  StepNextSubsite(MPICommandResult1,MPICommandResult2,MPICommandResult3,MPICommandResult4,MPICommandResult5)= False Then
			'			GoTo Measurement_err:
			'		End If

			'		If  StepNextDieWithBin(RandomBinNumber, MPICommandResult1,MPICommandResult2,DieColumn,DieRow)= False Then
			'	        GoTo Measurement_err:
			'	    End If

				' Prober is ready in first die
				If i <> 1 Then
					If  StepNextDie(MPICommandResult1,MPICommandResult2,DieColumn,DieRow)= False Then
						GoTo Measurement_err:
					End If
				End If

				If  MoveChuckContact(MPICommandResult1,MPICommandResult2,MPICommandResult3)= False Then
					GoTo Measurement_err:
				End If


				If  LightONOFF("Scope", "OFF")= False Then
					GoTo Measurement_err:
				End If

				'-----------measure-------------
				'If SetNoiseDataSavePath("D:\ApiTest\MeasData\Die_" & DieColumn & "_" & DieRow & "_\") <> 0 Then GoTo Measurement_err
				'RetInt = MeasureCurrentDevice(0)

				'If RetInt = 1 Then
					'user paused
					'skip it or GoTo MeasurementEnd
				'	GoTo Measurement_err
				'ElseIf RetInt = -1 Then
					'measurement fail
					'skip it or GoTo MeasurementEnd
				'	GoTo Measurement_err
				'Else
				'measurement pass
				'show status
				'Call ShowMeasureStatus("Shot_" & xx & "_" & yy & "_Die_" & DieAdd & "_Teg_" & TegAdd , "measurement completed...")
				'End If

				WaitSeconds(0.1)
				'-------------------------------

				If  LightONOFF("Scope", "ON")= False Then
					GoTo Measurement_err:
				End If


				If  MoveChuckSeparation(MPICommandResult1,MPICommandResult2,MPICommandResult3)= False Then
					GoTo Measurement_err:
				End If


			Next i
			'-------Stepping and measurement---------'


			'----Wafer unload----'
			If  WaferUnload()= False Then
				GoTo Measurement_err:
			End If

		End If


		Next j


	


    ReleaseProbeGPIB()
	GoTo Finish:
Measurement_err:
	If MPICommandResult1 <> "11" Then
    End If
    ReleaseProbeGPIB()
Finish:
	Call CloseMeasureNoiseWindow
End Sub
