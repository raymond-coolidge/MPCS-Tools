#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=htelnet_2.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GUIListBox.au3>
#include <Date.au3>
#include <File.au3>
#include <ListBoxConstants.au3>

Opt('WinWaitDelay',65)

Func _WinWaitActivate($title,$text,$timeout=1)
   WinWait($title,$text,$timeout)
   If Not WinActive($title,$text) Then WinActivate($title,$text)
   WinWaitActive($title,$text,$timeout)
EndFunc

#Region ### START Koda GUI section ### Form=c:\users\raymond.coolidge\desktop\mpcs tools gui.kxf
$hGUI1 = GUICreate("MPCS Tools", 410, 150, 563, 858)
$MenuItem1 = GUICtrlCreateMenu("&Serial Lists")
$MenuItem2 = GUICtrlCreateMenuItem("Send List to Window", $MenuItem1)
$MenuItem6 = GUICtrlCreateMenuItem("Reprint Traveler Nonconformance Sheets", $MenuItem1)
$MenuItem8 = GUICtrlCreateMenuItem("Unlock Plan List", $MenuItem1)
$MenuItem3 = GUICtrlCreateMenu("&Reports")
$MenuItem9 = GUICtrlCreateMenuItem("Nonconformance Report", $MenuItem3)
$MenuItem4 = GUICtrlCreateMenuItem("WIP Reports Plan", $MenuItem3)
$MenuItem7 = GUICtrlCreateMenuItem("WIP Reports by PN", $MenuItem3)
$MenuItem5 = GUICtrlCreateMenuItem("WJ Airflow", $MenuItem3)
GUISetIcon("C:\Users\raymond.coolidge\Desktop\htelnet_2.ico", -1)
$LaunchMPCS = GUICtrlCreateButton("Launch MPCS", 168, 8, 97, 33)
$SNView = GUICtrlCreateButton("Complete S/N", 280, 8, 97, 33)
$Tab1 = GUICtrlCreateTab(16, 32, 377, 81)
GUICtrlSetState(-1,$GUI_SHOW)
$TabSheet1 = GUICtrlCreateTabItem("Ticketing Stuff")
$Button1 = GUICtrlCreateCheckbox("E/M Ticket", 36, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$Button2 = GUICtrlCreateCheckbox("Clear Dev", 124, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$Button3 = GUICtrlCreateCheckbox("Reprint NC", 212, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$Button4 = GUICtrlCreateCheckbox("Scrap S/N", 300, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$TabSheet2 = GUICtrlCreateTabItem("QE Stuff")
$Button5 = GUICtrlCreateCheckbox("Unlock", 36, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$Button6 = GUICtrlCreateCheckbox("Orion Param", 124, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
$Button7 = GUICtrlCreateCheckbox("Part Program", 212, 65, 73, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX,$BS_PUSHLIKE))
GUICtrlCreateTabItem("")
#EndRegion ### END Koda GUI section ###

Func _MPCSLogin($sFunction="")
	Local $sShortcut = IniRead("MPCS Login.ini", "Shortcuts", $sFunction, Default)
	Local $sUsername = IniRead("MPCS Login.ini", "Credentials", "Username", Default)
	Local $sPassword = IniRead("MPCS Login.ini", "Credentials", "Password", Default)

	ShellExecute("C:\Program Files (x86)\HostExplorer.nt\user\alpha.hts")
	_WinWaitActivate("Connect","")
	Local $Conn_hWnd = WinGetHandle("Connect")
	ControlClick($Conn_hWnd, "OK", 1)
	_WinWaitActivate("WyseTerm","")
	Local $MPCS_hWnd = WinGetHandle("WyseTerm")

	If Not $sFunction = "" Then
		ControlSend($MPCS_hWNd, "", "", $sUsername & @CR & $sPassword & @CR & "mpcs" & @CR & $sShortcut & @CR)
	Else
		ControlSend($MPCS_hWNd, "", "", $sUsername & @CR & $sPassword & @CR & "mpcs" & @CR)
	EndIf

	Return $MPCS_hWnd
EndFunc

Global $iGUIWidth = 240, $iGUIHeight = 400, $buttonposition = $iGUIWidth/2 - 82.5

$hGUI2 = GUICreate("Auto Serial Entry", $iGUIWidth, $iGUIHeight, -1, -1, -1, $WS_EX_TOPMOST)
$idEdit_1 = GUICtrlCreateEdit("", 10, 10, $iGUIWidth-20, $iGUIHeight-60)
$idSend_Btn = GUICtrlCreateButton("S/N Only", $buttonposition, 360, 75, 30)
$idClear_Btn = GUICtrlCreateButton("S/N + CLR", $buttonposition + 90, 360, 75, 30)

$hGUI4 = GUICreate("Reprint NC Sheets", $iGUIWidth, $iGUIHeight, -1, -1, -1)
$idEdit_2 = GUICtrlCreateEdit("", 10, 10, $iGUIWidth-20, $iGUIHeight-60)
$idReprint_Btn = GUICtrlCreateButton("Reprint", $buttonposition, 360, 75, 30)
$idCancel_Btn2 = GUICtrlCreateButton("Cancel", $buttonposition + 90, 360, 75, 30)

$hGUI6 = GUICreate("Unlock Plan List", $iGUIWidth, $iGUIHeight, -1, -1, -1)
$idEdit_3 = GUICtrlCreateEdit("", 10, 10, $iGUIWidth-20, $iGUIHeight-60)
$idUnlock_Btn = GUICtrlCreateButton("Unlock", $buttonposition, 360, 75, 30)
$idCancel_Btn3 = GUICtrlCreateButton("Cancel", $buttonposition + 90, 360, 75, 30)

#Region ### START Koda GUI section ### Form=
$hGUI3 = GUICreate("WIP Reports", 202, 354, 441, 282)
$List1 = GUICtrlCreateList("", 8, 8, 185, 279, BitOR($LBS_EXTENDEDSEL, $LBS_NOTIFY, $WS_VSCROLL))
GUICtrlSetData(-1, "")
$idQuery_Btn = GUICtrlCreateButton("Query", 24, 304, 73, 33)
$idQuit_Btn = GUICtrlCreateButton("Quit", 104, 304, 73, 33)
#EndRegion ### END Koda GUI section ###

#Region ### START Koda GUI section ### Form=
$hGUI5 = GUICreate("WIP Reports", 202, 354, 441, 282)
$List2 = GUICtrlCreateList("", 8, 8, 185, 279, BitOR($LBS_EXTENDEDSEL, $LBS_NOTIFY, $WS_VSCROLL))
GUICtrlSetData(-1, "")
$idQuery_Btn2 = GUICtrlCreateButton("Query", 24, 304, 73, 33)
$idQuit_Btn2 = GUICtrlCreateButton("Quit", 104, 304, 73, 33)
#EndRegion ### END Koda GUI section ###

Func _SendArray($pWindowHandle, $idControl, $action=@CR)
	Local $iMax, $sInput
	Local $aArray = 0
	$sInput = GUICtrlRead($idControl)
	$aArray = StringSplit($sInput, @CR)
	WinActivate($pWindowHandle)
	If IsArray($aArray) Then
		$iMax = UBound($aArray)
		For $i = 1 to $iMax - 1
			ControlSend($pWindowHandle, "", "", $aArray[$i] & $action)
		Next
	EndIf
EndFunc

Func _UnlockPlanList()
	Local $pUnlock = _MPCSLogin("SPM1")

	Local $iMax, $sInput
	Local $aArray = 0
	$sInput = GUICtrlRead($idEdit_3)
	$aArray = StringSplit($sInput, @CR)
	WinActivate($pUnlock)

	Opt("SendKeyDelay", 50)
	If IsArray($aArray) Then
		$iMax = UBound($aArray)
		For $i = 1 to $iMax - 1
			ControlSend($pUnlock, "", "", "{TAB 2}" & $aArray[$i] & "{INSERT}{TAB}{SPACE}{F6}{END}{DEL}")
		Next
	EndIf
	Opt("SendKeyDelay", 5)

	Sleep(500)
	WinClose($pUnlock)
EndFunc

Global $PNList = 0, $PlanList = 0
_FileReadToArray("PN List.csv", $PNList, $FRTA_NOCOUNT,",")
_FileReadToArray("Plan List.csv", $PlanList, $FRTA_NOCOUNT,",")

Func _Plan_Populate()
	$iMax = UBound($PlanList, 1)
	If IsArray($PlanList) Then
		For $i = 0 to $iMax - 1
			GUICtrlSetData($List1, $PlanList[$i][2])
		Next
	EndIf
EndFunc

Func _PN_Populate()
	$iMax = UBound($PNList, 1)
	If IsArray($PNList) Then
		For $i = 0 to $iMax - 1
			GUICtrlSetData($List2, $PNList[$i][1])
		Next
	EndIf
EndFunc

Func _DateFormat($Date, $style)
    Local $hGui = GUICreate("My GUI get date", 200, 200, 800, 200)
    Local $idDate = GUICtrlCreateDate($Date, 10, 10, 185, 20)
    GUICtrlSendMsg($idDate, 0x1032, 0, $style)
    Local $sReturn = GUICtrlRead($idDate)
    GUIDelete($hGui)
    Return $sReturn
EndFunc

Func _AirflowReports()
	$planID = InputBox("Plan ID", "Enter plan ID of nozzle")
	$days = InputBox("Days", "How many days would you like to query?")

	$EndDate = _NowCalcDate()
	$StartDate = _DateAdd('d', -$days, _NowCalcDate())
	$sEndDate = _DateFormat($EndDate, "dd-MMM-yyyy")
	$sStartDate = _DateFormat($StartDate, "dd-MMM-yyyy")

	If $planID <> "" And $days <> "" Then
		$FNR53_hWnd = _MPCSLogin("FNR53")
		Opt("SendKeyDelay", 50)
		ControlSend($FNR53_hWnd, "", "", "98{ENTER}0{ENTER}END{ENTER}" & $planID & "{ENTER 2}1{ENTER 3}" & $sStartDate & "{ENTER}" & $sEndDate & "{ENTER}1{ENTER}END{ENTER}1{ENTER}")
		Sleep(250)
		ControlSend($FNR53_hWnd, "", "", "{ENTER}")
		WinClose($FNR53_hWnd)
		Opt("SendKeyDelay", 5)
	EndIf
EndFunc

Func _NonconformanceReport()
	$EndDate = _NowCalcDate()
	$sEndDate = _DateFormat($EndDate, "dd-MMM-yyyy")

	$FNR13_hWnd = _MPCSLogin("FNR13")

	Opt("SendKeyDelay", 50)
	ControlSend($FNR13_hWnd, "", "", "77{ENTER}2{ENTER}98{ENTER}0{ENTER}END{ENTER}1{ENTER}01-JAN-1990{ENTER}" & $sEndDate & "{ENTER}ALL{ENTER}ALL{ENTER 3}")
	Opt("SendKeyDelay", 5)

	Sleep(250)
	WinClose($FNR13_hWnd)
EndFunc

Func _ReprintTravelers()
	$ReprintPlan = InputBox("Plan ID", "Enter plan ID for serial list")
	$ReprintSheet = InputBox("Nonconformance Sheet", "Enter sheet number for nonconformance sheet for plan " & $ReprintPlan)

	If $ReprintPlan <> "" And $ReprintSheet <> "" Then
		$FWR2_hWnd = _MPCSLogin("FWR2")
		Opt("SendKeyDelay", 50)
		ControlSend($FWR2_hWnd, "", "", "2{ENTER}" & $ReprintPlan & "{ENTER}")
		_SendArray($FWR2_hWnd, $idEdit_2, "{ENTER}2{ENTER}1{ENTER}" & $ReprintSheet & "{ENTER}")
		ControlSend($FWR2_hWnd, "", "", "End{ENTER 2}3{ENTER 2}137{ENTER}")
		Opt("SendKeyDelay", 5)
	EndIf
EndFunc

Func _PlanListWIP()
	$PlanList_Sel = _GUICtrlListBox_GetSelItems($List1)
	$FWR14_hWnd = _MPCSLogin("FWR14")
	Opt("SendKeyDelay", 50)
	ControlSend($FWR14_hWnd, "", "", "98{ENTER}0{ENTER}END{ENTER}Y{ENTER}2{ENTER}")
		If IsArray($PlanList_Sel) Then
			$iMax = UBound($PlanList_Sel)
			For $i = 1 to $iMax - 1
				ControlSend($FWR14_hWnd, "", "", $PlanList[$PlanList_Sel[$i]][1] & "{ENTER 4}" & $PlanList[$PlanList_Sel[$i]][0] & "{ENTER}")
			Next
		EndIf
	ControlSend($FWR14_hWnd, "", "", "END{ENTER}")
		Sleep(500)
	ControlSend($FWR14_hWnd, "", "", "{ENTER}")
	WinClose($FWR14_hWnd)
	Opt("SendKeyDelay", 5)
EndFunc

Func _PNListWIP()
	$PNList_Sel = _GUICtrlListBox_GetSelItems($List2)
	$FWR14_hWnd = _MPCSLogin("FWR14")
	Opt("SendKeyDelay", 50)
	ControlSend($FWR14_hWnd, "", "", "98{ENTER}0{ENTER}END{ENTER}Y{ENTER}2{ENTER}")
		If IsArray($PNList_Sel) Then
			$iMax = UBound($PNList_Sel)
			For $i = 1 to $iMax - 1
				ControlSend($FWR14_hWnd, "", "", $PNList[$PNList_Sel[$i]][0] & "{ENTER 5}")
			Next
		EndIf
	ControlSend($FWR14_hWnd, "", "", "END{ENTER}")
		Sleep(500)
	ControlSend($FWR14_hWnd, "", "", "{ENTER}")
	WinClose($FWR14_hWnd)
	Opt("SendKeyDelay", 5)
EndFunc

_PN_Populate()
_Plan_Populate()
_MPCSTools()

Func _MPCSTools()
	GUISetState(@SW_SHOW, $hGUI1)
	While 1
		$aMsg = GUIGetMsg(1)
		Switch $aMsg[1]
			Case $hGUI1
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete()
						Exit
					Case $LaunchMPCS
						_MPCSLogin()
					Case $SNView
						$SNView_hWnd = _MPCSLogin("FWV6")
					Case $Button1
						Switch GUICtrlRead($Button1)
							Case $GUI_CHECKED
								Global $FNM4_hWnd = _MPCSLogin("FNM4")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button1, $GUI_CHECKED)
								WinActivate($FNM4_hWnd)
						EndSwitch
					Case $Button2
						Switch GUICtrlRead($Button2)
							Case $GUI_CHECKED
								Global $FNM12_hWnd = _MPCSLogin("FNM12")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button2, $GUI_CHECKED)
								WinActivate($FNM12_hWnd)
						EndSwitch
					Case $Button3
						Switch GUICtrlRead($Button3)
							Case $GUI_CHECKED
								Global $FNR45_hWnd = _MPCSLogin("FNR45")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button3, $GUI_CHECKED)
								WinActivate($FNR45_hWnd)
						EndSwitch
					Case $Button4
						Switch GUICtrlRead($Button4)
							Case $GUI_CHECKED
								Global $FNM5_hWnd = _MPCSLogin("FNM5")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button4, $GUI_CHECKED)
								WinActivate($FNM5_hWnd)
						EndSwitch
					Case $Button5
						Switch GUICtrlRead($Button5)
							Case $GUI_CHECKED
								Global $SPM1_hWnd = _MPCSLogin("SPM1")
								ControlSend($SPM1_hWnd, "", "", "{TAB 2}")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button5, $GUI_CHECKED)
								WinActivate($SPM1_hWnd)
						EndSwitch
					Case $Button6
						Switch GUICtrlRead($Button6)
							Case $GUI_CHECKED
								Global $CPM3_hWnd = _MPCSLogin("CPM3")
								ControlSend($CPM3_hWnd, "", "", "{TAB 2}")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button6, $GUI_CHECKED)
								WinActivate($CPM3_hWnd)
						EndSwitch
					Case $Button7
						Switch GUICtrlRead($Button7)
							Case $GUI_CHECKED
								Global $CPM20_hWnd = _MPCSLogin("CPM20")
							Case $GUI_UNCHECKED
								GUICtrlSetState($Button7, $GUI_CHECKED)
								WinActivate($CPM20_hWnd)
						EndSwitch
					Case $GUI_EVENT_SECONDARYUP
						$aCInfo = GUIGetCursorInfo($hGUI1)
						If GUICtrlRead($aCInfo[4]) = $GUI_CHECKED Then
							Switch $aCInfo[4]
								Case $Button1
									WinClose($FNM4_hWnd)
									GUICtrlSetState($Button1, $GUI_UNCHECKED)
								Case $Button2
									WinClose($FNM12_hWnd)
									GUICtrlSetState($Button2, $GUI_UNCHECKED)
								Case $Button3
									WinClose($FNR45_hWnd)
									GUICtrlSetState($Button3, $GUI_UNCHECKED)
								Case $Button4
									WinClose($FNM5_hWnd)
									GUICtrlSetState($Button4, $GUI_UNCHECKED)
								Case $Button5
									WinClose($SPM1_hWnd)
									GUICtrlSetState($Button5, $GUI_UNCHECKED)
								Case $Button6
									WinClose($CPM3_hWnd)
									GUICtrlSetState($Button6, $GUI_UNCHECKED)
								Case $Button7
									WinClose($CPM20_hWnd)
									GUICtrlSetState($Button7, $GUI_UNCHECKED)
							EndSwitch
						EndIf
					Case $MenuItem2
						GUISetState(@SW_SHOW, $hGUI2)
						$pSendTo = _MPCSLogin()
					Case $MenuItem6
						GUISetState(@SW_SHOW, $hGUI4)
					Case $MenuItem4
						GUISetState(@SW_SHOW, $hGUI3)
					Case $MenuItem7
						GUISetState(@SW_SHOW, $hGUI5)
					Case $MenuItem8
						GUISetState(@SW_SHOW, $hGUI6)
					Case $MenuItem5
						_AirflowReports()
					Case $MenuItem9
						_NonconformanceReport()
				EndSwitch
			Case $hGUI2
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI2)
						GUICtrlSetData($idEdit_1, "")
						WinClose($pSendTo)
					Case $idSend_Btn
						_SendArray($pSendTo, $idEdit_1)
					Case $idClear_Btn
						_SendArray($pSendTo, $idEdit_1, @CR & "CLR" & @CR)
				EndSwitch
			Case $hGUI4
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI4)
						GUICtrlSetData($idEdit_2, "")
					Case $idReprint_Btn
						_ReprintTravelers()
					Case $idCancel_Btn2
						GUISetState(@SW_HIDE, $hGUI4)
						GUICtrlSetData($idEdit_2, "")
				EndSwitch
			Case $hGUI3
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI3)
					Case $idQuery_Btn
						_PlanListWIP()
					Case $idQuit_Btn
						GUISetState(@SW_HIDE, $hGUI3)
				EndSwitch
			Case $hGUI5
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI5)
					Case $idQuery_Btn2
						_PNListWIP()
					Case $idQuit_Btn2
						GUISetState(@SW_HIDE, $hGUI5)
				EndSwitch
			Case $hGUI6
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI6)
						GUICtrlSetData($idEdit_3, "")
					Case $idUnlock_Btn
						_UnlockPlanList()
					Case $idCancel_Btn3
						GUISetState(@SW_HIDE, $hGUI6)
						GUICtrlSetData($idEdit_3, "")
				EndSwitch
		EndSwitch
	WEnd
EndFunc