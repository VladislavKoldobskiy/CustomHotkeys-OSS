#NoEnv
#Persistent
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%  

OnClipboardChange("ClipChange")
global strRequestId := "None"
global strRequestType := "None"
global strRequestDesc := "None"

global intLocalErrors := 0
global intMediaCtl := 1
global intMonitorCtl := 0
global intTooltipShown := 0

global strErrPath := "C:\_binaries\err.exe"
global strDDMPath := "C:\Program Files (x86)\Dell\Dell Display Manager\ddm.exe"
global strPSPath := "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
global strAppTitle := "CustomHotkeys"
global strVersion := "Version 0.1.8-oss"

global strMsgActionsErrors := "AltGr+J to open in //errors"
global strMsgActionsPhone := "AltGr+J to call"
global strMsgActionsRave := "AltGr+J to open in Rave"
global strMsgActionsSD := "AltGr+J to open SR in CaseBuddy`nAltGr+K to open SR in ServiceDesk`nAltGr+L to open SR in AzureSupportCenter`nAltGr+; to open DTM for this SR"
global strMsgActionsSDCase := "AltGr+{J/K/L/;} to open in CB/SD/ASC/DTM"
global strMsgActionsSDProblem := "AltGr+{J/K/L/;} to open related SR in CB/SD/ASC/DTM"
global strMsgActionsEvoSTS := "AltGr+J to look up in LogsMiner"
global strMsgActionsTenantID := "AltGr+J to open in Rave" ;NOT USED
global strMsgActionPaste := "Ctrl+Shift+V to paste this number"

global strMsgDisplaysHigh := "Displays set to HIGH brightness"
global strMsgDisplaysMed := "Displays set to medium brightness"
global strMsgDisplaysComf := "Displays set to ComfortView"

global strMsgDisplaysLEAOff := "Left display EasyArrange off" ;0
global strMsgDisplaysLEAMax := "Left display EasyArrange mode: single window maximized" ;1
global strMsgDisplaysLEA6W := "Left display EasyArrange mode: 6 windows" ;19
global strMsgDisplaysLEA4W := "Left display EasyArrange mode: 4 windows" ;10
global strMsgDisplaysLEA3WH := "Left display EasyArrange mode: 3 windows horizontally" ;4
global strMsgDisplaysLEA3WS := "Left display EasyArrange mode: 3 windows with 2 split vertically" ;7
global strMsgDisplaysLEA3WV := "Left display EasyArrange mode: 3 windows vertically" ;5
global strMsgDisplaysLEA2W := "Left display EasyArrange mode: 2 windows" ;14

global strMsgDisplaysREAOff := "Right display EasyArrange off" ;0
global strMsgDisplaysREAMax := "Right display EasyArrange mode: single window maximized" ;1
global strMsgDisplaysREA6W := "Right display EasyArrange mode: 6 windows" ;19
global strMsgDisplaysREA4W := "Right display EasyArrange mode: 4 windows" ;10
global strMsgDisplaysREA3WH := "Right display EasyArrange mode: 3 windows horizontally" ;4
global strMsgDisplaysREA3WS := "Right display EasyArrange mode: 3 windows with 2 split vertically" ;6
global strMsgDisplaysREA3WV := "Right display EasyArrange mode: 3 windows vertically" ;5
global strMsgDisplaysREA2W := "Right display EasyArrange mode: 2 windows" ;13

global strMsgUnderConstruction := "This function is under construction."

Menu, Tray, NoStandard
AppInit()

return
;;;;;;;;;;;;;;;;;;; MENU AND TIMER LABELS ;;;;;;;;;;;;;;;;;;;
lblExit:
ExitApp

lblNop:
return

;;;;;;;;;;;;;;;;;;; MEDIA CONTROL ;;;;;;;;;;;;;;;;;;;
RAlt & Volume_Down::HandleAltGrVolDn()
RAlt & Volume_Up::HandleAltGrVolUp()
;;;;;;;;;;;;;;;;;;; MONITOR CONTROL ;;;;;;;;;;;;;;;;;;;
RAlt & I::HandleAltGrI()
RAlt & O::HandleAltGrO()
RAlt & P::HandleAltGrP()
RAlt & 1::HandleAltGr1()
RAlt & 2::HandleAltGr2()
RAlt & 3::HandleAltGr3()
RAlt & 4::HandleAltGr4()
RAlt & Q::HandleAltGrQ()
RAlt & W::HandleAltGrW()
RAlt & E::HandleAltGrE()
RAlt & R::HandleAltGrR()
RAlt & A::HandleAltGrA()
RAlt & S::HandleAltGrS()
RAlt & D::HandleAltGrD()
RAlt & F::HandleAltGrF()
RAlt & Z::HandleAltGrZ()
RAlt & X::HandleAltGrX()
RAlt & C::HandleAltGrC()
RAlt & V::HandleAltGrV()

;;;;;;;;;;;;;;;;;;; APP SHORTCUTS ;;;;;;;;;;;;;;;;;;;
RAlt & J::HandleAltGrJ()
RAlt & K::HandleAltGrK()
RAlt & L::HandleAltGrL()
RAlt & SC027::HandleAltGrSC027() ;semicolon
RAlt & SC028::HandleAltGrSC028() ;quote

RAlt & \::
;;;; I wanted to use this hotkey to start/stop labor timer for a case number currently in clipboard.
;;;; This, however, would require me to make up some sort of a GUI, and I was too lazy to ever get this implemented
ToolTip %strMsgUnderConstruction%
return

;;;;;;;;;;;;;;;;;;; MISC ;;;;;;;;;;;;;;;;;;;
RAlt::HandleAltGr()
^+c::HandleCtrlShiftC()
^+v::HandleCtrlShiftV()

;;;;;;;;;;;;;;;;;;; TEXT REPLACEMENT ;;;;;;;;;;;;;;;;;;;
;;;; I put various links here so I can quickly paste them without going to browser favorites or onenote.
::~!fd::https://www.telerik.com/download/fiddler/fiddler4
::~!ut::https://docs.microsoft.com/azure/active-directory/users-groups-roles/domains-admin-takeover
return

;;;;;;;;;;;;;;;;;;; CLIPBOARD HANDLING ;;;;;;;;;;;;;;;;;;;
;;;; When copying something, we apply a bunch of regexes and try to find something that can be interpreted as something actionable.
;; Things to try:
;218120525001785003 - problem id
;119032522001555 - SR ID
;119011119537312
;13444634 - Rave case
;13537486
;+79162609254 - Phone number
;+7 (495) 9299947
;+49 (89) 31798765
;+498931712345
;+966 54 999 2113
;+966544442113
;+4352245223378
;(916)23344552
;+7 (495) 780-73-00 | 3160
;0x803f7001 LM_E_SATISFACTION_OUT_OF_RANGE
;-2147220495 DUI_E_INVALIDPROPVALUE	
;0x80070005 E_ACCESS_DENIED
;- Tenant Id: 457b58d2-8a32-47d9-a265-ea94ab9ddfdf
; Tenant Id: f8cdef31-a31e-4b4a-93e4-5f571e91255a

ClipChange(Type) {
	intTooltipShown := 1
	clipContent := clipboard
	if (RegExMatch(clipContent, "[2-3][1]\d{16}", strPossibleRequestId)) { ;MSSolve problem/task
		StringLeft strRequestId, strPossibleRequestId, 15
		strRequestId := RegExReplace(strRequestId, "^[2-3]", "1",,1)
		ToolTip This looks like an problem/task number: %strPossibleRequestId%`n%strMsgActionsSDProblem%`n%strMsgActionPaste%
		strRequestType := "SD"
	} else if (RegExMatch(clipContent, "[1][1]\d{13}", strPossibleRequestId)) { ;MSSolve case
		ToolTip This looks like an SR number: %strPossibleRequestId%`n%strMsgActionsSDCase%`n%strMsgActionPaste%
		strRequestType := "SD"
		strRequestId := strPossibleRequestId
	} else if (RegExMatch(clipContent, "(\+\d{1,3}|8)[\s.-]?\(?\d{2,3}\)?[\s.-]?\d{3}[\s.-]?\d{2}[\s.-]?\d{2,3}", strPossibleRequestId)) { ;phone number
		;TODO: clear formatting in phone number
		ToolTip This looks like a phone number: %strPossibleRequestId%`n%strMsgActionsPhone%`n%strMsgActionPaste%
		strRequestType := "Phone"
		strRequestId := strPossibleRequestId
	} else if (RegExMatch(clipContent, "\b[1]\d{7}\b", strPossibleRequestId)) { ;Rave case
		ToolTip This looks like a Rave case: %strPossibleRequestId%`n%strMsgActionsRave%`n%strMsgActionPaste%
		strRequestType := "Rave"
		strRequestId := strPossibleRequestId
	} else if (RegExMatch(clipContent, "0[xX][0-9a-fA-F]{1,8}", strPossibleRequestId)) { ;0xZZZZZZZZ error code
		if(intLocalErrors) {
			GetErrorDesc(strPossibleRequestId)
		}
		ToolTip This looks like an error code: %strPossibleRequestId%`n%strRequestDesc%`n%strMsgActionsErrors%`n%strMsgActionPaste%
		strRequestType := "Errors"
		strRequestId := strPossibleRequestId
	} else if (RegExMatch(clipContent, "Tenant Id:\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})", strPossibleRequestId)) { ;Tenant ID, copied from case details
		if(strPossibleRequestId1 = "f8cdef31-a31e-4b4a-93e4-5f571e91255a") {
			strRequestDesc := "`nThis is a default Microsoft Services tenant GUID"
		} else {
			strRequestDesc := ""
		}
		ToolTip This looks like a Tenant ID: %strPossibleRequestId1%%strRequestDesc%`n%strMsgActionsRave%`n%strMsgActionPaste%
		;BUG: ctrl+shift combo gets stuck when trying to Ctrl+Shift+V this
		strRequestType := "TenantID"
		strRequestId := strPossibleRequestId1
	} else if (RegExMatch(clipContent, "(Date\s(.*))|(Correlation Id\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}))", strPossibleRequestId)) { ;eSTS correlation ID from ASC
		;;;; Completely unfinished. Another regex is below, If I remember correctly this one supports both eSTS message format (ie what you see in the error page) and ASC
		;(Date\s(.*))|(Correlation Id:?\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}))|(Timestamp:\s(20[1-9]{2}-(0?[1-9]|1[0-2])-(0?[1-9]|[12][0-9]|3[01])T(0[0-9]|1[0-9]|2[0-3])(:[0-5][0-9]){2}.[0-9]{3}Z))
		; group 2 - date
		; group 4 - correlation id
		; group 6 - timestamp
		strRequestType := "eSTS"
		ToolTip This looks like is an evoSTS correlation ID from ASC, but these are not yet supported
	} else {
		ToolTip
		intTooltipShown := 0
	}
}
AppInit() {
	Menu, mFeatureFlags, Add, MonitorControl, lblNop
	Menu, mFeatureFlags, Disable, MonitorControl
	Menu, mFeatureFlags, Add, LocalErrors, lblNop
	Menu, mFeatureFlags, Disable, LocalErrors
	Menu, mFeatureFlags, Add, MediaControl, lblNop
	Menu, mFeatureFlags, Disable, MediaControl
	Menu, Tray, Add, %strVersion%, lblNop
	Menu, Tray, Disable, %strVersion%
	Menu, Tray, Add, Feature flags, :mFeatureFlags
	Menu, Tray, Add ;separator
	Menu, Tray, Add, Exit, lblExit
	Menu, Tray, Tip, %strAppTitle%

	if (FileExist(strDDMPath)) {
		intMonitorCtl := 1
		Menu, mFeatureFlags, Check, MonitorControl
	}
	
	if (FileExist(strErrPath)) {
		intLocalErrors := 1
		Menu, mFeatureFlags, Check, LocalErrors
	}
	
	if (intMediaCtl := 1) {
		;;;; This feature flag was mapped to alias check earlier, but now enabled for everyone. Disable in Globals if you don't like it.
		Menu, mFeatureFlags, Check, MediaControl
	}

	ToolTip %strAppTitle% started
	Sleep 2000
	ToolTip
	return
}
HandleAltGrI() {
	if (intMonitorCtl) {
		Run %strDDMPath% /SetNamedPreset Standard /SetBrightnessLevel 100
		;Run %strPSPath% -NoLogo -NoLogo -NonInteractive -WindowStyle Hidden -Command "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1`,100)",, Hidden
		;;;; I intentionally commented this out because PS window popping up for a split second wasn't looking nice, and my X1 Yoga LCD flickers at anything < 100% brightness.
		ToolTip %strMsgDisplaysHigh%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrO() {
	if (intMonitorCtl) {
		Run %strDDMPath% /SetNamedPreset Standard /SetBrightnessLevel 75
		;Run %strPSPath% -NoLogo -NoLogo -NonInteractive -WindowStyle Hidden -Command "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1`,75)",, Hidden
		ToolTip %strMsgDisplaysMed%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrP() {
	if (intMonitorCtl) {
		Run %strDDMPath% /SetNamedPreset ComfortView /SetBrightnessLevel 75
		;Run %strPSPath% -NoLogo -NoLogo -NonInteractive -WindowStyle Hidden -Command "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1`,75)",, Hidden
		ToolTip %strMsgDisplaysComf%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGr1() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 0
		ToolTip %strMsgDisplaysLEAOff%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGr2() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 1
		ToolTip %strMsgDisplaysLEAMax%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGr3() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 1
		ToolTip %strMsgDisplaysREAMax%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGr4() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 0
		ToolTip %strMsgDisplaysREAOff%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrQ() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 19
		ToolTip %strMsgDisplaysLEA6W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrW() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 10
		ToolTip %strMsgDisplaysLEA4W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrE() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 10
		ToolTip %strMsgDisplaysREA4W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrR() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 19
		ToolTip %strMsgDisplaysREA6W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrA() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 4
		ToolTip %strMsgDisplaysLEA3WV%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrS() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 7
		ToolTip %strMsgDisplaysLEA3WS%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrD() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 6
		ToolTip %strMsgDisplaysREA3WS%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrF() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 4
		ToolTip %strMsgDisplaysREA3WH%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrZ() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 5
		ToolTip %strMsgDisplaysLEA3WV%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrX() {
	if (intMonitorCtl) {
		Run %strDDMPath% /1:SetGridType 14
		ToolTip %strMsgDisplaysLEA2W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrC() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 13
		ToolTip %strMsgDisplaysREA2W%
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrV() {
	if (intMonitorCtl) {
		Run %strDDMPath% /2:SetGridType 5
		ToolTip %strMsgDisplaysREA3WV%
		Sleep 2000
		ToolTip
	}
	return
}

HandleAltGrJ() {
	if (strRequestType = "SD") {
		ToolTip Opening %strRequestId% in CaseBuddy
		Run "mscb:case?%strRequestId%"
	} else if (strRequestType = "Rave"||strRequestType = "TenantID") {
		ToolTip Opening %strRequestId% in Rave
		Run "https://rave.office.net/search?query=%strRequestId%"
	} else if (strRequestType = "Phone") {
		ToolTip Calling %strRequestId%
		Run "tel:%strRequestId%"
	} else if (strRequestType = "Errors") {
		ToolTip Opening //errors on %strRequestId%
		Run "http://errors.corp.microsoft.com/?%strRequestId%"
	} else if (strRequestType = "EvoSTS") {
		ToolTip Opening LogsMiner for this Correlation ID
	}
	Sleep 2000
	ToolTip
	return
}
HandleAltGrK() {
	if (strRequestType = "SD") {
		ToolTip Opening %strRequestId% in ServiceDesk
		Run "https://servicedesk.microsoft.com/#/customer/cases?caseNumber=%strRequestId%"
	}
	return
}
HandleAltGrL() {
	if (strRequestType = "SD") {
		ToolTip Opening %strRequestId% in Azure Support Center
		Run "https://azuresupportcenter.msftcloudes.com/ticket?srId=%strRequestId%"
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrSC027() {
	if (strRequestType = "SD") {
		ToolTip Opening DTM workspace for %strRequestId%
		Run "https://client.dtmnebula.microsoft.com/?srNumber=%strRequestId%"
		Sleep 2000
		ToolTip
	}
	return
}
HandleAltGrSC028() {
	;;;; This was used earlier when CaseParserPro was alive.
	return
}
HandleAltGrVolDn() {
	if (intMediaCtl) {
		ToolTip Switching to previous track
		Send {Media_Prev}
		Sleep 500
		ToolTip
	}
}
HandleAltGrVolUp() {
	if (intMediaCtl) {
		ToolTip Switching to next track
		Send {Media_Next}
		Sleep 500
		ToolTip
	}
}
HandleAltGr() {
	if (intTooltipShown = 0) {
		intTooltipShown := 1

		if (strRequestType = "SD") {
			ToolTip ServiceDesk case: %strRequestId%`nPossible actions:`n%strMsgActionsSD%
		} else if (strRequestType = "Rave") {
			ToolTip Rave case: %strRequestId%`nPossible actions:`n%strMsgActionsRave%
		} else if (strRequestType = "Phone") {
			ToolTip Phone number: %strRequestId%`nPossible actions:`n%strMsgActionsPhone%
		} else if (strRequestType = "Errors") {
			ToolTip Error code: %strRequestId%`n%strRequestDesc%`nPossible actions:`n%strMsgActionsErrors%
		} else if (strRequestType = "evoSTS") {
			ToolTip %strMsgUnderConstruction%
		} else if (strRequestType = "TenantID") {
			ToolTip Tenant ID: %strRequestId%`nPossible actions:`n%strMsgActionsRave%
		} else {
			ToolTip Copy something to clipboard first.
			Sleep 2000
			ToolTip
		}
	} else {
		intTooltipShown := 0
		ToolTip
	}
	return
}
HandleCtrlShiftC() {
	if (strRequestType != "None") {
		clipboard := strRequestId
		Sleep 50
		ToolTip %strRequestId% copied to clipboard
		Sleep 5000
		ToolTip
	}
	return
}
HandleCtrlShiftV() {
	if (strRequestType != "None") {
		SendInput {Text}%strRequestId%
	}
	return
}
GetErrorDesc(ErrorCode) {
	;;;; This, unfortunately, still blinks command line window for a moment.
	;;;; If you want to rewrite this, ideal approach would be adapt errors.h from //toolbox so there is no dependency on another binary at all
	;;;; However, this will require quite a lot of work to actually convert these data
	cmdline := strErrPath " " ErrorCode
	strRequestDesc := ComObjCreate("WScript.Shell").Exec(cmdline).StdOut.ReadAll()
}
