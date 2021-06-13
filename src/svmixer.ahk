/*
	Miroslav Bass
	bassmiroslav@gmail.com

	Dev-road:
		~ 23.02.2019 - v0.1 == Release
		~ 25.02.2019 - v0.2 == Auto-Update + Add selecter of all process/only multimedia
		~ 26.02.2019 - v0.3 == Change Volume in Active Window, Fix bugs
		~ 01.03.2019 - v0.3b == Fix autoupdate
		~ 05.03.2019 - v0.4 REBORN == Recoded profile system, Rewritten auto update system
		~ 12.03.2019 - v0.5 REBORN == Added save profile function, Fixed the minimizing a window of sound mixer
		~ 23.03.2019 - v0.6 RELEASE == Fix duplicate processes, Fix maximize a window

		* When loading a profile, the PID will be determined based on the application process name, and not from the specified .ini file.
		Now it makes sense to constantly use the load/unload profiles.

	Last stable version: [1.0] 3.23.2019
		~ Release
	Current version: [1.1] 6.13.2021
		~ [1.1.1] Code and tabulation optimization
		~ [1.1.2] Divided the functionality of the main file into several files for better code maintainability. Clear up code.
*/

#SingleInstance ignore
#Persistent
OnExit, ExitLabel

; Run as Admin
TryAgain:
if not A_IsAdmin
	Run *RunAs "%A_ScriptFullPath%",,UseErrorLevel
if (errorlevel)	{
	MsgBox, 262212, Sound Volume Mixer, SVMixer`, was NOT run as administrator.`n`n===`nTo make the functions work correctly, the script must be run by the administrator.`n`nPress Yes - to run as administrator.`nPress No - to continue.
	IfMsgBox Yes
		goto TryAgain
}

Gui, Margin, 10, 10
Gui, Add, ListView, w500 h600 vList +disabled +AltSubmit, PID|Process Name|Command Line

for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	LV_Add("", process.ProcessID, process.Name, process.CommandLine)
LV_ModifyCol() 
RemoveRecurStrings()

Gui, Add, Edit, x16 y650 w220 h20 vEdit +disabled, 
Gui, Add, Hotkey, x16 y625 w60 h20 gAWHotkey vAWHotkey +disabled, 
Gui, Add, Radio, x16 y675 w70 h20 vSelected gSelected +disabled, Selected
Gui, Add, Radio, x85 y675 w100 h20 vAW gAW +disabled, Active Window
Gui, Add, Button, x246 y625 w80 h20 +Disabled vEdit2 gEdit, Edit
Gui, Add, Button, x246 y650 w80 h20 vAddBtn gAddBtn, Add
Gui, Add, Button, x210 y675 w30 h20 vunload gunload , -
Gui, Add, Button, x180 y675 w30 h20 vload gload , +
Gui, Add, Button, x246 y675 w80 h20 +disabled vRemove gRemove, Remove
Gui, Add, GroupBox, x6 y610 w330 h95 , 
Gui, Add, GroupBox, x346 y610 w160 h95 , 
Gui, Add, DropDownList, x86 y625 w150 h20 R5 vDDsL hwndhDDL gDDL , 
Gui, Add, Radio, x352 y657 w75 w73 vAP gAllProcess +Checked, All process
Gui, Add, Radio, x430 y652 w80 w73 vOM gOnlyMedia, Only media
Gui, Add, Slider, x350 y680 w152 h20 vSlider Range0-100 ToolTip  +disabled, 0
Gui, Add, Button, x356 y625 w70 h20 vSave gSave +disabled, Save
Gui, Add, Button, x430 y625 w70 h20 vCancel  gCancel +disabled, Cancel

; To disable auto-update, comment out these lines
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("GET", "https://raw.githubusercontent.com/mrslv-bss/svmixer/main/README.md", false)
oWhr.Send()
RegExMatch(oWhr.ResponseText, "!\[Image alt\]\(https://img\.shields\.io/badge/Script%20Version-0.(\d)v-green\)", version)
Actualrelease := InStr(version1, 6) ? True : False
if !(Actualrelease)
	MsgBox, NEW VERSION
; To disable auto-update, comment out these lines

if (Actualrelease == "")
	Gui, Show, w510 h710, Sound Volume Mixer
else
	Gui, Show, w510 h710, Sound Volume Mixer | %version1%

Menu, tray, NoStandard
Menu, tray, add, Restore SVMixer, Restore 
Menu, tray, add 
Menu, tray, add, Quit, Exit  

#include svm_funcs.ahk
#include svm_profiles.ahk
return

; GUI All process Radio control
AllProcess:
LV_Delete()
for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	LV_Add("", process.ProcessID, process.Name, process.CommandLine)
LV_ModifyCol() 
RemoveRecurStrings()
return

; GUI Only Media Radio control
OnlyMedia:
LV_Delete()
Loop, Parse, % GetActive_Media(), `, 
{ 
	PATH := GetModuleFileNameEx(A_LoopField) 
	Loop, % PATH 
		Name := A_LoopFileName 
	LV_Add("", A_LoopField, Name, Path) 
}
LV_ModifyCol() 
RemoveRecurStrings()
return

; GUI Add button control
AddBtn:
GuiControl,enable,List
GuiControl, disable, load
GuiControl,disable, unload
add := true
return

; Launched in response to a right-click or press of the Apps key. Selecting process. ListView control.
GuiContextMenu:
selectedis := A_EventInfo
if (A_EventInfo != false)	{
	GuiControl, disable, AP
	GuiControl,disable,OM
	GuiControl, disable, DDsL
	GuiControl,disable,List
	GuiControl,disable,AddBtn
	GuiControl,enable,cancel
	GuiControl,enable,Selected
	GuiControl,enable,AW
	LV_GetText(selectedprocessPID, selectedis, 1)
	LV_GetText(selectedprocessName, selectedis, 2)
	gui, submit, nohide
	GuiControl,,edit, PID - %selectedprocessPID% | Process - %selectedprocessName%
}
return

; GUI Hotkey control
AWHotkey:
if (editvar != "save")	{
	GuiControl,enable,save
	GuiControl,enable,slider
}
return

; GUI Choosed process in ListView Radio control
Selected:
GuiControl,enable,AWHotkey
tosave := true
return

; GUI Active Window Radio control
AW:
GuiControl,enable,AWHotkey
tosave := "2"
return

; After filling in all the required fields, save the hotkey.
Save:
if (tosave = "2")	{
	selectedprocessPID = ActiveWindow
	selectedprocessName = ActiveWindow
}
if (AWHotkey != "")	{
	GuiControl,disable,AWHotkey
	gui, submit, nohide
	ItemText = %slider% | %AWHotkey% | %selectedprocessPID% | %selectedprocessName%
	SendMessage, 0x143, 0, &ItemText , , ahk_id %hDDL%  ;  CB_ADDSTRING 
	GuiControl,disable,Selected
	GuiControl,disable,AW
	GuiControl,disable,save
	GuiControl,disable, slider
	GuiControl,disable,cancel
	GuiControl,,AWHotkey,
	GuiControl, Enable, load
	GuiControl,Enable, unload
	LV_Modify(selectedis, "-Focus")
goto refreshprofiles
}
return

; Cancel save the hotkey
Cancel:
GuiControl,enable,DDsL
GuiControl,disable,Selected
GuiControl,disable,AW
GuiControl,disable,save
GuiControl,Enable,AddBtn
GuiControl,disable,cancel
GuiControl,,AWHotkey,
GuiControl,disable,AWHotkey
GuiControl,disable,slider
GuiControl,Enable,AP
GuiControl,Enable,OM
GuiControl, Enable, load
GuiControl,Enable, unload
LV_Modify(selectedis, "-Focus")
return

; Delete saved hotkey
remove:
gui, submit, nohide
RegExMatch(DDsL, "(.*) \Q|\E (.*) \Q|\E (.*) \Q|\E (.*)", rprofile)
hotkey, %rprofile2%, hotkey, off, UseErrorLevel
SendMessage, 0x158, 1, &DDsL, , ahk_id %hDDL%  ;  CB_FINDSTRINGEXACT
ItemIndex := ErrorLevel
SendMessage, 0x0144, ItemIndex, , , ahk_id %hDDL%  ;  CB_DELETESTRING
SendMessage, 0x014F, 1, , , ahk_id %hDDL%   ;	CB_SHOWDROPDOWN
GuiControl,disable,remove
GuiControl,disable,edit2
return

; DropDownList label, when we choose any saved hotkey.
DDL:
gui, submit, nohide
if (DDsL != "")	{
	GuiControl,enable,remove
	GuiControl,enable,edit2
}
return

; Edit saved hotkey
edit:
gui, submit, nohide
if (editvar != "save")	{
	RegExMatch(DDsL, "(.*)\Q|\E(.*)\Q|\E(.*)\Q|\E(.*)", profile)
		GuiControl,enable,slider
			GuiControl,enable,AWHotkey
				GuiControl,disable,list
				GuiControl,disable,remove
			GuiControl,disable,AddBtn
		GuiControl,, edit2, Save
	GuiControl, Disable, refreshprofiles
editvar := "save"
}	else	{
		if (AWHotkey = "")
			AWHotkey := profile2
		GuiControl,Disable,AWHotkey
		upd = %slider%|%AWHotkey%|%profile3%|%profile4%
			ItemText := DDsL
				Sleep 100
					SendMessage, 0x158, 1, &ItemText, , ahk_id %hDDL%  ;  CB_FINDSTRINGEXACT
						ItemIndex := ErrorLevel
							SendMessage, 0x0144, ItemIndex, , , ahk_id %hDDL%  ;  CB_DELETESTRING
							GuiControl,disable,slider
						ItemText := upd
					SendMessage, 0x014A, ItemIndex, &ItemText , , ahk_id %hDDL%  ;  CB_INSERTSTRING 
				GuiControl,, edit2, Edit
			GuiControl,disable,edit2
		GuiControl,Enable,AddBtn
	editvar = 0
SendMessage, 0x014F, 1, , , ahk_id %hDDL%   ;	CB_SHOWDROPDOWN
    }
return

; Running hotkey label
hotkey:
If (pid%A_ThisHotkey% = "ActiveWindow" Or process%A_ThisHotkey% = "ActiveWindow")
	WinGet, pid%A_ThisHotkey%, PID, A
SetAppVolume(pid%A_ThisHotkey%, %A_ThisHotkey%)
return

ExitLabel:
if A_ExitReason not in Logoff,Shutdown
{
    MsgBox, 262180, SoundMixerR | Exit, Are you sure you want to quit?
    IfMsgBox, No
        return
}
ExitApp
return

Restore:
Gui,Show
return

end::reload
!end::
Exit:
ExitApp