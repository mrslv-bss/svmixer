#SingleInstance ignore
#Persistent
OnExit, ExitLabel
ClientVersion := "6"

if not A_IsAdmin	; Запуск от им.админа
    Run *RunAs "%A_ScriptFullPath%",,UseErrorLevel

if (ErrorLevel)		{	; Запущена от админа ли
	MsgBox, 262160, SoundMixer Reborn, For the script to work properly`, you must run it with admin rights.
	ExitApp
}

;~  AutoUpdate
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
oWhr.Open("GET", "https://raw.githubusercontent.com/MirchikAhtung/soundmixer/master/readme.txt", false)
oWhr.Send()
RegExMatch(oWhr.ResponseText, "0.(.*)v`n`n  2.", version)

if (version1 != ClientVersion)	{
	MsgBox, 262212, Update released, 	Version of your client - 0.%ClientVersion%`nLatest version - 0.%version1%`n`nWant to download the new version now?`n`n"YES" - Open Browser Page`n"NO" - Open current version (0.%ClientVersion%)
	IfMsgBox Yes	
	{
		run, https://github.com/MirchikAhtung/soundmixer/blob/master/SoundMixer_R0.%version1%.exe
		ExitApp
	}
}
;~  AutoUpdate

;~ GUI
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
Gui, Show,w510 h710, SoundMixer Reborn | @bass_devware | Release [0.%ClientVersion%]
;~ GUI
;~ Tray
Menu, tray, NoStandard
	Menu, tray, add, @bass_devware, group 
		Menu, tray, add, Updates, updlist 
			Menu, tray, add, WebSite, WebSite
				Menu, tray, Disable, WebSite
			Menu, tray, add  
		Menu, tray, add, Restore SoundMixer, Restore 
	Menu, tray, add 
Menu, tray, add, Quit, Exit  
;~ Tray
return

AllProcess:
LV_Delete()
for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
	LV_Add("", process.ProcessID, process.Name, process.CommandLine)
LV_ModifyCol() 
RemoveRecurStrings()
return

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

AddBtn:
GuiControl,enable,List
GuiControl, disable, load
GuiControl,disable, unload
add := true
return

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

AWHotkey:
if (editvar != "save")	{
	GuiControl,enable,save
	GuiControl,enable,slider
}
return

Selected:
GuiControl,enable,AWHotkey
tosave := true
return

AW:
GuiControl,enable,AWHotkey
tosave := "2"
return

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

DDL:
gui, submit, nohide
if (DDsL != "")	{
	GuiControl,enable,remove
	GuiControl,enable,edit2
}
return

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

refreshprofiles:
GuiControl, Enable, AddBtn
	GuiControl, Enable, DDsL
		GuiControl, Enable, AP
			GuiControl, Enable, OM
		GuiControl, Enable, unload
	GuiControl, Enable, load
ControlGet, refprofiles, List,, ComboBox1, A
Loop, parse, refprofiles, `n, `r 
{
	RegExMatch(A_LoopField, "(.*) \Q|\E (.*) \Q|\E (.*) \Q|\E (.*)", profile%A_Index%)
	If (profile%A_Index%1 = "" Or profile%A_Index%2 = "" Or profile%A_Index%3 = "" Or profile%A_Index%4 = "")	{
		MsgBox The profile file is damaged. [ №%A_Index% ]
		return
	}
	tempvar := profile%A_Index%2
	pid%tempvar% := profile%A_Index%3
	process%tempvar% := profile%A_Index%4
	%tempvar% := profile%A_Index%1
	hotkey, %tempvar%, hotkey, on, UseErrorLevel
}
return

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

load:
GuiControl, Disable, AddBtn
	GuiControl, Disable, DDsL
		GuiControl, Disable, AP
			GuiControl, Disable, OM
		GuiControl, Disable, unload
	GuiControl, Disable, load
FileSelectFile, SelectedFile, 3, , Open a config with profiles., Text Documents (*.ini)
if (SelectedFile != "")	{
	FileRead, Profiles, %SelectedFile%
Loop, parse, Profiles, `n, `r
{
RegExMatch(A_LoopField, "(.*) \Q|\E (.*) \Q|\E (.*) \Q|\E (.*)", load)
If (load1 = "" Or load2 = "" Or load3 = "" Or load4 = "")	{
	MsgBox, 16, Warning!, The profile file is damaged | Line: %A_LoopField%
}	else	{
	if (load3 != "ActiveWindow")	{
		Process, Exist, % load4
		if ErrorLevel != % load3
			load3 := ErrorLevel
	}
	line = %load1% | %load2% | %load3% | %load4%
	SendMessage, 0x143, 0, &line , , ahk_id %hDDL%  ;  CB_ADDSTRING 
	}
} 
TrayTip, Load your profiles., Profiles have been loaded., 3
}
goto refreshprofiles
return

unload:
GuiControl, Disable, AddBtn
	GuiControl, Disable, DDsL
		GuiControl, Disable, AP
			GuiControl, Disable, OM
				GuiControl, Disable, unload
				GuiControl, Disable, load
			TrayTip, Saving your profile., Select a folder to save your profile., 5
		sleep 2000
	ControlGet, freeprofiles, List,, ComboBox1, A
FileSelectFolder, OutputVar, , 3
if (OutputVar != "")	{
	FormatTime, time,,hh.mm.ss
	FileAppend,%freeprofiles%,%OutputVar%\profilesettings_%time%.ini
	FileGetSize, size, %OutputVar%\profilesettings_%time%.ini
	if (size != false)
		MsgBox, 64, Excellent!, Profile has been saved.
	else
		MsgBox, 16, Warning!, Profile not saved.
}
else
	MsgBox, 48, Warning!, You have not selected a folder.
GuiControl, Enable, AddBtn
	GuiControl, Enable, DDsL
		GuiControl, Enable, AP
		GuiControl, Enable, OM
	GuiControl, Enable, unload
GuiControl, Enable, load
return

Restore:
Gui,Show
return

website:
updlist:
bassdevelopersoftware:
oWhr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	oWhr.Open("GET", "https://raw.githubusercontent.com/MirchikAhtung/soundmixer/master/update.log.txt", false)
		oWhr.Send()
	html := oWhr.ResponseText
MsgBox % html
return

group:
run, https://vk.com/bass_devware
return

Exit:
ExitApp
return

RemoveRecurStrings()	{
	global availablekeys
	availablekeys := "PIDPID"
	Loop % LV_GetCount()
	{
		count += 1
		LV_GetText(RetrievedText, count, 1)
		If InStr(availablekeys, "PID" RetrievedText "PID")	{
			LV_Delete(count)
			count -= 1
		}
		availablekeys := availablekeys . RetrievedText . "PID"
	}
}

SetAppVolume(pid, MasterVolume)	{
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+4*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 1, "UPtrP", IMMDevice, "UInt")
    ObjRelease(IMMDeviceEnumerator)

    VarSetCapacity(GUID, 16)
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}", "UPtr", &GUID)
    DllCall(NumGet(NumGet(IMMDevice+0)+3*A_PtrSize), "UPtr", IMMDevice, "UPtr", &GUID, "UInt", 23, "UPtr", 0, "UPtrP", IAudioSessionManager2, "UInt")
    ObjRelease(IMMDevice)

    DllCall(NumGet(NumGet(IAudioSessionManager2+0)+5*A_PtrSize), "UPtr", IAudioSessionManager2, "UPtrP", IAudioSessionEnumerator, "UInt")
    ObjRelease(IAudioSessionManager2)

    DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+3*A_PtrSize), "UPtr", IAudioSessionEnumerator, "UIntP", SessionCount, "UInt")
    Loop % SessionCount
    {
        DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+4*A_PtrSize), "UPtr", IAudioSessionEnumerator, "Int", A_Index-1, "UPtrP", IAudioSessionControl, "UInt")
        IAudioSessionControl2 := ComObjQuery(IAudioSessionControl, "{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}")
        ObjRelease(IAudioSessionControl)

        DllCall(NumGet(NumGet(IAudioSessionControl2+0)+14*A_PtrSize), "UPtr", IAudioSessionControl2, "UIntP", ProcessId, "UInt")
        If (pid == ProcessId)
        {
            ISimpleAudioVolume := ComObjQuery(IAudioSessionControl2, "{87CE5498-68D6-44E5-9215-6DA47EF883D8}")
            DllCall(NumGet(NumGet(ISimpleAudioVolume+0)+3*A_PtrSize), "UPtr", ISimpleAudioVolume, "Float", MasterVolume/100.0, "UPtr", 0, "UInt")
            ObjRelease(ISimpleAudioVolume)
        }
        ObjRelease(IAudioSessionControl2)
    }
    ObjRelease(IAudioSessionEnumerator)
}

GetModuleFileNameEx(p_pid)	{ 
	if A_OSVersion in WIN_95,WIN_98,WIN_ME 
	{ 
		MsgBox, This Windows version (%A_OSVersion%) is not supported. 
		return 
	} 
	h_process := DllCall( "OpenProcess", "uint", 0x10|0x400, "int", false, "uint", p_pid ) 
	if (ErrorLevel or h_process = 0) 
		return 
	name_size = 255 
	VarSetCapacity( name, name_size ) 
	result := DllCall( "psapi.dll\GetModuleFileNameEx" ( A_IsUnicode ? "W" : "A" ), "uint", h_process, "uint", 0, "str", name, "uint", name_size ) 
	DllCall( "CloseHandle", h_process ) 
	return, name 
}

GetActive_Media()	{ 
	IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}") 
	DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+4*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 1, "UPtrP", IMMDevice, "UInt") 
	ObjRelease(IMMDeviceEnumerator) 

	VarSetCapacity(GUID, 16) 
	DllCall("Ole32.dll\CLSIDFromString", "Str", "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}", "UPtr", &GUID) 
	DllCall(NumGet(NumGet(IMMDevice+0)+3*A_PtrSize), "UPtr", IMMDevice, "UPtr", &GUID, "UInt", 23, "UPtr", 0, "UPtrP", IAudioSessionManager2, "UInt") 
	ObjRelease(IMMDevice) 

	DllCall(NumGet(NumGet(IAudioSessionManager2+0)+5*A_PtrSize), "UPtr", IAudioSessionManager2, "UPtrP", IAudioSessionEnumerator, "UInt") 
	ObjRelease(IAudioSessionManager2) 

	DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+3*A_PtrSize), "UPtr", IAudioSessionEnumerator, "UIntP", SessionCount, "UInt") 
	Loop % SessionCount 
	{ 
	DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+4*A_PtrSize), "UPtr", IAudioSessionEnumerator, "Int", A_Index-1, "UPtrP", IAudioSessionControl, "UInt") 
	IAudioSessionControl2 := ComObjQuery(IAudioSessionControl, "{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}") 
	ObjRelease(IAudioSessionControl) 

	DllCall(NumGet(NumGet(IAudioSessionControl2+0)+14*A_PtrSize), "UPtr", IAudioSessionControl2, "UIntP", ProcessId, "UInt") 
	if (ProcessId) 
		PID .= ProcessId "," 
	ObjRelease(IAudioSessionControl2) 
	} 
	ObjRelease(IAudioSessionEnumerator) 
	StringTrimRight, PID, PID, 1 
	return PID 
}

end::reload
!end::ExitApp
