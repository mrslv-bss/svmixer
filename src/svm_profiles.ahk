/*
    
*/

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
	soundbeep, 100, 100
}
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