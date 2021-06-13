/*

    RemoveRecurStrings()

    SetAppVolume(pid, MasterVolume)

    GetModuleFileNameEx(p_pid)

    GetActive_Media()

*/

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