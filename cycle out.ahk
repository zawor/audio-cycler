; this work is based on this thread https://www.autohotkey.com/boards/viewtopic.php?t=49980 to some extent

#Include, %A_ScriptDir%\libs\tooltip.ahk

DetectHiddenWindows, On

; Set up tray tray menu

Menu, Tray, NoStandard
Menu, Tray, Add, &Next Output, CycleAutioOut
Menu, Tray, Default, &Next Output
Menu, Tray, Add, Exit, Exit
Menu, Tray, Click, 1

IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
; IMMDeviceEnumerator::GetDefaultAudioEndpoint
DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+4*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0, "UPtrP", IMMDevice, "UInt")
activeDevice := _getDeviceProps(IMMDevice)
ObjRelease(IMMDeviceEnumerator)

_ChangeTray(activeDevice.name, activeDevice.fullName)

Exit() {
    ExitApp
}

_ChangeTray(ico:="default", name:="Audio out cycler") {
    Menu, Tray, Tip, %name%
    if (FileExist("./icons/" . ico ".ico")) {
        Menu, Tray, Icon, icons/%ico%.ico
    }
    else {
        Menu, Tray, Icon, icons/default.ico
    }
}

ToolTipsPositionX := "LEFT"
ToolTipsPositionY := "BOTTOM"

_ShowTooltip(message:="") {
    params := {}
    params.message := message
    params.lifespan := 500
    params.position := { x: "LEFT", y: "BOTTOM"}
    params.fontSize := 12
    params.fontWeight := 400
    Toast(params)
}

_getDeviceId(IMMDevice) {
    DllCall(NumGet(NumGet(IMMDevice+0) + 5 * A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
    DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

    return DeviceID
}

_getDeviceProps(IMMDevice) {
    ; IMMDevice::OpenPropertyStore
    ; 0x0 = STGM_READ
    DllCall(NumGet(NumGet(IMMDevice+0) + 4 * A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")

    ; IPropertyStore::GetValue
    VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
    VarSetCapacity(PROPERTYKEY, 20)

    DllCall("Ole32.dll\CLSIDFromString", "Str", "{a45c254e-df1c-4efd-8020-67d146a850e0}", "UPtr", &PROPERTYKEY)
    NumPut(14, &PROPERTYKEY + 16, "UInt") ; Firendly Name
    DllCall(NumGet(NumGet(IPropertyStore + 0) + 5 * A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
    DeviceFullName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal

    NumPut(2, &PROPERTYKEY + 16, "UInt") ; Name
    DllCall(NumGet(NumGet(IPropertyStore + 0) + 5 * A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
    DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
    DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal

    ObjRelease(IPropertyStore)

    Props := {}

    Props.name := DeviceName
    Props.fullName := DeviceFullName

    return Props
}

CycleAutioOut() {
    DeviceNames := {}

    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    ; IMMDeviceEnumerator::GetDefaultAudioEndpoint
    DllCall(NumGet(NumGet(IMMDeviceEnumerator + 0) + 4 * A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0, "UPtrP", IMMDevice, "UInt")
    ActiveDeviceID := _getDeviceId(IMMDevice)

    ; IMMDeviceEnumerator::EnumAudioEndpoints
    ; eRender = 0, eCapture, eAll
    ; 0x1 = DEVICE_STATE_ACTIVE
    DllCall(NumGet(NumGet(IMMDeviceEnumerator + 0) + 3 * A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
    ObjRelease(IMMDeviceEnumerator)

    ; IMMDeviceCollection::GetCount
    DllCall(NumGet(NumGet(IMMDeviceCollection + 0) + 3 * A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
    Loop % (Count)
    {
        ; IMMDeviceCollection::Item
        DllCall(NumGet(NumGet(IMMDeviceCollection + 0) + 4 * A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")
        DeviceID := _getDeviceId(IMMDevice)
        DeviceNames[DeviceID] := _getDeviceProps(IMMDevice)
        ObjRelease(IMMDevice)
    }
    ObjRelease(IMMDeviceCollection)

    DeviceIds := {}
    NextDev := 0
    For DeviceID, DeviceName in DeviceNames {
        ObjRawSet(DeviceIds, A_Index - 1, DeviceID)
        if (DeviceID = ActiveDeviceID) {
            NextDev := Mod(A_Index, Count)
        }
    }

    ;IPolicyConfig::SetDefaultEndpoint
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
    DllCall(NumGet(NumGet(IPolicyConfig + 0) + 13 * A_PtrSize), "UPtr", IPolicyConfig, "Str", DeviceIds[NextDev], "UInt", 0, "UInt")
    ObjRelease(IPolicyConfig)

    _ChangeTray(DeviceNames[DeviceIds[NextDev]]["name"], DeviceNames[DeviceIds[NextDev]]["fullName"])
    _ShowTooltip(DeviceNames[DeviceIds[NextDev]]["fullName"])
}

#NumpadEnter:: CycleAutioOut()

^Volume_Up::
    Sleep, 200
    Send {Media_Next}
    Return
^Volume_Down::
    Sleep, 200
    Send {Media_Prev}
    Return