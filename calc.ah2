#SingleInstance Force
installKeybdHook
Persistent
#x::ExitApp ; Win+X to terminate script
^!r::Reload ; Ctrl+Alt+R to reload the script

#Include "C:\Users\Максим\Documents\AutoHotkey\Lib\CBR.ah2"

update_currencies("EUR", "CHF", "USD")
update_currencies(params*) {
    for key, value in params
        CBR(value, &value)

}

; g := Gui()
; g.AddText(, "EUR")
; g.AddText(, )