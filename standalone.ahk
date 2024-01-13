#SingleInstance Force
installKeybdHook
Persistent
#x::ExitApp ; Win+X to terminate script
^!r::Reload ; Ctrl+Alt+R to reload the script

#Include calculator2.ahk

#+D::bCalculator()
bCalculator()