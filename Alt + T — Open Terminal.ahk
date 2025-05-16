#NoTrayIcon
#SingleInstance Force

;
; This script toggles the Windows Terminal window:
; - Shows it if hidden
; - Hides it if active
; - Opens a new one if not running
; Triggered by: Alt + T or Alt + Enter
;

!T::  ; Alt + T
!Enter::  ; Alt + Enter
    ToggleTerminal()

ToggleTerminal() {
    matcher := "ahk_class CASCADIA_HOSTING_WINDOW_CLASS"
    DetectHiddenWindows, On
    if WinExist(matcher) {

        if !WinActive(matcher) {
            ; Hide first to enable cross-workspace raise
            HideTerminal()
            ShowTerminal()
        } else if WinExist(matcher) {
            HideTerminal()
        }

    } else {
        OpenNewTerminal()
    }
}

OpenNewTerminal() {
    Run C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps\wt.exe
    Sleep, 1000
    ShowTerminal()
}

ShowTerminal() {
    WinShow ahk_class CASCADIA_HOSTING_WINDOW_CLASS
    WinActivate ahk_class CASCADIA_HOSTING_WINDOW_CLASS
}

HideTerminal() {
    WinHide ahk_class CASCADIA_HOSTING_WINDOW_CLASS
}
