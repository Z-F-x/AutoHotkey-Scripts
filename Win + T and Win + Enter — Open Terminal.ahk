#NoTrayIcon
#SingleInstance Force

;
; Win + Enter: Opens Windows Terminal in default location or toggles existing terminal
; Alt + Enter: Opens Windows Terminal in current directory (or Desktop if none available)
;
; Supports multiple workspaces or virtual desktops.
;

#Enter::
    ToggleTerminal()

!Enter::
    OpenTerminalInCurrentDirectory()

ToggleTerminal() {
    matcher := "ahk_class CASCADIA_HOSTING_WINDOW_CLASS"
    DetectHiddenWindows, On
    if WinExist(matcher) {

        if !WinActive(matcher) {
            ; Hide it first to alow raising it later on a different workspace
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

OpenTerminalInCurrentDirectory() {
    currentDir := GetCurrentDirectory()
    if (currentDir = "") {
        ; Fallback to Desktop if no current directory found
        currentDir := A_Desktop
    }
    Run C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps\wt.exe -d "%currentDir%"
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

GetCurrentDirectory() {
    ; Try to get current directory from File Explorer
    currentDir := GetExplorerPath()
    if (currentDir != "") {
        return currentDir
    }

    ; Try to get current directory from VS Code
    currentDir := GetVSCodePath()
    if (currentDir != "") {
        return currentDir
    }

    ; Try to get current directory from other applications with address bar
    currentDir := GetAddressBarPath()
    if (currentDir != "") {
        return currentDir
    }

    return ""
}

GetExplorerPath() {
    ; Get path from Windows Explorer
    WinGetClass, winClass, A
    if (winClass = "CabinetWClass" || winClass = "ExploreWClass") {
        for window in ComObjCreate("Shell.Application").Windows {
            try {
                if (window.hwnd = WinExist("A")) {
                    return window.Document.Folder.Self.Path
                }
            }
        }
    }
    return ""
}

GetVSCodePath() {
    ; Get workspace path from VS Code window title
    WinGetClass, winClass, A
    if (InStr(winClass, "Chrome")) {
        WinGetTitle, title, A
        if (InStr(title, "Visual Studio Code")) {
            ; Extract path from VS Code title (usually shows workspace folder)
            if (RegExMatch(title, "([A-Za-z]:\\[^\\/:*?""<>|]+(?:\\[^\\/:*?""<>|]+)*)", match)) {
                if (FileExist(match)) {
                    return match
                }
            }
        }
    }
    return ""
}

GetAddressBarPath() {
    ; Try to get path from address bar of various applications
    WinGet, activeHwnd, ID, A
    ControlGetText, addressText, Edit1, ahk_id %activeHwnd%
    if (addressText != "" && FileExist(addressText)) {
        return addressText
    }

    ; Try different control names for address bars
    ControlGetText, addressText, ToolbarWindow323, ahk_id %activeHwnd%
    if (InStr(addressText, ":\") && FileExist(SubStr(addressText, InStr(addressText, ":\")-1))) {
        return SubStr(addressText, InStr(addressText, ":\")-1)
    }

    return ""
}
