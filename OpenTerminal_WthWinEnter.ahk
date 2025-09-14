
#Enter::
    WinKeyOpenTerminalInCurrentDirectory()


WinKeyOpenTerminalInCurrentDirectory() {
    currentDir := GetCurrentDirectory()
    if (currentDir = "") {
        ; Fallback to Desktop if no directory found
        currentDir := A_Desktop
    }
    ; Temporarily disable hotkey to prevent recursion
    Hotkey, #Enter, Off
    ; Use wt.exe directly (should be in PATH)
    Run, wt.exe -d "%currentDir%"
    Sleep, 1000
    ShowTerminal()
    Sleep, 500
    Hotkey, #Enter, On
}



ShowTerminal() {
    WinShow ahk_class CASCADIA_HOSTING_WINDOW_CLASS
    WinActivate ahk_class CASCADIA_HOSTING_WINDOW_CLASS
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
    if (winClass = "Chrome_WidgetWin_1") {
        WinGetTitle, title, A
        if (InStr(title, "Visual Studio Code")) {
            ; Extract path from VS Code title (usually shows workspace folder)
            if (RegExMatch(title, "([A-Za-z]:\\[^\\/:*?\""<>|]+(?:\\[^\\/:*?\""<>|]+)*)", match)) {
                ; Ensure it's a directory
                if (FileExist(match1) && InStr(FileExist(match1), "D")) {
                    return match1
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
    if (addressText != "" && FileExist(addressText) && InStr(FileExist(addressText), "D")) {
        return addressText
    }

    ; Try different control names for address bars
    ControlGetText, addressText, ToolbarWindow323, ahk_id %activeHwnd%
    if (InStr(addressText, ":\") && FileExist(SubStr(addressText, InStr(addressText, ":\")-1)) && InStr(FileExist(SubStr(addressText, InStr(addressText, ":\")-1)), "D")) {
        return SubStr(addressText, InStr(addressText, ":\")-1)
    }

    return ""
}


