#NoTrayIcon
#SingleInstance Force

; Ctrl + Alt + NumpadEnter: Open Windows Terminal in current directory (or Desktop if none found)
^!NumpadEnter::
    currentDir := GetCurrentDirectory()
    if (currentDir = "") {
        currentDir := A_Desktop
    }
    Run C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps\wt.exe -d "%currentDir%"
    return

GetCurrentDirectory() {
    ; Try to get current directory from File Explorer
    currentDir := GetExplorerPath()
    if (currentDir != "") {
        return currentDir
    }
    return ""
}

GetExplorerPath() {
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
