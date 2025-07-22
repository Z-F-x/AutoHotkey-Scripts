#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; Ctrl + Enter hotkey
^Enter::
; Try to get path from Explorer or Desktop
path := GetActiveExplorerPath()
if (path = "")
{
    MsgBox, 48, Error, No valid path found.`n`nPlease focus a File Explorer window or Desktop and try again.`n`nSupported windows:`n• File Explorer (any folder)`n• Desktop
    return
}

; Open Windows Terminal at the detected path
OpenTerminalAtPath(path)
return

; Windows + Enter hotkey (alternative)
#Enter::
; Try to get path from Explorer or Desktop
path := GetActiveExplorerPath()
if (path = "")
{
    MsgBox, 48, Error, No valid path found.`n`nPlease focus a File Explorer window or Desktop and try again.`n`nSupported windows:`n• File Explorer (any folder)`n• Desktop
    return
}

; Open Windows Terminal at the detected path
OpenTerminalAtPath(path)
return

; --- Functions ---

; Returns the path of the currently active File Explorer window or Desktop
GetActiveExplorerPath() {
    ; First, try to get the path from the currently active File Explorer window
    WinGet, hWnd, ID, A

    ; Check if the active window is File Explorer
    WinGetClass, WinClass, ahk_id %hWnd%
    if (WinClass = "CabinetWClass" || WinClass = "ExploreWClass") {
        ; Try to get path from the active Explorer window
        for window in ComObjCreate("Shell.Application").Windows {
            try {
                if (window.HWND = hWnd) {
                    return window.Document.Folder.Self.Path
                }
            } catch e {
                continue
            }
        }
    }

    ; Alternative method: Check all Explorer windows and find the active one
    for window in ComObjCreate("Shell.Application").Windows {
        try {
            ; Check if this window is the active window
            if (window.HWND = hWnd) {
                return window.Document.Folder.Self.Path
            }
        } catch e {
            continue
        }
    }

    ; Check if desktop is active
    if WinActive("ahk_class Progman") || WinActive("ahk_class WorkerW") {
        return A_Desktop
    }

    ; Final fallback: try to get any focused Explorer window
    for window in ComObjCreate("Shell.Application").Windows {
        try {
            if (window.Document.Focused) {
                return window.Document.Folder.Self.Path
            }
        } catch e {
            continue
        }
    }

    return ""
}

; Open Windows Terminal at the specified path
OpenTerminalAtPath(targetPath) {
    ; Escape any special characters in the path for command line usage
    escapedPath := StrReplace(targetPath, """", """""")

    ; Try multiple methods to open Windows Terminal

    ; Method 1: Try using wt.exe (Windows Terminal) directly
    try {
        Run, wt.exe -d "%escapedPath%", , Hide
        return
    } catch e {
        ; If wt.exe fails, try alternative methods
    }

    ; Method 2: Try using the Windows Terminal executable from Program Files
    try {
        Run, "%A_ProgramFiles%\WindowsApps\Microsoft.WindowsTerminal_*\wt.exe" -d "%escapedPath%", , Hide
        return
    } catch e {
        ; Continue to next method
    }

    ; Method 3: Try using the Microsoft Store app protocol
    try {
        Run, ms-windows-store://pdp/?productid=9N0DX20HK701, , Hide
        Sleep, 2000
        Run, wt.exe -d "%escapedPath%", , Hide
        return
    } catch e {
        ; Continue to next method
    }

    ; Method 4: Fallback to Command Prompt if Windows Terminal is not available
    try {
        Run, cmd.exe /k "cd /d ""%escapedPath%""", , Show
        return
    } catch e {
        ; Final fallback
    }

    ; Method 5: Ultimate fallback - PowerShell
    try {
        Run, powershell.exe -NoExit -Command "Set-Location '%escapedPath%'", , Show
    } catch e {
        MsgBox, 48, Error, Failed to open terminal at path: %targetPath%`n`nPlease ensure Windows Terminal or Command Prompt is installed.
    }
}
