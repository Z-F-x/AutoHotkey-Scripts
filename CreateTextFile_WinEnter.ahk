; -------------------------------------------------------------
;Create New Textfile Win + N — Desktop and Windows Explorer (With Copy Iteration)
; -------------------------------------------------------------

#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

#N::
; Try to get path from Explorer or Desktop
path := GetActiveExplorerPath()
if (path = "")
{
    MsgBox, 48, Error, No valid path found.`n`nPlease focus a File Explorer window or Desktop and try again.`n`nSupported windows:`n• File Explorer (any folder)`n• Desktop
    return
}

; Generate a new file name and create the file
newFileName := CreateUniqueFile(path, "New Text Document", "txt")
if (newFileName != "") {
    ; Optional: Refresh the current Explorer window to show the new file
    Send {F5}
}
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

; Create a unique file in the specified directory
CreateUniqueFile(directory, baseName, extension) {
    counter := 1
    Loop {
        fileName := (counter = 1)
            ? baseName . "." . extension
            : baseName . " (" . counter . ")." . extension
        fullPath := directory . "\" . fileName

        ; Check if file already exists
        if !FileExist(fullPath) {
            ; Try to create the file
            FileAppend, , %fullPath%

            ; Verify the file was created successfully
            if FileExist(fullPath) {
                ; Optional: Show success message with path
                ; ToolTip, Created: %fileName% in %directory%, , , 1
                ; SetTimer, RemoveToolTip, 2000
                return fileName
            } else {
                MsgBox, 48, Error, Failed to create file: %fullPath%
                return ""
            }
        }
        counter++

        ; Safety check to prevent infinite loop
        if (counter > 1000) {
            MsgBox, 48, Error, Too many files with similar names. Cannot create unique file.
            return ""
        }
    }
}

; Helper function to remove tooltip
RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
return
