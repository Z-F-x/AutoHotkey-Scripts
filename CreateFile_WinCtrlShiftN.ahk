#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; Win + Ctrl + Shift + N hotkey
#^+N::
; Try to get path from Explorer or Desktop
path := GetActiveExplorerPath()
if (path = "")
{
    MsgBox, 48, Error, No valid path found.`n`nPlease focus a File Explorer window or Desktop and try again.`n`nSupported windows:`n• File Explorer (any folder)`n• Desktop
    return
}

; Create a new file without extension and immediately rename it
newFileName := CreateNewFileAndRename(path)
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

; Create a new file without extension and immediately put it in rename mode
CreateNewFileAndRename(directory) {
    baseName := "New File"
    counter := 1

    Loop {
        fileName := (counter = 1) ? baseName : baseName . " (" . counter . ")"
        fullPath := directory . "\" . fileName

        ; Check if file already exists
        if !FileExist(fullPath) {
            ; Try to create the file
            FileAppend, , %fullPath%

            ; Verify the file was created successfully
            if FileExist(fullPath) {
                ; Refresh the Explorer window to show the new file
                Send {F5}

                ; Wait a moment for the refresh to complete
                Sleep, 200

                ; Select the new file and put it in rename mode
                SelectAndRenameFile(fileName)

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

; Select the newly created file and put it in rename mode
SelectAndRenameFile(fileName) {
    ; Wait a bit more to ensure the file appears in Explorer
    Sleep, 300

    ; Try to select the file by typing its name (works in most Explorer views)
    Send, %fileName%
    Sleep, 100

    ; Press F2 to enter rename mode
    Send, {F2}

    ; Alternative method if the above doesn't work:
    ; We could also try Ctrl+A to select all, then type the filename to select it
    ; But the above method should work in most cases
}
