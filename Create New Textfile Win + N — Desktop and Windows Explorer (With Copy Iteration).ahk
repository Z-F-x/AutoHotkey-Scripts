#N::

; Check if desktop is active
if WinActive("ahk_class Progman") {
    destinationDir := A_Desktop
} else {
    destinationDir := A_WorkingDir
}

; Generate a new file name and create the file
newFileName := CreateUniqueFile(destinationDir, "New Text Document", "txt")

return

; Function to create a unique file in the specified directory
CreateUniqueFile(directory, baseName, extension) {
    ; Initialize counter
    counter := 1

    ; Find a unique file name
    while FileExist(directory . "\" . baseName . " (" . counter . ")." . extension) {
        counter++
    }

    ; Create the unique file name
    uniqueFileName := baseName . " (" . counter . ")." . extension
    uniqueFilePath := directory . "\" . uniqueFileName

    ; Optionally, create the file
    FileAppend, `n, %uniqueFilePath%

    ; Return the unique file name
    return uniqueFileName
}
