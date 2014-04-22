Option Explicit

Const NAME_OF_FILESYSTEM = "Scripting.FileSystemObject"

Dim FileShell
Set FileShell = WScript.CreateObject(NAME_OF_FILESYSTEM)

Const FOR_READING_INCLUDE = 1
