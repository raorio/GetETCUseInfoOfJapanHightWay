'===============================================================================
' main
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
Const NAME_OF_FILESYSTEM = "Scripting.FileSystemObject"

Dim FileShell
Set FileShell = WScript.CreateObject(NAME_OF_FILESYSTEM)

Const FOR_READING_INCLUDE = 1


'*******************************************************************************
' read vbs file
'   @param FileName [in] read vbs file name
'   @retval nothing
'*******************************************************************************
Function ReadVBSFile(ByVal FileName)
  ReadVBSFile = FileShell.OpenTextFile(FileName, FOR_READING_INCLUDE, False).ReadAll()
End Function

Execute ReadVBSFile("IncludeConfig.vbs")
Execute ReadVBSFile("IncludeCommonConfig.vbs")
Execute ReadVBSFile("IncludeAPI.vbs")
Execute ReadVBSFile("IncludeCommonAPI.vbs")

' log file check
funcDummy = logFileCheck(LOG_FOLDER, logFilePath)

logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "start init program")

' set vbs timeout
If VBS_TIMEOUT > 0 Then
  WScript.timeout = vbsTimeoutValue
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "set vbs timeout: " & VBS_TIMEOUT)
End If

funcDummy = InitETCUseInfoOfJapanHightWay()

logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "end init program")
