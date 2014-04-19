'===============================================================================
' api
'===============================================================================
'-------------------------------------------------------------------------------
' common api
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' log api
'-------------------------------------------------------------------------------
'*******************************************************************************
' logOut
'   @param logLevel [in] log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOut(logLevel, message)
  Dim timeStr
  Dim logContext
  
  If logLevel <= logTargetLevel Then
    logDatetime = strLogDateTime
    
    logContext = logDatetime & DEFINE_SPACE & logLevelStrings(logLevel) & DEFINE_SPACE & message & DEFINE_CRLF
    
    appendFile(logFilePath, logContext)
  End If
  
  logContext = Nothing
  timeStr = Nothing
  
  'logOut = ""
End Function

'*******************************************************************************
' logFileCheck
'   @param logLevel [in] log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logFileCheck(logLevel, message)
  Dim timeStr
  Dim logContext
  
  If IsExistFolder(LOG_FOLDER) = true Then
    If IsExistFile(logFilePath) = false Then
      CreateFile(logFilePath)
    End If
  Else
    CreateFolder(LOG_FOLDER)
  End If
  
  logContext = Nothing
  timeStr = Nothing
  
  'logFileCheck = ""
End Function

'-------------------------------------------------------------------------------
' file system api
'-------------------------------------------------------------------------------
'---------------------------------------
' folder api
'---------------------------------------
'*******************************************************************************
' CreateFolder
'   @param folderPath [in] folder path
'   @retval nothing
'*******************************************************************************
Function CreateFolder(folderPath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  If objFileSys.FolderExists(folderPath) = false Then
    objFileSys.CreateFolder folderPath
  End If
  
  objFileSys = Nothing
  
  'CreateFolder = 
End Function

'*******************************************************************************
' IsExistFolder
'   @param folderPath [in] folder path
'   @retval true: exist
'   @retval false: don't exist
'*******************************************************************************
Function IsExistFolder(folderPath)
  Dim result
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  result objFileSys.FolderExists(folderPath)
  
  objFileSys = Nothing
  
  IsExistFolder = result
End Function

'---------------------------------------
' file api
'---------------------------------------
'*******************************************************************************
' CreateFile
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function CreateFile(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  If objFileSys.FileExists(filePath) = true Then
    objFileSys.DeleteFile filePath
  End If
  objFileSys.CreateTextFile filePath
  
  objFileSys = Nothing
  
  'CreateFolder = 
End Function

'*******************************************************************************
' IsExistFile
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function IsExistFile(filePath)
  Dim result
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  result objFileSys.FileExists(filePath)
  
  objFileSys = Nothing
  
  IsExistFile = result
End Function

'*******************************************************************************
' AppendFile
'   @param filePath [in] file path
'   @param context [in] context
'   @retval nothing
'*******************************************************************************
Function AppendFile(filePath)
  Dim objFileSys
  Dim resStream
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  Set resStream = objFSO.OpenTextFile(filePath, ForAppending)
  
  resStream.Write(context)
  resStream.Close
  
  resStream = Nothing
  objFileSys = Nothing
  
  'AppendFile = 
End Function

'*******************************************************************************
' ReadFileAllContext
'   @param filePath [in] file path
'   @retval file all context
'*******************************************************************************
Function ReadFileAllContext(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = logOut(logLevelError, "file don't exist: " & filePath)
    ReadFileAllContext = Nothing
  Else
    Dim resStream
    Dim resData
    
    Set resStream = objFSO.OpenTextFile(filePath, ForReading)
    
    resData = resStream.ReadAll
    resStream.Close
    
    resStream = Nothing
    objFileSys = Nothing
    
    ReadFileAllContext = resData
  End If
End Function

'*******************************************************************************
' ReadFileAllContext
'   @param filePath [in] file path
'   @retval file all context
'*******************************************************************************
Function ReadFileAllContextAfterDelete(filePath)
  Dim resData
  
  resData = ReadFileAllContext(filePath)
  If resData != Nothing Then
    DeleteFile(filePath)
  End If
  
  ReadFileAllContextAfterDelete = resData
End Function

'*******************************************************************************
' DeleteFile
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function DeleteFile(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  If objFileSys.FileExists(filePath) = true Then
    objFileSys.DeleteFile filePath
  End If
  
  objFileSys = Nothing
  
  'DeleteFile = 
End Function

'-------------------------------------------------------------------------------
' string api
'-------------------------------------------------------------------------------
'*******************************************************************************
' DeleteSpace
'   @param targetStrings [in] target strings
'   @retval replaced string
'*******************************************************************************
Function DeleteSpace(targetStrings)
  Dim replaceString
  
  replaceString = DeleteSpace2MoreSpace(targetStrings)
  replaceString = Replace(replaceString, DEFINE_SPACE, DEFINE_BRANK)
  
  DeleteSpace = replaceString
End Function

'*******************************************************************************
' DeleteSpace2MoreSpace
'   @param targetStrings [in] target strings
'   @retval replaced string
'*******************************************************************************
Function DeleteSpace2MoreSpace(targetStrings)
  Dim lengthStrings
  Dim replaceString
  
  lengthStrings = Len(targetStrings)
  replaceString = targetStrings
  
  For i = lengthStrings To 2 Step -1
    Dim strChars
    strChars = Space(i)
    replaceString Replace(replaceString, strChars, DEFINE_SPACE)
  Next
  
  DeleteSpace2MoreSpace = replaceString
End Function

