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
  
  If logLevel <= LOG_TARGET_LEVEL Then
    logDatetime = strLogDateTime
    
    logContext = logDatetime & DEFINE_SPACE & logLevelStrings(logLevel) & DEFINE_SPACE & message & DefineCrLf
    
    funcDummy = appendFile(logFilePath, logContext)
  End If
  
  Set logContext = Nothing
  Set timeStr = Nothing
  
  'logOut = ""
End Function

'*******************************************************************************
' logOutFatal
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutFatal(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_FATAL, message)
  
  'logOutFatal = ""
End Function

'*******************************************************************************
' logOutError
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutError(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_ERROR, message)
  
  'logOutError = ""
End Function

'*******************************************************************************
' logOutWarn
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutWarn(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_WARN, message)
  
  'logOutWarn = ""
End Function

'*******************************************************************************
' logOutInfo
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutInfo(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_INFO, message)
  
  'logOutInfo = ""
End Function

'*******************************************************************************
' logOutDebug
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutDebug(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_DEBUG, message)
  
  'logOutDebug = ""
End Function

'*******************************************************************************
' logOutDetailDebug
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function logOutDetailDebug(message)
  logReturnValueDummy = logOut(LOG_LEVEL_NUMBER_DETAIL_DEBUG, message)
  
  'logOutDetailDebug = ""
End Function

'*******************************************************************************
' logFileCheck
'   @param nothing
'   @retval nothing
'*******************************************************************************
Function logFileCheck()
  If IsExistFolder(LOG_FOLDER) = true Then
    If IsExistFile(logFilePath) = false Then
      CreateFile(logFilePath)
    End If
  Else
    CreateFolder(LOG_FOLDER)
    CreateFile(logFilePath)
  End If
    
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
  
  Set objFileSys = Nothing
  
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
  
  result = objFileSys.FolderExists(folderPath)
  
  Set objFileSys = Nothing
  
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
  
  Set objFileSys = Nothing
  
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
  
  result = objFileSys.FileExists(filePath)
  
  Set objFileSys = Nothing
  
  IsExistFile = result
End Function

'*******************************************************************************
' AppendFile
'   @param filePath [in] file path
'   @param context [in] context
'   @retval nothing
'*******************************************************************************
Function AppendFile(filePath, context)
  Dim objFileSys
  Dim resStream
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  Set resStream = objFileSys.OpenTextFile(filePath, ForAppending)
  
  resStream.Write(context)
  resStream.Close
  
  Set resStream = Nothing
  Set objFileSys = Nothing
  
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
    Set ReadFileAllContext = Nothing
  Else
    Dim resStream
    Dim resData
    
    Set resStream = objFileSys.OpenTextFile(filePath, ForReading)
    
    resData = resStream.ReadAll
    resStream.Close
    
    Set resStream = Nothing
    Set objFileSys = Nothing
    
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
  If resData <> Nothing Then
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
  
  Set objFileSys = Nothing
  
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

'-------------------------------------------------------------------------------
' object api
'-------------------------------------------------------------------------------
'*******************************************************************************
' DeleteSpace
'   @param isShowIEWindow [in] is show IE window(true or false)
'   @param url [in] url
'   @param waitTime [in] wait time
'   @retval IE object
'*******************************************************************************
Function CreateIEObject(isShowIEWindow, url, waitTime)
  Dim objIE
  
  Set objIE = WScript.CreateObject(NAME_OF_IE_APPLICATION)
  objIE.Visible = isShowIEWindow
  
  logReturnValueDummy = logOutDebug("web access start url: " & url)
  
  objIE.Navigate url
  
  If waitTime > 0 Then
    While objIE.ReadyState <> 4 Or objIE.Busy = True
      WScript.Sleep waitTime
    Wend
  End If
  
  logReturnValueDummy = logOutDebug("web access end url: " & url)
  
  CreateIEObject = objIE
End Function

