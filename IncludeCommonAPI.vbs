'Option Explicit

'-------------------------------------------------------------------------------
Const NAME_OF_FILESYSTEM_IN_COMMON_API = "Scripting.FileSystemObject"

Dim FileShellInCommonAPI
Set FileShellInCommonAPI = WScript.CreateObject(NAME_OF_FILESYSTEM_IN_COMMON_API)

Const FOR_READING_INCLUDE_IN_COMMON_API = 1

'*******************************************************************************
' read vbs file in common API
'   @param FileName [in] read vbs file name
'   @retval nothing
'*******************************************************************************
Function ReadVBSFileInCommonAPI(ByVal FileName)
  ReadVBSFileInCommonAPI = FileShellInCommonAPI.OpenTextFile(FileName, FOR_READING_INCLUDE_IN_COMMON_API, False).ReadAll()
End Function

'Execute ReadVBSFileInCommonAPI("IncludeCommonConfig.vbs")


'===============================================================================
' api
'===============================================================================
'-------------------------------------------------------------------------------
' common api
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' date api
'-------------------------------------------------------------------------------
'*******************************************************************************
' getDateTime
'   @param nothing
'   @retval date time
'*******************************************************************************
Function getDateTime()
  Dim result
  
  Dim yearValue
  Dim monthValue
  Dim dayValue
  Dim hourValue
  Dim minuteValue
  Dim secondValue
  yearValue = Year(Now)
  monthValue = Month(Now)
  dayValue = Day(Now)
  hourValue = Hour(Now)
  minuteValue = Minute(Now)
  secondValue = Second(Now)
  ' yyyy/mm/dd hh:mm:ss date time formate
  ' TODO
  'result = PaddingPrefixString(yearValue, "0", 4) & DEFINE_DELIM_DATE & PaddingPrefixString(monthValue, "0", 2) & DEFINE_DELIM_DATE & Day(Now) & DEFINE_DELIM_DATE_TIME & Hour(Now) & DEFINE_DELIM_TIME & Minute(Now) & DEFINE_DELIM_TIME & Second(Now)
  result = Year(Now) & DEFINE_DELIM_DATE & Month(Now) & DEFINE_DELIM_DATE & Day(Now) & DEFINE_DELIM_DATE_TIME & Hour(Now) & DEFINE_DELIM_TIME & Minute(Now) & DEFINE_DELIM_TIME & Second(Now)
  
  getDateTime = result
End Function

'*******************************************************************************
' getDateTimeAtISOFormat
'   @param nothing
'   @retval date time
'*******************************************************************************
Function getDateTimeAtISOFormat()
  Dim result
  
  ' yyyymmddThhmmss date time formate
  Dim strDateTimeISO
  ' TODO
  result = Year(Now) & Month(Now) & Day(Now) & DEFINE_DELIM_ISO_DATE_TIME & Hour(Now) & Minute(Now) & Second(Now)
  
  getDateTimeAtISOFormat = result
End Function


'-------------------------------------------------------------------------------
' log api
'-------------------------------------------------------------------------------
'*******************************************************************************
' LogOut
'   @param targetLogLevel [in] target log level
'   @param logLevel [in] log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOut(targetLogLevel, logLevel, message)
  Dim timeStr
  Dim logContext
  
  If logLevel <= targetLogLevel Then
    Dim logDateTime
    logDateTime = getDateTime()
    
    logContext = logDateTime & DEFINE_SPACE & logLevelStrings(logLevel) & DEFINE_SPACE & message & DefineCrLf
    
    funcDummy = appendFile(logFilePath, logContext)
  End If
  
  Set logContext = Nothing
  Set timeStr = Nothing
End Function

'*******************************************************************************
' logOutFatal
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutFatal(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_FATAL, message)
End Function

'*******************************************************************************
' logOutError
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutError(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_ERROR, message)
End Function

'*******************************************************************************
' logOutWarn
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutWarn(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_WARN, message)
End Function

'*******************************************************************************
' logOutInfo
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutInfo(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_INFO, message)
End Function

'*******************************************************************************
' logOutDebug
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutDebug(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_DEBUG, message)
End Function

'*******************************************************************************
' logOutDetailDebug
'   @param targetLogLevel [in] target log level
'   @param message [in] log message
'   @retval nothing
'*******************************************************************************
Function LogOutDetailDebug(targetLogLevel, message)
  logReturnValueDummy = LogOut(targetLogLevel, LOG_LEVEL_NUMBER_DETAIL_DEBUG, message)
End Function

'*******************************************************************************
' logFileCheck
'   @param logFolderPath [in] log folder path
'   @param logFilePath [in] log file path
'   @retval nothing
'*******************************************************************************
Function LogFileCheck(logFolderPath, logFilePath)
  If IsExistFolder(logFolderPath) = true Then
    If IsExistFile(logFilePath) = false Then
      CreateFile(logFilePath)
    End If
  Else
    CreateFolder(logFolderPath)
    CreateFile(logFilePath)
  End If
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
End Function

'*******************************************************************************
' IsExistFile
'   @param filePath [in] file path
'   @retval true/false true:exist, false:not exist
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
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
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
' OpenFileToRead
'   @param filePath [in] file path
'   @retval file object
'*******************************************************************************
Function OpenFileToRead(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
    Set OpenFileToRead = Nothing
  Else
    Set objFileSys = objFileSys.OpenTextFile(filePath, ForReading)
    
    Set OpenFileToRead = objFileSys
  End If
End Function

'*******************************************************************************
' OpenFileToWrite
'   @param filePath [in] file path
'   @retval file object
'*******************************************************************************
Function OpenFileToWrite(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  Set resStream = objFileSys.OpenTextFile(filePath, ForWriting)
End Function

'*******************************************************************************
' OpenFileToAppend
'   @param filePath [in] file path
'   @retval file object
'*******************************************************************************
Function OpenFileToAppend(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
    Set OpenFileToAppend = Nothing
  Else
    Set objFileSys = objFileSys.OpenTextFile(filePath, ForAppending)
    
    Set OpenFileToAppend = objFileSys
  End If
End Function

'*******************************************************************************
' WriteToObjectFile
'   @param objFileSys [in] object file system
'   @param context [in] context
'   @retval nothing
'*******************************************************************************
Function WriteToObjectFile(objFileSys, context)
  objFileSys.Write(context)
End Function

'*******************************************************************************
' ReadFromObjectFile
'   @param objFileSys [in] object file system
'   @retval line context
'*******************************************************************************
Function ReadFromObjectFile(objFileSys)
  Dim lineContext
  
  lineContext = objFileSys.ReadLine()
  
  ReadFromObjectFile = lineContext
End Function

'*******************************************************************************
' CloseObjectFile
'   @param objFileSys [in] object file system
'   @retval nothing
'*******************************************************************************
Function CloseObjectFile(objFileSys)
  objFileSys.Close()
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
End Function

'-------------------------------------------------------------------------------
' string api
'-------------------------------------------------------------------------------
'*******************************************************************************
' DeleteSpace
'   @param targetString [in] target string
'   @retval replaced string
'*******************************************************************************
Function DeleteSpace(targetString)
  Dim replaceString
  
  replaceString = DeleteSpace2MoreSpace(targetString)
  replaceString = Replace(replaceString, DEFINE_SPACE, DEFINE_BRANK)
  
  DeleteSpace = replaceString
End Function

'*******************************************************************************
' DeleteSpace2MoreSpace
'   @param targetString [in] target string
'   @retval replaced string
'*******************************************************************************
Function DeleteSpace2MoreSpace(targetString)
  Dim lengthString
  Dim replaceString
  
  lengthString = Len(targetString)
  replaceString = targetString
  
  Dim i
  For i = lengthString To 2 Step -1
    Dim strChars
    strChars = Space(i)
    replaceString = Replace(replaceString, strChars, DEFINE_SPACE)
  Next
  
  DeleteSpace2MoreSpace = replaceString
End Function

'*******************************************************************************
' PaddingPrefixString
'   @param targetString [in] target string
'   @param paddingChar [in] padding char
'   @param paddingSize [in] padding size
'   @retval padding string
'*******************************************************************************
Function PaddingPrefixString(targetString, paddingChar, paddingSize)
  Dim lengthString
  Dim paddingString
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingPrefixString start")
  
  paddingString = ""
  lengthString = Len(targetString)
  
  Dim i
  For i = paddingSize - 1 To lengthString Step -1
    paddingString = paddingString & paddingChar
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingPrefixString end")
  
  PaddingPrefixString = paddingString & targetString
End Function

'*******************************************************************************
' PaddingSuffixString
'   @param targetString [in] target string
'   @param paddingChar [in] padding char
'   @param paddingSize [in] padding size
'   @retval padding string
'*******************************************************************************
Function PaddingSuffixString(targetString, paddingChar, paddingSize)
  Dim lengthString
  Dim paddingString
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingSuffixString start")
  
  paddingString = ""
  lengthString = Len(targetString)
  
  Dim i
  For i = paddingSize - 1 To lengthString Step -1
    paddingString = paddingString & paddingChar
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingSuffixString end")
  
  PaddingSuffixString = targetString & paddingString
End Function


'-------------------------------------------------------------------------------
' file api
'-------------------------------------------------------------------------------
'TODO


'-------------------------------------------------------------------------------
' web api
'-------------------------------------------------------------------------------
'*******************************************************************************
' GetHttp
'   @param url [in] url
'   @param saveFilePath [in] save file path
'   @param porxyServer [in] proxy server(if brank, don't use proxy server)
'   @retval true/false true:success, false:failed
'*******************************************************************************
Function GetHttp(url, saveFilePath, proxyServer)
  Dim objHttp
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetHttp start")
  
  On Error Resume Next
    For Each objName In httpObjectList
      Set objHttp = CreateObject(objName)
      If Err.Number = 0 Then
        logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "http object: " & objName)
        Exit For
      End If
    Next
  On Error GoTo 0
  If IsNull(objXmlHttp) = True Then
    logReturnValueDummy = logOutError(LOG_TARGET_LEVEL, "create http object")
    GetHttp = False
  Else
    If IsExistFile(saveFilePath) = True Then
      DeleteFile(saveFilePath)
    End If
    
    If Len(proxyServer) <> 0 Then
      funcDummy = objHttp.SetProxy(2, proxyServer, "")
    End If
    
    funcDummy = objHttp.Open("GET", url, False)
    objHttp.Send()
    
    If objHttp.Status = 200 Then
      logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "success http get: " & url)
      
      Dim objStream
      Set objStream = CreateObject(ADODB_STREAM)
      objStream.Open()
      objStream.Type = 1
      objStream.Write(objHttp.responseBody)
      objStream.SaveToFile(saveFilePath)
      objStream.Close()
      
      GetHttp = True
    Else
      logReturnValueDummy = logOutError(LOG_TARGET_LEVEL, "failed http get: " & url)
      
      GetHttp = False
    End If
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetHttp end")
End Function


'-------------------------------------------------------------------------------
' object api
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' IE object api
'-------------------------------------------------------------------------------
'*******************************************************************************
' CreateIEObject
'   @param isShowIEWindow [in] is show IE window(true or false)
'   @param url [in] url
'   @param waitTime [in] wait time
'   @retval IE object
'*******************************************************************************
Function CreateIEObject(isShowIEWindow, url, waitTime)
  Dim objIE
  
  Set objIE = WScript.CreateObject(NAME_OF_IE_APPLICATION)
  objIE.Visible = isShowIEWindow
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "web access start url: " & url)
  
  objIE.Navigate url
  
  funcDummy = WaitIEObject(objIE, waitTime)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "web access end url: " & url)
  
  Set CreateIEObject = objIE
End Function

'*******************************************************************************
' WaitIEObject
'   @param objIE [in] object IE
'   @param waitTime [in] wait time
'   @retval nothing
'*******************************************************************************
Function WaitIEObject(objIE, waitTime)
  If waitTime > 0 Then
    While objIE.ReadyState <> 4 Or objIE.Busy = True
      WScript.Sleep waitTime
    Wend
  End If
End Function

