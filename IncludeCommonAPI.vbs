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
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  
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
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  
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
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  
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
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  
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
  Dim objFile
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  Set objFile = objFileSys.OpenTextFile(filePath, ForAppending)
  
  objFile.Write(context)
  objFile.Close
  
  Set objFile = Nothing
  Set objFileSys = Nothing
End Function

'*******************************************************************************
' ReadFileAllContext
'   @param filePath [in] file path
'   @retval file all context
'*******************************************************************************
Function ReadFileAllContext(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
    Set ReadFileAllContext = Nothing
  Else
    Dim objFile
    Dim resData
    
    Set objFile = objFileSys.OpenTextFile(filePath, ForReading)
    
    resData = objFile.ReadAll
    objFile.Close
    
    Set objFile = Nothing
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
  Dim objFile
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
    Set OpenFileToRead = Nothing
  Else
    Set objFile = objFileSys.OpenTextFile(filePath, ForReading)
    
    Set OpenFileToRead = objFile
  End If
End Function

'*******************************************************************************
' OpenFileToWrite
'   @param filePath [in] file path
'   @retval file object
'*******************************************************************************
Function OpenFileToWrite(filePath)
  Dim objFileSys
  Dim objFile
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  Set objFile = objFileSys.OpenTextFile(filePath, ForWriting)
  
  Set OpenFileToWrite = objFile
  Set objFileSys = Nothing
End Function

'*******************************************************************************
' OpenFileToAppend
'   @param filePath [in] file path
'   @retval file object
'*******************************************************************************
Function OpenFileToAppend(filePath)
  Dim objFileSys
  Dim objFile
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  If objFileSys.FileExists(filePath) = false Then
    logReturnValueDummy = LogOut(logLevelError, "file don't exist: " & filePath)
    Set OpenFileToAppend = Nothing
  Else
    Set objFile = objFileSys.OpenTextFile(filePath, ForAppending)
    
    Set OpenFileToAppend = objFile
  End If
End Function

'*******************************************************************************
' WriteToObjectFile
'   @param objFile [in] object file
'   @param context [in] context
'   @retval nothing
'*******************************************************************************
Function WriteToObjectFile(objFile, context)
  objFile.Write(context)
End Function

'*******************************************************************************
' WriteLineToObjectFile
'   @param objFile [in] object file
'   @param context [in] context
'   @retval nothing
'*******************************************************************************
Function WriteLineToObjectFile(objFile, context)
  objFile.Write(context & DefineCrLf)
End Function

'*******************************************************************************
' ReadFromObjectFile
'   @param objFile [in] object file
'   @retval line context
'*******************************************************************************
Function ReadFromObjectFile(objFile)
  Dim lineContext
  
  lineContext = objFile.ReadLine()
  
  ReadFromObjectFile = lineContext
End Function

'*******************************************************************************
' CloseObjectFile
'   @param objFile [in] object file
'   @retval nothing
'*******************************************************************************
Function CloseObjectFile(objFile)
  objFile.Close()
End Function

'*******************************************************************************
' DeleteFile
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function DeleteFile(filePath)
  Dim objFileSys
  
  Set objFileSys = WScript.CreateObject(NAME_OF_SCRIPTING_FILESYSTEMOBJECT)
  
  If objFileSys.FileExists(filePath) = true Then
    objFileSys.DeleteFile filePath
  End If
  
  Set objFileSys = Nothing
End Function

'-------------------
' ini file api
'-------------------
'*******************************************************************************
' CreateParameterDataHashFromIniFileContext
'   @param iniFileContext [in] ini file context
'   @retval parameter data hash
'*******************************************************************************
Function CreateParameterDataHashFromIniFileContext(iniFileContext)
  Dim parameterDataArray
  
  parameterDataArray = Split(iniFileContext, DefineCrLf)
  
  Dim parameterDataHash
  Set parameterDataHash = WScript.CreateObject(NAME_OF_SCRIPTING_DICTIONARY)
  
  Dim indexOfParameterData
  For indexOfParameterData = LBound(parameterDataArray) To UBound(parameterDataArray)
    Dim isExistConst
    isExistConst = InStr(parameterDataArray, "Const ")
    Dim parameterName
    Dim parameterValue
    If isExistCount <> 0 Then
      ' exist Const
      ' TODO
    Else
      ' not exist Const
      Dim noSpaceParameterData
      ' TODO consider include space
      noSpaceParameterData = DeleteSpace(parameterDataArray(indexOfParameterData))
      Dim parameterNameAndValue
      parameterNameAndValue = Split(noSpaceParamterData, DEFINE_EQUAL, 1)
      parameterName = parameterNameAndValue(0)
      parameterValue = parameterNameAndValue(1)
    End If
    funcDummy = SetValueToParameterDataHash(parameterName, parameterValue, parameterDataHash)
  Next
  
  Set CreateParameterDataHashFromIniFileContext = parameterDataHash
End Function

'*******************************************************************************
' SetValueToParameterDataHash
'   @param parameterName [in] parameter name
'   @param parameterValue [in] parameter value
'   @param parameterDataHash [in/out] parameter data hash
'   @retval nothing
'*******************************************************************************
Function SetValueToParameterDataHash(parameterName, parameterValue, parameterDataHash)
  If parameterDataHash.Exists(parameterName) = True Then
    parameterDataHash.Item(parameterName) = parameterValue
  Else
    funcDummy = parameterDataHash.Add(parameterName, parameterValue)
  End If
End Function

'*******************************************************************************
' SetValueToParameterDataHashForString
'   @param parameterName [in] parameter name
'   @param parameterValue [in] parameter value
'   @param parameterDataHash [in/out] parameter data hash
'   @retval nothing
'*******************************************************************************
Function SetValueToParameterDataHashForString(parameterName, parameterValue, parameterDataHash)
  Dim stringParameterValue
  stringParameterValue = DEFINE_DOUBLE_QUOTE & parameterValue & DEFINE_DOUBLE_QUOTE
  funcDummy = SetValueToParameterDataHash(parameterName, stringParameterValue, parameterDataHash)
End Function

'*******************************************************************************
' SaveParameterDataHashToIniFile
'   @param parameterDataHash [in] parameter data hash
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function SaveParameterDataHashToIniFile(parameterDataHash, filePath)
  Dim objFile
  Set objFile = CreateFile(filePath)
  
  Dim keys
  keys = parameterDataHash.Keys()
  For Each key In keys
    funcDummy = WriteToObjectFile(objFile, key & DEFINE_EQUAL & parameterDataHash.Item(key))
  Next
  
  CloseObjectFile(objFile)
  
  Set objFile = Nothing
End Function

'*******************************************************************************
' CreateParameterDataArrayFromIniFileContext
'   @param iniFileContext [in] ini file context
'   @retval parameter data hash
'*******************************************************************************
Function CreateParameterDataArrayFromIniFileContext(iniFileContext)
  Dim parameterDataArray
  
  parameterDataArray = Split(iniFileContext, DefineCrLf)
  
  ' TODO
  
  Set CreateParameterDataArrayFromIniFileContext = parameterDataArray
End Function

'*******************************************************************************
' SetParameterValue
'   @param parameterName [in] parameter name
'   @param parameterValue [in] parameter value
'   @param parameterDataArray [in/out] parameter data array
'   @retval nothing
'*******************************************************************************
Function SetParameterDataToValue(parameterName, parameterValue, parameterDataArray)
  For Each parameterData In parameterDataArray
    ' TODO consider include space
    Dim noSpaceParameterData
    noSpaceParameterData = DeleteSpace(parameterData)
    ' TODO
  Next
End Function

'*******************************************************************************
' SetParameterValueForString
'   @param parameterName [in] parameter name
'   @param parameterValue [in] parameter value
'   @param parameterDataArray [in/out] parameter data array
'   @retval nothing
'*******************************************************************************
Function SetValueToParameterDataHashForString(parameterName, parameterValue, parameterDataArray)
  Dim stringParameterValue
  stringParameterValue = DEFINE_DOUBLE_QUOTE & parameterValue & DEFINE_DOUBLE_QUOTE
  funcDummy = SetParameterDataToValue(parameterName, stringParameterValue, parameterDataArray)
End Function

'*******************************************************************************
' SaveParameterDataArrayToIniFile
'   @param parameterDataArray [in] parameter data array
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function SaveParameterDataArrayToIniFile(parameterDataArray, filePath)
  Dim objFile
  Set objFile = CreateFile(filePath)
  
  For Each parameterData In parameterDataArray
    funcDummy = WriteToObjectFile(objFile, parameterData)
  Next
  
  CloseObjectFile(objFile)
  
  Set objFile = Nothing
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

'*******************************************************************************
' MatchRegex
'   @param targetString [in] target string
'   @param regexString [in] regex string
'   @param ignoreCase [in] ignore case(true/false)
'   @retval collection
'*******************************************************************************
Function MatchRegex(targetString, regexString, ignoreCase)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "MatchRegex start")
  
  logReturnValueDummy = LogOutDetailDebug(LOG_TARGET_LEVEL, "MatchRegex targetString: " & targetString)
  logReturnValueDummy = LogOutDetailDebug(LOG_TARGET_LEVEL, "MatchRegex regexString: " & regexString)
  
  Dim regex
  Dim matches
  Set regex = New RegExp
  regex.Pattern = regexString
  regex.IgnoreCase = ignoreCase
  regex.Global = True
  Set matches = regex.Execute(targetString)
  
  For Each Match in Matches
    logReturnValueDummy = LogOutDetailDebug(LOG_TARGET_LEVEL, "MatchRegex Match.FirstIndex: " & Match.FirstIndex)
    logReturnValueDummy = LogOutDetailDebug(LOG_TARGET_LEVEL, "MatchRegex Match.Value: " & Match.Value)
  Next
  
  Set regex = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "MatchRegex end")
  
  Set MatchRegex = matches
End Function

'*******************************************************************************
' IsMatchRegex
'   @param targetString [in] target string
'   @param regexString [in] regex string
'   @param ignoreCase [in] ignore case(true/false)
'   @retval true/false true:matched false:not matched
'*******************************************************************************
Function IsMatchRegex(targetString, regexString, ignoreCase)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex start")
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex targetString: " & targetString)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex regexString: " & regexString)
  
  Dim regex
  Dim matches
  Set regex = New RegExp
  regex.Pattern = regexString
  regex.IgnoreCase = ignoreCase
  regex.Global = True
  Set matches = regex.Execute(targetString)
  
  For Each Match in Matches
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex Match.FirstIndex: " & Match.FirstIndex)
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex Match.Value: " & Match.Value)
  Next
  
  Dim isMatched
  If matches.Count = 0 Then
    isMatched = False
  Else
    isMatched = True
  End If
  
  Set matches = Nothing
  Set regex = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegex end")
  
  IsMatchRegex = isMatched
End Function

'*******************************************************************************
' IsMatchRegexArray
'   @param targetString [in] target string
'   @param regexStringArray [in] regex string array
'   @param ignoreCase [in] ignore case(true/false)
'   @retval true/false true:matched false:not matched
'*******************************************************************************
Function IsMatchRegexArray(targetString, regexStringArray, ignoreCase)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegexArray start")
  
  Dim resultIsMatched
  resultIsMatched = False
  For Each regexString in regexStringArray
    Dim isMatched
    isMatched = IsMatchRegex(targetString, regexString, ignoreCase)
    If resultIsMatched = False And isMatched = True Then
      resultIsMatched = True
    End If
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsMatchRegexArray end")
  
  IsMatchRegexArray = resultIsMatched
End Function


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
'---------------------------------------
' IE object api
'---------------------------------------
'*******************************************************************************
' CreateIEObject
'   @param isShowWindow [in] is show window(true or false)
'   @param url [in] url
'   @param waitTime [in] wait time
'   @retval IE object
'*******************************************************************************
Function CreateIEObject(isShowWindow, url, waitTime)
  Dim objIE
  
  Set objIE = WScript.CreateObject(NAME_OF_IE_APPLICATION)
  objIE.Visible = isShowWindow
  
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

'---------------------------------------
' excel object api
'---------------------------------------
'*******************************************************************************
' CreateEXCELObject
'   @param isShowWindow [in] is show window(true or false)
'   @retval excel object
'*******************************************************************************
Function CreateEXCELObject(isShowWindow)
  Dim objExcel
  
  Set objExcel = WScript.CreateObject(NAME_OF_EXCEL_APPLICATION)
  objExcel.Visible = isShowWindow
  
  Set CreateEXCELObject = objExcel
End Function

'*******************************************************************************
' OpenWorkBooksOfExcel
'   @param objExcel [in] object excel
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function OpenWorkBooksOfExcel(objExcel, filePath)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "OpenWorkBooksOfExcel start")
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "start excel open:" & filePath)
  
  objExcel.Workbooks.Open(filePath)
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "end excel open:" & filePath)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "OpenWorkBooksOfExcel end")
End Function

'*******************************************************************************
' SetCellsOfExcel
'   @param objExcel [in] object excel
'   @param numberOfWorkBook [in] number of work book
'   @param numberOfWorkSheet [in] number of work sheet
'   @param nameOfCellRow [in] name of cell row
'   @param nameOfCellColl [in] number of cell coll
'   @param valueOfCell [in] value of cell
'   @retval nothing
'*******************************************************************************
Function SetCellsOfExcel(objExcel, numberOfWorkBook, numberOfWorkSheet, nameOfCellRow, nameOfCellColl, valueOfCell)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SetCellsOfExcel start")
  
  objExcel.Workbooks(numberOfWorkBook).Worksheets(numberOfWorkSheet).Cells(nameOfCellColl, nameOfCellRow) = valueOfCell
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SetCellsOfExcel end")
End Function

'*******************************************************************************
' SaveOfExcel
'   @param objExcel [in] object excel
'   @param numberOfWorkBook [in] number of work book
'   @retval nothing
'*******************************************************************************
Function SaveOfExcel(objExcel, numberOfWorkBook)
  objExcel.Workbooks(numberOfWorkBook).Save()
End Function

'---------------------------------------
' dictionary object api
'---------------------------------------
'*******************************************************************************
' LogoutDictionaryObject
'   @param objDictionary [in] object dictionary
'   @param logoutPrefix [in] logout prefix
'   @retval nothing
'*******************************************************************************
Function LogoutDictionaryObject(objDictionary, logoutPrefix)
  For Each key In objDictionary
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, logoutPrefix & " key: " & key)
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, logoutPrefix & " value: " & objDictionary.Item(key))
  Next
End Function

'-------------------------------------------------------------------------------
' command api
'-------------------------------------------------------------------------------
'*******************************************************************************
' ExecCommand
'   @param command [in] command
'   @retval nothing
'*******************************************************************************
Function ExecCommand(command)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecCommand start")
  
  Dim WshShell
  Set WshShell = WScript.CreateObject(NAME_OF_WSCRIPT_SHELL)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecCommand command start: " & command)
  
  WshShell.Exec(command)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecCommand command end: " & command)
  
  Set WshShell = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecCommand end")
End Function

'*******************************************************************************
' ExecAndGetReturnCommand
'   @param command [in] command
'   @retval nothing
'*******************************************************************************
Function ExecAndGetReturnCommand(command)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndGetReturnCommand start")
  
  Dim WshShell
  Dim outExec
  Dim outStream
  Dim strOut
  Set WshShell = WScript.CreateObject(NAME_OF_WSCRIPT_SHELL)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndGetReturnCommand command start: " & command)
  
  Set outExec = WshShell.Exec(command)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndGetReturnCommand command end: " & command)
  
  Set outStream = outExec.StdOut
  strOut = ""
  Do While Not outStream.AtEndOfStream
    strOut = strOut & vbNewLine & outStream.ReadLine()
  Loop
  Set WshShell = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndGetReturnCommand end")
  
  ExecAndGetReturnCommand = strOut
End Function

'*******************************************************************************
' ExecAndWaitCommand
'   @param command [in] command
'   @retval nothing
'*******************************************************************************
Function ExecAndWaitCommand(command)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndWaitCommand start")
  
  Dim WshShell
  Dim result
  Set WshShell = WScript.CreateObject(NAME_OF_WSCRIPT_SHELL)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndWaitCommand command start: " & command)
  
  result = WshShell.Run(command, 1, true)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndWaitCommand command end: " & command)
  
  Set WshShell = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ExecAndWaitCommand end")
  
  ExecAndWaitCommand = result
End Function
