'Option Explicit

'-------------------------------------------------------------------------------
Const NAME_OF_FILESYSTEM_IN_API = "Scripting.FileSystemObject"

Dim FileShellInAPI
Set FileShellInAPI = WScript.CreateObject(NAME_OF_FILESYSTEM_IN_API)

Const FOR_READING_INCLUDE_IN_API = 1

'*******************************************************************************
' read vbs file in api
'   @param FileName [in] read vbs file name
'   @retval nothing
'*******************************************************************************
Function ReadVBSFileInAPI(ByVal FileName)
  ReadVBSFileInAPI = FileShellInAPI.OpenTextFile(FileName, FOR_READING_INCLUDE_IN_API, False).ReadAll()
End Function

'Execute ReadVBSFileInAPI("IncludeCommonAPI.vbs")
'Execute ReadVBSFileInAPI("IncludeConfig.vbs")


'===============================================================================
' api
'===============================================================================
'*******************************************************************************
' InitETCUseInfoOfJapanHightWay function
'   @param nothing
'   @retval nothing
'*******************************************************************************
Function InitETCUseInfoOfJapanHightWay()
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "InitETCUseInfoOfJapanHightWay start")
  
  ' TODO
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "InitETCUseInfoOfJapanHightWay end")
End Function

'*******************************************************************************
' GetETCUseInfoOfJapanHightWay function
'   @param nothing
'   @retval nothing
'*******************************************************************************
Function GetETCUseInfoOfJapanHightWay()
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetETCUseInfoOfJapanHightWay start")
  
  Dim periodParams
  Set periodParams = GetTargetPeriod(MODE_OF_AUTO_CALC_DATE)
  
  Dim targetToYear
  Dim targetToMonth
  Dim targetToDay
  Dim targetFromYear
  Dim targetFromMonth
  Dim targetFromDay
  targetToYear = periodParams.Item(NAME_OF_USE_TO_YEAR)
  targetToMonth = periodParams.Item(NAME_OF_USE_TO_MONTH)
  targetToDay = periodParams.Item(NAME_OF_USE_TO_DAY)
  targetFromYear = periodParams.Item(NAME_OF_USE_FROM_YEAR)
  targetFromMonth = periodParams.Item(NAME_OF_USE_FROM_MONTH)
  targetFromDay = periodParams.Item(NAME_OF_USE_FROM_DAY)
  
  ' get script file path
  Dim strScriptPath
  Dim strPeriodDate
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  strPeriodDate = targetFromYear & targetFromMonth & targetFromDay & DEFINE_HYPHEN & targetToYear & targetToMonth & targetToDay
  
  Dim userInfos
  userInfos = ReadUserInfoFile(FILE_NAME_OF_USER_INFO)
  If IsNull(userInfos) = True Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, "please confirm user info file: " & FILE_NAME_OF_USER_INFO)
    WScript.Quit 1
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "user info size: " & UBound(userInfos))
  Dim mainIEObj
  'For Each userInfo In userInfos
  For indexObUserInfos = 0 To UBound(userInfos) - 1
    Dim userInfo
    userInfo = userInfos(indexObUserInfos)
    If IsNull(userInfo) = True Then
      ' skip
    Else
      Dim carNumber
      Dim icCardNumber
      Dim otherInfo
      
      carNumber = userInfo(INDEX_OF_CAR_NUMBER)
      icCardNumber = userInfo(INDEX_OF_ID_CARD_NUMBER)
      otherInfo = userInfo(INDEX_OF_OTHER_INFO)
      
      Set mainIEObj = CreateIEObject(IS_SHOW_MAIN_WEB_GUI, URL_OF_ETC_SITE, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
      funcDummy = SetFormToIE(mainIEObj, periodParams, carNumber, icCardNumber)
      
      ' Enter
      mainIEObj.Document.forms(0).submit
      
      ' wait
      funcDummy = WaitIEObject(mainIEObj, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
      
      ' error check
      ' TODO
      
      Dim useResult
      Set useResult = CreateObject(NAME_OF_SCRIPTING_DICTIONARY)
      
      Dim currentPage
      currentPage = 1
      Dim sequenceNumber
      sequenceNumber = 1
      Dim isContinue
      isContinue = True
      ' request and parse, until reach end of link page
      Do Until isContinue = False
        isContinue = False
        isContinue = RequestAndParsePage(mainIEObj, sequenceNumber, useResult)
        sequenceNumber = sequenceNumber + 1
        
        Dim objAOfTag
        Set objAOfTag = mainIEObj.Document.getElementsByTagName(NAME_OF_A_NAME)
        Dim indexOfATag
        Dim hrefName
        Dim targetPage
        targetPage = -1
        Dim isClickNextPage
        isClickNextPage = False
        For indexOfATag = 0 To objAOfTag.Length - 1
          hrefName = objAOfTag(indexOfATag).getAttribute(NAME_OF_ATTRIBUTE_HREF)
          If hrefName <> DEFINE_BRANK Then
            Dim hrefNameParts
            hrefNameParts = Split(hrefName, NAME_OF_LINK_PAGE)
            If UBound(hrefNameParts) = 1 Then
              targetPage = CInt(hrefNameParts(1))
              logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "target page number is: " & targetPage & ". current page number is: " & currentPage)
            Else
              ' skip
            End If
            
            ' TODO
            
            If currentPage < targetPage Then
              logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "currentPage page is: " & currentPage & " < target page is: " & hrefName)
              isClickNextPage = True
              
              Exit For
            End If
          End If
        Next
        
        If isClickNextPage = True Then
          logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "isClickNextPage is: " & isClickNextPage)
          currentPage = targetPage
          objAOfTag(indexOfATag).Click
          indexOfATag = objAOfTag.Length
          
          funcDummy = WaitIEObject(mainIEObj, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
          
          isContinue = True
        End If
      Loop
      
      Dim summaryResult
      Set summaryResult = CreateObject(NAME_OF_SCRIPTING_DICTIONARY)
      funcDummy = CountUseInfo(MODE_OF_AUTO_CALC_DATE, useResult, summaryResult)
      
      ' print debug
      funcDummy = LogoutDictionaryObject(summaryResult, "summary")
      
      Dim strSaveFolderPath
      strSaveFolderPath = strScriptPath & strPeriodDate & DEFINE_DELIM_FOLDER & carNumber & DEFINE_HYPHEN & Right(icCardNumber, 4) & DEFINE_HYPHEN & otherInfo
      ExecAndWaitCommand("cmd /c mkdir " & strSaveFolderPath)
      
      If IS_SAVE_SUM_FILE = True Then
        Dim strSaveFilePath
        strSaveFilePath = strSaveFolderPath & DEFINE_DELIM_FOLDER & FILE_NAME_OF_SAVE_SUM_FILE
        CreateFolder(strSaveFolderPath)
        CreateFile(strSaveFilePath)
        Dim objFile
        Set objFile = OpenFileToWrite(strSaveFilePath)
        Dim explain
        explain = EXPLAIN_OF_GATES_IN_SUMMARY & DEFINE_DELIM_CANMA & EXPLAIN_OF_TOLL_IN_SUMMARY & DEFINE_DELIM_CANMA & EXPLAIN_OF_DATE_IN_SUMMARY & DEFINE_EQUAL & EXPLAIN_OF_COUNT_IN_SUMMARY
        funcDummy = WriteLineToObjectFile(objFile, explain)
        For Each key In summaryResult
          Dim context
          context = key & DEFINE_EQUAL & summaryResult.Item(key)
          funcDummy = WriteLineToObjectFile(objFile, context)
        Next
        CloseObjectFile(objFile)
        Set objFile = Nothing
      End If
      
      ' concat pdf
      If IS_DO_CONCAT_PDF = True Then
        ExecAndWaitCommand("cmd /c margePDF.bat")
      End If
      
      ' move to pdf
      ExecAndWaitCommand("cmd /c move /Y " & strScriptPath & "*.pdf " & strSaveFolderPath & DEFINE_DELIM_FOLDER)
      
      If IS_SAVE_EXCEL = True Then
        Dim strSaveExcelPath
        strSaveExcelPath = strSaveFolderPath & DEFINE_DELIM_FOLDER & FILE_NAME_OF_EXCEL
        ExecAndWaitCommand("cmd /c copy /Y " & strScriptPath & FILE_NAME_OF_EXCEL & DEFINE_SPACE & strSaveExcelPath)
        funcDummy = SaveSummaryInExcel(strSaveExcelPath, summaryResult)
      End If
      
      Set mainIEObj = Nothing
    End If
  Next
  
  ' TODO
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetETCUseInfoOfJapanHightWay end")
End Function

'*******************************************************************************
' ReadUserInfoFile 
'   @param filePath [in] get mode
'   @retval user info
'*******************************************************************************
Function ReadUserInfoFile(filePath)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ReadUserInfoFile start")
  
  Dim objFile
  
  Set objFile = OpenFileToRead(filePath)
  If IsNull(objFile) = True Then
    logReturnValueDummy = LogOutFatal(LOG_TARGET_LEVEL, "please confirm user info file: " & filePath)
    WScript.Quit 1
  End If
  
  ReDim userInfos(-1)
  Dim context
  Do Until objFile.AtEndOfLine = True
    context = ReadFromObjectFile(objFile)
    If IsNull(context) = True Then
      Exit Do
    End If
    
    Dim contextOfNoSpace
    Dim currentUserInfoSize
    Dim userInfo
    
    contextOfNoSpace = DeleteSpace(context)
    
    Dim firstChar
    firstChar = Left(context, 1)
    If firstChar = DEFINE_SINGLE_QUOTE Then
      ' skip comment
    Else
      userInfo = Split(contextOfNoSpace, DEFINE_DELIM_CANMA, SIZE_OF_USER_INFO_INDEX)
      If UBound(userInfo) = SIZE_OF_USER_INFO_INDEX Then
        ' skip invalid format
      Else
        currentUserInfoSize = UBound(userInfos)
        If currentUserInfoSize = -1 Then
          currentUserInfoSize = 0
        End If
        ReDim Preserve userInfos(currentUserInfoSize + 1)
        userInfos(currentUserInfoSize) = userInfo
      End If
    End If
  Loop
  
  CloseObjectFile(objFile)
  Set objFile = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ReadUserInfoFile end")
  
  ReadUserInfoFile = userInfos
End Function

'*******************************************************************************
' get target period
'   @param getDateMode [in] get mode
'   @retval resultPeriodHash result period hash
'*******************************************************************************
Function GetTargetPeriod(getDateMode)
  Dim getPeriodHash
  Dim resultPeriodHash
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetTargetPeriod start")
  
  If getDateMode = 1 Then
    ' "auto 20 day per a month"
    Set getPeriodHash = GetTargetPeriodByAuto20DayPerAMonth()
  ElseIf getDateMode = 2 Then
    ' "auto 20 day per a month(each toll)"
    Set getPeriodHash = GetTargetPeriodByAuto20DayPerAMonth()
  Else
    ' "auto 20 day per a month"
    Set getPeriodHash = GetTargetPeriodByAuto20DayPerAMonth()
  End If
  
  Set resultPeriodHash = PaddingTargetPeriod(getPeriodHash)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetTargetPeriod end")
  
  Set GetTargetPeriod = resultPeriodHash
End Function

'*******************************************************************************
' get target period by auto 20 day per a month
'   @param nothing
'   @retval resultPeriodHash result period hash
'*******************************************************************************
Function GetTargetPeriodByAuto20DayPerAMonth()
  Dim resultPeriodHash
  Dim currentMonth
  Dim currentDay
  Dim targetPrevYear
  Dim targetPrevMonth
  Dim targetPrevDay
  Dim targetCurrentYear
  Dim targetCurrentMonth
  Dim targetCurrentDay
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetTargetPeriodByAuto20DayPerAMonth start")
  
  currentMonth = Month(Now)
  currentDay = Day(Now)
  
  targetPrevYear = Year(Now)
  targetCurrentYear = Year(Now)
  
  ' detect month
  If currentDay > 20 Then
    targetPrevMonth = currentMonth - 1
    targetCurrentMonth = currentMonth
  Else
    targetPrevMonth = currentMonth - 2
    targetCurrentMonth = currentMonth - 1
  End If
  
  targetPrevDay = 21
  targetCurrentDay = 20
  If targetPrevMonth = 0 Then
    targetPrevYear = targetPrevYear - 1
    targetPrevMonth = 12
  ElseIf targetPrevMonth = -1 Then
    targetPrevYear = targetPrevYear - 1
    targetPrevMonth = 11
    targetCurrentYear = - 1
    targetCurrentMonth = 12
  End If
  
  Set resultPeriodHash = CreateObject(NAME_OF_SCRIPTING_DICTIONARY)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_YEAR, targetPrevYear)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_MONTH, targetPrevMonth)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_DAY, targetPrevDay)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_YEAR, targetCurrentYear)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_MONTH, targetCurrentMonth)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_DAY, targetCurrentDay)
  
  Set currentMonth = Nothing
  Set currentDay = Nothing
  Set targetPrevYear = Nothing
  Set targetPrevMonth = Nothing
  Set targetPrevDay = Nothing
  Set targetCurrentYear = Nothing
  Set targetCurrentMonth = Nothing
  Set targetCurrentDay = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetTargetPeriodByAuto20DayPerAMonth end")
  
  Set GetTargetPeriodByAuto20DayPerAMonth = resultPeriodHash
End Function

'*******************************************************************************
' padding target period
'   @param periodParams [in] period params
'   @retval resultPeriodHash result period hash
'*******************************************************************************
Function PaddingTargetPeriod(periodParams)
  Dim resultPeriodHash
  Dim currentMonth
  Dim currentDay
  Dim targetPrevYear
  Dim targetPrevMonth
  Dim targetPrevDay
  Dim targetCurrentYear
  Dim targetCurrentMonth
  Dim targetCurrentDay
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingTargetPeriod start")
  
  targetPrevYear = PaddingPrefixString(periodParams(NAME_OF_USE_FROM_YEAR), "0", 4)
  targetPrevMonth =  PaddingPrefixString(periodParams(NAME_OF_USE_FROM_MONTH), "0", 2)
  targetPrevDay = PaddingPrefixString(periodParams(NAME_OF_USE_FROM_DAY), "0", 2)
  targetCurrentYear = PaddingPrefixString(periodParams(NAME_OF_USE_TO_YEAR), "0", 4)
  targetCurrentMonth = PaddingPrefixString(periodParams(NAME_OF_USE_TO_MONTH), "0", 2)
  targetCurrentDay = PaddingPrefixString(periodParams(NAME_OF_USE_TO_DAY), "0", 2)
  
  Set resultPeriodHash = CreateObject(NAME_OF_SCRIPTING_DICTIONARY)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_YEAR, targetPrevYear)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_MONTH, targetPrevMonth)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_FROM_DAY, targetPrevDay)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_YEAR, targetCurrentYear)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_MONTH, targetCurrentMonth)
  funcDummy = resultPeriodHash.Add(NAME_OF_USE_TO_DAY, targetCurrentDay)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "PaddingTargetPeriod end")
  
  Set PaddingTargetPeriod = resultPeriodHash
End Function

'*******************************************************************************
' set form to IE
'   @param objIE [in] object IE
'   @param periodParams [in] period params
'   @param carNumber [in] car number
'   @param icCardNumber [in] ic card number
'   @retval nothing
'*******************************************************************************
Function SetFormToIE(objIE, periodParams, carNumber, icCardNumber)
  Dim resultPeriodHash
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SetFormToIE start")
  
  Dim objCarNumber
  Dim objICCardNumber
  Dim objFromYear
  Dim objFromMonth
  Dim objFromDay
  Dim objToYear
  Dim objToMonth
  Dim objToDay
  
  Set objCarNumber = objIE.Document.getElementsByName(NAME_OF_USE_CAR_NUMBER)
  Set objICCardNumber = objIE.Document.getElementsByName(NAME_OF_USE_ETC_CARD_NUMBER)
  Set objFromYear = objIE.Document.getElementsByName(NAME_OF_USE_FROM_YEAR)
  Set objFromMonth = objIE.Document.getElementsByName(NAME_OF_USE_FROM_MONTH)
  Set objFromDay = objIE.Document.getElementsByName(NAME_OF_USE_FROM_DAY)
  Set objToYear = objIE.Document.getElementsByName(NAME_OF_USE_TO_YEAR)
  Set objToMonth = objIE.Document.getElementsByName(NAME_OF_USE_TO_MONTH)
  Set objToDay = objIE.Document.getElementsByName(NAME_OF_USE_TO_DAY)
  
  Dim errorMessage
  errorMessage = "サイト内容が変わったか、メンテナンス中です。表示されている内容を確認し、必要があれば、開発者に問い合わせてください。"
  If objCarNumber.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_CAR_NUMBER)
    WScript.Quit 1
  End If
  If objICCardNumber.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_ETC_CARD_NUMBER)
    WScript.Quit 1
  End If
  If objFromYear.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_FROM_YEAR)
    WScript.Quit 1
  End If
  If objFromMonth.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_FROM_MONTH)
    WScript.Quit 1
  End If
  If objFromDay.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_FROM_DAY)
    WScript.Quit 1
  End If
  If objToYear.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_TO_YEAR)
    WScript.Quit 1
  End If
  If objToMonth.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_TO_MONTH)
    WScript.Quit 1
  End If
  If objToDay.Length = 0 Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, errorMessage & NAME_OF_USE_TO_DAY)
    WScript.Quit 1
  End If
  
  Dim setParameterLogMessage
  setParameterLogMessage = "from: " & periodParams(NAME_OF_USE_FROM_YEAR) & DEFINE_DELIM_DATE & periodParams(NAME_OF_USE_FROM_MONTH) & DEFINE_DELIM_DATE & periodParams(NAME_OF_USE_FROM_DAY)
  setParameterLogMessage = setParameterLogMessage & " to: " & periodParams(NAME_OF_USE_TO_YEAR) & DEFINE_DELIM_DATE & periodParams(NAME_OF_USE_TO_MONTH) & DEFINE_DELIM_DATE & periodParams(NAME_OF_USE_TO_DAY)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, setParameterLogMessage)
  
  objCarNumber(0).Value = carNumber
  Set objCarNumber = Nothing
  objICCardNumber(0).Value = icCardNumber
  Set objICCardNumber = Nothing
  objFromYear(0).Value = periodParams(NAME_OF_USE_FROM_YEAR)
  Set objFromYear = Nothing
  objFromMonth(0).Value = periodParams(NAME_OF_USE_FROM_MONTH)
  Set objFromMonth = Nothing
  objFromDay(0).Value = periodParams(NAME_OF_USE_FROM_DAY)
  Set objFromDay = Nothing
  objToYear(0).Value = periodParams(NAME_OF_USE_TO_YEAR)
  Set objToYear = Nothing
  objToMonth(0).Value = periodParams(NAME_OF_USE_TO_MONTH)
  Set objToMonth = Nothing
  objToDay(0).Value = periodParams(NAME_OF_USE_TO_DAY)
  Set objToDay = Nothing
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SetFormToIE end")
  
  'SetFormToIE = 
End Function

'*******************************************************************************
' RequestAndParsePage
'   @param objIE [in] IE object
'   @param sequenceNumber [in] sequence number
'   @param useResult [in/out] use result
'   @retval true/false true:continue, false:not continue
'*******************************************************************************
Function RequestAndParsePage(objIE, sequenceNumber, useResult)
  Dim result
  result = False
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "RequestAndParsePage start")
  
  Dim objInputTags
  Set objInputTags = objIE.Document.getElementsByTagName(NAME_OF_INPUT)
  
  If objInputTags.Length = 0 Then
    Dim errorMessage
    errorMessage = "入力に不正があるか、メンテナンス中です。表示されている内容を確認し、必要があれば、開発者に問い合わせてください。"
    logReturnValueDummy = logOutError(LOG_TARGET_LEVEL, errorMessage)
    WScript.Quit 1
  Else
    Dim indexOfInputTag
    For indexOfInputTag = 0 To objInputTags.Length - 1
      Dim typeOfAttrName
      Dim nameOfAttrName
      typeOfAttrName = objInputTags(indexOfInputTag).getAttribute(NAME_OF_ATTRIBUTE_TYPE)
      nameOfAttrName = objInputTags(indexOfInputTag).getAttribute(NAME_OF_ATTRIBUTE_NAME)
      If typeOfAttrName = NAME_OF_CHECK_BOX Then
        CheckHightWayUse(objInputTags(indexOfInputTag))
      End If
    Next
  End If
  
  If IS_CONFORM_BEFORE_HIGHT_WAY_USE_DETERM = true Then
    MsgBox "料金計算発行を実施します。" & DefineCrLr & "自動チェック内容を確認し、必要があればチェック操作してください。" & DefineCrLf & "確認完了後、「OK」を押してください。"
  End If
  
  Dim bodyOfHtml
  bodyOfHtml = objIE.Document.body.InnerHtml
  
  Dim isExistCheck
  isExistCheck = ParseBodyOfHtml(bodyOfHtml, objIE, useResult)
  
  If IS_SAVE_USE_CONTEXT_PDF = True AND isExistCheck = True Then
    objIE.Document.forms(0).submit
    
    funcDummy = WaitIEObject(objIE, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
    
    ' TODO windows control
    Dim objShell
    Dim objPDFOfIE
    Set objShell = CreateObject(NAME_OF_SHELL_APPLICATION)
    Set objPDFOfIE = objShell.Windows.Item(objShell.Windows.Count - 1)
    
    funcDummy = WaitIEObject(objPDFOfIE, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
    
    Dim locationURL
    locationURL = objPDFOfIE.LocationURL
    funcDummy = GetHttp(locationURL, SAVE_PREFIX_OF_USE_CONTEXT_PDF & sequenceNumber & SAVE_SUFFIX_OF_USE_CONTEXT_PDF, PROXY_SERVER)
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "RequestAndParsePage end")
  
  ' TODO
  ' result
  RequestAndParsePage = result
End Function

'*******************************************************************************
' CheckMatching
'   @param targetString [in] target string
'   @param regexStringOfConfig [in] regex string of config
'   @retval true/false true:match false:not match
'*******************************************************************************
Function CheckMatching(targetString, regexStringOfConfig)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckMatching start")
  
  Dim isMatch
  isMatch = False
  If Len(targetString) <> 0 And Len(regexStringOfConfig) <> 0 Then
    Dim regexArray
    regexArray = GetRegexArray(regexStringOfConfig)
    isMatch = IsMatchRegexArray(targetString, regexArray, true)
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckMatching end")
  
  CheckMatching = isMatch
End Function

'*******************************************************************************
' CheckMatchingAll
'   @param hightWayUseInfo [in] hight way use info
'   @retval true/false true:match false:not match
'*******************************************************************************
Function CheckMatchingAll(hightWayUseInfo)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckMatchingAll start")
  
  ' TODO check match
  '   hightWayUseInfo(NUMBER_OF_DATE_AT_SUMMARY)
  '   hightWayUseInfo(NUMBER_OF_TIME_AT_SUMMARY)
  
  ' gate's
  Dim isMatchOfGates
  Dim isMatchOfGatesReverse
  isMatchOfGates = False
  isMatchOfGatesReverse = False
  If Len(NAMES_OF_USE_TARGET_GATE) <> 0 Then
    isMatchOfGates = CheckMatching(hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY) & DEFINE_HYPHEN & hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY), NAMES_OF_USE_TARGET_GATE)
    isMatchOfGatesReverse = CheckMatching(hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY) & DEFINE_HYPHEN & hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY), NAMES_OF_USE_TARGET_GATE)
  End If
  
  ' first gate
  Dim isMatchOfFirstGate
  isMatchOfFirstGate = False
  If Len(FIRST_NAME_OF_USE_TARGET_GATE) <> 0 Then
    isMatchOfFirstGate = CheckMatching(hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY), FIRST_NAME_OF_USE_TARGET_GATE)
  End If
  
  ' second gate
  Dim isMatchOfSecondGate
  isMatchOfSecondGate = False
  If Len(SECOND_NAME_OF_USE_TARGET_GATE) <> 0 Then
    isMatchOfSecondGate = CheckMatching(hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY), SECOND_NAME_OF_USE_TARGET_GATE)
  End If
  
  ' toll
  Dim isMatchOfTollGate
  isMatchOfTollGate = False
  If Len(TOLL_OF_USE_TARGET) <> 0 Then
    isMatchOfTollGate = CheckMatching(hightWayUseInfo(NUMBER_OF_TOLL_AT_SUMMARY), TOLL_OF_USE_TARGET)
  End If
  
  ' gate's exclude
  Dim isMatchOfGatesExclude
  Dim isMatchOfGatesReverseExclude
  isMatchOfGatesExclude = False
  isMatchOfGatesReverseExclude = False
  If Len(NAMES_OF_USE_EXCLUDE_GATE) <> 0 Then
    isMatchOfGatesExclude = CheckMatching(hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY) & DEFINE_HYPHEN & hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY), NAMES_OF_USE_EXCLUDE_GATE)
    isMatchOfGatesReverseExclude = CheckMatching(hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY) & DEFINE_HYPHEN & hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY), NAMES_OF_USE_EXCLUDE_GATE)
  End If
  
  ' first gate exclude exclude
  Dim isMatchOfFirstGateExclude
  isMatchOfFirstGateExclude = False
  If Len(FIRST_NAME_OF_EXCLUDE_GATE) <> 0 Then
    isMatchOfFirstGateExclude = CheckMatching(hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY), FIRST_NAME_OF_EXCLUDE_GATE)
  End If
  
  ' second gate exclude exclude
  Dim isMatchOfSecondGateExclude
  isMatchOfSecondGateExclude = False
  If Len(SECOND_NAME_OF_EXCLUDE_GATE) <> 0 Then
    isMatchOfSecondGateExclude = CheckMatching(hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY), SECOND_NAME_OF_EXCLUDE_GATE)
  End If
  
  ' toll exclude exclude
  Dim isMatchOfTollGateExclude
  isMatchOfTollGateExclude = False
  If Len(TOLL_OF_EXCLUDE) <> 0 Then
    isMatchOfTollGateExclude = CheckMatching(hightWayUseInfo(NUMBER_OF_TOLL_AT_SUMMARY), TOLL_OF_EXCLUDE)
  End If
  
  Dim isMatch
  If isMatchOfGates = True Or isMatchOfGatesReverse = True Or isMatchOfFirstGate = True Or isMatchOfSecondGate = True Or isMatchOfTollGate = True Then
    isMatch = True
  Else
    isMatch = False
  End If
  Dim isMatchExclude
  If isMatchOfGatesExclude = True Or isMatchOfGatesReverseExclude = True Or isMatchOfFirstGateExclude = True Or isMatchOfSecondGateExclude = True Or isMatchOfTollGateExclude = True Then
    isMatchExclude = True
  Else
    isMatchExclude = False
  End If
  If isMatch = True And isMatchExclude = True Then
    isMatch = False
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckMatchingAll end")
  
  CheckMatchingAll = isMatch
End Function

'*******************************************************************************
' CheckHightWayUse
'   @param objElement [in] object element
'   @retval nothing
'*******************************************************************************
Function CheckHightWayUse(objElement)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckHightWayUse start")
  
  Dim key
  key = GetKeyFromBodyOfHtml(objElement)
  
  If IsNull(key) = True Then
    ' invalid key
    ' skip
  Else
    ' valid key
    Dim hightWayUseInfo
    hightWayUseInfo = GetHightWayUseInfoFromKey(key)
    
    Dim matched
    matched = CheckMatchingAll(hightWayUseInfo)
    
    If matched = True Then
      ' match
      Dim version
      version = GetIEVersion(objElement)
      
      If version = NUMBER_OF_IE10_VERSION Then
        funcDummy = objElement.SetAttribute(NAME_OF_CHECKED, NAME_OF_CHECKED_VALUE)
      ElseIf version = NUMBER_OF_IE8_VERSION Then
        objElement.Click()
      End If
    Else
      ' not match
    End If
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CheckHightWayUse end")
End Function

'*******************************************************************************
' GetIEVersion
'   @param objElement [in] object element
'   @retval version
'*******************************************************************************
Function GetIEVersion(objElement)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetIEVersion start")
  
  Dim version
  
  Dim targetHightWayUse
  targetHightWayUse = objElement.parentNode.parentNode.innerText
  
  Dim targetHightWayUseParts
  targetHightWayUseParts = Split(targetHightWayUse, DefineCrLf)
  If UBound(targetHightWayUseParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE10 Then
    version = 10
  ElseIf UBound(targetHightWayUseParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE8 Then
    version = 8
  Else
    ' invalid format
    ' skip
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetIEVersion end")
  
  GetIEVersion = version
End Function

'*******************************************************************************
' ParseBodyOfHtml
'   @param bodyOfHtml [in] body of html
'   @param objIE [in] object IE
'   @param useResult [in/out] use result
'   @retval true/false true:exist check, false:not exist check
'*******************************************************************************
Function ParseBodyOfHtml(bodyOfHtml, objIE, useResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ParseBodyOfHtml start")
  
  Dim isExistCheck
  isExistCheck = False
  
  Dim objInputTags
  Set objInputTags = objIE.Document.getElementsByTagName(NAME_OF_INPUT)
  Dim indexOfInput
  For indexOfInput = 0 To objInputTags.Length - 1
    Dim typeName
    typeName = objInputTags(indexOfInput).getAttribute(NAME_OF_ATTRIBUTE_TYPE)
    If typeName = NAME_OF_CHECK_BOX Then
      Dim isCheckedAttribute
      isCheckedAttribute = objInputTags(indexOfInput).getAttribute(NAME_OF_CHECKED)
      ' if detect by true/false or checked value, don't detect checked. there for check by not brank
      If isCheckedAttribute <> DEFINE_BRANK And isCheckedAttribute <> False Then
      'If isCheckedAttribute = NAME_OF_CHECKED_VALUE Then
      'If isCheckedAttribute = True Then
        Dim key
        key = GetKeyFromBodyOfHtml(objInputTags(indexOfInput))
        
        If IsNull(key) = True Then
          ' invalid key
          ' skip
        Else
          ' valid key
          funcDummy = useResult.Add(key, True)
          
          If isExistCheck = False Then
            isExistCheck = True
          End If
        End If
      Else
        ' didn't check
        ' skip
      End If
    Else
      ' don't checkbox
      ' skip
    End If
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ParseBodyOfHtml end")
  
  ParseBodyOfHtml = isExistCheck
End Function

'*******************************************************************************
' GetKeyFromBodyOfHtml
'   @param objElement [in] object element
'   @retval key
'*******************************************************************************
Function GetKeyFromBodyOfHtml(objElement)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetKeyFromBodyOfHtml start")
  
  Dim key
  
  ' checked
  Dim inputBody
  inputBody = objElement.parentNode.parentNode.innerText
  
  key = GetKeyFromSingleHightWayUse(inputBody)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetKeyFromBodyOfHtml end")
  
  GetKeyFromBodyOfHtml = key
End Function

'*******************************************************************************
' GetKeyFromSingleHightWayUse
'   @param textOfSingleHightWayUse [in] text of single hight way use
'   @retval key
'*******************************************************************************
Function GetKeyFromSingleHightWayUse(textOfSingleHightWayUse)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetKeyFromSingleHightWayUse start")
  
  Dim key
  
  Dim inputBodyParts
  inputBodyParts = Split(textOfSingleHightWayUse, DefineCrLf)
  Dim i
  For i = 0 To UBound(inputBodyParts) - 1 Step 1
    If Len(inputBodyParts(i)) >= 2 Then
      inputBodyParts(i) = DeleteHeadTailChar(inputBodyParts(i))
    End If
  Next
  
  Dim secondGateAndDateOfHightWayUse
  Dim dateOfFirstGateHightWayUse
  Dim timeOfFirstGateHightWayUse
  Dim dateOfSecondGateHightWayUse
  Dim timeOfSecondGateHightWayUse
  Dim firstGateOfHightWayUse
  Dim secondGateOfHightWayUse
  Dim tollOfHightWayUse
  Dim tollOfHightWayUseDeleteSpaceAndYenAndConma
  Dim tollOfHightWayUseDeleteSpaceAndYen
  Dim tollOfHightWayUseDeleteSpace
  Dim tollOfHightWayUseParts
  If UBound(inputBodyParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE10 Then
    ' valid format
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & inputBodyParts(NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10) & inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE10))
    secondGateAndDateOfHightWayUse = Split(inputBodyParts(NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10), DEFINE_SPACE & DEFINE_SPACE)
    secondGateOfHightWayUse = secondGateAndDateOfHightWayUse(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST)
    dateOfFirstGateHightWayUse = secondGateAndDateOfHightWayUse(NUMBER_OF_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST)
    timeOfFirstGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_TIME_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10))
    dateOfSecondGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10))
    timeOfSecondGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10))
    firstGateOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10))
    tollOfHightWayUseDeleteSpace = DeleteSpace2MoreSpace(inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE10))
    tollOfHightWayUseDeleteSpaceAndYen = Replace(tollOfHightWayUseDeleteSpace, PRISE_PREFIX_VALUE, DEFINE_BRANK)
    tollOfHightWayUseDeleteSpaceAndYenAndConma = Replace(tollOfHightWayUseDeleteSpaceAndYen, DEFINE_DELIM_CANMA, DEFINE_BRANK)
    tollOfHightWayUseParts = Split(tollOfHightWayUseDeleteSpaceAndYenAndConma, DEFINE_SPACE)
    tollOfHightWayUse = tollOfHightWayUseParts(NUMBER_OF_TOLL_PARTS_IN_TOLL)
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & dateOfSecondGateHightWayUse & DEFINE_SPACE & timeOfSecondGateHightWayUse & DEFINE_SPACE & dateOfFirstGateHightWayUse & DEFINE_SPACE & timeOfFirstGateHightWayUse & DEFINE_SPACE & firstGateOfHightWayUse & DEFINE_SPACE & secondGateOfHightWayUse & DEFINE_SPACE & tollOfHightWayUse)
    
    key = CreateKeyFromHightWayUseInfo(dateOfFirstGateHightWayUse, timeOfFirstGateHightWayUse, dateOfSecondGateHightWayUse, timeOfSecondGateHightWayUse, firstGateOfHightWayUse, secondGateOfHightWayUse, tollOfHightWayUse)
  ElseIf UBound(inputBodyParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE8 Then
    ' valid format
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & inputBodyParts(NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8) & inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE8))
    secondGateAndDateOfHightWayUse = Split(inputBodyParts(NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8), DEFINE_SPACE & DEFINE_SPACE)
    secondGateOfHightWayUse = secondGateAndDateOfHightWayUse(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST)
    dateOfFirstGateHightWayUse = secondGateAndDateOfHightWayUse(NUMBER_OF_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST)
    timeOfFirstGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS_FOR_IE8))
    dateOfSecondGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8))
    timeOfSecondGateHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8))
    firstGateOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8))
    tollOfHightWayUseDeleteSpace = DeleteSpace2MoreSpace(inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE8))
    tollOfHightWayUseDeleteSpaceAndYen = Replace(tollOfHightWayUseDeleteSpace, PRISE_PREFIX_VALUE, DEFINE_BRANK)
    tollOfHightWayUseDeleteSpaceAndYenAndConma = Replace(tollOfHightWayUseDeleteSpaceAndYen, DEFINE_DELIM_CANMA, DEFINE_BRANK)
    tollOfHightWayUseParts = Split(tollOfHightWayUseDeleteSpaceAndYenAndConma, DEFINE_SPACE)
    tollOfHightWayUse = tollOfHightWayUseParts(NUMBER_OF_TOLL_PARTS_IN_TOLL)
    logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & dateOfSecondGateHightWayUse & DEFINE_SPACE & timeOfSecondGateHightWayUse & DEFINE_SPACE & dateOfFirstGateHightWayUse & DEFINE_SPACE & timeOfFirstGateHightWayUse & DEFINE_SPACE & firstGateOfHightWayUse & DEFINE_SPACE & secondGateOfHightWayUse & DEFINE_SPACE & tollOfHightWayUse)
    
    key = CreateKeyFromHightWayUseInfo(dateOfFirstGateHightWayUse, timeOfFirstGateHightWayUse, dateOfSecondGateHightWayUse, timeOfSecondGateHightWayUse, firstGateOfHightWayUse, secondGateOfHightWayUse, tollOfHightWayUse)
  Else
    ' invalid format
    ' skip
    logReturnValueDummy = logOutWarn(LOG_TARGET_LEVEL, "skip, becaouseinvalid format")
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetKeyFromSingleHightWayUse end")
  
  GetKeyFromSingleHightWayUse = key
End Function

'*******************************************************************************
' CreateKeyFromHightWayUseInfo
'   @param dateOfFirstGate [in] date of first gate
'   @param timeOfFirstGate [in] time of first gate
'   @param dateOfSecondGate [in] date of second gate
'   @param timeOfSecondGate [in] time of second gate
'   @param firstGate [in] first gate
'   @param secondGate [in] second gate
'   @param toll [in] toll
'   @retval key
'*******************************************************************************
Function CreateKeyFromHightWayUseInfo(dateOfFirstGate, timeOfFirstGate, dateOfSecondGate, timeOfSecondGate, firstGate, secondGate, toll)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromHightWayUseInfo start")
  
  Dim key
  key = firstGate & DELIM_OF_GATE & secondGate & DELIM_OF_CATEGORY & toll & DELIM_OF_CATEGORY & dateOfFirstGate & DEFINE_SPACE & timeOfFirstGate & DELIM_OF_GATE_TIME & dateOfSecondGate & DEFINE_SPACE & timeOfSecondGate
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromHightWayUseInfo key: " & key)

  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromHightWayUseInfo end")
  
  CreateKeyFromHightWayUseInfo = key
End Function

'*******************************************************************************
' GetHightWayUseInfoFromKey
'   @param key [in] key
'   @retval key
'*******************************************************************************
Function GetHightWayUseInfoFromKey(key)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetHightWayUseInfoFromKey start")
  
  Dim categoryParts
  categoryParts = Split(key, DELIM_OF_CATEGORY)
  Dim gateParts
  gateParts = Split(categoryParts(NUMBER_OF_GATE_AT_KEY), DELIM_OF_GATE)
  Dim dateTimeParts
  dateTimeParts = Split(categoryParts(NUMBER_OF_DATE_TIME_AT_KEY), DEFINE_SPACE)
  
  ' TODO
  ReDim Preserve hightWayUseInfo(NUMBER_OF_SUMMARY_SIZE)
  hightWayUseInfo(NUMBER_OF_FIRST_GATE_AT_SUMMARY) = gateParts(0)
  hightWayUseInfo(NUMBER_OF_SECOND_GATE_AT_SUMMARY) = gateParts(1)
  hightWayUseInfo(NUMBER_OF_TOLL_AT_SUMMARY) = categoryParts(NUMBER_OF_TOLL_AT_KEY)
  hightWayUseInfo(NUMBER_OF_DATE_AT_SUMMARY) = dateTimeParts(0)
  hightWayUseInfo(NUMBER_OF_TIME_AT_SUMMARY) = dateTimeParts(1)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetHightWayUseInfoFromKey end")
  
  GetHightWayUseInfoFromKey = hightWayUseInfo
End Function

'*******************************************************************************
' CountUseInfo
'   @param getDateMode [in] get date mode
'   @param useResult [in] use result
'   @param summaryResult [in/out] summary result
'   @retval nothing
'*******************************************************************************
Function CountUseInfo(getDateMode, useResult, summaryResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfo start")
  
  If getDateMode = 1 Then
    ' "auto 20 day per a month"
    funcDummy = CountUseInfoByAuto20DayPerAMonth(useResult, summaryResult)
  ElseIf getDateMode = 2 Then
    funcDummy = CountUseInfoByAuto20DayPerAMonthAndEachToll(useResult, summaryResult)
  Else
    ' "auto 20 day per a month"
    funcDummy = CountUseInfoByAuto20DayPerAMonth(useResult, summaryResult)
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfo end")
End Function

'*******************************************************************************
' CountUseInfoByAuto20DayPerAMonth
'   @param useResult [in] use result
'   @param summaryResult [in/out] summary result
'   @retval nothing
'*******************************************************************************
Function CountUseInfoByAuto20DayPerAMonth(useResult, summaryResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfoByAuto20DayPerAMonth start")
  
  Dim keys
  keys = useResult.Keys()
  
  For Each key In keys
    Dim useInfos
    useInfos = GetHightWayUseInfoFromKey(key)
    
    Dim key
    key = CreateKeyFromAuto20DayPerAMonth(useInfos)
    
    If summaryResult.Exists(key) = True Then
      ' exist
      Dim useCount
      useCount = summaryResult.Item(key)
      useCount = useCount + 1
      'funcDummy = summaryResult.Add(key, useCount)
      summaryResult.Item(key) = useCount
    Else
      ' don't exist
      Dim firstUseCount
      firstUseCount = 1
      'funcDummy = summaryResult.Add(key, firstUseCount)
      summaryResult.Item(key) = firstUseCount
    End If
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfoByAuto20DayPerAMonth end")
End Function

'*******************************************************************************
' CreateKeyFromHightWayUseInfo
'   @param useInfos [in] use info
'   @retval key
'*******************************************************************************
Function CreateKeyFromAuto20DayPerAMonth(useInfos)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromAuto20DayPerAMonth start")
  
  Dim month
  Dim dateParts
  dateParts = Split(useInfos(NUMBER_OF_DATE_AT_SUMMARY), DELIM_OF_DATE_AT_ETC_SITE)
  
  Dim key
  key = useInfos(NUMBER_OF_FIRST_GATE_AT_SUMMARY) & DELIM_OF_GATE & useInfos(NUMBER_OF_SECOND_GATE_AT_SUMMARY) & DELIM_OF_CATEGORY & useInfos(NUMBER_OF_TOLL_AT_SUMMARY) & DELIM_OF_CATEGORY & dateParts(NUMBER_OF_YEAR_AT_DATE) & DEFINE_SPACE & dateParts(NUMBER_OF_MONTH_AT_DATE)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromAuto20DayPerAMonth end")
  
  CreateKeyFromAuto20DayPerAMonth = key
End Function

'*******************************************************************************
' CountUseInfoByAuto20DayPerAMonthAndEachToll
'   @param useResult [in] use result
'   @param summaryResult [in/out] summary result
'   @retval nothing
'*******************************************************************************
Function CountUseInfoByAuto20DayPerAMonthAndEachToll(useResult, summaryResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfoByAuto20DayPerAMonthAndEachToll start")
  
  Dim keys
  keys = useResult.Keys()
  
  For Each key In keys
    Dim useInfos
    useInfos = GetHightWayUseInfoFromKey(key)
    
    Dim key
    key = CreateKeyFromAuto20DayPerAMonthAndEachToll(useInfos)
    
    If summaryResult.Exists(key) = True Then
      ' exist
      Dim useCount
      useCount = summaryResult.Item(key)
      useCount = useCount + 1
      'funcDummy = summaryResult.Add(key, useCount)
      summaryResult.Item(key) = useCount
    Else
      ' don't exist
      Dim firstUseCount
      firstUseCount = 1
      'funcDummy = summaryResult.Add(key, firstUseCount)
      summaryResult.Item(key) = firstUseCount
    End If
  Next
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CountUseInfoByAuto20DayPerAMonthAndEachToll end")
End Function

'*******************************************************************************
' CreateKeyFromHightWayUseInfoAndEachToll
'   @param useInfos [in] use info
'   @retval key
'*******************************************************************************
Function CreateKeyFromAuto20DayPerAMonthAndEachToll(useInfos)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromHightWayUseInfoAndEachToll start")
  
  Dim month
  Dim dateParts
  dateParts = Split(useInfos(NUMBER_OF_DATE_AT_SUMMARY), DELIM_OF_DATE_AT_ETC_SITE)
  
  Dim key
  key = DEFINE_ASTERISK & DELIM_OF_GATE & DEFINE_ASTERISK & DELIM_OF_CATEGORY & useInfos(NUMBER_OF_TOLL_AT_SUMMARY) & DELIM_OF_CATEGORY & dateParts(NUMBER_OF_YEAR_AT_DATE) & DEFINE_SPACE & dateParts(NUMBER_OF_MONTH_AT_DATE)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromHightWayUseInfoAndEachToll end")
  
  CreateKeyFromAuto20DayPerAMonthAndEachToll = key
End Function

'*******************************************************************************
' SaveSummaryInExcel
'   @param filePath [in] file path
'   @param summaryResult [in] summary
'   @retval key
'*******************************************************************************
Function SaveSummaryInExcel(filePath, summaryResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SaveSummaryInExcel start")
  
  Dim objExcel
  Set objExcel = CreateEXCELObject(IS_SHOW_EXCEL_WINDOW)
  
  ' open
  funcDummy = OpenWorkBooksOfExcel(objExcel, filePath)
  
  ' set
  Dim collOfCell
  collOfCell = 1
  ' explanation
  '   gates
  funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_GATES_CELL, collOfCell, EXPLAIN_OF_GATES_IN_EXCEL)
  '   toll
  funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_TOLL_CELL, collOfCell, EXPLAIN_OF_TOLL_IN_EXCEL)
  '   date
  funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_DATE_CELL, collOfCell, EXPLAIN_OF_DATE_IN_EXCEL)
  '   count
  funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_COUNT_CELL, collOfCell, EXPLAIN_OF_COUNT_IN_EXCEL)
  collOfCell = collOfCell + 1

  Dim valueOfCell
  For Each key In summaryResult
    Dim keyParts
    keyParts = Split(key, DEFINE_DELIM_CANMA)
    
    Dim gates
    Dim toll
    Dim date
    Dim count
    gates = keyParts(0)
    toll = keyParts(1)
    date = keyParts(2)
    count = summaryResult.Item(key)
    
    ' gates
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_GATES_CELL, collOfCell, gates)
    ' toll
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_TOLL_CELL, collOfCell, toll)
    ' date
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_DATE_CELL, collOfCell, date)
    ' count
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_COUNT_CELL, collOfCell, count)
    
    collOfCell = collOfCell + 1
  Next
  
  ' save
  funcDummy = SaveOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SaveSummaryInExcel end")
End Function

'*******************************************************************************
' SaveSummaryToSpecifyCellInExcel
'   @param filePath [in] file path
'   @param summaryResult [in] summary
'   @retval key
'*******************************************************************************
Function SaveSummaryToSpecifyCellInExcel(filePath, summaryResult)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SaveSummaryToSpecifyCellInExcel start")
  
  Dim objExcel
  Set objExcel = CreateEXCELObject(IS_SHOW_EXCEL_WINDOW)
  
  ' open
  funcDummy = OpenWorkBooksOfExcel(objExcel, filePath)
  
  ' set
  Dim collOfCell
  collOfCell = 1
  
  'Dim currentMonthNormalToll
  'Dim currentMonthPlaseDownToll
  'Dim currentMonthNormalCount
  'Dim currentMonthPlaseDownCount
  'Dim lastMonthNormalToll
  'Dim lastMonthPlaseDownToll
  'Dim lastMonthNormalCount
  'Dim lastMonthPlaseDownCount
  'TODO
  
  Dim valueOfCell
  For Each key In summaryResult
    Dim keyParts
    keyParts = Split(key, DEFINE_DELIM_CANMA)
    
    Dim gates
    Dim toll
    Dim date
    Dim count
    gates = keyParts(0)
    toll = keyParts(1)
    date = keyParts(2)
    count = summaryResult.Item(key)
    
    ' gates
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_GATES_CELL, collOfCell, gates)
    ' toll
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_TOLL_CELL, collOfCell, toll)
    ' date
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_DATE_CELL, collOfCell, date)
    ' count
    funcDummy = SetCellsOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK, NUMBER_OF_FIRST_WORKSHEET, ROW_OF_COUNT_CELL, collOfCell, count)
    
    collOfCell = collOfCell + 1
  Next
  
  ' save
  funcDummy = SaveOfExcel(objExcel, NUMBER_OF_FIRST_WORKBOOK)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "SaveSummaryToSpecifyCellInExcel end")
End Function

'*******************************************************************************
' GetRegexArray
'   @param regexStringOfConfig [in] regex config
'   @retval regex array
'*******************************************************************************
Function GetRegexArray(regexStringOfConfig)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetRegexArray start")
  
  Dim regexArray
  regexArray = Split(regexStringOfConfig, DEFINE_DELIM_CANMA)
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "CreateKeyFromAuto20DayPerAMonth end")
  
  GetRegexArray = regexArray
End Function

