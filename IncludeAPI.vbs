'===============================================================================
' api
'===============================================================================
'-------------------------------------------------------------------------------
' main api
'-------------------------------------------------------------------------------
'*******************************************************************************
' GetETCUseInfoOfJapanHightWay function
'   @param nothing
'   @retval nothing
'*******************************************************************************
Function GetETCUseInfoOfJapanHightWay()
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetETCUseInfoOfJapanHightWay start")
  
  ' get script file path
  Dim strSaveFilePath
  Dim strScriptPath
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  strSaveFilePath = strScriptPath & targetCurrentYear & targetCurrentMonth & targetCurrentDay
  CreateFolder(strSaveFilePath)
  CreateFile(strSaveFilePath & FILE_NAME_OF_SAVE_SUM_FILE)
  
  Dim mainIEObj
  mainIEObj = CreateIEObject(isDispExecIE, URL_OF_ETC_SITE, webSleepTime)
  
  Dim periodParams
  Set periodParams = GetTargetPeriod(MODE_OF_AUTO_CALC_DATE)
  
  Dim userInfos
  userInfos = ReadUserInfoFile(FILE_NAME_OF_USER_INFO)
  If IsNull(userInfos) = True Then
    logReturnValueDummy = logOutFatal(LOG_TARGET_LEVEL, "please confirm user info file: " & FILE_NAME_OF_USER_INFO)
    WScript.Quit 1
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "user info size: " & UBound(userInfos))
  'For Each userInfo In userInfos
  For indexObUserInfos = 0 To UBound(userInfos) - 1
    Dim userInfo
    userInfo = userInfos(indexObUserInfos)
    If IsNull(userInfo) = True Then
      ' skip
    Else
      Dim carNumber
      Dim icCardNumber
      
      carNumber = userInfo(INDEX_OF_CAR_NUMBER)
      icCardNumber = userInfo(INDEX_OF_ID_CARD_NUMBER)
      
      Set mainIEObj = CreateIEObject(IS_SHOW_MAIN_WEB_GUI, URL_OF_ETC_SITE, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
      funcDummy = SetFormToIE(mainIEObj, periodParams, carNumber, icCardNumber)
      
      ' Enter
      mainIEObj.Document.forms(0).submit
      
      ' wait
      funcDummy = WaitIEObject(mainIEObj, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
      
      ' error check
      ' TODO
      
      Dim currentPage
      currentPage = 1
      Dim sequenceNumber
      sequenceNumber = 0
      Dim isContinue
      isContinue = True
      ' request and parse
      Do Until isContinue = False
        isContinue = False
        isContinue = RequestAndParsePage(mainIEObj, sequenceNumber)
        sequenceNumber = sequenceNumber + 1
        
        Dim objAOfTag
        Set objAOfTag = mainIEObj.Document.getElementsByTagName(NAME_OF_A_NAME)
        Dim indexOfATag
        Dim hrefName
        For indexOfATag = 0 To objAOfTag.Length - 1
          hrefName = objAOfTag(indexOfATag).getAttribute(NAME_OF_ATTRIBUTE_HREF)
          If hrefName <> DEFINE_BRANK Then
            Dim targetPage
            targetPage = -1
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
              logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "target page is: " & hrefName)
              currentPage = targetPage
              objAOfTag(indexOfATag).Click
              indexOfATag = objAOfTag.Length
              
              funcDummy = WaitIEObject(mainIEObj, SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI)
              
              isContinue = True
              
              Exit For
            End If
          End If
        Next
      Loop
      
      Set mainIEObj = Nothing
    End If
  Next
  
  ' TODO
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetETCUseInfoOfJapanHightWay end")
  
  'GetETCUseInfoOfJapanHightWay = TODO
End Function


'-------------------------------------------------------------------------------
' other api
'-------------------------------------------------------------------------------
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
  'ReadUserInfoFile = resultUserInfos
End Function

'*******************************************************************************
' get target period
'   @param getMode [in] get mode
'   @retval resultPeriodHash result period hash
'*******************************************************************************
Function GetTargetPeriod(getMode)
  Dim getPeriodHash
  Dim resultPeriodHash
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "GetTargetPeriod start")
  
  If getMode = 1 Then
    ' "auto 20 day per a month"
    Set getPeriodHash = GetTargetPeriodByAuto20DayPerAMonth
  'ElseIf getTargetMode = "" Then
  '  Set getPeriodHash = GetTargetPeriodByTODO
  Else
    ' "auto 20 day per a month"
    Set getPeriodHash = GetTargetPeriodByAuto20DayPerAMonth
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
  
  Set resultPeriodHash = CreateObject("Scripting.Dictionary")
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
  
  Set resultPeriodHash = CreateObject("Scripting.Dictionary")
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
'   @retval true/false true:continue, false:
'*******************************************************************************
Function RequestAndParsePage(objIE, sequenceNumber)
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
        If IsCheckHightWayUse(objInputTags(indexOfInputTag)) = True Then
          ' target hight way
          ' if check by Click(), don't detect checked. there for check by SetAttribute
          'objInputTags(indexOfInputTag).Click()
          funcDummy = objInputTags(indexOfInputTag).SetAttribute(NAME_OF_CHECKED, NAME_OF_CHECKED_VALUE)
        Else
          ' not target hight way
          ' skip
        End If
      End If
    Next
  End If
  
  If IS_CONFORM_BEFORE_HIGHT_WAY_USE_DETERM = true Then
    MsgBox "料金計算発行を実施します。" & DefineCrLr & "自動チェック内容を確認し、必要があればチェック操作してください。" & DefineCrLf & "確認完了後、「OK」を押してください。"
  End If
  
  Dim bodyOfHtml
  bodyOfHtml = objIE.Document.body.InnerHtml
  
  funcDummy = ParseBodyOfHtml(bodyOfHtml, objIE)
  
  If IS_SAVE_USE_CONTEXT_PDF = True Then
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
' IsTargetHightWayUse
'   @param objElement [in] objElement
'   @retval true/false true:target, false:not target
'*******************************************************************************
Function IsCheckHightWayUse(objElement)
  Dim result
  result = False
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsCheckHightWayUse start")
  
  Dim targetHightWayUse
  targetHightWayUse = objElement.parentNode.parentNode.innerText
  
  Dim targetHightWayUseParts
  targetHightWayUseParts = Split(targetHightWayUse, DefineCrLf)
  If UBound(targetHightWayUseParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS Then
    ' Check
    ' TODO
    
    result = True
  End If
  
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "IsCheckHightWayUse end")
  
  IsCheckHightWayUse = result
End Function

'*******************************************************************************
' ParseBodyOfHtml
'   @param bodyOfHtml [in] body of html
'   @param objIE [in] object IE
'   @retval 
'*******************************************************************************
Function ParseBodyOfHtml(bodyOfHtml, objIE)
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ParseBodyOfHtml start")
  
  Dim objInputTags
  Set objInputTags = objIE.Document.getElementsByTagName(NAME_OF_INPUT)
  Dim indexOfInput
  For indexOfInput = 0 To objInputTags.Length - 1
logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ParseBodyOfHtml index: " & indexOfInput)
    Dim typeName
    typeName = objInputTags(indexOfInput).getAttribute(NAME_OF_ATTRIBUTE_TYPE)
    If typeName = NAME_OF_CHECK_BOX Then
      Dim isCheckedAttribute
      isCheckedAttribute = objInputTags(indexOfInput).getAttribute(NAME_OF_CHECKED)
      ' if detect by true/false, don't detect checked. there for check by checked value(brank)
      If isCheckedAttribute = NAME_OF_CHECKED_VALUE Then
      'If isCheckedAttribute = true Then
      'If isCheckedAttribute <> DEFINE_BRANK Then
        ' checked
        Dim inputBody
        inputBody = objInputTags(indexOfInput).parentNode.parentNode.innerText
logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "ParseBodyOfHtml body: " & inputBody)
        Dim inputBodyParts
        inputBodyParts = Split(inputBody, DefineCrLf)
        If UBound(inputBodyParts) = NUMBER_OF_HIGHT_WAY_USE_PARTS Then
          ' valid format
          logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & inputBodyParts(NUMBER_OF_DATE_HIGHT_WAY_USE_PARTS) & inputBodyParts(NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS) & inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS) & inputBodyParts(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS) & inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS))
          Dim dateOfHightWayUse
          Dim timeOfHightWayUse
          Dim firstGateOfHightWayUse
          Dim secondGateOfHightWayUse
          Dim tollOfHightWayUse
          dateOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_DATE_HIGHT_WAY_USE_PARTS))
          timeOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS))
          firstGateOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS))
          secondGateOfHightWayUse = DeleteSpace(inputBodyParts(NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS))
          Dim tollOfHightWayUseDeleteSpace
          tollOfHightWayUseDeleteSpace = DeleteSpace2MoreSpace(inputBodyParts(NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS))
          Dim tollOfHightWayUseDeleteSpaceAndYen
          tollOfHightWayUseDeleteSpaceAndYen = Replace(tollOfHightWayUseDeleteSpace, PRISE_PREFIX_VALUE, DEFINE_BRANK)
          Dim tollOfHightWayUseParts
          tollOfHightWayUseParts = Split(tollOfHightWayUseDeleteSpaceAndYen, DEFINE_SPACE)
          tollOfHightWayUse = tollOfHightWayUseParts(NUMBER_OF_TOLL_PARTS_IN_TOLL)
          logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "one of hight way use info: " & dateOfHightWayUse & DEFINE_SPACE & timeOfHightWayUse & DEFINE_SPACE & firstGateOfHightWayUse & DEFINE_SPACE & secondGateOfHightWayUse & DEFINE_SPACE & tollOfHightWayUse)
          
          ' TODO
        Else
          ' invalid format
          ' skip
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
  
  'ParseBodyOfHtml = 
End Function

