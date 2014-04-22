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
  Dim targetPrevYear
  Dim targetPrevMonth
  Dim targetPrevDay
  Dim targetCurrentYear
  Dim targetCurrentMonth
  Dim targetCurrentDay
  targetPeriodHash = GetTargetPeriod()
  
  logDummy = logOutDebug("GetETCUseInfoOfJapanHightWay start")
  
  ' get script file path
  Dim strSaveFilePath
  Dim strScriptPath
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  strSaveFilePath = strScriptPath & targetCurrentYear & targetCurrentMonth & targetCurrentDay
  CreateFolder(strSaveFilePath)
  CreateFile(strSaveFilePath & DEFINE_DELIM_FOLDER & saveSumFile)
  
  Dim mainIEObj
  mainIEObj = CreateIEObject(isDispExecIE, URL_OF_ETC_SITE, webSleepTime)
  
  Dim periodParams
  periodParams = GetTargetPeriod()
  
  Dim carNumber
  Dim icCardNumber
  
  ' TODO ファイルから番号を取得し、繰り返す
  
  funcDummy = SetFormToIE(mainIEObj, periodParams, carNumber, icCardNumber)
  
  ' TODO
  
  logDummy = logOutDebug("GetETCUseInfoOfJapanHightWay end")
  
  'GetETCUseInfoOfJapanHightWay = TODO
End Function


'-------------------------------------------------------------------------------
' other api
'-------------------------------------------------------------------------------
'*******************************************************************************
' get target period
'   @param nothing
'   @retval resultPeriodHash result period hash
'*******************************************************************************
Function GetTargetPeriod()
  Dim resultPeriodHash
  
  If MODE_OF_AUTO_CALC_DATE = 1 Then
    ' "auto 20 day per a month"
    resultPeriodHash = GetTargetPeriodByAuto20DayPerAMonth
  'ElseIf getTargetMode = "" Then
  '  resultPeriodHash = GetTargetPeriodByTODO
  Else
    ' "auto 20 day per a month"
    resultPeriodHash = GetTargetPeriodByAuto20DayPerAMonth
  End If

  GetTargetPeriod = resultPeriodHash
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
  
  currentMonth = Month(Now)
  currentDay = Day(Now)
  
  targetPrevYear = Year(Now)
  targetCurrentYear = Year(Now)
  
  ' detect month
  If currentMonth > 20 Then
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
  
  ' TODO
  funcDummy = resultPeriodHash(PREV_YEAR, targetPrevYear)
  
  currentMonth = Nothing
  currentDay = Nothing
  targetPrevYear = Nothing
  targetPrevMonth = Nothing
  targetPrevDay = Nothing
  targetCurrentYear = Nothing
  targetCurrentMonth = Nothing
  targetCurrentDay = Nothing
  
  GetTargetPeriod = resultPeriodHash
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
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_CAR_NUMBER)
    WScript.Quit 1
  End If
  If objICCardNumber.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_ETC_CARD_NUMBER)
    WScript.Quit 1
  End If
  If objFromYear.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_FROM_YEAR)
    WScript.Quit 1
  End If
  If objFromMonth.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_FROM_MONTH)
    WScript.Quit 1
  End If
  If objFromDay.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_FROM_DAY)
    WScript.Quit 1
  End If
  If objToYear.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_TO_YEAR)
    WScript.Quit 1
  End If
  If objToMonth.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_TO_MONTH)
    WScript.Quit 1
  End If
  If objToDay.Length = 0 Then
    logDummy = logOutFatal(errorMessage & NAME_OF_USE_TO_DAY)
    WScript.Quit 1
  End If
  
  objCarNumber(0).Value = carNumber
  Set objCarNumber = Nothing
  objICCardNumber(0).Value = carNumber
  Set objICCardNumber = Nothing
  objFromYear(0).Value = carNumber
  Set objFromYear = Nothing
  objFromMonth(0).Value = carNumber
  Set objFromMonth = Nothing
  objFromDay(0).Value = carNumber
  Set objFromDay = Nothing
  objToYear(0).Value = carNumber
  Set objToYear = Nothing
  objToMonth(0).Value = carNumber
  Set objToMonth = Nothing
  objToDay(0).Value = carNumber
  Set objToDay = Nothing
  
  'SetFormToIE = 
End Function

