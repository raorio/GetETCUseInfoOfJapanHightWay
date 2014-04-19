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
  ' create log
  logCreate()
  
  logDummy = logOutInfo("start program")
  
  ' set vbs timeout
  If vbsTimeoutValue > 0 Then
    WScript.timeout = vbsTimeoutValue
    logDummy = logOutDebug("set vbs timeout: " & vbsTimeoutValue)
  End If
  
  Dim targetPrevYear
  Dim targetPrevMonth
  Dim targetPrevDay
  Dim targetCurrentYear
  Dim targetCurrentMonth
  Dim targetCurrentDay
  targetPeriodHash = GetTargetPeriod()
  
  ' get script file path
  Dim strSaveFilePath
  Dim strScriptPath
  strScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
  strSaveFilePath = strScriptPath & targetCurrentYear & targetCurrentMonth & targetCurrentDay
  CreateFolder(strSaveFilePath)
  CreateFile(strSaveFilePath & DEFINE_DELIM_FOLDER & )
  
  
  ' TODO
  
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
  
  If getTargetMode = "auto 20 day per a month" Then
    resultPeriodHash = GetTargetPeriodByAuto20DayPerAMonth
  'ElseIf getTargetMode = "" Then
  '  resultPeriodHash = GetTargetPeriodByTODO
  Else
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
  resultPeriodHash(PREV_YEAR, targetPrevYear)
  
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


