'===============================================================================
' test
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

Execute ReadVBSFile("..\IncludeConfig.vbs")
Execute ReadVBSFile("..\IncludeCommonConfig.vbs")
Execute ReadVBSFile("..\IncludeAPI.vbs")
Execute ReadVBSFile("..\IncludeCommonAPI.vbs")

' log file check
funcDummy = logFileCheck(LOG_FOLDER, logFilePath)

logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "start test api program")

' set vbs timeout
If VBS_TIMEOUT > 0 Then
  WScript.timeout = vbsTimeoutValue
  logReturnValueDummy = logOutDebug(LOG_TARGET_LEVEL, "set vbs timeout: " & VBS_TIMEOUT)
End If

Dim resultTest


resultTest = TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE8FirstGate_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE8BothGate_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE8BothGateNoSecondDateTime_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1 result:" & resultTest)


resultTest = TestGetKeyFromSingleHightWayUse_DataIE8BothGateAndDiscount_1()
logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1 result:" & resultTest)


logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "end test api program")


'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "" & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & "   14/05/01 " & DefineCRLf & " 05:00 " & DefineCRLf & " è¿ìcâ∫ÇË " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \370  1   "
  Dim expectTest
  expectTest = "è¿ìcâ∫ÇË-,370,14/05/01 05:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE10FirstGate_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "" & DefineCRLf & " 14/04/01 " & DefineCRLf & " 10:00 " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \1,000  1   "
  Dim expectTest
  expectTest = "éuòa-çLìá,1000,14/04/01 10:00-14/04/01 20:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE10BothGate_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "" & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \1,000  1   "
  Dim expectTest
  expectTest = "éuòa-çLìá,1000,14/04/01 20:00-"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE10BothGateNoSecondDateTime_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "" & DefineCRLf & " 14/04/01 " & DefineCRLf & " 10:00 " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & " (\1,000) " & DefineCRLf & " (\-700) " & DefineCRLf & " \300  1  018 "
  Dim expectTest
  expectTest = "éuòa-çLìá,300,14/04/01 10:00-14/04/01 20:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE10BothGateAndDiscount_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE8FirstGate_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE8FirstGate_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "  " & DefineCRLf & "  " & DefineCRLf & "   14/05/01 " & DefineCRLf & " 05:00 " & DefineCRLf & " è¿ìcâ∫ÇË " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \370  1   "
  Dim expectTest
  expectTest = "è¿ìcâ∫ÇË-,370,14/05/01 05:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE8FirstGate_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE8BothGate_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE8BothGate_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = " 14/04/01 " & DefineCRLf & " 10:00 " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \1,000  1   "
  Dim expectTest
  expectTest = "éuòa-çLìá,1000,14/04/01 10:00-14/04/01 20:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE8BothGate_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE8BothGateNoSecondDateTime_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE8BothGateNoSecondDateTime_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = "  " & DefineCRLf & "  " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & "  " & DefineCRLf & "  " & DefineCRLf & " \1,000  1   "
  Dim expectTest
  expectTest = "éuòa-çLìá,1000,14/04/01 20:00-"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE8BothGateNoSecondDateTime_1 = resultTest
End Function

'*******************************************************************************
' TestGetKeyFromSingleHightWayUse_DataIE8BothGateAndDiscount_1
'   @param nothing
'   @retval result test
'*******************************************************************************
Function TestGetKeyFromSingleHightWayUse_DataIE8BothGateAndDiscount_1()
  Dim resultTest
  
  Dim inputTextBody
  inputTextBody = " 14/04/01 " & DefineCRLf & " 10:00 " & DefineCRLf & " éuòa  14/04/01 " & DefineCRLf & " 20:00 " & DefineCRLf & " çLìá " & DefineCRLf & " (\1,000) " & DefineCRLf & " (\-700) " & DefineCRLf & " \300  1  018 "
  Dim expectTest
  expectTest = "éuòa-çLìá,300,14/04/01 10:00-14/04/01 20:00"
  
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "input arg1: " & inputTextBody)
  resultTest = GetKeyFromSingleHightWayUse(inputTextBody)
  logReturnValueDummy = logOutInfo(LOG_TARGET_LEVEL, "output result: " & resultTest)
  If resultTest = expectTest Then
    resultTest = "ok"
  Else
    resultTest = "ng"
  End If
  
  TestGetKeyFromSingleHightWayUse_DataIE8BothGateAndDiscount_1 = resultTest
End Function
