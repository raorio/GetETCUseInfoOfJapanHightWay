'===============================================================================
' main
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
Dim FileShell
Set FileShell = Wscript.CreateObject('Scripting FileSystemObject')

Const ForReadingInclude = 1


'*******************************************************************************
' read vbs file
'   @param FileName [in] read vbs file name
'   @retval nothing
'*******************************************************************************
Function ReadVBSFile(ByVal FileName)
  ReadFile = FileShell.OpenTextFile(FileName, ForReadingInclude, False).ReadAll()
End Function


Execute ReadVBSFile('IncludeConfig.vbs')
Execute ReadVBSFile('IncludeCommonConfig.vbs')
Execute ReadVBSFile('IncludeAPI')
Execute ReadVBSFile('IncludeCommonAPI')


GetETCInfoOfJapanHightWay()

