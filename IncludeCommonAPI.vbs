'===============================================================================
' api
'===============================================================================
'-------------------------------------------------------------------------------
' common api
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' log api
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' file api
'-------------------------------------------------------------------------------
'*******************************************************************************
' CreateFolder
'   @param folderPath [in] folder path
'   @retval nothing
'*******************************************************************************
Function CreateFolder(folderPath)
  Dim objFileSys
  Dim objOutFile
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  If objFileSys.FolderExists(folderPath) = false Then
    objFileSys.CreateFolder folderPath
  End If
  
  objOutFile = Nothing
  objFileSys = Nothing
  
  CreateFolder = resultPeriodHash
End Function


'*******************************************************************************
' CreateFile
'   @param filePath [in] file path
'   @retval nothing
'*******************************************************************************
Function CreateFile(filePath)
  Dim objFileSys
  Dim objOutFile
  
  Set objFileSys = WScript.CreateObject("Scripting.FileSystemObject")
  
  If objFileSys.FileExists(filePath) = false Then
    objFileSys.DeleteFile filePath
  End If
  objFileSys.CreateTextFile filePath
  
  objOutFile = Nothing
  objFileSys = Nothing
  
  CreateFolder = resultPeriodHash
End Function


