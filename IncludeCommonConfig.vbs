'===============================================================================
' variavle
'===============================================================================
'-------------------------------------------------------------------------------
' common parameter
'-------------------------------------------------------------------------------
'---------------------------------------
' define common string
'---------------------------------------
' crlf
Dim Define_CrLf
DefineCrLf = vbCrLf

' space
Const DEFINE_SPACE = " "

' hyphen
Const DEFINE_HYPHEN = "-"

' squal
Const DEFINE_EQUAL = "="

' hyphen
Const DEFINE_SLASH = "/"

' colon
Const DEFINE_COLON = ":"

' brank
Const DEFINE_BRANK = ""

' single qulote
Const DEFINE_SINGLE_QUOTE = "'"

' double qulote
Const DEFINE_DOUBLE_QUOTE = """"

' delim folder
Const DEFINE_DELIM_FOLDER = "\"

' delim date time
Const DEFINE_DELIM_ISO_DATE_TIME = "T"

' delim date time
Const DEFINE_DELIM_DATE_TIME = " "

' delim date
Const DEFINE_DELIM_DATE = "/"

' delim time
Const DEFINE_DELIM_TIME = ":"

' delim CONMA
Const DEFINE_DELIM_CANMA = ","

' func dummy
Dim funcDummy

'---------------------------------------
' date time
'---------------------------------------


'---------------------------------------
' file system
'---------------------------------------
' file open option reading
Const ForReading = 1

' file open option writing
Const ForWriting = 2

' file open option appending
Const ForAppending = 8

'---------------------------------------
' object
'---------------------------------------
' name of scripting file system object
Const NAME_OF_SCRIPTING_FILESYSTEMOBJECT = "Scripting.FileSystemObject"

' name of scripting dictionary
Const NAME_OF_SCRIPTING_DICTIONARY = "Scripting.Dictionary"

' name of IE application
Const NAME_OF_IE_APPLICATION = "InternetExplorer.Application"

' name of shell application
Const NAME_OF_SHELL_APPLICATION = "Shell.Application"

' name of excel application
Const NAME_OF_EXCEL_APPLICATION = "Excel.Application"

' name of wscript shell
Const NAME_OF_WSCRIPT_SHELL = "WScript.Shell"

'-------------------
' http object
'-------------------
' MSXML2.SERVERXMLHTTP.4.0
Const MSXML2_SERVERXMLHTTP_4_0 = "MSXML2.SERVERXMLHTTP.4.0"

' MSXML2.XMLHTTP.3.0
Const MSXML2_XMLHTTP_3_0 = "MSXML2.XMLHTTP.3.0"

' MSXML.XMLHTTPRequest
Const MSXML_XMLHTTPREQUEST = "MSXML.XMLHTTPRequest"

' Microsoft.XMLHTTP
Const MICROSOFT_XMLHTTP = "Microsoft.XMLHTTP"

' http object list
Dim httpObjectList(4)
httpObjectList(0) = MSXML2_SERVERXMLHTTP_4_0
httpObjectList(1) = MSXML2_XMLHTTP_3_0
httpObjectList(2) = MSXML_XMLHTTPREQUEST
httpObjectList(3) = MICROSOFT_XMLHTTP

' ADOBJ.Stream
Const ADODB_STREAM = "ADODB.Stream"

'---------------------------------------
' log
'---------------------------------------
' log time
Dim strLogDateTime
strLogDateTime = strLogDateTime

' log folder
Const LOG_FOLDER = "log"

' log file name
Const LOG_FILE_NAME = "logfile.log"

' log file path
Dim logFilePath
logFilePath = LOG_FOLDER & DEFINE_DELIM_FOLDER & LOG_FILE_NAME

' log return value dummy
Dim logReturnValueDummy

'-------------------
' log level string
'-------------------
' log level string fatal
Const LOG_LEVEL_FATAL = "Fatal"
' log level string error
Const LOG_LEVEL_ERROR = "Error"
' log level string warn
Const LOG_LEVEL_WARN = "Warn"
' log level string info
Const LOG_LEVEL_INFO = "Info"
' log level string debug
Const LOG_LEVEL_DEBUB = "Debug"
' log level string detail debug
Const LOG_LEVEL_DETAIL_DEBUG = "DetailDebug"

' log level number
Const LOG_LEVEL_NUMBER_FATAL = 0
Const LOG_LEVEL_NUMBER_ERROR = 1
Const LOG_LEVEL_NUMBER_WARN = 2
Const LOG_LEVEL_NUMBER_INFO = 3
Const LOG_LEVEL_NUMBER_DEBUG = 4
Const LOG_LEVEL_NUMBER_DETAIL_DEBUG = 5

' log level strings
Dim logLevelStrings(6)
logLevelStrings(LOG_LEVEL_NUMBER_FATAL) = LOG_LEVEL_FATAL
logLevelStrings(LOG_LEVEL_NUMBER_ERROR) = LOG_LEVEL_ERROR
logLevelStrings(LOG_LEVEL_NUMBER_WARN) = LOG_LEVEL_WARN
logLevelStrings(LOG_LEVEL_NUMBER_INFO) = LOG_LEVEL_INFO
logLevelStrings(LOG_LEVEL_NUMBER_DEBUG) = LOG_LEVEL_DEBUB
logLevelStrings(LOG_LEVEL_NUMBER_DETAIL_DEBUG) = LOG_LEVEL_DETAIL_DEBUG

