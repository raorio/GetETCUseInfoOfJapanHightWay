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
Const DEFINE_CRLF = vbcrlf

' space
Const DEFINE_SPACE = " "

' hyphen
Const DEFINE_HYPHEN = "-"

' colon
Const DEFINE_COLON = ":"

' brank
Const DEFINE_BRANK = ""

' delim folder
Const DEFINE_DELIM_FOLDER = "\"

' delim date time
Const DEFINE_DELIM_DATE_TIME = "T"

'---------------------------------------
' date time
'---------------------------------------
' yyyy-mm-dd hh-mm-ss date time formate
Dim strDateTimeSystemTime
strLogDateTime = Year(Now) & DEFINE_DELIM_DATE & Month(Now) & DEFINE_DELIM_DATE & Day(Now) & DEFINE_DELIM_TIME & Hour(Now) & DEFINE_DELIM_TIME & Minute(Now) & DEFINE_DELIM_TIME & Second(Now)

' yyyymmddThhmmss date time formate
Dim strDateTimeISO
strLogDateTime = Year(Now) & Month(Now) & Day(Now) & DEFINE_DELIM_DATE_TIME & Hour(Now) & Minute(Now) & Second(Now)

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

' log level strings
Dim logLevelStrings(6)
logLevelStrings(0) = LOG_LEVEL_FATAL
logLevelStrings(1) = LOG_LEVEL_ERROR
logLevelStrings(2) = LOG_LEVEL_WARN
logLevelStrings(3) = LOG_LEVEL_INFO
logLevelStrings(4) = LOG_LEVEL_DEBUB
logLevelStrings(5) = LOG_LEVEL_DETAIL_DEBUG

