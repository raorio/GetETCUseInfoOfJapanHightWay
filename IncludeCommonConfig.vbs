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

' func dummy
Dim funcDummy

'---------------------------------------
' date time
'---------------------------------------
' yyyy-mm-dd hh-mm-ss date time formate
Dim strDateTimeSystemTime
strDateTimeSystemTime = Year(Now) & DEFINE_DELIM_DATE & Month(Now) & DEFINE_DELIM_DATE & Day(Now) & DEFINE_DELIM_DATE_TIME & Hour(Now) & DEFINE_DELIM_TIME & Minute(Now) & DEFINE_DELIM_TIME & Second(Now)

' yyyymmddThhmmss date time formate
Dim strDateTimeISO
strDateTimeISO = Year(Now) & Month(Now) & Day(Now) & DEFINE_DELIM_ISO_DATE_TIME & Hour(Now) & Minute(Now) & Second(Now)

' log date time
strLogDateTime = strDateTimeSystemTime

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
' name of IE application
Const NAME_OF_IE_APPLICATION = "InternetExplorer.Application"

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

