'===============================================================================
' variavle
'===============================================================================
'---------------------------------------
' custom parameter
'---------------------------------------
'-------------------
' application
'-------------------
' application name
Const APPLICATION_NAME = "GetETCUseInfoObJapanHightWay"

' vbs timeout
'   if less then 0, not set timeout
Const VBS_TIMEOUT = -1

' log target level
'   0: fatal
'   1: error
'   2: warn
'   3: info
'   4: debug
'   5: detail debug
Const LOG_TARGET_LEVEL = 4

'-------------------
' date
'-------------------
' mode of auto calc date
'   0: no auto calc date
'   1: auto 20 day per a month
Const MODE_OF_AUTO_CALC_DATE = 1

' year of use from
Dim YEAR_OF_USE_FROM
YEAR_OF_USE_FROM = "2014"

' month of use from
Dim MONTH_OF_USE_FROM
MONTH_OF_USE_FROM = "3"

' day of use from
Dim DAY_OF_USE_FROM
DAY_OF_USE_FROM = "21"

' year of use to
Dim YEAR_OF_USE_TO
YEAR_OF_USE_TO = "2014"

' month of use to
Dim MONTH_OF_USE_TO
MONTH_OF_USE_TO = "4"

' day of use to
Dim DAY_OF_USE_TO
DAY_OF_USE_TO = "20"

'-------------------
' hight way
'-------------------
' in name's of use target gate
'   delimiter is ,
'   support * waild card, which is all(only *, don't support è¿ìc* and *çLìá)
'   ex: IN_NAMES_OF_USE_TARGET_GATE = "è¿ìc,çLìá"
Const IN_NAMES_OF_USE_TARGET_GATE = "*"

' out name's of use target gate
'   delimiter is ,
'   support * waild card, which is all(only *, don't support è¿ìc* and *çLìá)
'   ex: OUT_NAMES_OF_USE_TARGET_GATE = "è¿ìc,çLìá"
Const OUT_NAMES_OF_USE_TARGET_GATE = "*"

' prise's of use target
'   delimiter is ,
'   support * waild card, which is all(only *, don't support 410* and *410)
'   ex: PRISES_OF_USE_TARGET = "410,370"
Const PRISES_OF_USE_TARGET = "*"

' in name's of exclude gate
'   delimiter is ,
'   support * waild card, which is all(only *, don't support éuòa* and *êºïóêVìs)
'   ex: IN_NAMES_OF_EXCLUDE_GATE = "éuòa,êºïóêVìs"
Const IN_NAMES_OF_EXCLUDE_GATE = ""

' out name's of exclude gate
'   delimiter is ,
'   support * waild card, which is all(only *, don't support éuòa* and *êºïóêVìs)
'   ex: OUT_NAMES_OF_EXCLUDE_GATE = "éuòa,êºïóêVìs"
Const OUT_NAMES_OF_EXCLUDE_GATE = ""

' prise's of exclude
'   delimiter is ,
'   support * waild card, which is all(only *, don't support 820* and *640)
'   ex: PRISES_OF_EXCLUDE = "820,640"
Const PRISES_OF_EXCLUDE = ""

'---------------------------------------
' input parameter
'---------------------------------------
' file name of user info
Const FILE_NAME_OF_USER_INFO = "UserInfo.ini"

'---------------------------------------
' view parameter
'---------------------------------------
' sleep time to wait show Web GUI
Const SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI = 500

'---------------------------------------
' etc site parameter
'---------------------------------------
' url of etc site
Const URL_OF_ETC_SITE = "https://www2.etc-user.jp/NASapp/etc/Etc-User?funccode=1011000000&nextfunc=1011100000"

' name of use car number
Const NAME_OF_USE_CAR_NUMBER = "sharyo_no"

' name of use etc card number
Const NAME_OF_USE_ETC_CARD_NUMBER = "iccard_no"

' name of use from year
Const NAME_OF_USE_FROM_YEAR = "riyou_year_from"

' name of use from month
Const NAME_OF_USE_FROM_MONTH = "riyou_month_from"

' name of use from day
Const NAME_OF_USE_FROM_DAY = "riyou_day_from"

' name of use to year
Const NAME_OF_USE_TO_YEAR = "riyou_year_to"

' name of use to month
Const NAME_OF_USE_TO_MONTH = "riyou_month_to"

' name of use to day
Const NAME_OF_USE_TO_DAY = "riyou_day_to"

' prise prefix value
Const PRISE_PREFIX_VALUE = " \"

' prise suffix value
Const PRISE_SUFFIX_VALUE = ""

'---------------------------------------
' excel parameter
'---------------------------------------


'---------------------------------------
' save parameter
'---------------------------------------
'-------------------
' pdf
'-------------------
' is save use context pdf
'   true or false
Const IS_SAVE_USE_CONTEXT_PDF = true

' save prefix of use contex pdf
Const SAVE_PREFIX_OF_USE_CONTEXT_PDF = "etc-pdf-html-file-"

' save suffix of use contex pdf
Const SAVE_SUFFIX_OF_USE_CONTEXT_PDF = ".pdf"

'-------------------
' debug
'-------------------
' is save use context html
'   true or false
Const IS_SAVE_USE_CONTEXT_HTML = false

' save prefix of use context html
Const SAVE_PREFIX_OF_USE_CONTEXT_HTML = "etc-html-file-"

' save suffix of use context html
Const SAVE_SUFFIX_OF_USE_CONTEXT_HTML = ".html"

' is save use context txt
'   true or false
Const IS_SAVE_USE_CONTEXT_TXT = false

' save prefix of use context txt
Const SAVE_PREFIX_OF_USE_CONTEXT_TXT = "etc-file-"

' save suffix of use context txt
Const SAVE_SUFFIX_OF_USE_CONTEXT_TXT = ".txt"

' file name of save sum file
Const FILE_NAME_OF_SAVE_SUM_FILE = "sum-file.log"


