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
Const LOG_TARGET_LEVEL = 3
'Const LOG_TARGET_LEVEL = 4
'Const LOG_TARGET_LEVEL = 5

'-------------------
' date
'-------------------
' mode of auto calc date
'   0: no auto calc date
'   1: auto 20 day per a month
'   2: auto 20 day per a month(and each toll)
Const MODE_OF_AUTO_CALC_DATE_NO = 0
Const MODE_OF_AUTO_CALC_DATE_AUTO_20DAY_PER_MONTH = 1
Const MODE_OF_AUTO_CALC_DATE_AUTO_20DAY_PER_MONTH_EACH_TALL = 2
Dim MODE_OF_AUTO_CALC_DATE
MODE_OF_AUTO_CALC_DATE = MODE_OF_AUTO_CALC_DATE_AUTO_20DAY_PER_MONTH
'MODE_OF_AUTO_CALC_DATE = MODE_OF_AUTO_CALC_DATE_AUTO_20DAY_PER_MONTH_EACH_TALL

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
' name's of use target gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: NAMES_OF_USE_TARGET_GATE = "è¿ìc,çLìá"
Const NAMES_OF_USE_TARGET_GATE = "è¿ìcè„ÇË,è¿ìcâ∫ÇË"

' first name of use target gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: FIRST_NAME_OF_USE_TARGET_GATE = "è¿ìc,çLìá"
Const FIRST_NAME_OF_USE_TARGET_GATE = "è¿ìcè„ÇË,è¿ìcâ∫ÇË"

' second name of use target gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: SECOND_NAME_OF_USE_TARGET_GATE = "è¿ìc,çLìá"
Const SECOND_NAME_OF_USE_TARGET_GATE = "è¿ìcè„ÇË,è¿ìcâ∫ÇË"

' toll of use target
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: TOLL_OF_USE_TARGET = "410,370,310,280"
Const TOLL_OF_USE_TARGET = ""

' name's of exclude gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: NAMES_OF_USE_EXCLUDE_GATE = "éuòa-êºïóêVìs,éuòa-çLìá,êºïóêVìs-çLìá"
Const NAMES_OF_USE_EXCLUDE_GATE = ""

' first name of exclude gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: FIRST_NAME_OF_EXCLUDE_GATE = "éuòa,êºïóêVìs,çLìá"
Const FIRST_NAME_OF_EXCLUDE_GATE = ""

' second name of exclude gate
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: SECOND_NAME_OF_EXCLUDE_GATE = "éuòa,êºïóêVìs,çLìá"
Const SECOND_NAME_OF_EXCLUDE_GATE = ""

' toll of exclude
'   delimiter is ,
'   support regex(http://msdn.microsoft.com/ja-jp/library/ms974570.aspx)
'   ex: TOLL_OF_EXCLUDE = "820,640"
Const TOLL_OF_EXCLUDE = ""

'-------------------
' user info
'-------------------
' index of car number
Const INDEX_OF_CAR_NUMBER = 0

' index of ic card number
Const INDEX_OF_ID_CARD_NUMBER = 1

' index of other info
Const INDEX_OF_OTHER_INFO = 2

' size of user info index
Const SIZE_OF_USER_INFO_INDEX = 3

'---------------------------------------
' input parameter
'---------------------------------------
' file name of user info
Const FILE_NAME_OF_USER_INFO = "UserInfo.ini"

' proxy server(if not use, brank)
Const PROXY_SERVER = ""

'---------------------------------------
' view parameter
'---------------------------------------
' sleep time to wait show Web GUI
Const SLEEP_TIME_TO_WAIT_SHOW_WEB_GUI = 500

' show main Web GUI
Const IS_SHOW_MAIN_WEB_GUI = true

' is conform before hight way use determ
Const IS_CONFORM_BEFORE_HIGHT_WAY_USE_DETERM = true

'---------------------------------------
' etc site parameter
'---------------------------------------
' url of etc site
Const URL_OF_ETC_SITE = "https://www2.etc-user.jp/NASApp/etc/Etc-User?funccode=1011000000&nextfunc=1011100000"

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

' name of input
Const NAME_OF_INPUT = "INPUT"

' name of input check box type
Const NAME_OF_CHECK_BOX = "checkbox"

' name of attribute type
Const NAME_OF_ATTRIBUTE_TYPE = "type"

' name of attribute name
Const NAME_OF_ATTRIBUTE_NAME = "name"

' name of A name
Const NAME_OF_A_NAME = "A"

' name of attribute href
Const NAME_OF_ATTRIBUTE_HREF = "href"

' name of link page
Const NAME_OF_LINK_PAGE = "&page="

' number of IE 10 version
Const NUMBER_OF_IE10_VERSION = 10

' number of IE 8 version
Const NUMBER_OF_IE8_VERSION = 8

' number of hight way use parts for IE 10
Const NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE10 = 8

' number of date in hight way use parts for IE 10
Const NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 1

' number of time in hight way use parts for IE 10
Const NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 2

' number of second gate and date in hight way use parts for IE 10
Const NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 3

' number of time in hight way use parts for IE 10
Const NUMBER_OF_TIME_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 4

' number of first gate in hight way use parts for IE 10
Const NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 5

' number of second gate in hight way use parts for IE 10
Const NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE10 = 6

' number of toll in hight way use parts for IE 10
Const NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE10 = 8

' number of hight way use parts for IE 8
Const NUMBER_OF_HIGHT_WAY_USE_PARTS_FOR_IE8 = 7

' number of date in hight way use parts for IE 8
Const NUMBER_OF_DATE_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8 = 0

' number of time in hight way use parts for IE 8
Const NUMBER_OF_TIME_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8 = 1

' number of second gate and date first gate in hight way use parts for IE 8
Const NUMBER_OF_SECOND_GATE_AND_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8 = 2

' number of time in hight way use parts for IE 8
Const NUMBER_OF_TIME_HIGHT_WAY_USE_PARTS_FOR_IE8 = 3

' number of first gate in hight way use parts for IE 8
Const NUMBER_OF_FIRST_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8 = 4

' number of second gate in hight way use parts for IE 8
Const NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_FOR_IE8 = 5

' number of toll in hight way use parts for IE 8
Const NUMBER_OF_TOLL_HIGHT_WAY_USE_PARTS_FOR_IE8 = 7

' number of second gate hight way use parts when second gate and first gate
Const NUMBER_OF_SECOND_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST = 0

' number of date first gate hight way use parts when second gate and first gate
Const NUMBER_OF_DATE_FIRST_GATE_HIGHT_WAY_USE_PARTS_WHEN_SECOND_GATE_AND_FIRST = 1

' name of checked
Const NAME_OF_CHECKED = "CHECKED"

' name of checked value
Const NAME_OF_CHECKED_VALUE = "CHECKED_VALUE"

' number of toll in toll parts
Const NUMBER_OF_TOLL_PARTS_IN_TOLL = 0


'---------------------------------------
' summary parameter
'---------------------------------------
' delim of category
Const DELIM_OF_CATEGORY = ","

' delim of gate
Const DELIM_OF_GATE = "-"

' delim of date time
Const DELIM_OF_GATE_TIME = "-"

' delim of key and value in param's
Const DELIM_OF_KEY_AND_VALUE_IN_PARAMS = "==="

' delim of entry(key and value)'s in param's
Const DELIM_OF_ENTRY_IN_PARAMS = "<<<>>>"

' number of key size
Const NUMBER_OF_KEY_SIZE = 4

' number of gate at key
Const NUMBER_OF_GATE_AT_KEY = 0

' number of toll at key
Const NUMBER_OF_TOLL_AT_KEY = 1

' number of date time at key
Const NUMBER_OF_DATE_TIME_AT_KEY = 2

' number of param's at key
Const NUMBER_OF_PARAMS_AT_KEY = 3

' number of gate size
Const NUMBER_OF_GATE_SIZE = 2

' number of first gate at gate
Const NUMBER_OF_FIRST_GATE_AT_GATE = 0

' number of second gate at gate
Const NUMBER_OF_SECOND_GATE_AT_GATE = 1

' number of date time size
Const NUMBER_OF_DATE_TIME_SIZE = 2

' number of date at date time
Const NUMBER_OF_DATE_AT_DATE_TIME = 0

' number of time at date time
Const NUMBER_OF_TIME_AT_DATE_TIME = 1

' number of date siez
Const NUMBER_OF_DATE_SIZE = 3

' number of year at date
Const NUMBER_OF_YEAR_AT_DATE = 0

' number of month at date
Const NUMBER_OF_MONTH_AT_DATE = 1

' number of day at date
Const NUMBER_OF_DAY_AT_DATE = 2

' number of summary size
Const NUMBER_OF_SUMMARY_SIZE = 6

' number of first gate at summary
Const NUMBER_OF_FIRST_GATE_AT_SUMMARY = 0

' number of second gate at summary
Const NUMBER_OF_SECOND_GATE_AT_SUMMARY = 1

' number of toll at summary
Const NUMBER_OF_TOLL_AT_SUMMARY = 2

' number of date at summary
Const NUMBER_OF_DATE_AT_SUMMARY = 3

' number of time at summary
Const NUMBER_OF_TIME_AT_SUMMARY = 4

' number of param's at summary
Const NUMBER_OF_PARAMS_AT_SUMMARY = 5

' delim of date at etc site
Const DELIM_OF_DATE_AT_ETC_SITE = "/"

' explain of gates in summary
Const EXPLAIN_OF_GATES_IN_SUMMARY = "gates"

' explain of toll in summary
Const EXPLAIN_OF_TOLL_IN_SUMMARY = "toll"

' explain of date in summary
Const EXPLAIN_OF_DATE_IN_SUMMARY = "date"

' explain of count in summary
Const EXPLAIN_OF_COUNT_IN_SUMMARY = "count"

' key of discount in param's
Const KEY_OF_DISCOUNT_IN_PARAMS = "DISCOUNT"


'---------------------------------------
' excel parameter
'---------------------------------------
' is save excel
Const IS_SAVE_EXCEL = true

' mode of save excel
'   0: list
'   1: specify cell
Const MODE_OF_SAVE_EXCEL_LIST = 0
Const MODE_OF_SAVE_EXCEL_SPECIFY_CELL = 1
Dim MODE_OF_SAVE_EXCEL
MODE_OF_SAVE_EXCEL = MODE_OF_SAVE_EXCEL_LIST
'MODE_OF_SAVE_EXCEL = MODE_OF_SAVE_EXCEL_SPECIFY_CELL

' config of mode list
' file name of excel
Const FILE_NAME_OF_EXCEL = "UseSummary.xlsx"

' is show excel window
Const IS_SHOW_EXCEL_WINDOW = true

' number of first workbook
Const NUMBER_OF_FIRST_WORKBOOK = 1

' number of first worksheet
Const NUMBER_OF_FIRST_WORKSHEET = 1

' row of gates cell
Const ROW_OF_GATES_CELL = 1

' row of toll cell
Const ROW_OF_TOLL_CELL = 2

' row of date cell
Const ROW_OF_DATE_CELL = 3

' row of count cell
Const ROW_OF_COUNT_CELL = 4

' explain of gates in excel
Const EXPLAIN_OF_GATES_IN_EXCEL = "gates"

' explain of toll in excel
Const EXPLAIN_OF_TOLL_IN_EXCEL = "toll"

' explain of date in excel
Const EXPLAIN_OF_DATE_IN_EXCEL = "date"

' explain of count in excel
Const EXPLAIN_OF_COUNT_IN_EXCEL = "count"

' config of mode specify cell
' file name of excel mode specify cell
Const FILE_NAME_OF_EXCEL_MODE_SPECIFY_CELL = "UseSummary.xlsx"

' is show excel window mode specify cell
Const IS_SHOW_EXCEL_WINDOW_MODE_SPECIFY_CELL = true

' number of first workbook mode specify cell
Const NUMBER_OF_FIRST_WORKBOOK_MODE_SPECIFY_CELL = 1

' number of first worksheet mode specify cell
Const NUMBER_OF_FIRST_WORKSHEET_MODE_SPECIFY_CELL = 1

' row of count normal price last 1 month cell mode specify cell
Const ROW_OF_COUNT_NORMAL_PRICE_LAST_1_MONTH_CELL_MODE_SPECIFY_CELL = 4

' column of count normal price last 1 month cell mode specify cell
Const COLUMN_OF_COUNT_NORMAL_PRICE_LAST_1_MONTH_CELL_MODE_SPECIFY_CELL = 4

' row of count discount price last 1 month cell mode specify cell
Const ROW_OF_COUNT_DISCOUNT_PRICE_LAST_1_MONTH_CELL_MODE_SPECIFY_CELL = 4

' column of count discount price last 1 month cell mode specify cell
Const COLUMN_OF_COUNT_DISCOUNT_PRICE_LAST_1_MONTH_CELL_MODE_SPECIFY_CELL = 5

' row of count normal price last 2 month cell mode specify cell
Const ROW_OF_COUNT_NORMAL_PRICE_LAST_2_MONTH_CELL_MODE_SPECIFY_CELL = 4

' column of count normal price last 2 month cell mode specify cell
Const COLUMN_OF_COUNT_NORMAL_PRICE_LAST_2_MONTH_CELL_MODE_SPECIFY_CELL = 6

' row of count discount price last 2 month cell mode specify cell
Const ROW_OF_COUNT_DISCOUNT_PRICE_LAST_2_MONTH_CELL_MODE_SPECIFY_CELL = 4

' column of count discount price last 2 month cell mode specify cell
Const COLUMN_OF_COUNT_DISCOUNT_PRICE_LAST_2_MONTH_CELL_MODE_SPECIFY_CELL = 7


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
' concat pdf
'-------------------
' is do concat pdf
Const IS_DO_CONCAT_PDF = True


'-------------------
' text
'-------------------
' is save sum file
Const IS_SAVE_SUM_FILE = true

' file name of save sum file
Const FILE_NAME_OF_SAVE_SUM_FILE = "sum-file.log"


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


