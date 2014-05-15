GetETCUseInfoOfJapanHightWay
============================


GetETCUseInfoOfJapanHightWay


setup
  specify below environment in file IncludeConfig.vbs
    if use proxy server
      PROXY_SERVER
    date
      MODE_OF_AUTO_CALC_DATE
      if MODE_OF_AUTO_CALC_DATE is 0
        YEAR_OF_USE_FROM
        MONTH_OF_USE_FROM
        DAY_OF_USE_FROM
        YEAR_OF_USE_TO
        MONTH_OF_USE_TO
        DAY_OF_USE_TO

  specify user info in file UserInfo.ini

  setup pop up block from ie option, below 
    インターネットオプション -> プライバシー -> ポップアップブロック -> 設定 -> 許可するWebサイトのアドレス(W)
      *.etc-user.jp
        追加 -> 閉じる -> OK


special thanks
  ujihara san, who created concat pdf


refer
  etc user for japan
    http://www.etc-user.jp/
  concat pdf(LGPL lisence)
    http://www.ujihara.jp/ConcatPDF/en/
    http://www.ujihara.jp/ConcatPDF/
