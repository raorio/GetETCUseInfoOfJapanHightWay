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


寄付のお願い
  このツールにより得られた利益がありましたら、何%でもよいので、寄付していただければ幸いです。
  寄付していただいた金額は、まず世界の恵まれない人々への寄付にあてさせていただきます。
    例：world vision等 http://www.worldvision.jp/
  私が寄付金控除を行い、控除額を、こちらの開発等の諸経費として使わせていただきます。
  寄付される場合は、以下の口座へよろしくお願いします。
    TODO
  寄付された金額や寄付者の情報は、基本的に、寄付履歴にupいたします。
    寄付履歴ページ：TODO


refer
  etc user for japan
    http://www.etc-user.jp/
  concat pdf(LGPL lisence)
    http://www.ujihara.jp/ConcatPDF/en/
    http://www.ujihara.jp/ConcatPDF/
