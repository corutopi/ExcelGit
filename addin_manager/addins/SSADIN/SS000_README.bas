Attribute VB_Name = "SS000_README"
'***課題************************************************************************
'・ダブルクォーテーションを考慮したCSVデータ加工（メソッド〇, DataFrame）
'*******************************************************************************



'*******************************************************************************
'--comment_version_0.1.0------------------
'共通アドインモジュール.
'各種エクセルツールで使いまわせる様々な関数・クラスを作成することを目的とした
'アドイン.
'構築ルールは以下の通り.
'
'標準モジュール:
'   ・モジュール名はヘッダ「SSXXX_」で始める.XXXは3桁0埋の数字.
'   ・XXXの項番ルールは以下の通り.
'           0XX : 処理に直接関係しないお作法的な関数群.ログ出力など.
'           1XX : VBAコーディングでよく使用する関数群.
'           2XX : シート上のフォームコントロールで使用する関数群.
'           3XX :
'           4XX :
'           5XX : テストで使用するファイルなど、コーディング作業の支援関数.
'           6XX :
'           7XX : 単体で一つのツール or 機能として使用する関数群.
'           8XX :
'           9XX : 上記以外の関数群.未完成のコードなど.
'   ・7XXの関数群でフォームを伴う場合は同じフォーム名に同じヘッダをつける.
'   ・【単体で完結】させる.(単一で別Excelに移しても動作する.)
'   ・ただし、7XXのモジュールのみ連動するフォームの存在を許容する.
'クラス:
'   ・クラス名はヘッダ「SSC_」で始める.
'   ・単体で完結させる.
'-----------------------------------------
'--更新履歴-------------------------------
'yyyymmdd   :xxx            :[更新内容]
'*******************************************************************************


'*******************************************************************************
'--comment_version_0.2.0------------------
'Model Comment.
'public な関数にはコメントを書くようにしたい.
'コメントの形を変えたくなった時のためにバージョン付けとこう.
'絶対なルールではなく、ゆる〜い感じで守る.
'Gitに登録する際に文字化けするのでできる限り英語を使う。
'→ と思ったけど日本語もちゃんと表記されるようになったっぽい？でもついでだからなるべく英語で書く。
'
'-----------------------------------------
'argment    :hikihiki       :discription return
'return     :discription return
'-----------------------------------------
'--change log-------------------------------
'yyyymmdd   :xxx            : [discription about update]
'*******************************************************************************
