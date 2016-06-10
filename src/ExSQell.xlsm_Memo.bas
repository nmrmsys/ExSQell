'========================================================
' ExSQell.xlsm - Copyright (C) 2016 nmrmsys
' sqltools.xls - Copyright (C) 2005 M.Nomura
'========================================================
'
'changelog
' v0.0.1 設定シート基本構成、SQL実行周り ExecSql ST_Query ST_Plus
' v0.0.2 システムユーティリティの拡充 GetCfgVal GetConStr Propクラス
' v0.0.3 件数確認ロジック追加 ST_Query ExecSql、コメント操作 [GS]etCmtVal
' v0.0.4 マクロ起動方式の変更 Aplication.OnKey、GetCfgSheet
' v0.0.5 結果シート作成まわりを整理 GetNewSheet MakeList
' v0.0.6 定型SQLのシート管理、プレースホルダ置換 GetLibSql SetSqlParam
' v0.0.7 ST_Whichの実装 オブジェクト情報取得 GetDBObjectType/List/Info
' v0.0.8 GetSql拡張 文途中SQL取得、構文タイプ判断、タイトル取得、SELECT補完
' v0.0.9 縞模様を条件付き書式で実現、フィルタ絞込み時も縞々になるようにした
' v0.1.0 DML/DDL/DCL実行 ST_Query ExecSql
' v0.1.1 件数確認のみ or サンプル抽出 ST_Query MakeList
' v0.1.2 拡張一覧項目 ST_Which GetDBObjectList MakeList
' v0.1.3 便利なショートカットキー追加 シート削除、設定シート、前シート
' v0.1.4 拡張情報取得/表示 ST_ExtInfo GetDBObjectExtInfo
' v0.1.5 オブジェクト情報取得 INDEX, PACKAGE追加 GetDBObjectInfo
' v0.1.6 メンテナンスモード ST_Maintain(途中まで)、MakeList 空白列 行連番
' v0.1.7 定型ＳＱＬメニュー表示 ST_Library(途中まで)
' v0.1.8 BugFix MakeList 項目名での値取得を序数に、同名だとと取得できない為
' v0.1.9 クエリ実行時の縞模様書式のセットを設定で制御できるようにした
' v0.2.0 メンテナンスモードの更新SQL生成、実行部分を仮作成
' v0.2.1 クエリ実行時の件数、秒数のコメント付加を設定で制御できるようにした
' v0.2.2 メンテモード更新 ステータスバー表示、同条件で再抽出、自動再抽出
' v0.2.3 オートフィルタ条件全解除のショートカットを追加
'          DisplayFormulas = true（Ctrl+Shift+@）で数式を可視化できるとわ・・
' v0.2.4 改行有りSQL文をセル行にセットする関数、空白セルでの Ctrl+Wの無視
' v0.2.5 シート名は設定項目じゃなく定数に変更、行番号表示 値/数式/非表示を選択
' v0.2.6 メンテシートにイベントマクロを付加して自動で U or I マークをセット
' v0.2.7 抽出条件シート表示、入力後 Ctrl + Q でクエリ実行
' v0.2.8 新規作業シート作成のショートカット追加
' v0.2.9 表引き可能オブジェクトのクエリ実行時は抽出条件シートを表示
' v0.3.0 メンテナンスシートのキー項目タイトルを太字表示
' v0.3.1 定型SQLポップアップメニュー 実行結果表示、定型SQLシート表示
' v0.3.2 抽出条件シートを非表示にして条件指定を使い回すモード追加
' v0.3.3 シート非表示、非表示シート一覧ショートカットキー追加
' v0.3.4 抽出条件の入力を強化、NULL指定、IN指定、追加SQL文のコメント対応
' v0.3.5 コードの整理、MakeListでデータ抽出時に進捗状況をステータス表示
' v0.3.6 ライブラリにお気に入りブックを登録しポップアップメニュー表示
' v0.3.7 GetSqlParams,GetSql SQL文中に&置換変数があれば入力DLG表示
' v0.3.8 ST_Template,GetTemplateSql テンプレートSQL文の生成
' v0.3.9 オブジェクト一覧表示でタイプ/名称もフィルタ
' v0.4.0 抽出条件シート 項目日本語名称を表示名に、テーブル名エイリアス付加
' v0.4.1 クエリ結果シートのデータ型文字色、SQL*Plusのウインドウタイトル変更
' v0.4.2 データ貼り付け時、書式設定を数値列以外にだけ文字列書式を列設定
' v0.4.3 選択範囲のSQLを実行 SELECTのみ 文頭の括弧がある時は対応する括弧まで
'          文末の実行対象外ブロックの判定は難しいので取りあえず放置
' v0.4.4 ExecSql Scroll Lock時は実行SQLをウインドウ表示
' v0.4.5 テーブル表示時のシート名はプレフィックス + テーブル名で上書き作成
'          メンテナンスシートも同様にシート名のプレフィックスを定数化
' v0.4.6 コード中のSQL抽出、SQLのコード化
' v0.4.7 SQLの整形 取りあえず SqlFmt.dll使用の機能を付けとく、そのうち換装
' v0.4.8 更新SQLの実行設定に実行しないを付加、お断りメッセージを表示させる
' v0.4.9 実行SQLのウインドウ表示にクエリを実行する/しないのボタンを追加
' v0.5.0 クエリ結果シートの A1 or 空白セルで Ctrl + Q 時は再抽出をおこなう
' v0.5.1 更新SQLに ;改行 がある場合は分割してバッチ実行、選択範囲指定も可
' v0.5.2 条件付書式で無い縞模様、列幅自動調整にデータ行/列全体の設定を追加
' v0.5.3 隠し機能の抽出条件シートのロック、項目名称表示、SQL表示を公開する
' v0.5.4 単独の更新SQL実行がバグっていたのを修正、ライブラリURLメニュー追加
' v0.5.5 式項目のデータ型判定がおかしいので oo4oの時は .OraIDataTypeを使用
' v0.5.6 データ貼り付け時のESC押下で処理中断、メンテ時は０件チェック無し
' v0.5.7 抽出条件 追加SQLのAND付加を修正、空行挿入、条件入力でOR条件を表現
' v0.5.8 OO4O版のレコードセット取得、更新SQL発行の非同期実行 ESCキーで中断
' v0.5.9 更新SQL実行時の影響レコード件数表示 ０件なら表示しないようにした
' v0.6.0 選択範囲のSQL文を取得して編集ウインドウを表示、修正した内容を再展開
' v0.6.1 編集ウインドウの機能強化、空白セルを強調表示、重複チェック列の追加
' v0.6.2 定義一覧表示時に存在するオブジェクトなら自動でフィルタ選択のみする
' v0.6.3 ストアドの再コンパイル、範囲選択でバッチ再コンパイルも可能
' v0.6.4 oo4oのCreateDynaset時のオプション指定がいろいろ不味かったので修正
' v0.6.5 ストアドのコンパイル、エラー情報をソース行色付け、コメントで表示
' v0.6.6 ストアドの実行
' v0.6.7 ライブラリ系コマンド キーアサイン変更、登録SQL/リンク追加
' v0.6.8 接続オブジェクトのプーリング、コミットモード/メニューの追加
' v0.6.9 AUTOTRACEを設定した状態での Plus起動
' v0.7.0 コミットメニュー 更新SQLを実行している時は * を後ろに表示
' v0.7.1 Worksheet for Oracle Utilityを搭載
' v0.7.2 抽出条件 日付型NULL指定不具合、複数値ワイルドカードの場合LIKE OR展開
' v0.7.3 ライブラリにナイスミドルなSQLを追加、クエリ実行で置換変数入力DLG表示
' v0.7.4 件数確認SQL文生成時にORDER BYが含まれている場合は以降を削除
'
'todo
'  ストアドのパラメータ情報取得 oo4o:USER_ARGUMENTS テンプレ展開
'  ストアドの実行 DBMS_OUTPUTをGET_LINESで取得し、結果表示
'  メンテ/抽出条件シートのSQL生成時にシート名からテーブル名を取得しない
'  オブジェクト一覧表示で Like演算子を利用した選択表現
'  拡張一覧表示方法変更 MakeList時に処理実行、行再描画
'  拡張情報の取り扱い PropのKeysプロパティ使用
'  GetDBObject* 別スキーマ/ＤＢ対応
'  SQL*Plusを裏で実行して結果を取得 実行計画、トレース、その他で使用
'  コミットモード 接続オブジェクトの管理 コンサバモードで実現 memo参照
'  キー中断のロジックが誤動作してエラー時に無限ループになってた
'    マクロ実行中の中断キー許可/エラー制御の方針を整理する必要あり
'  SQL発行を非同期実行にしてキー中断を可能にする
'    oo4o CreateSQL ORASQL_NONBLK(&H4) で OraSqlStmtを作成し NonBlockingState
'    ado  adAsyncExecute adAsyncFetchを使用し Stateを監視 Cancelで中断
'  SQL整形エンジン ReSQL
'  LogParserのレコードセットも取り扱えるようにする
'    COM入力形式プラグインで各種ドキュメント形式も検索可能
'
'コマンド一覧(未実装も含む)
'  Ctrl + Q SQL実行
'    SHIFT同時 件数確認DLG表示
'    ALT同時 件数確認、抽出条件入力キャンセル
'    表、ビューはSELECT自動補完、抽出条件入力、強制件数確認
'  Ctrl + M データメンテナンス
'    SHIFT同時 件数確認DLG表示
'    ALT同時 件数確認、抽出条件入力キャンセル
'    表、ビューはSELECT自動補完、抽出条件入力、強制件数確認
'  Ctrl + P SQL*PLUS起動
'    SHIFT同時 実行計画: 未
'    ALT同時 トレース: 未
'  Ctrl + W オブジェクト一覧/定義(表定義、ソース)
'    オブジェクトを特定できた場合は定義表示
'    SHIFT同時 一覧表示キャンセル
'    ALT同時 拡張一覧項目も表示、SHIFT同時で一覧表示：未
'    選択セルを1つづつ処理：未
'  Ctrl + E 拡張情報をコメントで付加
'    SHIFT同時 バルーン表示 チャートオブジェクトの方がいいかも
'    ALT同時 定義書シート: 未
'    選択セルを1つづつ処理：未
'  Ctrl + T テンプレート生成
'    SHIFT同時 SQL整形
'    ALT同時 SQL文字列
'  Ctrl + L 便利な一覧をポップアップメニュー選択し表示
'    SHIFT同時 SQL挿入
'    ALT同時 定型SQLシート表示
'  Ctrl + K オンラインマニュアル：未
'  Ctrl+Shift+INS     新規作業シート作成
'  Ctrl+Shift+DEL     現在シートの削除
'  Ctrl+Shift+BS      直前のシートを表示
'  Ctrl+Shift+TAB     設定シートを表示
'  Ctrl+Shift+ *      フィルタ/抽出条件のクリア
'  Ctrl+Shift+ -      シートの非表示
'  Ctrl+Shift+ ^      非表示シート一覧
'  Alt+Shift+方向キー 現在セルと値の違うセルに移動：未
'
'memo
'
'共通：
'  接続関連
'    A1セルに接続文字列だけでなく キー:値 形式を許可
'    コミットモード 接続オブジェクトの管理
'      コネクションサーバ ConServ接続モード 略してコンサバモード
'      SQL発行部分を MultiUseの1スレッドプールな ActiveX.EXEで実装
'      接続キーを指定する事により異なるクライアントからもコネクション
'      トランザクションを共有できる タスクトレイから現在の接続一覧、
'      コミット/ロールバック制御、実行SQL履歴の確認 マルチスレッド化
'      似非SQL*Plus風のHTAクライアント SQL*Minus
'      http://www.int21.co.jp/pcdn/vb/noriolib/vbmag/9809/com/
'      http://www.int21.co.jp/pcdn/vb/noriolib/vbmag/9810/db_solu/
'      http://www.koalanet.ne.jp/~akiya/vbtaste/vbp/#else03
'      http://www2.plala.or.jp/k-world/vbasic/vbasic009.html
'      http://pooh3.dip.jp/vbvcjava/window/tasktray.html
'      http://www2.netf.org/freesoft.html#IPCATL
'  定型SQL管理
'    &変数取得用 GetSqlParams acceptのpromptがある場合はそれも取得
'    acceptはあんまり意味が無かった無名ブロックの結果を取得できない
'      DBMS_OUTPUT.GET_LINEで取得できる事が解ったがストアド実行は
'      oo4oでやることにしたのであまり使い道が無い
'    SQL文生成コードからのSQL文抽出をするときに変数名に&をつけて展開
'  結果シート生成
'    1 行ごとの色塗りや重複行検知を条件付書式に
'      http://www.excel7.com/chotto16.htm
'      http://www.excel7.com/chotto17.htm
'      前行と同じ場合は非表示 データ行A列 条件付き書式で =A2=A1
'      メンテナンスシート行追加時対応用の数式No.列
'  便利機能
'    元セル位置にジャンプ HYPERLINK関数
'    重複行検知書式セット、シート同士の差分比較
'
'ST_Query：
'  単純SELECT文の実行
'  DMLを判別し実行 INSERT UPDATE DELETE MERGE
'  DDLを判別し実行 CREATE ALTER DROP RENAME TRUNCATE COMMENT
'  DCLを判別し実行 GRANT REVOKE
'  オブジェクト名だけの場合は自動でSQL文展開
'    TABLE, VIEW SELECT文展開、抽出条件入力シート表示
'      http://chibie2000.hp.infoseek.co.jp/excelvba/
'      ユーザーダイアログはグリッド表示ができないので止め
'    PROCEDURE, FUNCTION EXEC文：未
'  メンテナンスシートを判別しDB更新
'  SPを判別し引数入力､実行､結果表示:未
'  SPを判別しSHIFT同時押しで再コンパイル:未
'    ALTER PROCEDUR/FUNCTION objname COMPILE
'ST_Which：
'  オブジェクト情報の取得
'    OO4O: ディクショナリ or OraMetaData(key,index取得不可)
'      シノニムの解決
'      別スキーマ情報の取得 DBA_* owner指定要
'      別データベース情報の取得 OraMetaData or 別接続
'      拡張情報の表示
'      PUBLICを検索対象にするしない
'    ADO:  ディクショナリ (接続先がoracleの場合のみ) or ADOX
'  Private Function GetDBObjectType(argObjName) As String
'    OO4O: カタログの *_OBJECTS
'    ADO:   ADOXで Tables, Views, Procedures
'  Private Function GetDBObjectList(argObjName, argRs, argFlds, argOpts)
'    OO4O: カタログの *_OBJECTS、拡張情報は *_*S とジョイン
'    ADO:   ADOXで Tables, Views, Procedures
'  Private Function GetDBObjectInfo(argObjName, argRs, argFlds, argOpts)
'     OO4O: カタログの *_*S
'     ADO:   ADOXで Table(Columns, Indexes, Keys), View / Procedure(Command)
'  定義シート TABLE, VIEW, INDEX
'    定義シートは先頭セルがオブジェクト名でフィルタ可能
'    名称､空白
'    名称、行数、マーク列、項目定義 or コード
'    条件付き書式で2行目以降は空白
'  ソースシート PROCEDURE, FUNCTION, PACKAGE
'  その他シート SEQENCE, TRIGGER, SNAPSHOT, DBLINK
'  拡張情報 シノニム実体 件数 行数 サイズ
'  Ctrl+Alt+Shift+W は定義表示をキャンセル
'  セル範囲選択時 Ctrl+W は一覧表示キャンセル
'  DBA_* で ディクショナリテーブルの情報が取得できるか 取れない
'ST_ExtInfo：
'  拡張情報をコメント付加
'  拡張情報をバルーン表示
'  定義書生成
'  選択セルを1つづつ処理:   未
'ST_Maintenance：
'  データ編集シート
'    A列を非表示､A1に接続先､テーブル､キー
'      それ以降にはTAB区切りの更新キー値を保存
'    B列で i: 追加 U: 更新 D: 削除 を指定
'    ワークシートイベントでセルの更新を取得
'  更新キー列はROWID
'  VBAからのイベントマクロ挿入 http://www.mrexcel.com/archive/VBA/14890.html
'ST_Plus：
'  接続文字列をウインドウタイトル表示:
'    http://homepage2.nifty.com/sak/w_sak3/doc/sysbrd/vb_t02.htm
'  SHIFT同時で実行計画:   未
'  ALT同時でトレース表示:   未
'ST_Template：
'  表、ビューはSELECT/INSERT/UPDATE/DELETE/CREATE
'  SPは実行用の無名ブロックを生成
'  その他のオブジェクトはDDLを生成
'  SQL整形 SqlFmt.dllはいま一つ、自前で実装
'  SQL文字列化/解除
'ST_Library：
'  コマンドバーでポップアップメニュー
'  ユーザ､セッション､使用率
'  定型SQLシートのFLGがONのもののみ表示 一覧 Or 挿入
'ST_Knowledge：
'  http://www.oracle.co.jp/ultra_owa/ULTRA/usearch?rdo_1=1 マニュアル/製品情報 検索
'  http://otn.oracle.co.jp/otn_pl/otn_tool/sqlconstraction SQL構文検索
'  http://otn.oracle.co.jp/otn_pl/otn_tool/err_code_proc   エラー・メッセージ検索
'  http://otn.oracle.co.jp/otn_pl/otn_tool2/search_forum   OTN掲示板検索
'
'そのうち：
'  項目名自動日本語化
'  アドイン化
'
