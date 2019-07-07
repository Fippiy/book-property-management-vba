Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    '===↓VBA全体オブジェクト設定↓===
        
'        'IEオブジェクト
'        '●●クラスオブジェクトへ載せ替え
'        Dim objIE As InternetExplorer 'IEオブジェクトを準備
'        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
'        objIE.Visible = False 'IEを表示
        'IEオブジェクト（クラス使用
        Dim waitObjIE As waitObjIE 'IE読み込み待ちモジュール作成
        Set waitObjIE = New waitObjIE
        Set waitObjIE.objIE = CreateObject("Internetexplorer.Application")
'        waitObjIE.objIE.Visible = True
        waitObjIE.objIE.Visible = False
        
        'HTMLオブジェクト
        '●●残置
        '中で宣言しているので外はなし
'        Dim htmlDoc As HTMLDocument 'HTML全体
        
'        'データ取得URL
'        '●●ログイン処理内でURLを利用設定とする
'        Dim Domain As String 'Webドメイン名
'        Dim ProcessDir As String '処理実施ディレクトリ
'        Domain = "https://protected-fortress-61913.herokuapp.com/"
'        ProcessDir = "book/isbn_some_input"
'        'VBA動作初回ログインチェック
'        Dim CheckFirstLogin As Boolean 'ログインチェックフラグ
'        CheckFirstLogin = True
        'ログイン・URL設定(クラス使用
        Dim Login As BookdataLogin 'ログインクラスモジュール作成
        Set Login = New BookdataLogin
        Login.Domain = "https://protected-fortress-61913.herokuapp.com/" 'ドメイン格納
        Login.ProcessDir = "book/isbn_some_input" 'ディレクトリ指定
        Login.CheckFirstLogin = True 'ログインチェックフラグ
        Set Login.waitObjIE = waitObjIE 'IEオブジェクトをLoginに引渡
            
    '===↑VBA全体オブジェクト設定↑===
            
    'ログイン状態チェック
    '●●クラス処理に変更
'    Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
    Login.CheckLogin
    
    '===↓処理用オブジェクト設定↓===

        '作業ワークシート設定
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
            
        '登録ISBNコード取得設定
        Dim InputISBN As Collection 'データ取得
        Set InputISBN = New Collection
        Dim ISBNAllCount As Integer 'ISBN総数
        Const LimitEntry As Integer = 20 'フォーム入力ISBN上限
        Dim EntryISBN() As String 'フォーム入力用csv(上限毎)
        Dim MaxRepeat As Long 'ISBN処理回数
        Dim LastISBNCount As Integer '最終ISBN件数
        Dim ElementCounter As Long '要素取得カウンタ
            
        '繰り返し処理
        Dim i As Integer
        Dim j As Integer
        
        '出力メッセージ
        Dim ExitMsg As String
        
    '===↑処理用オブジェクト設定↑===
    
    
    'ワークシートからISBNコード取得
    i = 2 '1行目インデックスなので2行目から
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop
    ISBNAllCount = InputISBN.Count 'ISBNコード総数


    'Web上限毎に登録処理をするカンマ区切りテキストを準備
    
        'ISBN処理回数算出
        MaxRepeat = Application.RoundUp(ISBNAllCount / LimitEntry, 0) '繰り返し回数
        LastISBNCount = ISBNAllCount Mod LimitEntry '繰り返しラスト取得件数
        ReDim EntryISBN(1 To MaxRepeat) '配列として要素指定して再宣言

        ElementCounter = 1 '要素取得カウンタ初期値
        
        'Web処理上限毎に処理できるようにする
        For j = 1 To MaxRepeat
        
            i = 1 '繰り返し変数初期化
            'カンマ区切りテキスト生成(全ISBN or 上限件数まで)
            Do
                EntryISBN(j) = EntryISBN(j) & InputISBN(ElementCounter) 'ISBNコードを要素として追加
                '処理上限orISBN総数ラストはカンマなし
                If ElementCounter <> ISBNAllCount And i <> LimitEntry Then EntryISBN(j) = EntryISBN(j) & ","
                ElementCounter = ElementCounter + 1
                i = i + 1
            Loop Until i > LimitEntry Or ElementCounter > ISBNAllCount
        Next j
        
    '全件処理完了まで繰り返し
        
        
    'カンマ区切りテキストを全て反映させる
    i = 2 '結果出力テキスト挿入位置初期化
        
    For j = 1 To MaxRepeat
        
        'ログイン状態チェックとHTML取得
        '●●クラス処理に変更
'        Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
        Login.CheckLogin
        
        'フォーム入力
        '●●クラス処理に変更
'        htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN(j)
'        htmlDoc.getElementsByClassName("send isbn")(0).Click
        Login.htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN(j)
        Login.htmlDoc.getElementsByClassName("send isbn")(0).Click
    
        'フォーム結果HTML取得
        '●●クラス処理に変更
'        Call WaitResponse(objIE) '読み込み待ち
        waitObjIE.WaitResponse
'        Set htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        Set Login.htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        
        'フォーム処理結果取得
        '●●クラス処理に変更
'        Call getISBNAnswers(htmlDoc, ISSheet, i)
        Call getISBNAnswers(Login.htmlDoc, ISSheet, i)
    Next j

    '全件処理完了まで繰り返し

    'VBA終了処理
    waitObjIE.objIE.Quit 'objIEを終了させる
    ExitMsg = "登録処理が完了しました。"
    MsgBox ExitMsg

End Sub

Sub getISBNAnswers(htmlDoc As HTMLDocument, ISSheet As Worksheet, i As Integer)
    
    '結果処理変数
    Dim ResultRecord As HTMLDivElement 'Record単位データ
    Dim ResultTitle As HTMLDivElement 'タイトル名
    Dim ResultText As HTMLDivElement '結果テキスト
    
    For Each ResultRecord In htmlDoc.getElementsByClassName("isbn-result__box")
    
        Set ResultText = ResultRecord.getElementsByClassName("isbn-result__box--msg")(0)
        Set ResultTitle = ResultRecord.getElementsByClassName("isbn-result__box--title")(0)
        
        ISSheet.Cells(i, 4).Value = ResultText.innerText
        
        If (ResultTitle Is Nothing) = False Then
            ISSheet.Cells(i, 3).Value = ResultTitle.innerText
        End If
                
        i = i + 1
    Next ResultRecord

End Sub

'Sub CheckLogin(objIE As InternetExplorer, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
'
'    'オブジェクト設定
'
'        'ログイン設定(ディレクトリ)
'        Dim LoginDir As String 'ログインディレクトリ
'        LoginDir = "login" 'ログインディレクトリ指定
'        Dim LoginPageURL As String 'ログインページURL
'        LoginPageURL = Domain & LoginDir 'ログインページURL生成
'        'ログイン設定(Web送信情報)
'        Dim LoginEmail As String 'ログインメールアドレス
'        Dim LoginPassword As String 'ログインパスワード
'        LoginEmail = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 1).Value 'Email
'        LoginPassword = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 2).Value 'Password
'        'URL取得設定
'        Dim ProcessPageURL As String '処理実施ページURL
'        Dim ResponseURL As String '処理実施ページ表示後URL取得
'
'    '処理実施ページ決定
'    If CheckFirstLogin = True Then
'        ProcessPageURL = LoginPageURL 'ログインページURL生成
'    Else
'        ProcessPageURL = Domain & ProcessDir '処理実施ページURL生成
'    End If
'
'    '処理実施ページへアクセス後、URL取得
'    objIE.navigate ProcessPageURL 'IEで開く
'    Call WaitResponse(objIE) '読み込み待ち
'    ResponseURL = objIE.document.URL 'URL取得
'
'    'ログイン画面表示時はログイン処理
'    If ResponseURL = LoginPageURL Then
'        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'        'ログインフォーム入力
'        htmlDoc.getElementsByName("email")(0).Value = LoginEmail
'        htmlDoc.getElementsByName("password")(0).Value = LoginPassword
'        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
'
'        'ログイン結果確認
'        Call WaitResponse(objIE) '読み込み待ち
'        ResponseURL = objIE.document.URL '読み込み後のURL取得
'        If ResponseURL = LoginPageURL Then 'ログインURLのままは失敗
'            'ログイン失敗として終了する
'            objIE.Quit 'objIEを終了させる
'            MsgBox "ログインに失敗しました。"
'            End '続きの処理はせずに終了
'        End If
'    End If
'
'    'ログイン済みorログイン後サイトのHTMLオブジェクト取得
'    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'
'    '初回処理終了処理
'    If CheckFirstLogin = True Then CheckFirstLogin = False
'
'End Sub
'
'Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
'    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
'        DoEvents
'    Loop
'End Sub
