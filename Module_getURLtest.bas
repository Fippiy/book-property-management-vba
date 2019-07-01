Attribute VB_Name = "Module_getURLtest"
Option Explicit

Sub getURLtest()
    
    'オブジェクト設定
        
        'IEオブジェクト
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
'        objIE.Visible = True 'IEを表示
        
        'HTMLオブジェクト
        Dim htmlDoc As HTMLDocument 'HTML全体
'        Dim HTTPStatus As Integer 'HTTPリクエストステータス
        
        'URL設定
        Dim Domain As String 'Webドメイン名
        Domain = "https://protected-fortress-61913.herokuapp.com/" 'ドメイン格納
        Dim ProcessDir As String '処理実施ディレクトリ
        ProcessDir = "book" 'ディレクトリ指定
                
        'VBA動作初回ログインチェック
        Dim CheckFirstLogin As Boolean 'ログインチェックフラグ
        CheckFirstLogin = True
                
    'ログイン状態チェック
    Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)

    'VBA各種処理の実施
    
        'navigate時にログイン状態を確認として挿入
        Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
    
    'VBA各種処理完了
    
    objIE.Quit 'objIEを終了させる
    MsgBox "処理が完了しました。"

End Sub

Sub CheckLogin(objIE As InternetExplorer, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
        
    'オブジェクト設定
        
        'ログイン設定
        Dim LoginDir As String 'ログインディレクトリ
        LoginDir = "login" 'ログインディレクトリ指定
        Dim LoginPageURL As String 'ログインページURL
        LoginPageURL = Domain & LoginDir 'ログインページURL生成
        Dim LoginEmail As String 'ログインメールアドレス
        Dim LoginPassword As String 'ログインパスワード
        
        Dim LoginAnswer As String 'ログイン結果確認用
        Dim ExitMsg As String 'メッセージ表示用
        
        'URL取得設定
        Dim ProcessPageURL As String '処理実施ページURL
        Dim ResponseURL As String '処理実施ページ表示後URL取得
        
    '処理実施ページ決定
    If CheckFirstLogin = True Then
        ProcessPageURL = LoginPageURL 'ログインページURL生成
    Else
        ProcessPageURL = Domain & ProcessDir '処理実施ページURL生成
    End If
    
    '処理実施ページへアクセス後、URL取得
    objIE.navigate ProcessPageURL 'IEで開く
    Call WaitResponse(objIE) '読み込み待ち
    ResponseURL = objIE.document.URL 'URL取得
    
    'ログイン画面表示時はログイン処理
    If ResponseURL = LoginPageURL Then
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        'フォーム入力
        htmlDoc.getElementsByName("email")(0).Value = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 1)
        htmlDoc.getElementsByName("password")(0).Value = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 2)
        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
        
        'ログイン結果確認
        Call WaitResponse(objIE) '読み込み待ち
        ResponseURL = objIE.document.URL '読み込み後のURL取得
        Debug.Print ResponseURL 'デバッグ確認
        If ResponseURL = LoginPageURL Then
            LoginAnswer = "ログイン失敗"
            'ログイン失敗時はアラートをメッセージとして返す
            ExitMsg = "ログインに失敗しました。"
            MsgBox ExitMsg
            '続きの処理はせずに終了
            End
        Else
            LoginAnswer = "ログイン成功"
        End If
    Else
        LoginAnswer = "ログイン済み"
    End If
    
    'ログイン済みorログイン後サイトのHTMLオブジェクト取得
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    '結果確認情報
    Debug.Print "結果表示開始"
    Debug.Print "CheckFirstLogin " & CheckFirstLogin
    Debug.Print LoginAnswer
    Debug.Print "ProcessPageURL " & ProcessPageURL
    Debug.Print "ProcessDir " & ProcessDir
    Debug.Print "Web表示Titleタグ " & htmlDoc.getElementsByTagName("title")(0).innerText
    Debug.Print "結果表示終了"
    Debug.Print ""
    
    '初回処理終了処理
    If CheckFirstLogin = True Then CheckFirstLogin = False
    
End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
