Attribute VB_Name = "Module_getURLtest"
Option Explicit

Sub getURLtest()
    
    'クラスモジュール化テスト(IE読み込みチェック
    Dim waitObjIE As waitObjIE 'IE読み込み待ちモジュール作成
    Set waitObjIE = New waitObjIE
    Set waitObjIE.objIE = CreateObject("Internetexplorer.Application")
'    waitObjIE.objIE.Visible = True
    waitObjIE.objIE.Visible = False
    
    'クラスモジュール化テスト(ログイン
    Dim Login As BookdataLogin 'ログインクラスモジュール作成
    Set Login = New BookdataLogin
    Login.Domain = "https://protected-fortress-61913.herokuapp.com/" 'ドメイン格納
    Login.ProcessDir = "book" 'ディレクトリ指定
    Login.CheckFirstLogin = True 'ログインチェックフラグ
    Set Login.waitObjIE = waitObjIE 'IEオブジェクトをLoginに引渡
    
    'オブジェクト設定
        
        'IEオブジェクト
'        Dim objIE As InternetExplorer 'IEオブジェクトを準備
'        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
'        objIE.Visible = False 'IEを表示
'        objIE.Visible = True 'IEを表示
        
        'HTMLオブジェクト
        Dim htmlDoc As HTMLDocument 'HTML全体
        
        'URL設定
'        Dim Domain As String 'Webドメイン名
'        Domain = "https://protected-fortress-61913.herokuapp.com/" 'ドメイン格納
'        Dim ProcessDir As String '処理実施ディレクトリ
'        ProcessDir = "book" 'ディレクトリ指定
                
        'VBA動作初回ログインチェック
'        Dim CheckFirstLogin As Boolean 'ログインチェックフラグ
'        CheckFirstLogin = True
                
    'ログイン状態チェック
'    Call CheckLogin(ObjIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
'    Call CheckLogin(waitObjIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
    Login.CheckLogin
    
    'VBA各種処理の実施
    
        'navigate時にログイン状態を確認として挿入
'        Call CheckLogin(ObjIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
'        Call CheckLogin(waitObjIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
        Login.CheckLogin
    
    'VBA各種処理完了
    
'    objIE.Quit 'objIEを終了させる
    waitObjIE.objIE.Quit 'objIEを終了させる
    MsgBox "処理が完了しました。"

End Sub
'
''Sub CheckLogin(objIE As InternetExplorer, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
'Sub CheckLogin(waitObjIE As waitObjIE, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
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
'        '処理結果確認
'        Dim LoginAnswer As String 'ログイン結果確認用
'        Dim ExitMsg As String 'メッセージ表示用
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
''    objIE.navigate ProcessPageURL 'IEで開く
'    waitObjIE.objIE.navigate ProcessPageURL
''    Call WaitResponse(objIE) '読み込み待ち
'    waitObjIE.WaitResponse
''    ResponseURL = objIE.document.URL 'URL取得
'    ResponseURL = waitObjIE.objIE.document.URL 'URL取得
'
'    'ログイン画面表示時はログイン処理
'    If ResponseURL = LoginPageURL Then
''        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'        Set htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'        'フォーム入力
'        htmlDoc.getElementsByName("email")(0).Value = LoginEmail
'        htmlDoc.getElementsByName("password")(0).Value = LoginPassword
'        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
'
'        'ログイン結果確認
''        Call WaitResponse(objIE) '読み込み待ち
'        waitObjIE.WaitResponse
''        ResponseURL = objIE.document.URL '読み込み後のURL取得
'        ResponseURL = waitObjIE.objIE.document.URL '読み込み後のURL取得
''        Debug.Print ResponseURL 'デバッグ確認
'        If ResponseURL = LoginPageURL Then
'            LoginAnswer = "ログイン失敗"
'            'オブジェクト終了処理を実施しておく
''            objIE.Quit 'objIEを終了させる
'            waitObjIE.objIE.Quit 'objIEを終了させる
'            'ログイン失敗時はアラートをメッセージとして返す
'            ExitMsg = "ログインに失敗しました。"
'            MsgBox ExitMsg
'            '続きの処理はせずに終了
'            End
'        Else
'            LoginAnswer = "ログイン成功"
'        End If
'    Else
'        LoginAnswer = "ログイン済み"
'    End If
'
'    'ログイン済みorログイン後サイトのHTMLオブジェクト取得
''    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'    Set htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'
'    '結果確認情報
'    Debug.Print "結果表示開始"
'    Debug.Print "CheckFirstLogin " & CheckFirstLogin
'    Debug.Print LoginAnswer
'    Debug.Print "ProcessPageURL " & ProcessPageURL
'    Debug.Print "ProcessDir " & ProcessDir
'    Debug.Print "Web表示Titleタグ " & htmlDoc.getElementsByTagName("title")(0).innerText
'    Debug.Print "結果表示終了"
'    Debug.Print ""
'
'    '初回処理終了処理
'    If CheckFirstLogin = True Then CheckFirstLogin = False
'
'End Sub

'Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
'    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
'        DoEvents
'    Loop
'End Sub

'ワークシート内オブジェクト取得テスト
Sub testobjname()
    Dim strObjName() As String
    Dim intObj As Integer
    Dim i As Integer

    'アクティブシートのShapes数をカウント
    intObj = ActiveSheet.Shapes.Count
    '配列を再宣言
    ReDim strObjName(intObj)

    '配列strObjNameにオブジェクト名を代入
    For i = 1 To intObj
        strObjName(i) = ActiveSheet.Shapes(i).Name
    Next i

    '配列strObjNameに代入されたオブジェクト名を表示
    For i = 1 To intObj
'        MsgBox strObjName(i)
        Debug.Print strObjName(i)
    Next i
End Sub
