Attribute VB_Name = "Module_getURLtest"
Option Explicit

Sub getURLtest()
    
    'オブジェクト設定
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = True 'IEを表示
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim HTTPStatus As Integer 'HTTPリクエストステータス
        'URL設定
        Dim Domain As String 'Web操作ドメイン名
        Dim OpenPage As String '操作URL
        Dim ResponseURL As String '表示サイトURL
        Domain = "https://protected-fortress-61913.herokuapp.com/"
                
'        'URLを開いてオブジェクト取得
'        OpenPage = Domain '削除書籍URL取得
'        objIE.navigate OpenPage 'IEでURLを開く
'        Call WaitResponse(objIE) '読み込み待ち
'        Debug.Print objIE.document.URL '読み込み後のURL取得
        
        'URLを開いてオブジェクト取得
        OpenPage = Domain & "login" '削除書籍URL取得
        objIE.navigate OpenPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        ResponseURL = objIE.document.URL '読み込み後のURL取得
        Debug.Print ResponseURL 'デバッグ確認
        
        'ログインURL設定
        Dim LoginURL As String
        LoginURL = Domain & "login"
        
        'ログイン画面時はログイン処理
        If ResponseURL = LoginURL Then
            Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
'            'フォーム入力
            htmlDoc.getElementsByName("email")(0).Value = "test"
            htmlDoc.getElementsByName("password")(0).Value = "test"
            htmlDoc.getElementsByClassName("form-group__submit")(0).Click
            
            'ログイン結果確認
            Call WaitResponse(objIE) '読み込み待ち
            ResponseURL = objIE.document.URL '読み込み後のURL取得
            Debug.Print ResponseURL 'デバッグ確認
            If ResponseURL = LoginURL Then
                Debug.Print "ログイン失敗"
            Else
                Debug.Print "ログイン成功"
            End If
        End If


'        'URLを開いてオブジェクト取得
'        OpenPage = Domain & "book" '削除書籍URL取得
'        objIE.navigate OpenPage 'IEでURLを開く
'        Call WaitResponse(objIE) '読み込み待ち
'        Debug.Print objIE.document.URL '読み込み後のURL取得


End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
