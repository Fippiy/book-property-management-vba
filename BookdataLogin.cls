VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookdataLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public htmlDoc As HTMLDocument
Public Domain As String
Public ProcessDir As String
Public CheckFirstLogin As Boolean
Public waitObjIE As waitObjIE

'Sub CheckLogin(waitObjIE As waitObjIE, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
Sub CheckLogin()
        
    'オブジェクト設定
        
        'ログイン設定(ディレクトリ)
        Dim LoginDir As String 'ログインディレクトリ
        LoginDir = "login" 'ログインディレクトリ指定
        Dim LoginPageURL As String 'ログインページURL
        LoginPageURL = Domain & LoginDir 'ログインページURL生成
        'ログイン設定(Web送信情報)
        Dim LoginEmail As String 'ログインメールアドレス
        Dim LoginPassword As String 'ログインパスワード
        LoginEmail = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 1).Value 'Email
        LoginPassword = ThisWorkbook.Worksheets("ログイン設定").Cells(2, 2).Value 'Password
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
'    waitObjIE.navigate ProcessPageURL 'IEで開く
    waitObjIE.objIE.navigate ProcessPageURL
'    Call WaitResponse(objIE) '読み込み待ち
    waitObjIE.WaitResponse

'    ResponseURL = waitObjIE.document.URL 'URL取得
    ResponseURL = waitObjIE.objIE.document.URL 'URL取得
    
    'ログイン画面表示時はログイン処理
    If ResponseURL = LoginPageURL Then
'        Set htmlDoc = waitObjIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        Set htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        'ログインフォーム入力
        htmlDoc.getElementsByName("email")(0).Value = LoginEmail
        htmlDoc.getElementsByName("password")(0).Value = LoginPassword
        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
        
        'ログイン結果確認
'        Call WaitResponse(objIE) '読み込み待ち
         waitObjIE.WaitResponse

'        ResponseURL = waitObjIE.document.URL '読み込み後のURL取得
        ResponseURL = waitObjIE.objIE.document.URL '読み込み後のURL取得
        If ResponseURL = LoginPageURL Then 'ログインURLのままは失敗
            'ログイン失敗として終了する
'            waitObjIE.Quit 'objIEを終了させる
            waitObjIE.objIE.Quit 'objIEを終了させる
            MsgBox "ログインに失敗しました。"
            End '続きの処理はせずに終了
        End If
    End If
    
    'ログイン済みorログイン後サイトのHTMLオブジェクト取得
'    Set htmlDoc = waitObjIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    Set htmlDoc = waitObjIE.objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    '初回処理終了処理
    If CheckFirstLogin = True Then CheckFirstLogin = False
    
End Sub

