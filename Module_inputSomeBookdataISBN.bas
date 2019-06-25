Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        '作業ワークシート
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
        '登録ISBN
        Dim InputISBN As Collection 'データ取得
        Set InputISBN = New Collection
        Dim EntryISBN As String 'フォーム入力用csv
        Const LimitEntry As Integer = 20 'フォーム入力ISBN上限
        'データ取得URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn_some_input"
        '繰り返し処理
        Dim i As Integer
        i = 2
        '処理完了メッセージ
        Dim ExitMsg As String
        
    'URLを開いてオブジェクト取得
    objIE.navigate InputISBNPage 'IEでURLを開く
    Call WaitResponse(objIE) '読み込み待ち
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

    'ISBNコード取得
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop

    'カンマ区切りテキスト生成(全ISBN or 上限件数まで)
    i = 1 '繰り返し変数初期化
    Do
        EntryISBN = EntryISBN & InputISBN(i)
        If i <> InputISBN.Count Then EntryISBN = EntryISBN & ","
        i = i + 1
    Loop Until i > InputISBN.Count Or i > LimitEntry

    'フォーム入力
    htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN
    htmlDoc.getElementsByClassName("send isbn")(0).Click

    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    ExitMsg = "本を登録しました。"
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
