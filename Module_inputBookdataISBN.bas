Attribute VB_Name = "Module_inputBookdataISBN"
Option Explicit


Sub inputBookdataISBN()


    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
'        objIE.Visible = False 'IEを表示
        objIE.Visible = True 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        '作業ワークシート
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
        'データ取得URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn"
        '繰り返し処理
        Dim i As Integer
        i = 1
        '処理完了メッセージ
        Dim ExitMsg As String
        
    'URLを開いてオブジェクト取得
    objIE.navigate InputISBNPage 'IEでURLを開く
    Call WaitResponse(objIE) '読み込み待ち
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

    'フォーム入力
    htmlDoc.getElementsByClassName("form-input__input")(0).Value = "1234567890123"

    'VBA終了処理
'    objIE.Quit 'objIEを終了させる
    ExitMsg = "test"
    MsgBox ExitMsg


End Sub


Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
