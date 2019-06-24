Attribute VB_Name = "Module_deleteBookDataID"
Option Explicit

Sub deleteBookdataISBN()

    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        '作業ワークシート
        Dim DelBookSheet As Worksheet 'DeleteBookWorksheet
        Set DelBookSheet = ThisWorkbook.Worksheets("書籍情報削除")
        '削除ID
        Dim DelID As Long
        DelID = DelBookSheet.Cells(2, 1).Value
        'データ取得URL
        Dim DelBookPageBase As String
        Dim DelBookPage As String
        DelBookPageBase = "https://protected-fortress-61913.herokuapp.com/book/"
        DelBookPage = DelBookPageBase & DelID
        '繰り返し処理
        Dim i As Integer
        i = 1
        '処理完了メッセージ
        Dim ExitMsg As String
        
    'URLを開いてオブジェクト取得
    objIE.navigate DelBookPage 'IEでURLを開く
    Call WaitResponse(objIE) '読み込み待ち
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

'    '書籍を消す
    htmlDoc.getElementsByClassName("nav-btn delete")(0).Click

    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    ExitMsg = "test"
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
