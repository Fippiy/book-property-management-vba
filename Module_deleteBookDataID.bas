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
        Dim DelID As Collection
        Set DelID = New Collection
        'データ取得URL
        Dim DelBookPageBase As String
        Dim DelBookPage As String
        DelBookPageBase = "https://protected-fortress-61913.herokuapp.com/book/"
        '繰り返し処理
        Dim i As Integer
        i = 2 '2行目から数値取得
        '処理完了メッセージ
        Dim ExitMsg As String
        
    '削除ID取得
    Do Until DelBookSheet.Cells(i, 1).Value = ""
        DelID.Add DelBookSheet.Cells(i, 1).Value
        i = i + 1
    Loop
        
        
    '削除ID毎に処理
    If DelID.Count = 0 Then
        
        ExitMsg = "削除IDがありません"
    
    Else
        
        i = 1 '繰り返し初期化
        Do
            DelBookPage = DelBookPageBase & DelID(i)
            'URLを開いてオブジェクト取得
            objIE.navigate DelBookPage 'IEでURLを開く
            Call WaitResponse(objIE) '読み込み待ち
            Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
            '書籍を消す
            htmlDoc.getElementsByClassName("nav-btn delete")(0).Click
            i = i + 1
        Loop Until i > DelID.Count
        
        ExitMsg = "書籍情報を削除しました"
    
    End If
        

    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
