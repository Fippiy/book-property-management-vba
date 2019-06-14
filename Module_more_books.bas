Attribute VB_Name = "Module_more_books"
Option Explicit

Sub getBookdatas()

    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = True 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        Dim PagiLink As HTMLAnchorElement '次ページリンク
        '作業ワークシート
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("テスト")
        'データ取得URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        '繰り返し処理
        Dim i As Integer
        i = 1
    
    'OpenPageがある間はループして続ける
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IEでURLを開く
        
        Call WaitResponse(objIE) '読み込み待ち
        
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        
        OpenPage = ""
        
        '書籍情報取得処理
        Call getBookList(SWSheet, htmlDoc, i)
        
        'クラス名(pagination)の取得
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            'ページネーションがある場合は取得処理
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    OpenPage = PagiLink.href
                End If
            Next PagiLink
        
        End If
        
    Loop 'ループエンド

    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"

End Sub

Sub getBookList(SWSheet As Worksheet, htmlDoc As HTMLDocument, i As Integer)
    
    Dim Bookdata As HTMLDivElement 'レコード単位データ
    Dim detailField As HTMLDivElement '詳細フィールドデータ
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        SWSheet.Cells(i + 1, 1).Value = i
        
        Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0) '--detailを取得

        SWSheet.Cells(i + 1, 2).Value = detailField.getElementsByClassName("list-book-title")(0).innerText
        i = i + 1
    Next Bookdata

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

