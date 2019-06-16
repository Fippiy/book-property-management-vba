Attribute VB_Name = "Module_PropetyDatas"
Option Explicit

Sub getBookdatasDatail()

    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        Dim PagiLink As HTMLAnchorElement '次ページリンク
        '作業ワークシート
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("スクレイピング")
        Dim DWSheet As Worksheet 'DetailWorksheet
        Set DWSheet = ThisWorkbook.Worksheets("詳細ページ情報")
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
        OpenPage = "" 'データ取得URL初期化
        

        '詳細ページURL取得
        Call getBookList(DWSheet, htmlDoc, i)
        
        
        'ページネーション処理
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            'ページネーションがある場合は取得処理
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    OpenPage = PagiLink.href
                End If
            Next PagiLink
        
        End If
        
    Loop 'OpenPageループエンド


    '取得した詳細ページURLから詳細ページ情報を取得する
    Call getDetailBookdata(SWSheet, DWSheet)


    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"

End Sub

Sub getBookList(DWSheet As Worksheet, htmlDoc As HTMLDocument, i As Integer)
    
    '詳細ページURLを取得
    Dim Bookdata As HTMLDivElement 'レコード単位データ
    Dim detailField As HTMLDivElement '詳細フィールドデータ
    Dim BookdataURL As String '詳細ページURL
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        
        '--detail情報からデータ取得
        
            '--detailを取得
            Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0)
    
            '詳細ページURL
            BookdataURL = detailField.getElementsByTagName("a")(0) 'URL取得
            DWSheet.Cells(i, 1).Value = BookdataURL  '取得URL反映
        '--detail情報からデータ取得ここまで
        
        '列番号処理
        i = i + 1
    Next Bookdata

End Sub
Sub getDetailBookdata(SWSheet As Worksheet, DWSheet As Worksheet)

    '詳細ページURLから詳細内容を取得

End Sub
Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub



