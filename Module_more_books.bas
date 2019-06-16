Attribute VB_Name = "Module_more_books"
Option Explicit

Sub getBookdatas()

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
    
    Dim BookdataURL As String '詳細ページURL
    Dim BookdataURLSplit() As String '詳細ページURL要素
    Dim BookdataURLBound As Long 'URL要素数
    Dim BookdataID As Integer 'ID番号
    Dim BookdataImg As HTMLImg 'IMGタグ情報
    Dim ImgURL As String '画像URL
    Dim ActCell As Range '画像出力セル


    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        
        '--detail情報からデータ取得
        
            '--detailを取得
            Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0)
    
            'タイトル名取得
            SWSheet.Cells(i + 1, 2).Value = detailField.getElementsByClassName("list-book-title")(0).innerText
            
            '詳細テキスト
            SWSheet.Cells(i + 1, 3).Value = detailField.getElementsByClassName("list-book-detail")(0).innerText
            
            '詳細ページURL
            BookdataURL = detailField.getElementsByTagName("a")(0) 'URL取得
            SWSheet.Cells(i + 1, 4).Value = BookdataURL  '取得URL反映
            
            'Bookdata_ID取得
            BookdataURLSplit = Split(BookdataURL, "/")  'URL要素分割
            BookdataURLBound = UBound(BookdataURLSplit)  'URL要素数確認
            BookdataID = BookdataURLSplit(BookdataURLBound)  'ID番号取得
            SWSheet.Cells(i + 1, 1).Value = BookdataID
        
        '--detail情報からデータ取得ここまで
        
        
        '画像処理

            Set BookdataImg = Bookdata.getElementsByTagName("img")(0)  '画像取得
            ImgURL = BookdataImg.src '画像URL
            Set ActCell = SWSheet.Cells(i + 1, 5) '出力セル

            '画像出力セルのピクセルを指定して表示
            SWSheet.Shapes.AddPicture _
                fileName:=ImgURL, _
                    LinkToFile:=True, _
                        SaveWithDocument:=True, _
                        Left:=ActCell.Left, _
                        Top:=ActCell.Top, _
                        Width:=100, _
                        Height:=100

        '画像処理ここまで
        
        '列番号処理
        i = i + 1
    Next Bookdata

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

