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
        'URL取得繰り返し処理
        Dim i As Integer
        i = 1
    
    'OpenPageがある間はループして続ける
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        OpenPage = "" 'データ取得URL初期化
        

        '詳細ページURL取得
'        Call getBookList(DWSheet, htmlDoc, i)
        
        
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
    Call getDetailBookdata(SWSheet, DWSheet, objIE)


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

Sub getDetailBookdata(SWSheet As Worksheet, DWSheet As Worksheet, objIE As InternetExplorer)

    '詳細ページURLから詳細内容を取得
    
    'データ取得URL
    Dim OpenPage As String
    Dim htmlDoc As HTMLDocument 'HTML全体
    Dim DocContent As HTMLDivElement 'HTMLコンテンツ処理
    Dim DocColumn As HTMLDivElement 'column情報
    Dim i As Long, j As Long '書き出し用行列処理
    i = 2
    
    Dim URLi As Long '詳細URL読み込み行番号処理
    URLi = 1
    
    '最初のページ
    OpenPage = DWSheet.Cells(URLi, 1).Value
    
    '画像処理
    Dim DocPicture As Variant
    Dim ImgURL As Variant
    Dim actcell As Variant
    
    '詳細ページを開いて中のデータを取得
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        j = 2
        
        '画像取得処理
        Set DocPicture = htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
        SWSheet.Cells(i, j).Value = ImgURL.src
        Set actcell = SWSheet.Cells(i, j)
        
        SWSheet.Shapes.AddPicture _
          fileName:=ImgURL.src, _
            LinkToFile:=True, _
            SaveWithDocument:=True, _
            Left:=actcell.Left, _
            Top:=actcell.Top, _
            Width:=100, _
            Height:=100
        j = j + 1
                
        '詳細ページHTMLからデータ取得
        'document-contentを取得
        For Each DocContent In htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        '詳細ページURL全取得で終了
        i = i + 1
        URLi = URLi + 1
        
        '次ページURL取得
        OpenPage = DWSheet.Cells(URLi, 1).Value
    Loop
    
End Sub
Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

