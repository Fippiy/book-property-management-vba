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
        'データ取得URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        '繰り返し処理
        Dim i As Integer
        i = 1
    
        'URLコレクション
        Dim URLCol As Collection
        Set URLCol = New Collection
        
        '処理完了メッセージ
        Dim ExitMsg As String

    'OpenPageがある間はループして続ける
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        OpenPage = "" 'データ取得URL初期化
        
        '詳細ページURL取得
        Call getBookList(htmlDoc, i, URLCol)
        
        
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

    
    '詳細ページURLがなければ終了する
    If URLCol.Count > 0 Then
        Call getDetailBookdata(SWSheet, objIE, URLCol)
        ExitMsg = "データ取得が完了しました。"
    Else
        ExitMsg = "取得データがありません"
    End If


    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    MsgBox ExitMsg

End Sub

Sub getBookList(htmlDoc As HTMLDocument, i As Integer, URLCol As Collection)
    
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
            URLCol.Add BookdataURL
        
        '--detail情報からデータ取得ここまで
        
        '列番号処理
        i = i + 1
    Next Bookdata

End Sub

Sub getDetailBookdata(SWSheet As Worksheet, objIE As InternetExplorer, URLCol As Collection)

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

    'URL取得総数確認
    Dim fornumber As Long
    fornumber = URLCol.Count
    
    '画像処理
    Dim DocPicture As HTMLDivElement
    Dim ImgURL As HTMLImg
    Dim actcell As Range

    'ID取得
    Dim GetUrl As String '詳細ページURL
    Dim GetUrlData() As String '詳細ページURL,Splitデータ
    Dim GetUrlElement As Integer 'URLSplit要素数
    Dim GetID As Integer 'ID番号

    '詳細ページを開いて中のデータを取得
    Do
        
        '次ページURL取得
        OpenPage = URLCol(URLi)
        
        objIE.navigate OpenPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        j = 1
        
        '1列目にID番号表示
        GetUrl = OpenPage 'URL取得
        GetUrlData = Split(GetUrl, "/")  'URL要素取得
        GetUrlElement = UBound(GetUrlData)  'URL要素確認
        GetID = GetUrlData(GetUrlElement)  'URLから番号取得
        SWSheet.Cells(i, j).Value = GetID
        j = j + 1
        
        '2列目に画像表示
        Set DocPicture = htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
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
        
        
        '3列目以降にテキスト表示
        For Each DocContent In htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        
        'カウント追加
        i = i + 1
        URLi = URLi + 1
        
    'URL要素数を超える場合はループ終了
    Loop Until URLi > fornumber
    
End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
