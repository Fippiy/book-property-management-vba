Attribute VB_Name = "Module_PropetyDatas"
Option Explicit

Sub getBookdatasDatail()

    '===↓VBA全体オブジェクト設定↓===

        'IEオブジェクト
        Dim waitObjIE As waitObjIE 'IE読み込み待ちモジュール作成
        Set waitObjIE = New waitObjIE
        Set waitObjIE.objIE = CreateObject("Internetexplorer.Application")
        waitObjIE.objIE.Visible = False
        
        'データ取得URL
        Dim Login As BookdataLogin 'ログインクラスモジュール作成
        Set Login = New BookdataLogin
        Login.Domain = "https://protected-fortress-61913.herokuapp.com/" 'ドメイン格納
        Login.ProcessDir = "book" 'ディレクトリ指定
        Login.CheckFirstLogin = True 'ログインチェックフラグ
        Set Login.waitObjIE = waitObjIE 'IEオブジェクトをLoginに引渡
            
    '===↑VBA全体オブジェクト設定↑===
        
    'ログイン状態チェック
    Login.CheckLogin
        
    '===↓処理用オブジェクト設定↓===
        
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        Dim PagiLink As HTMLAnchorElement '次ページリンク
        '作業ワークシート
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("スクレイピング")
        Dim MaxRow As Long 'レコード数確認
        'ワークシートID取得
        Dim books As Range '取得済書籍一覧
        Dim arrBooksId As Variant 'ID配列
        '繰り返し処理
        Dim i As Integer
        i = 1
    
        'URLコレクション
        Dim URLCol As Collection
        Set URLCol = New Collection
        
        '処理完了メッセージ
        Dim ExitMsg As String

    '===↑処理用オブジェクト設定↑===

    'ワークシート書籍情報取得
    MaxRow = SWSheet.Cells(Rows.Count, 1).End(xlUp).Row 'ワークシート要素の最終セル
    Set books = SWSheet.Range(Cells(2, 1), Cells(MaxRow, 1)) 'ワークシートID一覧取得
    arrBooksId = books '配列化
    
    'OpenPageがある間はループして続ける
    Do Until Login.ProcessDir = ""
        
        Login.CheckLogin
        
        '詳細ページURL取得
        Call getBookList(Login.htmlDoc, i, URLCol, Login, arrBooksId)
        
        'ページネーション処理
        Set Pagination = Login.htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            '判定用にリセット
            Login.ProcessDir = ""
            'ページネーションがある場合は取得処理
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    Login.ProcessDir = Replace(PagiLink.href, Login.Domain, "")
                End If
            Next PagiLink
        
        End If
        
    Loop 'OpenPageループエンド

    
    '詳細ページURLがなければ終了する
    If URLCol.Count > 0 Then
        Call getDetailBookdata(SWSheet, waitObjIE.objIE, URLCol, Login, MaxRow)
        ExitMsg = "データ取得が完了しました。"
    Else
        ExitMsg = "新規取得データはありませんでした"
    End If


    'VBA終了処理
    waitObjIE.objIE.Quit 'objIEを終了させる
    MsgBox ExitMsg

End Sub

Sub getBookList(htmlDoc As HTMLDocument, i As Integer, URLCol As Collection, Login As BookdataLogin, arrBooksId As Variant)
    
    '詳細ページURLを取得
    Dim Bookdata As HTMLDivElement 'レコード単位データ
    Dim detailField As HTMLDivElement '詳細フィールドデータ
    Dim BookdataURL As String '詳細ページURL
    Dim BookdataURLDir As String '詳細ページディレクトリ
    Dim checkId As Variant '書籍ID
    Dim checkIdFlag As Boolean '書籍有無
    checkIdFlag = False
    
    'URLからID取得
    Dim getIdData As Variant
    Dim getIdElement As Long
    Dim getBookdataId As Long
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        
        '--detail情報からデータ取得
        
            '--detailを取得
            Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0)
    
            '詳細ページURL
            BookdataURL = detailField.getElementsByTagName("a")(0) 'URL取得
            BookdataURLDir = Replace(BookdataURL, Login.Domain, "") 'ディレクトリ取得
            
            'ディレクトリからID取得
            getIdData = Split(BookdataURLDir, "/") 'URL要素取得
            getIdElement = UBound(getIdData)  'URL要素確認
            getBookdataId = getIdData(getIdElement) 'URLから番号取得
            
            'arrBooksIdにある場合は書籍があるので除外
            For Each checkId In arrBooksId
                If checkId = getBookdataId Then
                    checkIdFlag = True
                    Exit For
                End If
            Next checkId
            
            If checkIdFlag = False Then URLCol.Add BookdataURLDir
            checkIdFlag = False
        '--detail情報からデータ取得ここまで
        
        '列番号処理
        i = i + 1
    Next Bookdata

End Sub

Sub getDetailBookdata(SWSheet As Worksheet, objIE As InternetExplorer, URLCol As Collection, Login As BookdataLogin, i As Long)

    '詳細ページURLから詳細内容を取得
    
    'データ取得URL
    Dim DocContent As HTMLDivElement 'HTMLコンテンツ処理
    Dim DocColumn As HTMLDivElement 'column情報
    Dim j As Long '書き出し用行列処理

    Dim URLi As Long '詳細URL読み込み行番号処理
    URLi = 1

    'URL取得総数確認
    Dim fornumber As Long
    fornumber = URLCol.Count
    
    '画像処理
    Dim DocPicture As HTMLDivElement
    Dim ImgURL As HTMLImg
    Dim ActCell As Range

    'ID取得
    Dim GetUrl As String '詳細ページURL
    Dim GetUrlData() As String '詳細ページURL,Splitデータ
    Dim GetUrlElement As Integer 'URLSplit要素数
    Dim GetID As Integer 'ID番号

    '詳細ページを開いて中のデータを取得
    Do
        i = i + 1
        '次ページURL取得
        Login.ProcessDir = URLCol(URLi)
        
        '次ページへアクセス
        Login.CheckLogin
        j = 1
        
        '1列目にID番号表示
        GetUrl = Login.ProcessDir 'URL取得
        GetUrlData = Split(GetUrl, "/")  'URL要素取得
        GetUrlElement = UBound(GetUrlData)  'URL要素確認
        GetID = GetUrlData(GetUrlElement)  'URLから番号取得
        SWSheet.Cells(i, j).Value = GetID
        j = j + 1
        
        '2列目に画像表示
        Set DocPicture = Login.htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
        Set ActCell = SWSheet.Cells(i, j)
        
        SWSheet.Shapes.AddPicture _
          fileName:=ImgURL.src, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=ActCell.Left, _
            Top:=ActCell.Top, _
            Width:=100, _
            Height:=100
        j = j + 1
        
        '3列目以降にテキスト表示
        For Each DocContent In Login.htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        'カウント追加
        URLi = URLi + 1
        
    'URL要素数を超える場合はループ終了
    Loop Until URLi > fornumber
    
End Sub
