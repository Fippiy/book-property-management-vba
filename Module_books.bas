Attribute VB_Name = "Module_books"
Option Explicit


Sub getBookdata()
    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
    objIE.Visible = False 'IEを表示
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IEでURLを開く
    
    Call WaitResponse(objIE) '読み込み待ち
    
    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    Dim Str As Object
    Dim i As Integer
    i = 1
    
    'レコード単位出力(テストシート)
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("book-table__list")
        Worksheets("テスト").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
    Next Str
    
    ' シート上に指定クラスの全取得テキストを表示
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("スクレイピング").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
    Next Str

    '書籍詳細
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("スクレイピング").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str


    'URL
    Dim GetUrl As String
    Dim GetUrlData() As String
    Dim GetUrlElement As Integer
    Dim GetID As Integer

    i = 1
    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
        GetUrl = Str.getElementsByTagName("a")  'URL取得
        Worksheets("スクレイピング").Cells(i + 1, 4).Value = GetUrl  '取得URL反映
        GetUrlData = Split(GetUrl, "/")  'URL要素取得
        GetUrlElement = UBound(GetUrlData)  'URL要素確認
        GetID = GetUrlData(GetUrlElement)  'URLから番号取得
        Worksheets("スクレイピング").Cells(i + 1, 1).Value = GetID  'ワークシートへ反映
        i = i + 1  '次の行指定
    Next Str

    '画像用変数
    Dim imgURL As String '画像URL
    Dim Img As Object '画像オブジェクト
    Dim ActCell As Object '画像出力セル

    i = 1
    For Each Img In htmlDoc.images '画像要素取得
        imgURL = Img.src '画像URL
        Set ActCell = Worksheets("スクレイピング").Cells(i + 1, 5)

        '画像出力セルのピクセルを指定して表示
        Worksheets("スクレイピング").Shapes.AddPicture _
            fileName:=imgURL, _
                LinkToFile:=True, _
                    SaveWithDocument:=True, _
                    Left:=ActCell.Left, _
                    Top:=ActCell.Top, _
                    Width:=100, _
                    Height:=100

        i = i + 1
    Next Img


    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"
End Sub
Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
