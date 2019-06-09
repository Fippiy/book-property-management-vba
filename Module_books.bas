Attribute VB_Name = "Module_books"
Option Explicit


Sub getBookdata()
    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
    objIE.Visible = True 'IEを表示
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IEでURLを開く
    
    Call WaitResponse(objIE) '読み込み待ち
    
    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    ' シート上に指定クラスの全取得テキストを表示
    Dim Str As Object
    Dim i As Integer
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
    i = 1
    Dim GetUrl As String
    Dim GetUrlData() As String
    Dim GetUrlElement As Integer
    Dim GetID As Integer
    
    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
        GetUrl = Str.getElementsByTagName("a")  'URL取得
        Worksheets("スクレイピング").Cells(i + 1, 4).Value = GetUrl  '取得URL反映
        GetUrlData = Split(GetUrl, "/")  'URL要素取得
        GetUrlElement = UBound(GetUrlData)  'URL要素確認
        GetID = GetUrlData(GetUrlElement)  'URLから番号取得
        Worksheets("スクレイピング").Cells(i + 1, 1).Value = GetID  'ワークシートへ反映
        i = i + 1  '次の行指定
    Next Str



    '画像URL取得
    
    '画像用変数
    Dim imgURL As String '画像URL
    Dim Img As Object '画像オブジェクト
    Dim toppix As Long '位置ピクセル

'    '1個をサンプル取得
'    imgURL = htmlDoc.images(0).src
'    Worksheets("スクレイピング").Cells(2, 5).Value = imgURL
'    Worksheets("スクレイピング").Shapes.AddPicture _
'        fileName:=imgURL, _
'            LinkToFile:=True, _
'                SaveWithDocument:=True, _
'                Left:=0, _
'                Top:=0, _
'                Width:=100, _
'                Height:=80
    
'    'URLのみ取得
'    i = 1
'    For Each IMG In htmlDoc.images 'イメージ取得
'        imgURL = IMG.src '変数格納
'        Worksheets("スクレイピング").Cells(i + 1, 5).Value = imgURL '取得URL反映
'        i = i + 1
'    Next IMG
    
    
    
    Dim ActCell As Object

    i = 1
    toppix = 0
    For Each Img In htmlDoc.images 'イメージ取得
        imgURL = Img.src '変数格納
        Set ActCell = Worksheets("スクレイピング").Cells(i + 1, 5)
        ActCell.Value = imgURL  '取得URL反映

        '画像を表示
        Worksheets("スクレイピング").Shapes.AddPicture _
            fileName:=imgURL, _
                LinkToFile:=True, _
                    SaveWithDocument:=True, _
                    Left:=0, _
                    Top:=0 + toppix, _
                    Width:=100, _
                    Height:=100

'        '画像を表示、セルピクセル取得
'        Worksheets("スクレイピング").Shapes.AddPicture _
'            fileName:=imgURL, _
'                LinkToFile:=True, _
'                    SaveWithDocument:=True, _
'                    Left:=ActCell.Left, _
'                    Top:=ActCell.Top, _
'                    Width:=100, _
'                    Height:=100

        i = i + 1
        toppix = toppix + 100
    Next Img


    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"
End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
Sub picture1()
    'ローカルディスク上の画像ファイルを表示
    Worksheets("スクレイピング").Shapes.AddPicture _
        fileName:="Z:\FierVega\ariawase-master\bin\test.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0, _
                Width:=100, _
                Height:=80
End Sub
Sub picture2()
    'coverURLを指定してファイル表示
    Worksheets("スクレイピング").Shapes.AddPicture _
        fileName:="https://cover.openbd.jp/9784797398892.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0, _
                Width:=100, _
                Height:=80
End Sub
Sub picture3()
    'coverURLを指定してファイル表示
    Worksheets("スクレイピング").Shapes.AddPicture _
        fileName:="https://cover.openbd.jp/9784797398892.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=330, _
                Top:=40, _
                Width:=100, _
                Height:=80
End Sub
