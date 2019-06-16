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
    
    '作業ワークシート指定
    Dim SWSheet As Worksheet 'ScrapingWorksheet
    Set SWSheet = ThisWorkbook.Worksheets("スクレイピング")
    
    
    Dim Bookdata As Object 'レコード単位データ
    Dim detailField As Variant '詳細フィールドデータ
    Dim geturl As String '詳細ページURL
    Dim GetUrlData() As String '詳細ページURL,Splitデータ
    Dim GetUrlElement As Integer 'URLSplit要素数
    Dim GetID As Integer 'ID番号
    
    Dim ImgURL As String '画像URL
    Dim Img As Variant '画像オブジェクト
    Dim ActCell As Object '画像出力セル
    
    Dim i As Integer
    i = 1
    
    
    'レコード単位でデータ出力
    ' book-table__listの要素をBookdataとして処理
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
    
        '--detail部を取得してそれぞれ反映
            
            detailField = Bookdata.getElementsByClassName("book-table__list--detail") '--detailを取得
            
            'タイトル名
            SWSheet.Cells(i + 1, 2).Value = detailField.getElementsByClassName("list-book-title")(0).innerText
            
            '詳細テキスト
            SWSheet.Cells(i + 1, 3).Value = detailField.getElementsByClassName("list-book-detail")(0).innerText
            
            
            '詳細ページURL
            geturl = detailField.getElementsByTagName("a") 'URL取得
            SWSheet.Cells(i + 1, 4).Value = geturl  '取得URL反映
            GetUrlData = Split(geturl, "/")  'URL要素取得
            GetUrlElement = UBound(GetUrlData)  'URL要素確認
            GetID = GetUrlData(GetUrlElement)  'URLから番号取得
            SWSheet.Cells(i + 1, 1).Value = GetID
        
        '--detail部を取得してそれぞれ反映ここまで
        
                
        '画像処理

        Img = Bookdata.getElementsByTagName("img")  '画像取得
        ImgURL = Img.src '画像URL
        Set ActCell = SWSheet.Cells(i + 1, 5)

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
        
        
        '次のレコードの行番号
        i = i + 1
    Next Bookdata

    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"

End Sub
Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
