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
    
'    test = htmlDoc.getElementsByClassName("list-book-title")
    
'     メッセージボックスに取得クラスの最初の文字を出力
'    MsgBox htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
    
'     メッセージボックスに取得クラスの子要素を取得出力
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).outerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-detail")(0).innerHTML




'    クラス名内の要素配下のタグ名の中身を取得
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0).innerHTML




'    クラス名内の要素配下のaタグ要素を取得(URL)
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0)




'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).outerHTML.getElementsByTagName("a")
    
'    データ取得まとめ
'    Debug.Print "1." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print "2." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print "3." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("h3")
'    Debug.Print "4." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("p")
    
    
    ' イミディエイトに指定クラスの全取得テキストを表示
'    Dim Str As Variant
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
''        Debug.Print "出力：" & Str.innerHTML
'    Next Str




'    test = htmlDoc.getElementsByClassName("list-book-title")
'    Debug.Print UBound(test, 2)








''    ' シート上に指定クラスの全取得テキストを表示
    Dim Str As Variant
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
'        Worksheets("スクレイピング").Cells(i + 1, 1).Value = i
        Worksheets("スクレイピング").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
'        Debug.Print "出力：" & Str.innerHTML
    Next Str


    '書籍詳細
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("スクレイピング").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str


'    i = 1
'    Dim Arr As Variant
'    ReDim Arr(20)
'    Arr = htmlDoc.getElementsByClassName("list-book-detail")
'    Range("A2").Value = Arr


    'URL
    i = 1
'    htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0)
'    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
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


'    Dim test As String
''    test = "aa/bb/cc/dd"
'    test = "https://protected-fortress-61913.herokuapp.com/book"
'    test2 = Split(test, "/")
'    Debug.Print test
'    Debug.Print test2(0)
'    Debug.Print test2(1)
'    Debug.Print test2(2)
'    Debug.Print test2(3)
'    Debug.Print UBound(test2)


'book-table__list--detail




    objIE.Quit 'objIEを終了させる
    Debug.Print "データ取得が完了しました。"
End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
