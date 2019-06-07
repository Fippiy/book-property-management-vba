Attribute VB_Name = "Module_books"
Option Explicit

Sub getBookdata()
    'URL取得
    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
    
    objIE.Visible = False 'IEを表示
    
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IEでURLを開く
    
    Call WaitResponse(objIE) '読み込み待ち

    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    'メッセージボックスに取得クラスの最初の文字を出力
'    MsgBox htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-title").innerHTML
'    Cells(2, 2).Value = htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML

    ' シート上に指定クラスの全取得テキストを表示
'    Dim Str As Object
''    Dim i As Integer
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
'        Debug.Print "出力：" & Str.innerHTML
'    Next Str

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
    
''    データ取得まとめ
'    Debug.Print "1." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print "2." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print "3." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("h3")
'    Debug.Print "4." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("p")
    
    
    ' イミディエイトに指定クラスの全取得テキストを表示
'    Dim Str As Variant
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
''        Debug.Print "出力：" & Str.innerHTML
'    Next Str




'    ' シート上に指定クラスの全取得テキストを表示
    Dim Str As Object
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("スクレイピング").Cells(i + 1, 1).Value = i
        Worksheets("スクレイピング").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
'        Debug.Print "出力：" & Str.innerHTML
    Next Str

'    i = 1
'    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
'        Worksheets("スクレイピング").Cells(i + 1, 3).Value = Str.innerHTML
'        i = i + 1
'    Next Str

'    i = 1
''    htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0)
''    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
'    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
'        Worksheets("スクレイピング").Cells(i + 1, 4).Value = Str.getElementsByTagName("a")
'        i = i + 1
'    Next Str


'book-table__list--detail


    objIE.Quit 'objIEを終了させる
'    Debug.Print "データ取得が完了しました。"
End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

