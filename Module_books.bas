Attribute VB_Name = "Module_books"
Option Explicit


Sub getBookdata()
    '旧コメント削除

    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
    
    objIE.Visible = False 'IEを表示、FalseでIE表示なし
    
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IEでURLを開く
    
    Call WaitResponse(objIE) '読み込み待ち

    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
  
''    データ取得まとめ
'    Debug.Print "1." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print "2." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print "3." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("h3")
'    Debug.Print "4." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("p")
    

    ' シート上に指定クラスの全取得テキストを表示
    Dim Str As Object
    Dim i As Integer
    
    'タイトル名取得
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("スクレイピング").Cells(i + 1, 1).Value = i
        Worksheets("スクレイピング").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
    Next Str

    'detail取得
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("スクレイピング").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str

    'URL取得
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
        Worksheets("スクレイピング").Cells(i + 1, 4).Value = Str.getElementsByTagName("a")
        i = i + 1
    Next Str

    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"

End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub

