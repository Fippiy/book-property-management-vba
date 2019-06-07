Attribute VB_Name = "Module_books"
Option Explicit

Sub getBookdata()
    'ベース作成+セル表示
    Dim objIE As InternetExplorer 'IEオブジェクトを準備
    Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット

    objIE.Visible = True 'IEを表示

    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IEでURLを開く

    Call WaitResponse(objIE) '読み込み待ち

    Dim htmlDoc As HTMLDocument 'HTMLドキュメントオブジェクトを準備
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

    'メッセージボックスに取得クラスの最初の文字を出力
'     MsgBox htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
    
    
    'イミディエイトに指定クラスの全取得テキストを表示
'    Dim Str As Variant
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
'        Debug.Print "出力：" & Str.innerHTML
'    Next Str
    
    
    'シート上に指定クラスの全取得テキストを表示
    Dim Str As Variant
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("スクレイピング").Cells(i + 1, 1).Value = i
        Worksheets("スクレイピング").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
'        Debug.Print "出力：" & Str.innerHTML
    Next Str
    
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("スクレイピング").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str
'    Debug.Print "データ取得が完了しました。"
'    objIE.Quit
End Sub

Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
'        Debug.Print objIE.Busy
'        Debug.Print objIE.readyState
        DoEvents
    Loop
End Sub

