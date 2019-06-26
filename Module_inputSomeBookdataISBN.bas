Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        '作業ワークシート
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
        '登録ISBN
        Dim InputISBN As Collection 'データ取得
        Set InputISBN = New Collection
        Dim EntryISBN As String 'フォーム入力用csv
        Const LimitEntry As Integer = 20 'フォーム入力ISBN上限
        'データ取得URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn_some_input"
        '繰り返し処理
        Dim i As Integer
        i = 2
        '処理完了メッセージ
        Dim ExitMsg As String
        
    'URLを開いてオブジェクト取得
    objIE.navigate InputISBNPage 'IEでURLを開く
    Call WaitResponse(objIE) '読み込み待ち
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット

    'ISBNコード取得
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop

    'カンマ区切りテキスト生成(全ISBN or 上限件数まで)
    i = 1 '繰り返し変数初期化
    Do
        EntryISBN = EntryISBN & InputISBN(i)
        If i <> InputISBN.Count Then EntryISBN = EntryISBN & ","
        i = i + 1
    Loop Until i > InputISBN.Count Or i > LimitEntry

    'フォーム入力
    htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN
    htmlDoc.getElementsByClassName("send isbn")(0).Click

    'フォーム結果HTML取得
    Call WaitResponse(objIE) '読み込み待ち
    Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
    
    'フォーム処理結果取得
    Call getISBNAnswers(htmlDoc, ISSheet)

    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    ExitMsg = "登録処理が完了しました。"
    MsgBox ExitMsg

End Sub

Sub getISBNAnswers(htmlDoc As HTMLDocument, ISSheet As Worksheet)
    
    '結果処理変数
    Dim ResultRecord As HTMLDivElement 'Record単位データ
    Dim ResultTitle As HTMLDivElement 'タイトル名
    Dim ResultText As HTMLDivElement '結果テキスト
    Dim i As Long
    i = 2
    
    For Each ResultRecord In htmlDoc.getElementsByClassName("isbn-result__box")
    
        Set ResultText = ResultRecord.getElementsByClassName("isbn-result__box--msg")(0)
        Set ResultTitle = ResultRecord.getElementsByClassName("isbn-result__box--title")(0)
        
        ISSheet.Cells(i, 4).Value = ResultText.innerText
        
        If (ResultTitle Is Nothing) = False Then
            ISSheet.Cells(i, 3).Value = ResultTitle.innerText
        End If
                
        i = i + 1
    Next ResultRecord

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub
