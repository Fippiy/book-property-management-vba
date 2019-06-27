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
        Dim ISBNAllCount As Integer 'ISBN総数
        Const LimitEntry As Integer = 20 'フォーム入力ISBN上限
        Dim EntryISBN() As String 'フォーム入力用csv(上限毎)
        Dim MaxRepeat As Long 'ISBN処理回数
        Dim LastISBNCount As Integer '最終ISBN件数
        Dim ElementCounter As Long '要素取得カウンタ
        'データ取得URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn_some_input"
        '繰り返し処理
        Dim i As Integer
        Dim j As Integer
        i = 2
        '処理完了メッセージ
        Dim ExitMsg As String
        
    'ISBNコード取得
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop
    ISBNAllCount = InputISBN.Count 'ISBNコード総数

    'Web上限毎に登録処理をするカンマ区切りテキストを準備
    
        'ISBN処理回数算出
        MaxRepeat = Application.RoundUp(ISBNAllCount / LimitEntry, 0) '繰り返し回数
        LastISBNCount = ISBNAllCount Mod LimitEntry '繰り返しラスト取得件数
        ReDim EntryISBN(MaxRepeat - 1) '配列として要素指定して再宣言
        ElementCounter = 1 '要素取得カウンタ初期値
        
        'Web処理上限毎に処理できるようにする
        For j = 0 To MaxRepeat - 1
        
            i = 1 '繰り返し変数初期化
            'カンマ区切りテキスト生成(全ISBN or 上限件数まで)
            Do
                EntryISBN(j) = EntryISBN(j) & InputISBN(ElementCounter) 'ISBNコードを要素として追加
                '処理上限orISBN総数ラストはカンマなし
                If ElementCounter <> ISBNAllCount And i <> LimitEntry Then EntryISBN(j) = EntryISBN(j) & ","
                ElementCounter = ElementCounter + 1
                i = i + 1
            Loop Until i > LimitEntry Or ElementCounter > ISBNAllCount
        Next j
        
    '全件処理完了まで繰り返し
        
    'カンマ区切りテキストを全て反映させる
    i = 2 '結果出力テキスト挿入位置初期化
        
    For j = 0 To MaxRepeat - 1
        
        'フォームを開く
        objIE.navigate InputISBNPage 'IEでURLを開く
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        
        'フォーム入力
        htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN(j)
        htmlDoc.getElementsByClassName("send isbn")(0).Click
    
        'フォーム結果HTML取得
        Call WaitResponse(objIE) '読み込み待ち
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        
        'フォーム処理結果取得
        Call getISBNAnswers(htmlDoc, ISSheet, i)
    Next j

    '全件処理完了まで繰り返し

    'VBA終了処理
    objIE.Quit 'objIEを終了させる
    ExitMsg = "登録処理が完了しました。"
    MsgBox ExitMsg

End Sub

Sub getISBNAnswers(htmlDoc As HTMLDocument, ISSheet As Worksheet, i As Integer)
    
    '結果処理変数
    Dim ResultRecord As HTMLDivElement 'Record単位データ
    Dim ResultTitle As HTMLDivElement 'タイトル名
    Dim ResultText As HTMLDivElement '結果テキスト
    
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
