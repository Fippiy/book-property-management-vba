Attribute VB_Name = "Module_pagination"
Option Explicit

Sub pagecheck()
    
    'オブジェクト設定
        'IE
        Dim objIE As InternetExplorer 'IEオブジェクトを準備
        Set objIE = CreateObject("Internetexplorer.Application") '新しいIEオブジェクトを作成してセット
        objIE.Visible = False 'IEを表示
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML全体
        Dim Pagination As HTMLUListElement 'HTMLページネーション
        Dim PagiLink As HTMLAnchorElement '次ページリンク
        '作業ワークシート
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("テスト")
        'データ取得URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        '繰り返し処理
        Dim i As Integer
        Dim page As Integer
        i = 2
        page = 1
    
    'OpenPageがある間はループして続ける
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IEでURLを開く
        
        Call WaitResponse(objIE) '読み込み待ち
        
        Set htmlDoc = objIE.document 'objIEで読み込まれているHTMLドキュメントをセット
        
        OpenPage = ""
        
        'クラス名(pagination)の取得
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            'ページネーションがある場合は取得処理
            For Each PagiLink In Pagination.getElementsByTagName("a")
                Cells(i, 1).Value = page
                Cells(i, 2).Value = PagiLink.outerHTML
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    OpenPage = PagiLink.href
                    Cells(i, 3).Value = OpenPage
                End If
                i = i + 1
            Next PagiLink
        
        End If
        

        page = page + 1
    Loop 'ループエンド

    objIE.Quit 'objIEを終了させる
    MsgBox "データ取得が完了しました。"

End Sub

Sub WaitResponse(objIE As Object) 'Webブラウザ表示完了待ち
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
End Sub


