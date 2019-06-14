Attribute VB_Name = "Module_more_books"
Option Explicit

Sub getBookdatas()

    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = True 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        Dim Pagination As HTMLUListElement 'HTML�y�[�W�l�[�V����
        Dim PagiLink As HTMLAnchorElement '���y�[�W�����N
        '��ƃ��[�N�V�[�g
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("�e�X�g")
        '�f�[�^�擾URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        '�J��Ԃ�����
        Dim i As Integer
        i = 1
    
    'OpenPage������Ԃ̓��[�v���đ�����
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IE��URL���J��
        
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        
        OpenPage = ""
        
        '���Џ��擾����
        Call getBookList(SWSheet, htmlDoc, i)
        
        '�N���X��(pagination)�̎擾
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            '�y�[�W�l�[�V����������ꍇ�͎擾����
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    OpenPage = PagiLink.href
                End If
            Next PagiLink
        
        End If
        
    Loop '���[�v�G���h

    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"

End Sub

Sub getBookList(SWSheet As Worksheet, htmlDoc As HTMLDocument, i As Integer)
    
    Dim Bookdata As HTMLDivElement '���R�[�h�P�ʃf�[�^
    Dim detailField As HTMLDivElement '�ڍ׃t�B�[���h�f�[�^
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        SWSheet.Cells(i + 1, 1).Value = i
        
        Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0) '--detail���擾

        SWSheet.Cells(i + 1, 2).Value = detailField.getElementsByClassName("list-book-title")(0).innerText
        i = i + 1
    Next Bookdata

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub

