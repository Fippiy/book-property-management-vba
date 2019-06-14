Attribute VB_Name = "Module_pagination"
Option Explicit

Sub pagecheck()
    
    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
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
        Dim page As Integer
        i = 2
        page = 1
    
    'OpenPage������Ԃ̓��[�v���đ�����
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IE��URL���J��
        
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        
        OpenPage = ""
        
        '�N���X��(pagination)�̎擾
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            '�y�[�W�l�[�V����������ꍇ�͎擾����
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
    Loop '���[�v�G���h

    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub


