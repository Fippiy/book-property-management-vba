Attribute VB_Name = "Module_deleteBookDataID"
Option Explicit

Sub deleteBookdataISBN()

    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = True 'IE��\��
'        objIE.Visible = False 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        '��ƃ��[�N�V�[�g
        Dim DelBookSheet As Worksheet 'DeleteBookWorksheet
        Set DelBookSheet = ThisWorkbook.Worksheets("���Џ��폜")
        '�폜ID
        Dim DelID As Collection
        Set DelID = New Collection
        '�f�[�^�擾URL
        Dim DelBookPageBase As String
        Dim DelBookPage As String
        Dim DelBookURL As String
        DelBookPageBase = "https://protected-fortress-61913.herokuapp.com/book/"
        '�J��Ԃ�����
        Dim i As Integer
        i = 2 '2�s�ڂ��琔�l�擾
        '�����������b�Z�[�W
        Dim ExitMsg As String
        
    '�폜ID�擾
    Do Until DelBookSheet.Cells(i, 1).Value = ""
        DelID.Add DelBookSheet.Cells(i, 1).Value
        i = i + 1
    Loop
        
        
    '�폜ID���ɏ���
    If DelID.Count = 0 Then
        
        ExitMsg = "�폜ID������܂���"
    
    Else
        
        i = 1 '�J��Ԃ�������
        Dim objHTTP As Object 'HTTP�`�F�b�N�p�I�u�W�F�N�g
        Dim HTTPStatus As Integer
        Do
            DelBookPage = DelBookPageBase & DelID(i) '�폜����URL�擾
            
            'URL�w���̊m�F
            Set objHTTP = CreateObject("MSXML2.XMLHTTP") 'IXMLHTTPRequest�I�u�W�F�N�g����(���C�u�����Ȃ�)
            objHTTP.Open "HEAD", DelBookPage, False 'IXMLHTTPRequest�I�u�W�F�N�g������
            objHTTP.send 'IXMLHTTPRequest���N�G�X�g���M
            Do While objHTTP.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
                DoEvents
            Loop
            HTTPStatus = objHTTP.Status 'HTTP���N�G�X�g���ʊi�[
        
            
            
            'URL���J���ăI�u�W�F�N�g�擾
            objIE.navigate DelBookPage 'IE��URL���J��
            Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
            Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
                
            '���Ђ�����
            htmlDoc.getElementsByClassName("nav-btn delete")(0).Click
            '�폜��̏���
            Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
            DelBookURL = objIE.document.URL & "/" '�ǂݍ��݌��URL�擾
            
            If DelBookURL = DelBookPageBase Then
                DelBookSheet.Cells(i + 1, 2).Value = "�폜���܂���"
            Else
                DelBookSheet.Cells(i + 1, 2).Value = "�폜�ł��܂���ł���"
            End If
            
            i = i + 1 '���f�[�^�����J�n����
        Loop Until i > DelID.Count
        
        ExitMsg = "�폜�������������܂���"
    
    End If
        

    'VBA�I������
'    objIE.Quit 'objIE���I��������
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
