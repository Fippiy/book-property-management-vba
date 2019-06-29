Attribute VB_Name = "Module_deleteBookDataID"
Option Explicit

Sub deleteBookdataISBN()
    
    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        '��ƃ��[�N�V�[�g
        Dim DelBookSheet As Worksheet 'DeleteBookWorksheet
        Set DelBookSheet = ThisWorkbook.Worksheets("���Џ��폜")
        '�폜ID
        Dim DelID As Collection
        Set DelID = New Collection
        '�ڍ׃y�[�WURL�x�[�X
        Dim DelBookPageBase As String
        DelBookPageBase = "https://protected-fortress-61913.herokuapp.com/book/"
        '�J��Ԃ�����
        Dim i As Integer
        i = 2 '2�s�ڂ��琔�l�擾
        '�����������b�Z�[�W
        Dim ExitMsg As String
    
    '�폜ID�����[�N�V�[�g����擾
    Do Until DelBookSheet.Cells(i, 1).Value = ""
        DelID.Add DelBookSheet.Cells(i, 1).Value
        i = i + 1
    Loop
    
    '�擾����ID�R���N�V�������珈�������{
    If DelID.Count = 0 Then
        
        ExitMsg = "�폜ID������܂���"
    
    Else
        
        '�I�u�W�F�N�g�錾
        Dim BookProcess As Range '�������ʊi�[�Z��
        Dim DelBookPage As String '�폜���Ѓy�[�W
        Dim DelBookURLAfter As String '�폜��J�ڂ���T�C�g��URL
        Dim HTTPStatus As Integer 'HTTP���N�G�X�g�X�e�[�^�X
        
        i = 1 '�J��Ԃ�������
        
        'URL���ɍ폜�����{
        
        Do
            
            DelBookPage = DelBookPageBase & DelID(i) '�폜����URL�擾
            Set BookProcess = DelBookSheet.Cells(i + 1, 2) '�������ʔ��f�Z��
            
            'HTTP���N�G�X�g�X�e�[�^�X���m�F
            Call CheckHTTPRequest(DelBookPage, HTTPStatus)
            
            'HTTP���N�G�X�g=200�Ȃ�A�폜���������{
            If HTTPStatus = 200 Then
                
                'URL���J���ăI�u�W�F�N�g�擾
                objIE.navigate DelBookPage 'IE��URL���J��
                Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
                Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
                
                '���Ђ�����
                htmlDoc.getElementsByClassName("nav-btn delete")(0).Click
                '�폜��̏���
                Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
                DelBookURLAfter = objIE.document.URL & "/" '�ǂݍ��݌��URL�擾
                
                '���ʂ����[�N�V�[�g�֏o��
                If DelBookURLAfter = DelBookPageBase Then
                    DelBookSheet.Range(BookProcess.Address).Value = "�폜���܂���"
                Else
                    DelBookSheet.Range(BookProcess.Address).Value = "�폜�ł��܂���ł���"
                End If
            
            'HTTP���N�G�X�g<>200�́A�G���[�Ƃ��Č��ʂ�Ԃ�
            Else
                
                DelBookSheet.Range(BookProcess.Address).Value = "�ڑ��G���[(" & HTTPStatus & ")"
            
            End If
            
            i = i + 1 '���f�[�^�����J�n����
        
        Loop Until i > DelID.Count
        
        ExitMsg = "�폜�������������܂���"
    
    End If
    
    'VBA�I������
    objIE.Quit 'objIE���I��������
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub

Sub CheckHTTPRequest(DelBookPage As String, HTTPStatus As Integer)
    Dim objHTTP As Object 'HTTP�`�F�b�N�p�I�u�W�F�N�g

    Set objHTTP = CreateObject("MSXML2.XMLHTTP") 'IXMLHTTPRequest�I�u�W�F�N�g����(���C�u�����Ȃ�)
    objHTTP.Open "HEAD", DelBookPage, False 'IXMLHTTPRequest�I�u�W�F�N�g������
    objHTTP.send 'IXMLHTTPRequest���N�G�X�g���M
    Do While objHTTP.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    HTTPStatus = objHTTP.Status 'HTTP���N�G�X�g���ʊi�[
End Sub
