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
        '�f�[�^�擾URL
        Dim DelBookPageBase As String
        Dim DelBookPage As String
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
        Do
            DelBookPage = DelBookPageBase & DelID(i)
            'URL���J���ăI�u�W�F�N�g�擾
            objIE.navigate DelBookPage 'IE��URL���J��
            Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
            Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
            '���Ђ�����
            htmlDoc.getElementsByClassName("nav-btn delete")(0).Click
            i = i + 1
        Loop Until i > DelID.Count
        
        ExitMsg = "���Џ����폜���܂���"
    
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
