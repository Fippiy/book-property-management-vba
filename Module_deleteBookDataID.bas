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
        Dim DelID As Long
        DelID = DelBookSheet.Cells(2, 1).Value
        '�f�[�^�擾URL
        Dim DelBookPageBase As String
        Dim DelBookPage As String
        DelBookPageBase = "https://protected-fortress-61913.herokuapp.com/book/"
        DelBookPage = DelBookPageBase & DelID
        '�J��Ԃ�����
        Dim i As Integer
        i = 1
        '�����������b�Z�[�W
        Dim ExitMsg As String
        
    'URL���J���ăI�u�W�F�N�g�擾
    objIE.navigate DelBookPage 'IE��URL���J��
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g

'    '���Ђ�����
    htmlDoc.getElementsByClassName("nav-btn delete")(0).Click

    'VBA�I������
    objIE.Quit 'objIE���I��������
    ExitMsg = "test"
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
