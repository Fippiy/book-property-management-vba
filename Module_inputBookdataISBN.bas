Attribute VB_Name = "Module_inputBookdataISBN"
Option Explicit


Sub inputBookdataISBN()


    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
'        objIE.Visible = False 'IE��\��
        objIE.Visible = True 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        Dim Pagination As HTMLUListElement 'HTML�y�[�W�l�[�V����
        '��ƃ��[�N�V�[�g
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
        '�f�[�^�擾URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn"
        '�J��Ԃ�����
        Dim i As Integer
        i = 1
        '�����������b�Z�[�W
        Dim ExitMsg As String
        
    'URL���J���ăI�u�W�F�N�g�擾
    objIE.navigate InputISBNPage 'IE��URL���J��
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g

    '�t�H�[������
    htmlDoc.getElementsByClassName("form-input__input")(0).Value = "1234567890123"

    'VBA�I������
'    objIE.Quit 'objIE���I��������
    ExitMsg = "test"
    MsgBox ExitMsg


End Sub


Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
