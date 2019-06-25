Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        Dim Pagination As HTMLUListElement 'HTML�y�[�W�l�[�V����
        '��ƃ��[�N�V�[�g
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
        '�o�^ISBN
        Dim InputISBN As Collection '�f�[�^�擾
        Set InputISBN = New Collection
        Dim EntryISBN As String '�t�H�[�����͗pcsv
        Const LimitEntry As Integer = 20 '�t�H�[������ISBN���
        '�f�[�^�擾URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn_some_input"
        '�J��Ԃ�����
        Dim i As Integer
        i = 2
        '�����������b�Z�[�W
        Dim ExitMsg As String
        
    'URL���J���ăI�u�W�F�N�g�擾
    objIE.navigate InputISBNPage 'IE��URL���J��
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g

    'ISBN�R�[�h�擾
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop

    '�J���}��؂�e�L�X�g����(�SISBN or ��������܂�)
    i = 1 '�J��Ԃ��ϐ�������
    Do
        EntryISBN = EntryISBN & InputISBN(i)
        If i <> InputISBN.Count Then EntryISBN = EntryISBN & ","
        i = i + 1
    Loop Until i > InputISBN.Count Or i > LimitEntry

    '�t�H�[������
    htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN
    htmlDoc.getElementsByClassName("send isbn")(0).Click

    'VBA�I������
    objIE.Quit 'objIE���I��������
    ExitMsg = "�{��o�^���܂����B"
    MsgBox ExitMsg

End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
