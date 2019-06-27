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
        Dim ISBNAllCount As Integer 'ISBN����
        Const LimitEntry As Integer = 20 '�t�H�[������ISBN���
        Dim EntryISBN() As String '�t�H�[�����͗pcsv(�����)
        Dim MaxRepeat As Long 'ISBN������
        Dim LastISBNCount As Integer '�ŏIISBN����
        Dim ElementCounter As Long '�v�f�擾�J�E���^
        '�f�[�^�擾URL
        Dim InputISBNPage As String
        InputISBNPage = "https://protected-fortress-61913.herokuapp.com/book/isbn_some_input"
        '�J��Ԃ�����
        Dim i As Integer
        Dim j As Integer
        i = 2
        '�����������b�Z�[�W
        Dim ExitMsg As String
        
    'ISBN�R�[�h�擾
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop
    ISBNAllCount = InputISBN.Count 'ISBN�R�[�h����

    'Web������ɓo�^����������J���}��؂�e�L�X�g������
    
        'ISBN�����񐔎Z�o
        MaxRepeat = Application.RoundUp(ISBNAllCount / LimitEntry, 0) '�J��Ԃ���
        LastISBNCount = ISBNAllCount Mod LimitEntry '�J��Ԃ����X�g�擾����
        ReDim EntryISBN(MaxRepeat - 1) '�z��Ƃ��ėv�f�w�肵�čĐ錾
        ElementCounter = 1 '�v�f�擾�J�E���^�����l
        
        'Web����������ɏ����ł���悤�ɂ���
        For j = 0 To MaxRepeat - 1
        
            i = 1 '�J��Ԃ��ϐ�������
            '�J���}��؂�e�L�X�g����(�SISBN or ��������܂�)
            Do
                EntryISBN(j) = EntryISBN(j) & InputISBN(ElementCounter) 'ISBN�R�[�h��v�f�Ƃ��Ēǉ�
                '�������orISBN�������X�g�̓J���}�Ȃ�
                If ElementCounter <> ISBNAllCount And i <> LimitEntry Then EntryISBN(j) = EntryISBN(j) & ","
                ElementCounter = ElementCounter + 1
                i = i + 1
            Loop Until i > LimitEntry Or ElementCounter > ISBNAllCount
        Next j
        
    '�S�����������܂ŌJ��Ԃ�
        
    '�J���}��؂�e�L�X�g��S�Ĕ��f������
    i = 2 '���ʏo�̓e�L�X�g�}���ʒu������
        
    For j = 0 To MaxRepeat - 1
        
        '�t�H�[�����J��
        objIE.navigate InputISBNPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        
        '�t�H�[������
        htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN(j)
        htmlDoc.getElementsByClassName("send isbn")(0).Click
    
        '�t�H�[������HTML�擾
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        
        '�t�H�[���������ʎ擾
        Call getISBNAnswers(htmlDoc, ISSheet, i)
    Next j

    '�S�����������܂ŌJ��Ԃ�

    'VBA�I������
    objIE.Quit 'objIE���I��������
    ExitMsg = "�o�^�������������܂����B"
    MsgBox ExitMsg

End Sub

Sub getISBNAnswers(htmlDoc As HTMLDocument, ISSheet As Worksheet, i As Integer)
    
    '���ʏ����ϐ�
    Dim ResultRecord As HTMLDivElement 'Record�P�ʃf�[�^
    Dim ResultTitle As HTMLDivElement '�^�C�g����
    Dim ResultText As HTMLDivElement '���ʃe�L�X�g
    
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

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
