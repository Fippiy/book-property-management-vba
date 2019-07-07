Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    '===��VBA�S�̃I�u�W�F�N�g�ݒ聫===
        
        'IE�I�u�W�F�N�g
        Dim waitObjIE As waitObjIE 'IE�ǂݍ��ݑ҂����W���[���쐬
        Set waitObjIE = New waitObjIE
        Set waitObjIE.objIE = CreateObject("Internetexplorer.Application")
        waitObjIE.objIE.Visible = False
        
        '�f�[�^�擾URL
        Dim Login As BookdataLogin '���O�C���N���X���W���[���쐬
        Set Login = New BookdataLogin
        Login.Domain = "https://protected-fortress-61913.herokuapp.com/" '�h���C���i�[
        Login.ProcessDir = "book/isbn_some_input" '�f�B���N�g���w��
        Login.CheckFirstLogin = True '���O�C���`�F�b�N�t���O
        Set Login.waitObjIE = waitObjIE 'IE�I�u�W�F�N�g��Login�Ɉ��n
            
    '===��VBA�S�̃I�u�W�F�N�g�ݒ聪===
            
    '���O�C����ԃ`�F�b�N
    Login.CheckLogin
    
    '===�������p�I�u�W�F�N�g�ݒ聫===

        '��ƃ��[�N�V�[�g�ݒ�
        Dim ISSheet As Worksheet 'ISBNWorksheet
        Set ISSheet = ThisWorkbook.Worksheets("ISBN")
            
        '�o�^ISBN�R�[�h�擾�ݒ�
        Dim InputISBN As Collection '�f�[�^�擾
        Set InputISBN = New Collection
        Dim ISBNAllCount As Integer 'ISBN����
        Const LimitEntry As Integer = 20 '�t�H�[������ISBN���
        Dim EntryISBN() As String '�t�H�[�����͗pcsv(�����)
        Dim MaxRepeat As Long 'ISBN������
        Dim LastISBNCount As Integer '�ŏIISBN����
        Dim ElementCounter As Long '�v�f�擾�J�E���^
            
        '�J��Ԃ�����
        Dim i As Integer
        Dim j As Integer
        
        '�o�̓��b�Z�[�W
        Dim ExitMsg As String
        
    '===�������p�I�u�W�F�N�g�ݒ聪===
    
    
    '���[�N�V�[�g����ISBN�R�[�h�擾
    i = 2 '1�s�ڃC���f�b�N�X�Ȃ̂�2�s�ڂ���
    Do Until ISSheet.Cells(i, 2).Value = ""
        InputISBN.Add ISSheet.Cells(i, 2).Value
        i = i + 1
    Loop
    ISBNAllCount = InputISBN.Count 'ISBN�R�[�h����


    'Web������ɓo�^����������J���}��؂�e�L�X�g������
    
        'ISBN�����񐔎Z�o
        MaxRepeat = Application.RoundUp(ISBNAllCount / LimitEntry, 0) '�J��Ԃ���
        LastISBNCount = ISBNAllCount Mod LimitEntry '�J��Ԃ����X�g�擾����
        ReDim EntryISBN(1 To MaxRepeat) '�z��Ƃ��ėv�f�w�肵�čĐ錾

        ElementCounter = 1 '�v�f�擾�J�E���^�����l
        
        'Web����������ɏ����ł���悤�ɂ���
        For j = 1 To MaxRepeat
        
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
        
    For j = 1 To MaxRepeat
        
        '���O�C����ԃ`�F�b�N��HTML�擾
        Login.CheckLogin
        
        '�t�H�[������
        Login.htmlDoc.getElementsByClassName("form-input__detail")(0).Value = EntryISBN(j)
        Login.htmlDoc.getElementsByClassName("send isbn")(0).Click
    
        '�t�H�[������HTML�擾
        waitObjIE.WaitResponse '�ǂݍ��ݑ҂�
        Set Login.htmlDoc = waitObjIE.objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        
        '�t�H�[���������ʎ擾
        Call getISBNAnswers(Login.htmlDoc, ISSheet, i)
    
    Next j

    '�S�����������܂ŌJ��Ԃ�

    'VBA�I������
    waitObjIE.objIE.Quit 'objIE���I��������
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
