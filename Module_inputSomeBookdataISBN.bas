Attribute VB_Name = "Module_inputSomeBookdataISBN"
Option Explicit

Sub inputSomeBookdataISBN()

    '===��VBA�S�̃I�u�W�F�N�g�ݒ聫===
        
        'IE�I�u�W�F�N�g
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
        'HTML�I�u�W�F�N�g
        Dim htmlDoc As HTMLDocument 'HTML�S��
        '�f�[�^�擾URL
        Dim Domain As String 'Web�h���C����
        Dim ProcessDir As String '�������{�f�B���N�g��
        Domain = "https://protected-fortress-61913.herokuapp.com/"
        ProcessDir = "book/isbn_some_input"
        'VBA���쏉�񃍃O�C���`�F�b�N
        Dim CheckFirstLogin As Boolean '���O�C���`�F�b�N�t���O
        CheckFirstLogin = True
            
    '===��VBA�S�̃I�u�W�F�N�g�ݒ聪===
            
    '���O�C����ԃ`�F�b�N
    Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
        
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
'        ReDim EntryISBN(MaxRepeat - 1) '�z��Ƃ��ėv�f�w�肵�čĐ錾
'            '������1 to MaxRepeat��1�n�܂�̏I���l�w��Ŕz��錾�ł��邼�H
        ReDim EntryISBN(1 To MaxRepeat) '�z��Ƃ��ėv�f�w�肵�čĐ錾

        ElementCounter = 1 '�v�f�擾�J�E���^�����l
        
        'Web����������ɏ����ł���悤�ɂ���
'        For j = 0 To MaxRepeat - 1
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
        
'    For j = 0 To MaxRepeat - 1
    For j = 1 To MaxRepeat
        
'        '�t�H�[�����J��
'        objIE.navigate InputISBNPage 'IE��URL���J��
'        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
'        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
            '������t�H�[���W�J�S�̂̓��O�C���v���V�[�W���ւ܂����āAHTML���擾���Ă���
        '���O�C����ԃ`�F�b�N
        Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
        
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

Sub CheckLogin(objIE As InternetExplorer, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
        
    '�I�u�W�F�N�g�ݒ�
        
        '���O�C���ݒ�(�f�B���N�g��)
        Dim LoginDir As String '���O�C���f�B���N�g��
        LoginDir = "login" '���O�C���f�B���N�g���w��
        Dim LoginPageURL As String '���O�C���y�[�WURL
        LoginPageURL = Domain & LoginDir '���O�C���y�[�WURL����
        '���O�C���ݒ�(Web���M���)
        Dim LoginEmail As String '���O�C�����[���A�h���X
        Dim LoginPassword As String '���O�C���p�X���[�h
        LoginEmail = ThisWorkbook.Worksheets("���O�C���ݒ�").Cells(2, 1).Value 'Email
        LoginPassword = ThisWorkbook.Worksheets("���O�C���ݒ�").Cells(2, 2).Value 'Password
        '�������ʊm�F
'        Dim LoginAnswer As String '���O�C�����ʊm�F�p
        Dim ExitMsg As String '���b�Z�[�W�\���p
        'URL�擾�ݒ�
        Dim ProcessPageURL As String '�������{�y�[�WURL
        Dim ResponseURL As String '�������{�y�[�W�\����URL�擾
        
    '�������{�y�[�W����
    If CheckFirstLogin = True Then
        ProcessPageURL = LoginPageURL '���O�C���y�[�WURL����
    Else
        ProcessPageURL = Domain & ProcessDir '�������{�y�[�WURL����
    End If
    
    '�������{�y�[�W�փA�N�Z�X��AURL�擾
    objIE.navigate ProcessPageURL 'IE�ŊJ��
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    ResponseURL = objIE.document.URL 'URL�擾
    
    '���O�C����ʕ\�����̓��O�C������
    If ResponseURL = LoginPageURL Then
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        '�t�H�[������
        htmlDoc.getElementsByName("email")(0).Value = LoginEmail
        htmlDoc.getElementsByName("password")(0).Value = LoginPassword
        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
        
        '���O�C�����ʊm�F
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        ResponseURL = objIE.document.URL '�ǂݍ��݌��URL�擾
        Debug.Print ResponseURL '�f�o�b�O�m�F
        If ResponseURL = LoginPageURL Then
'            LoginAnswer = "���O�C�����s"
            '�I�u�W�F�N�g�I�����������{���Ă���
            objIE.Quit 'objIE���I��������
            '���O�C�����s���̓A���[�g�����b�Z�[�W�Ƃ��ĕԂ�
            ExitMsg = "���O�C���Ɏ��s���܂����B"
            MsgBox ExitMsg
            '�����̏����͂����ɏI��
            End
        Else
'            LoginAnswer = "���O�C������"
        End If
    Else
'        LoginAnswer = "���O�C���ς�"
    End If
    
    '���O�C���ς�or���O�C����T�C�g��HTML�I�u�W�F�N�g�擾
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
'    '���ʊm�F���
'    Debug.Print "���ʕ\���J�n"
'    Debug.Print "CheckFirstLogin " & CheckFirstLogin
'    Debug.Print LoginAnswer
'    Debug.Print "ProcessPageURL " & ProcessPageURL
'    Debug.Print "ProcessDir " & ProcessDir
'    Debug.Print "Web�\��Title�^�O " & htmlDoc.getElementsByTagName("title")(0).innerText
'    Debug.Print "���ʕ\���I��"
'    Debug.Print ""
    
    '���񏈗��I������
    If CheckFirstLogin = True Then CheckFirstLogin = False
    
End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
