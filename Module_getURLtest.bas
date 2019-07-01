Attribute VB_Name = "Module_getURLtest"
Option Explicit

Sub getURLtest()
    
    '�I�u�W�F�N�g�ݒ�
        
        'IE�I�u�W�F�N�g
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
'        objIE.Visible = True 'IE��\��
        
        'HTML�I�u�W�F�N�g
        Dim htmlDoc As HTMLDocument 'HTML�S��
'        Dim HTTPStatus As Integer 'HTTP���N�G�X�g�X�e�[�^�X
        
        'URL�ݒ�
        Dim Domain As String 'Web�h���C����
        Domain = "https://protected-fortress-61913.herokuapp.com/" '�h���C���i�[
        Dim ProcessDir As String '�������{�f�B���N�g��
        ProcessDir = "book" '�f�B���N�g���w��
                
        'VBA���쏉�񃍃O�C���`�F�b�N
        Dim CheckFirstLogin As Boolean '���O�C���`�F�b�N�t���O
        CheckFirstLogin = True
                
    '���O�C����ԃ`�F�b�N
    Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)

    'VBA�e�폈���̎��{
    
        'navigate���Ƀ��O�C����Ԃ��m�F�Ƃ��đ}��
        Call CheckLogin(objIE, htmlDoc, Domain, ProcessDir, CheckFirstLogin)
    
    'VBA�e�폈������
    
    objIE.Quit 'objIE���I��������
    MsgBox "�������������܂����B"

End Sub

Sub CheckLogin(objIE As InternetExplorer, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
        
    '�I�u�W�F�N�g�ݒ�
        
        '���O�C���ݒ�
        Dim LoginDir As String '���O�C���f�B���N�g��
        LoginDir = "login" '���O�C���f�B���N�g���w��
        Dim LoginPageURL As String '���O�C���y�[�WURL
        LoginPageURL = Domain & LoginDir '���O�C���y�[�WURL����
        Dim LoginEmail As String '���O�C�����[���A�h���X
        Dim LoginPassword As String '���O�C���p�X���[�h
        
        Dim LoginAnswer As String '���O�C�����ʊm�F�p
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
        htmlDoc.getElementsByName("email")(0).Value = ThisWorkbook.Worksheets("���O�C���ݒ�").Cells(2, 1)
        htmlDoc.getElementsByName("password")(0).Value = ThisWorkbook.Worksheets("���O�C���ݒ�").Cells(2, 2)
        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
        
        '���O�C�����ʊm�F
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        ResponseURL = objIE.document.URL '�ǂݍ��݌��URL�擾
        Debug.Print ResponseURL '�f�o�b�O�m�F
        If ResponseURL = LoginPageURL Then
            LoginAnswer = "���O�C�����s"
            '���O�C�����s���̓A���[�g�����b�Z�[�W�Ƃ��ĕԂ�
            ExitMsg = "���O�C���Ɏ��s���܂����B"
            MsgBox ExitMsg
            '�����̏����͂����ɏI��
            End
        Else
            LoginAnswer = "���O�C������"
        End If
    Else
        LoginAnswer = "���O�C���ς�"
    End If
    
    '���O�C���ς�or���O�C����T�C�g��HTML�I�u�W�F�N�g�擾
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
    '���ʊm�F���
    Debug.Print "���ʕ\���J�n"
    Debug.Print "CheckFirstLogin " & CheckFirstLogin
    Debug.Print LoginAnswer
    Debug.Print "ProcessPageURL " & ProcessPageURL
    Debug.Print "ProcessDir " & ProcessDir
    Debug.Print "Web�\��Title�^�O " & htmlDoc.getElementsByTagName("title")(0).innerText
    Debug.Print "���ʕ\���I��"
    Debug.Print ""
    
    '���񏈗��I������
    If CheckFirstLogin = True Then CheckFirstLogin = False
    
End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
