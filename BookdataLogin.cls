VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookdataLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public htmlDoc As HTMLDocument
Public Domain As String
Public ProcessDir As String
Public CheckFirstLogin As Boolean
Public waitObjIE As waitObjIE

'Sub CheckLogin(waitObjIE As waitObjIE, htmlDoc As HTMLDocument, Domain As String, ProcessDir As String, CheckFirstLogin As Boolean)
Sub CheckLogin()
        
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
'    waitObjIE.navigate ProcessPageURL 'IE�ŊJ��
    waitObjIE.objIE.navigate ProcessPageURL
'    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    waitObjIE.WaitResponse

'    ResponseURL = waitObjIE.document.URL 'URL�擾
    ResponseURL = waitObjIE.objIE.document.URL 'URL�擾
    
    '���O�C����ʕ\�����̓��O�C������
    If ResponseURL = LoginPageURL Then
'        Set htmlDoc = waitObjIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        Set htmlDoc = waitObjIE.objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        '���O�C���t�H�[������
        htmlDoc.getElementsByName("email")(0).Value = LoginEmail
        htmlDoc.getElementsByName("password")(0).Value = LoginPassword
        htmlDoc.getElementsByClassName("form-group__submit")(0).Click
        
        '���O�C�����ʊm�F
'        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
         waitObjIE.WaitResponse

'        ResponseURL = waitObjIE.document.URL '�ǂݍ��݌��URL�擾
        ResponseURL = waitObjIE.objIE.document.URL '�ǂݍ��݌��URL�擾
        If ResponseURL = LoginPageURL Then '���O�C��URL�̂܂܂͎��s
            '���O�C�����s�Ƃ��ďI������
'            waitObjIE.Quit 'objIE���I��������
            waitObjIE.objIE.Quit 'objIE���I��������
            MsgBox "���O�C���Ɏ��s���܂����B"
            End '�����̏����͂����ɏI��
        End If
    End If
    
    '���O�C���ς�or���O�C����T�C�g��HTML�I�u�W�F�N�g�擾
'    Set htmlDoc = waitObjIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    Set htmlDoc = waitObjIE.objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
    '���񏈗��I������
    If CheckFirstLogin = True Then CheckFirstLogin = False
    
End Sub
