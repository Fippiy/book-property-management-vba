Attribute VB_Name = "Module_getURLtest"
Option Explicit

Sub getURLtest()
    
    '�I�u�W�F�N�g�ݒ�
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = True 'IE��\��
        Dim htmlDoc As HTMLDocument 'HTML�S��
        Dim HTTPStatus As Integer 'HTTP���N�G�X�g�X�e�[�^�X
        'URL�ݒ�
        Dim Domain As String 'Web����h���C����
        Dim OpenPage As String '����URL
        Dim ResponseURL As String '�\���T�C�gURL
        Domain = "https://protected-fortress-61913.herokuapp.com/"
                
'        'URL���J���ăI�u�W�F�N�g�擾
'        OpenPage = Domain '�폜����URL�擾
'        objIE.navigate OpenPage 'IE��URL���J��
'        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
'        Debug.Print objIE.document.URL '�ǂݍ��݌��URL�擾
        
        'URL���J���ăI�u�W�F�N�g�擾
        OpenPage = Domain & "login" '�폜����URL�擾
        objIE.navigate OpenPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        ResponseURL = objIE.document.URL '�ǂݍ��݌��URL�擾
        Debug.Print ResponseURL '�f�o�b�O�m�F
        
        '���O�C��URL�ݒ�
        Dim LoginURL As String
        LoginURL = Domain & "login"
        
        '���O�C����ʎ��̓��O�C������
        If ResponseURL = LoginURL Then
            Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
'            '�t�H�[������
            htmlDoc.getElementsByName("email")(0).Value = "test"
            htmlDoc.getElementsByName("password")(0).Value = "test"
            htmlDoc.getElementsByClassName("form-group__submit")(0).Click
            
            '���O�C�����ʊm�F
            Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
            ResponseURL = objIE.document.URL '�ǂݍ��݌��URL�擾
            Debug.Print ResponseURL '�f�o�b�O�m�F
            If ResponseURL = LoginURL Then
                Debug.Print "���O�C�����s"
            Else
                Debug.Print "���O�C������"
            End If
        End If


'        'URL���J���ăI�u�W�F�N�g�擾
'        OpenPage = Domain & "book" '�폜����URL�擾
'        objIE.navigate OpenPage 'IE��URL���J��
'        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
'        Debug.Print objIE.document.URL '�ǂݍ��݌��URL�擾


End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
