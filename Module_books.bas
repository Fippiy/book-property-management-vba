Attribute VB_Name = "Module_books"
Option Explicit

Sub getBookdata()
    '�x�[�X�쐬+�Z���\��
    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g

    objIE.Visible = True 'IE��\��

    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IE��URL���J��

    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�

    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g

    '���b�Z�[�W�{�b�N�X�Ɏ擾�N���X�̍ŏ��̕������o��
'     MsgBox htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
    
    
    '�C�~�f�B�G�C�g�Ɏw��N���X�̑S�擾�e�L�X�g��\��
'    Dim Str As Variant
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
'        Debug.Print "�o�́F" & Str.innerHTML
'    Next Str
    
    
    '�V�[�g��Ɏw��N���X�̑S�擾�e�L�X�g��\��
    Dim Str As Variant
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 1).Value = i
        Worksheets("�X�N���C�s���O").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
'        Debug.Print "�o�́F" & Str.innerHTML
    Next Str
    
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str
'    Debug.Print "�f�[�^�擾���������܂����B"
'    objIE.Quit
End Sub

Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
'        Debug.Print objIE.Busy
'        Debug.Print objIE.readyState
        DoEvents
    Loop
End Sub

