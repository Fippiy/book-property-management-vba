Attribute VB_Name = "Module_books"
Option Explicit


Sub getBookdata()
    '���R�����g�폜

    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
    
    objIE.Visible = False 'IE��\���AFalse��IE�\���Ȃ�
    
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IE��URL���J��
    
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�

    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
  
''    �f�[�^�擾�܂Ƃ�
'    Debug.Print "1." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print "2." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print "3." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("h3")
'    Debug.Print "4." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("p")
    

    ' �V�[�g��Ɏw��N���X�̑S�擾�e�L�X�g��\��
    Dim Str As Object
    Dim i As Integer
    
    '�^�C�g�����擾
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 1).Value = i
        Worksheets("�X�N���C�s���O").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
    Next Str

    'detail�擾
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str

    'URL�擾
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 4).Value = Str.getElementsByTagName("a")
        i = i + 1
    Next Str

    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"

End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub

