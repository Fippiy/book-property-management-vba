Attribute VB_Name = "Module_books"
Option Explicit

Sub getBookdata()
    'URL�擾
    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
    
    objIE.Visible = False 'IE��\��
    
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IE��URL���J��
    
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�

    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
    '���b�Z�[�W�{�b�N�X�Ɏ擾�N���X�̍ŏ��̕������o��
'    MsgBox htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-title").innerHTML
'    Cells(2, 2).Value = htmlDoc.getElementsByClassName("list-book-title")(0).innerHTML

    ' �V�[�g��Ɏw��N���X�̑S�擾�e�L�X�g��\��
'    Dim Str As Object
''    Dim i As Integer
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
'        Debug.Print "�o�́F" & Str.innerHTML
'    Next Str

'     ���b�Z�[�W�{�b�N�X�Ɏ擾�N���X�̎q�v�f���擾�o��
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).outerHTML
'    Debug.Print htmlDoc.getElementsByClassName("list-book-detail")(0).innerHTML


'    �N���X�����̗v�f�z���̃^�O���̒��g���擾
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0).innerHTML


'    �N���X�����̗v�f�z����a�^�O�v�f���擾(URL)
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0)


'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print htmlDoc.getElementsByClassName("book-table__list--detail")(0).outerHTML.getElementsByTagName("a")
    
''    �f�[�^�擾�܂Ƃ�
'    Debug.Print "1." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).innerHTML
'    Debug.Print "2." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")
'    Debug.Print "3." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("h3")
'    Debug.Print "4." & htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("p")
    
    
    ' �C�~�f�B�G�C�g�Ɏw��N���X�̑S�擾�e�L�X�g��\��
'    Dim Str As Variant
'    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
''        Debug.Print "�o�́F" & Str.innerHTML
'    Next Str




'    ' �V�[�g��Ɏw��N���X�̑S�擾�e�L�X�g��\��
    Dim Str As Object
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 1).Value = i
        Worksheets("�X�N���C�s���O").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
'        Debug.Print "�o�́F" & Str.innerHTML
    Next Str

'    i = 1
'    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
'        Worksheets("�X�N���C�s���O").Cells(i + 1, 3).Value = Str.innerHTML
'        i = i + 1
'    Next Str

'    i = 1
''    htmlDoc.getElementsByClassName("book-table__list--detail")(0).getElementsByTagName("a")(0)
''    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
'    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
'        Worksheets("�X�N���C�s���O").Cells(i + 1, 4).Value = Str.getElementsByTagName("a")
'        i = i + 1
'    Next Str


'book-table__list--detail


    objIE.Quit 'objIE���I��������
'    Debug.Print "�f�[�^�擾���������܂����B"
End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub

