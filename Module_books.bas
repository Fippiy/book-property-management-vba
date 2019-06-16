Attribute VB_Name = "Module_books"
Option Explicit

Sub getBookdata()
    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
    objIE.Visible = False 'IE��\��
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IE��URL���J��
    
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    
    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
    '��ƃ��[�N�V�[�g�w��
    Dim SWSheet As Worksheet 'ScrapingWorksheet
    Set SWSheet = ThisWorkbook.Worksheets("�X�N���C�s���O")
    
    
    Dim Bookdata As Object '���R�[�h�P�ʃf�[�^
    Dim detailField As Variant '�ڍ׃t�B�[���h�f�[�^
    Dim geturl As String '�ڍ׃y�[�WURL
    Dim GetUrlData() As String '�ڍ׃y�[�WURL,Split�f�[�^
    Dim GetUrlElement As Integer 'URLSplit�v�f��
    Dim GetID As Integer 'ID�ԍ�
    
    Dim ImgURL As String '�摜URL
    Dim Img As Variant '�摜�I�u�W�F�N�g
    Dim ActCell As Object '�摜�o�̓Z��
    
    Dim i As Integer
    i = 1
    
    
    '���R�[�h�P�ʂŃf�[�^�o��
    ' book-table__list�̗v�f��Bookdata�Ƃ��ď���
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
    
        '--detail�����擾���Ă��ꂼ�ꔽ�f
            
            detailField = Bookdata.getElementsByClassName("book-table__list--detail") '--detail���擾
            
            '�^�C�g����
            SWSheet.Cells(i + 1, 2).Value = detailField.getElementsByClassName("list-book-title")(0).innerText
            
            '�ڍ׃e�L�X�g
            SWSheet.Cells(i + 1, 3).Value = detailField.getElementsByClassName("list-book-detail")(0).innerText
            
            
            '�ڍ׃y�[�WURL
            geturl = detailField.getElementsByTagName("a") 'URL�擾
            SWSheet.Cells(i + 1, 4).Value = geturl  '�擾URL���f
            GetUrlData = Split(geturl, "/")  'URL�v�f�擾
            GetUrlElement = UBound(GetUrlData)  'URL�v�f�m�F
            GetID = GetUrlData(GetUrlElement)  'URL����ԍ��擾
            SWSheet.Cells(i + 1, 1).Value = GetID
        
        '--detail�����擾���Ă��ꂼ�ꔽ�f�����܂�
        
                
        '�摜����

        Img = Bookdata.getElementsByTagName("img")  '�摜�擾
        ImgURL = Img.src '�摜URL
        Set ActCell = SWSheet.Cells(i + 1, 5)

        '�摜�o�̓Z���̃s�N�Z�����w�肵�ĕ\��
        SWSheet.Shapes.AddPicture _
            fileName:=ImgURL, _
                LinkToFile:=True, _
                    SaveWithDocument:=True, _
                    Left:=ActCell.Left, _
                    Top:=ActCell.Top, _
                    Width:=100, _
                    Height:=100

        '�摜���������܂�
        
        
        '���̃��R�[�h�̍s�ԍ�
        i = i + 1
    Next Bookdata

    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"

End Sub
Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
