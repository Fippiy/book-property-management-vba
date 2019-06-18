Attribute VB_Name = "Module_PropetyDatas"
Option Explicit

Sub getBookdatasDatail()

    '�I�u�W�F�N�g�ݒ�
        'IE
        Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
        Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
        objIE.Visible = False 'IE��\��
        'HTML
        Dim htmlDoc As HTMLDocument 'HTML�S��
        Dim Pagination As HTMLUListElement 'HTML�y�[�W�l�[�V����
        Dim PagiLink As HTMLAnchorElement '���y�[�W�����N
        '��ƃ��[�N�V�[�g
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("�X�N���C�s���O")
        Dim DWSheet As Worksheet 'DetailWorksheet
        Set DWSheet = ThisWorkbook.Worksheets("�ڍ׃y�[�W���")
        '�f�[�^�擾URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        'URL�擾�J��Ԃ�����
        Dim i As Integer
        i = 1
    
    'OpenPage������Ԃ̓��[�v���đ�����
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        OpenPage = "" '�f�[�^�擾URL������
        

        '�ڍ׃y�[�WURL�擾
'        Call getBookList(DWSheet, htmlDoc, i)
        
        
        '�y�[�W�l�[�V��������
        Set Pagination = htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            '�y�[�W�l�[�V����������ꍇ�͎擾����
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    OpenPage = PagiLink.href
                End If
            Next PagiLink
        
        End If
        
    Loop 'OpenPage���[�v�G���h


    '�擾�����ڍ׃y�[�WURL����ڍ׃y�[�W�����擾����
    Call getDetailBookdata(SWSheet, DWSheet, objIE)


    'VBA�I������
    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"

End Sub

Sub getBookList(DWSheet As Worksheet, htmlDoc As HTMLDocument, i As Integer)
    
    '�ڍ׃y�[�WURL���擾
    Dim Bookdata As HTMLDivElement '���R�[�h�P�ʃf�[�^
    Dim detailField As HTMLDivElement '�ڍ׃t�B�[���h�f�[�^
    Dim BookdataURL As String '�ڍ׃y�[�WURL
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        
        '--detail��񂩂�f�[�^�擾
        
            '--detail���擾
            Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0)
    
            '�ڍ׃y�[�WURL
            BookdataURL = detailField.getElementsByTagName("a")(0) 'URL�擾
            DWSheet.Cells(i, 1).Value = BookdataURL  '�擾URL���f
        '--detail��񂩂�f�[�^�擾�����܂�
        
        '��ԍ�����
        i = i + 1
    Next Bookdata

End Sub

Sub getDetailBookdata(SWSheet As Worksheet, DWSheet As Worksheet, objIE As InternetExplorer)

    '�ڍ׃y�[�WURL����ڍד��e���擾
    
    '�f�[�^�擾URL
    Dim OpenPage As String
    Dim htmlDoc As HTMLDocument 'HTML�S��
    Dim DocContent As HTMLDivElement 'HTML�R���e���c����
    Dim DocColumn As HTMLDivElement 'column���
    Dim i As Long, j As Long '�����o���p�s�񏈗�
    i = 2
    
    Dim URLi As Long '�ڍ�URL�ǂݍ��ݍs�ԍ�����
    URLi = 1
    
    '�ŏ��̃y�[�W
    OpenPage = DWSheet.Cells(URLi, 1).Value
    
    '�摜����
    Dim DocPicture As Variant
    Dim ImgURL As Variant
    Dim actcell As Variant
    
    '�ڍ׃y�[�W���J���Ē��̃f�[�^���擾
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        j = 2
        
        '�摜�擾����
        Set DocPicture = htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
        SWSheet.Cells(i, j).Value = ImgURL.src
        Set actcell = SWSheet.Cells(i, j)
        
        SWSheet.Shapes.AddPicture _
          fileName:=ImgURL.src, _
            LinkToFile:=True, _
            SaveWithDocument:=True, _
            Left:=actcell.Left, _
            Top:=actcell.Top, _
            Width:=100, _
            Height:=100
        j = j + 1
                
        '�ڍ׃y�[�WHTML����f�[�^�擾
        'document-content���擾
        For Each DocContent In htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        '�ڍ׃y�[�WURL�S�擾�ŏI��
        i = i + 1
        URLi = URLi + 1
        
        '���y�[�WURL�擾
        OpenPage = DWSheet.Cells(URLi, 1).Value
    Loop
    
End Sub
Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub

