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
        '�f�[�^�擾URL
        Dim OpenPage As String
        OpenPage = "https://protected-fortress-61913.herokuapp.com/book"
        '�J��Ԃ�����
        Dim i As Integer
        i = 1
    
        'URL�R���N�V����
        Dim URLCol As Collection
        Set URLCol = New Collection
        
        '�����������b�Z�[�W
        Dim ExitMsg As String

    'OpenPage������Ԃ̓��[�v���đ�����
    Do Until OpenPage = ""
        
        objIE.navigate OpenPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        OpenPage = "" '�f�[�^�擾URL������
        
        '�ڍ׃y�[�WURL�擾
        Call getBookList(htmlDoc, i, URLCol)
        
        
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

    
    '�ڍ׃y�[�WURL���Ȃ���ΏI������
    If URLCol.Count > 0 Then
        Call getDetailBookdata(SWSheet, objIE, URLCol)
        ExitMsg = "�f�[�^�擾���������܂����B"
    Else
        ExitMsg = "�擾�f�[�^������܂���"
    End If


    'VBA�I������
    objIE.Quit 'objIE���I��������
    MsgBox ExitMsg

End Sub

Sub getBookList(htmlDoc As HTMLDocument, i As Integer, URLCol As Collection)
    
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
            URLCol.Add BookdataURL
        
        '--detail��񂩂�f�[�^�擾�����܂�
        
        '��ԍ�����
        i = i + 1
    Next Bookdata

End Sub

Sub getDetailBookdata(SWSheet As Worksheet, objIE As InternetExplorer, URLCol As Collection)

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

    'URL�擾�����m�F
    Dim fornumber As Long
    fornumber = URLCol.Count
    
    '�摜����
    Dim DocPicture As HTMLDivElement
    Dim ImgURL As HTMLImg
    Dim actcell As Range

    'ID�擾
    Dim GetUrl As String '�ڍ׃y�[�WURL
    Dim GetUrlData() As String '�ڍ׃y�[�WURL,Split�f�[�^
    Dim GetUrlElement As Integer 'URLSplit�v�f��
    Dim GetID As Integer 'ID�ԍ�

    '�ڍ׃y�[�W���J���Ē��̃f�[�^���擾
    Do
        
        '���y�[�WURL�擾
        OpenPage = URLCol(URLi)
        
        objIE.navigate OpenPage 'IE��URL���J��
        Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
        Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
        j = 1
        
        '1��ڂ�ID�ԍ��\��
        GetUrl = OpenPage 'URL�擾
        GetUrlData = Split(GetUrl, "/")  'URL�v�f�擾
        GetUrlElement = UBound(GetUrlData)  'URL�v�f�m�F
        GetID = GetUrlData(GetUrlElement)  'URL����ԍ��擾
        SWSheet.Cells(i, j).Value = GetID
        j = j + 1
        
        '2��ڂɉ摜�\��
        Set DocPicture = htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
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
        
        
        '3��ڈȍ~�Ƀe�L�X�g�\��
        For Each DocContent In htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        
        '�J�E���g�ǉ�
        i = i + 1
        URLi = URLi + 1
        
    'URL�v�f���𒴂���ꍇ�̓��[�v�I��
    Loop Until URLi > fornumber
    
End Sub

Sub WaitResponse(objIE As Object) 'Web�u���E�U�\�������҂�
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
