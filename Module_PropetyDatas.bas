Attribute VB_Name = "Module_PropetyDatas"
Option Explicit

Sub getBookdatasDatail()

    '===��VBA�S�̃I�u�W�F�N�g�ݒ聫===

        'IE�I�u�W�F�N�g
        Dim waitObjIE As waitObjIE 'IE�ǂݍ��ݑ҂����W���[���쐬
        Set waitObjIE = New waitObjIE
        Set waitObjIE.objIE = CreateObject("Internetexplorer.Application")
        waitObjIE.objIE.Visible = False
        
        '�f�[�^�擾URL
        Dim Login As BookdataLogin '���O�C���N���X���W���[���쐬
        Set Login = New BookdataLogin
        Login.Domain = "https://protected-fortress-61913.herokuapp.com/" '�h���C���i�[
        Login.ProcessDir = "book" '�f�B���N�g���w��
        Login.CheckFirstLogin = True '���O�C���`�F�b�N�t���O
        Set Login.waitObjIE = waitObjIE 'IE�I�u�W�F�N�g��Login�Ɉ��n
            
    '===��VBA�S�̃I�u�W�F�N�g�ݒ聪===
        
    '���O�C����ԃ`�F�b�N
    Login.CheckLogin
        
    '===�������p�I�u�W�F�N�g�ݒ聫===
        
        Dim Pagination As HTMLUListElement 'HTML�y�[�W�l�[�V����
        Dim PagiLink As HTMLAnchorElement '���y�[�W�����N
        '��ƃ��[�N�V�[�g
        Dim SWSheet As Worksheet 'ScrapingWorksheet
        Set SWSheet = ThisWorkbook.Worksheets("�X�N���C�s���O")
        Dim MaxRow As Long '���R�[�h���m�F
        '���[�N�V�[�gID�擾
        Dim books As Range '�擾�Ϗ��Јꗗ
        Dim arrBooksId As Variant 'ID�z��
        '�J��Ԃ�����
        Dim i As Integer
        i = 1
    
        'URL�R���N�V����
        Dim URLCol As Collection
        Set URLCol = New Collection
        
        '�����������b�Z�[�W
        Dim ExitMsg As String

    '===�������p�I�u�W�F�N�g�ݒ聪===

    '���[�N�V�[�g���Џ��擾
    MaxRow = SWSheet.Cells(Rows.Count, 1).End(xlUp).Row '���[�N�V�[�g�v�f�̍ŏI�Z��
    Set books = SWSheet.Range(Cells(2, 1), Cells(MaxRow, 1)) '���[�N�V�[�gID�ꗗ�擾
    arrBooksId = books '�z��
    
    'OpenPage������Ԃ̓��[�v���đ�����
    Do Until Login.ProcessDir = ""
        
        Login.CheckLogin
        
        '�ڍ׃y�[�WURL�擾
        Call getBookList(Login.htmlDoc, i, URLCol, Login, arrBooksId)
        
        '�y�[�W�l�[�V��������
        Set Pagination = Login.htmlDoc.getElementsByClassName("pagination")(0)
        
        If Not Pagination Is Nothing Then
            '����p�Ƀ��Z�b�g
            Login.ProcessDir = ""
            '�y�[�W�l�[�V����������ꍇ�͎擾����
            For Each PagiLink In Pagination.getElementsByTagName("a")
                If InStr(PagiLink.outerHTML, "rel=""next") > 0 Then
                    Login.ProcessDir = Replace(PagiLink.href, Login.Domain, "")
                End If
            Next PagiLink
        
        End If
        
    Loop 'OpenPage���[�v�G���h

    
    '�ڍ׃y�[�WURL���Ȃ���ΏI������
    If URLCol.Count > 0 Then
        Call getDetailBookdata(SWSheet, waitObjIE.objIE, URLCol, Login, MaxRow)
        ExitMsg = "�f�[�^�擾���������܂����B"
    Else
        ExitMsg = "�V�K�擾�f�[�^�͂���܂���ł���"
    End If


    'VBA�I������
    waitObjIE.objIE.Quit 'objIE���I��������
    MsgBox ExitMsg

End Sub

Sub getBookList(htmlDoc As HTMLDocument, i As Integer, URLCol As Collection, Login As BookdataLogin, arrBooksId As Variant)
    
    '�ڍ׃y�[�WURL���擾
    Dim Bookdata As HTMLDivElement '���R�[�h�P�ʃf�[�^
    Dim detailField As HTMLDivElement '�ڍ׃t�B�[���h�f�[�^
    Dim BookdataURL As String '�ڍ׃y�[�WURL
    Dim BookdataURLDir As String '�ڍ׃y�[�W�f�B���N�g��
    Dim checkId As Variant '����ID
    Dim checkIdFlag As Boolean '���ЗL��
    checkIdFlag = False
    
    'URL����ID�擾
    Dim getIdData As Variant
    Dim getIdElement As Long
    Dim getBookdataId As Long
    
    For Each Bookdata In htmlDoc.getElementsByClassName("book-table__list")
        
        '--detail��񂩂�f�[�^�擾
        
            '--detail���擾
            Set detailField = Bookdata.getElementsByClassName("book-table__list--detail")(0)
    
            '�ڍ׃y�[�WURL
            BookdataURL = detailField.getElementsByTagName("a")(0) 'URL�擾
            BookdataURLDir = Replace(BookdataURL, Login.Domain, "") '�f�B���N�g���擾
            
            '�f�B���N�g������ID�擾
            getIdData = Split(BookdataURLDir, "/") 'URL�v�f�擾
            getIdElement = UBound(getIdData)  'URL�v�f�m�F
            getBookdataId = getIdData(getIdElement) 'URL����ԍ��擾
            
            'arrBooksId�ɂ���ꍇ�͏��Ђ�����̂ŏ��O
            For Each checkId In arrBooksId
                If checkId = getBookdataId Then
                    checkIdFlag = True
                    Exit For
                End If
            Next checkId
            
            If checkIdFlag = False Then URLCol.Add BookdataURLDir
            checkIdFlag = False
        '--detail��񂩂�f�[�^�擾�����܂�
        
        '��ԍ�����
        i = i + 1
    Next Bookdata

End Sub

Sub getDetailBookdata(SWSheet As Worksheet, objIE As InternetExplorer, URLCol As Collection, Login As BookdataLogin, i As Long)

    '�ڍ׃y�[�WURL����ڍד��e���擾
    
    '�f�[�^�擾URL
    Dim DocContent As HTMLDivElement 'HTML�R���e���c����
    Dim DocColumn As HTMLDivElement 'column���
    Dim j As Long '�����o���p�s�񏈗�

    Dim URLi As Long '�ڍ�URL�ǂݍ��ݍs�ԍ�����
    URLi = 1

    'URL�擾�����m�F
    Dim fornumber As Long
    fornumber = URLCol.Count
    
    '�摜����
    Dim DocPicture As HTMLDivElement
    Dim ImgURL As HTMLImg
    Dim ActCell As Range

    'ID�擾
    Dim GetUrl As String '�ڍ׃y�[�WURL
    Dim GetUrlData() As String '�ڍ׃y�[�WURL,Split�f�[�^
    Dim GetUrlElement As Integer 'URLSplit�v�f��
    Dim GetID As Integer 'ID�ԍ�

    '�ڍ׃y�[�W���J���Ē��̃f�[�^���擾
    Do
        i = i + 1
        '���y�[�WURL�擾
        Login.ProcessDir = URLCol(URLi)
        
        '���y�[�W�փA�N�Z�X
        Login.CheckLogin
        j = 1
        
        '1��ڂ�ID�ԍ��\��
        GetUrl = Login.ProcessDir 'URL�擾
        GetUrlData = Split(GetUrl, "/")  'URL�v�f�擾
        GetUrlElement = UBound(GetUrlData)  'URL�v�f�m�F
        GetID = GetUrlData(GetUrlElement)  'URL����ԍ��擾
        SWSheet.Cells(i, j).Value = GetID
        j = j + 1
        
        '2��ڂɉ摜�\��
        Set DocPicture = Login.htmlDoc.getElementsByClassName("book-detail__picture")(0)
        Set ImgURL = DocPicture.getElementsByTagName("img")(0)
        Set ActCell = SWSheet.Cells(i, j)
        
        SWSheet.Shapes.AddPicture _
          fileName:=ImgURL.src, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=ActCell.Left, _
            Top:=ActCell.Top, _
            Width:=100, _
            Height:=100
        j = j + 1
        
        '3��ڈȍ~�Ƀe�L�X�g�\��
        For Each DocContent In Login.htmlDoc.getElementsByClassName("document-content")
            Set DocColumn = DocContent.getElementsByClassName("document-content__column")(0)
            SWSheet.Cells(i, j).Value = DocColumn.innerHTML
            j = j + 1
        Next DocContent
        
        '�J�E���g�ǉ�
        URLi = URLi + 1
        
    'URL�v�f���𒴂���ꍇ�̓��[�v�I��
    Loop Until URLi > fornumber
    
End Sub
