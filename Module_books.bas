Attribute VB_Name = "Module_books"
Option Explicit


Sub getBookdata()
    Dim objIE As InternetExplorer 'IE�I�u�W�F�N�g������
    Set objIE = CreateObject("Internetexplorer.Application") '�V����IE�I�u�W�F�N�g���쐬���ăZ�b�g
    objIE.Visible = True 'IE��\��
    objIE.navigate "https://protected-fortress-61913.herokuapp.com/book" 'IE��URL���J��
    
    Call WaitResponse(objIE) '�ǂݍ��ݑ҂�
    
    Dim htmlDoc As HTMLDocument 'HTML�h�L�������g�I�u�W�F�N�g������
    Set htmlDoc = objIE.document 'objIE�œǂݍ��܂�Ă���HTML�h�L�������g���Z�b�g
    
    ' �V�[�g��Ɏw��N���X�̑S�擾�e�L�X�g��\��
    Dim Str As Object
    Dim i As Integer
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-title")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 2).Value = Str.innerHTML
        i = i + 1
    Next Str

    '���Џڍ�
    i = 1
    For Each Str In htmlDoc.getElementsByClassName("list-book-detail")
        Worksheets("�X�N���C�s���O").Cells(i + 1, 3).Value = Str.innerHTML
        i = i + 1
    Next Str


    'URL
    i = 1
    Dim GetUrl As String
    Dim GetUrlData() As String
    Dim GetUrlElement As Integer
    Dim GetID As Integer
    
    For Each Str In htmlDoc.getElementsByClassName("book-table__list--detail")
        GetUrl = Str.getElementsByTagName("a")  'URL�擾
        Worksheets("�X�N���C�s���O").Cells(i + 1, 4).Value = GetUrl  '�擾URL���f
        GetUrlData = Split(GetUrl, "/")  'URL�v�f�擾
        GetUrlElement = UBound(GetUrlData)  'URL�v�f�m�F
        GetID = GetUrlData(GetUrlElement)  'URL����ԍ��擾
        Worksheets("�X�N���C�s���O").Cells(i + 1, 1).Value = GetID  '���[�N�V�[�g�֔��f
        i = i + 1  '���̍s�w��
    Next Str



    '�摜URL�擾
    
    '�摜�p�ϐ�
    Dim imgURL As String '�摜URL
    Dim Img As Object '�摜�I�u�W�F�N�g
    Dim toppix As Long '�ʒu�s�N�Z��

'    '1���T���v���擾
'    imgURL = htmlDoc.images(0).src
'    Worksheets("�X�N���C�s���O").Cells(2, 5).Value = imgURL
'    Worksheets("�X�N���C�s���O").Shapes.AddPicture _
'        fileName:=imgURL, _
'            LinkToFile:=True, _
'                SaveWithDocument:=True, _
'                Left:=0, _
'                Top:=0, _
'                Width:=100, _
'                Height:=80
    
'    'URL�̂ݎ擾
'    i = 1
'    For Each IMG In htmlDoc.images '�C���[�W�擾
'        imgURL = IMG.src '�ϐ��i�[
'        Worksheets("�X�N���C�s���O").Cells(i + 1, 5).Value = imgURL '�擾URL���f
'        i = i + 1
'    Next IMG
    
    
    
    Dim ActCell As Object

    i = 1
    toppix = 0
    For Each Img In htmlDoc.images '�C���[�W�擾
        imgURL = Img.src '�ϐ��i�[
        Set ActCell = Worksheets("�X�N���C�s���O").Cells(i + 1, 5)
        ActCell.Value = imgURL  '�擾URL���f

        '�摜��\��
        Worksheets("�X�N���C�s���O").Shapes.AddPicture _
            fileName:=imgURL, _
                LinkToFile:=True, _
                    SaveWithDocument:=True, _
                    Left:=0, _
                    Top:=0 + toppix, _
                    Width:=100, _
                    Height:=100

'        '�摜��\���A�Z���s�N�Z���擾
'        Worksheets("�X�N���C�s���O").Shapes.AddPicture _
'            fileName:=imgURL, _
'                LinkToFile:=True, _
'                    SaveWithDocument:=True, _
'                    Left:=ActCell.Left, _
'                    Top:=ActCell.Top, _
'                    Width:=100, _
'                    Height:=100

        i = i + 1
        toppix = toppix + 100
    Next Img


    objIE.Quit 'objIE���I��������
    MsgBox "�f�[�^�擾���������܂����B"
End Sub
Sub WaitResponse(objIE As Object)
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
End Sub
Sub picture1()
    '���[�J���f�B�X�N��̉摜�t�@�C����\��
    Worksheets("�X�N���C�s���O").Shapes.AddPicture _
        fileName:="Z:\FierVega\ariawase-master\bin\test.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0, _
                Width:=100, _
                Height:=80
End Sub
Sub picture2()
    'coverURL���w�肵�ăt�@�C���\��
    Worksheets("�X�N���C�s���O").Shapes.AddPicture _
        fileName:="https://cover.openbd.jp/9784797398892.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=0, _
                Top:=0, _
                Width:=100, _
                Height:=80
End Sub
Sub picture3()
    'coverURL���w�肵�ăt�@�C���\��
    Worksheets("�X�N���C�s���O").Shapes.AddPicture _
        fileName:="https://cover.openbd.jp/9784797398892.jpg", _
            LinkToFile:=True, _
                SaveWithDocument:=True, _
                Left:=330, _
                Top:=40, _
                Width:=100, _
                Height:=80
End Sub
