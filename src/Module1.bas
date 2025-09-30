Attribute VB_Name = "Module1"
Option Explicit
' �������X�g�́u���f�[�^�͈́v��Ԃ��i�w�b�_�[���O�j
' �f�[�^��������� Nothing ��Ԃ�
Public Function GetOrderDataRange(ByVal ws As Worksheet) As Range
    On Error GoTo ExitFunc

    ' �e�[�u���iListObject�j������ꍇ�FDataBodyRange ���g��
    If ws.ListObjects.Count > 0 Then
        Dim lo As ListObject
        Set lo = ws.ListObjects(1)
        If Not (lo.DataBodyRange Is Nothing) Then
            If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then
                Set GetOrderDataRange = lo.DataBodyRange
                Exit Function
            End If
        End If
        ' �e�[�u�������f�[�^�Ȃ�
        Set GetOrderDataRange = Nothing
        Exit Function
    End If

    ' �e�[�u���������ꍇ�FA��̍ŏI�s���擾�iA1�̓w�b�_�[�z��j
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        Set GetOrderDataRange = Nothing
    Else
        ' A?C��̂ǂ����ɒl�����邩�i�񐔂͕K�v�ɉ����Ē����j
        Dim rng As Range
        Set rng = ws.Range("A2:C" & lastRow)
        If Application.WorksheetFunction.CountA(rng) > 0 Then
            Set GetOrderDataRange = rng
        Else
            Set GetOrderDataRange = Nothing
        End If
    End If
    Exit Function

ExitFunc:
    Set GetOrderDataRange = Nothing
End Function



' --- �����I�Ƀt�H���_�����ׂč�� ---
Private Function EnsureFolder(ByVal path As String) As Boolean
    Dim p As String, i As Long, parts() As String
    
    path = Trim(path)
    path = Replace(path, "/", "\")
    ' ����\�͗��Ƃ��i���肵�₷�����邽�߁j
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    If Len(path) = 0 Then EnsureFolder = False: Exit Function
    
    ' UNC or �h���C�u���^�[�Ή�
    If Left$(path, 2) = "\\" Then
        ' \\server\share\dir\dir...
        parts = Split(Mid$(path, 3), "\")
        p = "\\" & parts(0) & "\" & parts(1)   ' \\server\share
        i = 2
    Else
        ' C:\dir\dir...
        parts = Split(path, "\")
        p = parts(0) & "\"                      ' C:
        i = 1
    End If
    
    On Error Resume Next
    For i = i To UBound(parts)
        p = p & "\" & parts(i)
        If Dir(p, vbDirectory) = "" Then MkDir p
        If Err.Number <> 0 Then EnsureFolder = False: Err.Clear: Exit Function
    Next i
    On Error GoTo 0
    
    EnsureFolder = (Dir(path, vbDirectory) <> "")
End Function


Sub UpdateStock()
    Dim wsM As Worksheet, wsIO As Worksheet, wsS As Worksheet, wsSet As Worksheet
    Dim lastM As Long, lastIO As Long, i As Long, r As Long
    Dim sku As String
    Dim qty As Double
    Dim defaultSafe As Double, safe As Double, cur As Double, lack As Double
    Dim dictCur As Object 'SKU���Ƃ̌��݌�
    Set dictCur = CreateObject("Scripting.Dictionary")
    
    Set wsM = ThisWorkbook.Sheets("�i�ڃ}�X�^")
    Set wsIO = ThisWorkbook.Sheets("���o��")
    Set wsS = ThisWorkbook.Sheets("�݌�")
    Set wsSet = ThisWorkbook.Sheets("�ݒ�") ' A1=���S�݌Ɋ���l / B1=�l
    
    ' �ݒ�F���S�݌ɂ̊���l
    If IsNumeric(wsSet.Range("B1").Value) Then
        defaultSafe = CDbl(wsSet.Range("B1").Value)
    Else
        defaultSafe = 0
    End If
    
    ' --- �}�X�^�ǂݍ��݁FSKU�������i���݌�=0�j
    lastM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastM
        sku = Trim(CStr(wsM.Cells(i, 1).Value))
        If Len(sku) > 0 Then
            dictCur(sku) = 0
        End If
    Next i
    
    ' --- ���o�ɏW�v�F����+ / �o��-
    lastIO = wsIO.Cells(wsIO.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastIO
        sku = Trim(CStr(wsIO.Cells(i, 3).Value))
        If dictCur.Exists(sku) Then
            qty = CDbl(Val(wsIO.Cells(i, 4).Value))
            Select Case Replace(Trim(CStr(wsIO.Cells(i, 2).Value)), " ", "")
                Case "��", "����": dictCur(sku) = dictCur(sku) + qty
                Case "�o", "�o��": dictCur(sku) = dictCur(sku) - qty
                Case Else
                    ' ��ʂ��z��O�̓X�L�b�v�i�K�v�Ȃ烍�O���j
            End Select
        Else
            ' ���o�^SKU�̓X�L�b�v�i�K�v�Ȃ烍�O���j
        End If
    Next i
    
    ' --- �݌ɏo��
    wsS.Cells.ClearContents
    wsS.Cells.Interior.ColorIndex = xlNone
    wsS.Range("A1:E1").Value = Array("SKU", "�i��", "���݌�", "���S�݌�", "�s����")
    
    r = 2
    For i = 2 To lastM
        sku = Trim(CStr(wsM.Cells(i, 1).Value))
        If Len(sku) > 0 Then
            cur = 0
            If dictCur.Exists(sku) Then cur = CDbl(dictCur(sku))
            
            ' ���S�݌ɁF�i�ڃ}�X�^C�񂪋�/0�Ȃ�ݒ������g�p
            If IsNumeric(wsM.Cells(i, 3).Value) Then
                safe = CDbl(wsM.Cells(i, 3).Value)
            Else
                safe = 0
            End If
            If safe = 0 Then safe = defaultSafe
            
            lack = Application.Max(0, safe - cur)
            
            wsS.Cells(r, 1).Value = sku
            wsS.Cells(r, 2).Value = wsM.Cells(i, 2).Value
            wsS.Cells(r, 3).Value = cur
            wsS.Cells(r, 4).Value = safe
            wsS.Cells(r, 5).Value = lack
            
            If lack > 0 Then
                wsS.Range(wsS.Cells(r, 1), wsS.Cells(r, 5)).Interior.Color = RGB(255, 200, 200)
            End If
            r = r + 1
        End If
    Next i
    
    MsgBox "�݌ɍX�V���������܂����i�s��: " & (r - 2) & "�j", vbInformation
End Sub

Sub CreateOrderList()
    Dim wsS As Worksheet, wsO As Worksheet
    Dim lastS As Long, i As Long, r As Long
    Dim sku As String, pname As String
    Dim lack As Double
    
    Set wsS = ThisWorkbook.Sheets("�݌�")
    Set wsO = ThisWorkbook.Sheets("�������X�g")
    
    ' �������X�g��������
    wsO.Cells.ClearContents
    wsO.Range("A1:C1").Value = Array("SKU", "�i��", "������")
    
    ' �݌ɃV�[�g�̍ŏI�s���擾
    lastS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    
    r = 2
    For i = 2 To lastS
        sku = wsS.Cells(i, 1).Value
        pname = wsS.Cells(i, 2).Value
        lack = wsS.Cells(i, 5).Value ' �s������
        
        If lack > 0 Then
            wsO.Cells(r, 1).Value = sku
            wsO.Cells(r, 2).Value = pname
            wsO.Cells(r, 3).Value = lack
            r = r + 1
        End If
    Next i
    
    MsgBox "�������X�g���쐬���܂����i" & (r - 2) & "���j", vbInformation
End Sub


Sub SaveOrderListCsv()
    Dim wsSet As Worksheet, wsO As Worksheet
    Dim outDir As String, outFile As String
    Dim wbTmp As Workbook
    Dim usedR As Range
    
    Set wsSet = ThisWorkbook.Sheets("�ݒ�")
    Set wsO = ThisWorkbook.Sheets("�������X�g")
    
    ' �o�̓t�H���_�i��FC:\temp\inventory\�j
outDir = Trim(CStr(wsSet.Range("B2").Value))

' �@ B2�����͂̂Ƃ��̈ē�����̉�
If Len(outDir) = 0 Then
    MsgBox "�ݒ�V�[�g��B2�Z���ɕۑ���t�H���_����͂��Ă��������B" & vbCrLf & _
           "�i��FC:\temp\inventory�j", vbExclamation
    Exit Sub
End If

' �A �p�X�̐��K���F/ �� \ �ɕϊ� �{ ������ \ ��t����
outDir = Replace(outDir, "/", "\")
If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"

' �B �t�H���_�쐬�i���s���̃��b�Z�[�W����̉��j
If Not EnsureFolder(outDir) Then
    MsgBox "�ۑ���t�H���_���쐬�܂��͗��p�ł��܂���ł����B" & vbCrLf & _
           "�p�X��A�N�Z�X�����A�Z�L�����e�B�\�t�g�̐ݒ���m�F���Ă��������B" & vbCrLf & _
           "�Ώ�: " & outDir, vbExclamation
    Exit Sub
End If

' �C �������X�g��`�F�b�N�FCells��UsedRange�ɕύX�i�������Ӑ}�ɍ����j
Dim dataR As Range
Set dataR = GetOrderDataRange(wsO)
If dataR Is Nothing Then
    MsgBox "�������X�g����ł��B��Ɂm�������X�g�쐬�n�����s���Ă��������B", vbExclamation
    Exit Sub
End If

Set usedR = dataR

    
    ' �ꎞ�u�b�N�ɓ\��t����CSV�ۑ��iUTF-8�j
    Set wbTmp = Application.Workbooks.Add
    usedR.Copy wbTmp.Sheets(1).Range("A1")
    
    outFile = outDir & "����_" & Format(Now, "yyyymmdd_hhnn") & ".csv"
    On Error GoTo ErrH
    wbTmp.SaveAs Filename:=outFile, FileFormat:=62 ' xlCSVUTF8=62
    wbTmp.Close SaveChanges:=False
    MsgBox "CSV��ۑ����܂���: " & outFile, vbInformation
    Exit Sub
ErrH:
    On Error Resume Next
    wbTmp.Close SaveChanges:=False
    MsgBox "CSV�ۑ��ŃG���[: " & Err.Description, vbExclamation
End Sub

Sub SanityCheck()
    On Error Resume Next
    Dim msg As String
    msg = ""
    msg = msg & "�V�[�g����: " & _
        CStr(Not ThisWorkbook.Sheets("�i�ڃ}�X�^") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("���o��") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("�݌�") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("�ݒ�") Is Nothing) & vbCrLf
    
    Dim lastM As Long, lastIO As Long
    lastM = ThisWorkbook.Sheets("�i�ڃ}�X�^").Cells(Rows.Count, 1).End(xlUp).Row
    lastIO = ThisWorkbook.Sheets("���o��").Cells(Rows.Count, 1).End(xlUp).Row
    msg = msg & "�ŏI�s: �}�X�^=" & lastM & " / ���o��=" & lastIO & vbCrLf
    
    Dim outDir As String
    outDir = CStr(ThisWorkbook.Sheets("�ݒ�").Range("B2").Value)
    msg = msg & "�o�̓t�H���_: [" & outDir & "]" & vbCrLf
    
    MsgBox msg, vbInformation, "SanityCheck"
End Sub


Option Explicit

' ���ʁF1���̃V�[�g�Ɋ�{�t�H�[�}�b�g��K�p
Private Sub FS_FormatCommon(ByVal ws As Worksheet)
    With ws.Cells
        .Font.name = "Meiryo"
        .Font.Size = 10
        .RowHeight = 20
    End With
    ws.Rows(1).RowHeight = 24

    ' �擪�s�Œ�i���S�Ɂj
    ws.Activate
    On Error Resume Next
    ws.Range("A2").Select
    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 1
        .FreezePanes = True
    End With
    On Error GoTo 0
End Sub

' �݌� A:E�iSKU, �i��, ���݌�, ���S�݌�, �s�����j
Private Sub FS_FormatStock(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:E1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C:E").ColumnWidth = 10
    ws.Columns("C:E").NumberFormatLocal = "#,##0"

    ' �s��(E��>0)�𔖐Ԃŋ���
    Dim lastRow As Long, rng As Range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1000
    Set rng = ws.Range("A2:E" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$E2>0")
        .Interior.Color = RGB(255, 200, 200)
    End With
End Sub

' �������X�g A:C�iSKU, �i��, �������j
Private Sub FS_FormatOrder(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:C1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("C").NumberFormatLocal = "#,##0"
End Sub

' �ݒ� A:B�iA1=���S�݌Ɋ���l, A2=�o�̓t�H���_�j
Private Sub FS_FormatSetting(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:A2")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 48
End Sub

' �i�ڃ}�X�^ A:C�iSKU, �i��, ���S�݌Ɂj
Private Sub FS_FormatMaster(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:C1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12   ' SKU
    ws.Columns("B").ColumnWidth = 28   ' �i��
    ws.Columns("C").ColumnWidth = 10   ' ���S�݌�
    ws.Columns("C").NumberFormatLocal = "#,##0"
End Sub

' ���o�Ɂi�z���FA=���t, B=���(��/�o), C=SKU, D=����, E=���l�j
Private Sub FS_FormatIO(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:E1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12   ' ���t
    ws.Columns("B").ColumnWidth = 8    ' ���
    ws.Columns("C").ColumnWidth = 12   ' SKU
    ws.Columns("D").ColumnWidth = 10   ' ����
    ws.Columns("E").ColumnWidth = 24   ' ���l

    ' �\���`��
    ws.Columns("A").NumberFormatLocal = "yyyy/mm/dd"
    ws.Columns("D").NumberFormatLocal = "#,##0"

    ' �i�C�Ӂj��ʂ̓��͋K���i��/�o�̂݁j
    On Error Resume Next
    With ws.Range("B2:B1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="��,�o"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "�I��"
        .InputMessage = "�� or �o ��I��"
        .ErrorTitle = "���̓G���["
        .ErrorMessage = "�u���v�܂��́u�o�v�ȊO�͓��͂ł��܂���B"
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Public Sub FormatScreensWithReport_NoHelper()
    Dim ws As Worksheet
    Dim done As String, missing As String

    ' �݌�
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("�݌�"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatStock(ws): done = done & "�E�݌�" & vbCrLf
    Else
        missing = missing & "�E�݌�" & vbCrLf
    End If

    ' �������X�g
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("�������X�g"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatOrder(ws): done = done & "�E�������X�g" & vbCrLf
    Else
        missing = missing & "�E�������X�g" & vbCrLf
    End If

    ' �ݒ�
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("�ݒ�"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatSetting(ws): done = done & "�E�ݒ�" & vbCrLf
    Else
        missing = missing & "�E�ݒ�" & vbCrLf
    End If

    ' �i�ڃ}�X�^
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("�i�ڃ}�X�^"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatMaster(ws): done = done & "�E�i�ڃ}�X�^" & vbCrLf
    Else
        missing = missing & "�E�i�ڃ}�X�^" & vbCrLf
    End If

    ' ���o��
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("���o��"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatIO(ws): done = done & "�E���o��" & vbCrLf
    Else
        missing = missing & "�E���o��" & vbCrLf
    End If

    ' ���|�[�g
    Dim msg As String: msg = "���`���������܂����B" & vbCrLf & vbCrLf
    If Len(done) > 0 Then msg = msg & "���`�����V�[�g�F" & vbCrLf & done & vbCrLf
    If Len(missing) > 0 Then msg = msg & "������Ȃ������V�[�g�F" & vbCrLf & missing & _
        "(�^�u���̕\�L�h���X�y�[�X�E�S�p/���p���m�F���Ă�������)"
    MsgBox msg, vbInformation, "FormatScreens"
End Sub


