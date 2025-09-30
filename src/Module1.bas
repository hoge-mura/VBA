Attribute VB_Name = "Module1"
Option Explicit
' 発注リストの「実データ範囲」を返す（ヘッダー除外）
' データが無ければ Nothing を返す
Public Function GetOrderDataRange(ByVal ws As Worksheet) As Range
    On Error GoTo ExitFunc

    ' テーブル（ListObject）がある場合：DataBodyRange を使う
    If ws.ListObjects.Count > 0 Then
        Dim lo As ListObject
        Set lo = ws.ListObjects(1)
        If Not (lo.DataBodyRange Is Nothing) Then
            If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then
                Set GetOrderDataRange = lo.DataBodyRange
                Exit Function
            End If
        End If
        ' テーブルだがデータなし
        Set GetOrderDataRange = Nothing
        Exit Function
    End If

    ' テーブルが無い場合：A列の最終行を取得（A1はヘッダー想定）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        Set GetOrderDataRange = Nothing
    Else
        ' A?C列のどこかに値があるか（列数は必要に応じて調整）
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



' --- 逐次的にフォルダをすべて作る ---
Private Function EnsureFolder(ByVal path As String) As Boolean
    Dim p As String, i As Long, parts() As String
    
    path = Trim(path)
    path = Replace(path, "/", "\")
    ' 末尾\は落とす（判定しやすくするため）
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    If Len(path) = 0 Then EnsureFolder = False: Exit Function
    
    ' UNC or ドライブレター対応
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
    Dim dictCur As Object 'SKUごとの現在庫
    Set dictCur = CreateObject("Scripting.Dictionary")
    
    Set wsM = ThisWorkbook.Sheets("品目マスタ")
    Set wsIO = ThisWorkbook.Sheets("入出庫")
    Set wsS = ThisWorkbook.Sheets("在庫")
    Set wsSet = ThisWorkbook.Sheets("設定") ' A1=安全在庫既定値 / B1=値
    
    ' 設定：安全在庫の既定値
    If IsNumeric(wsSet.Range("B1").Value) Then
        defaultSafe = CDbl(wsSet.Range("B1").Value)
    Else
        defaultSafe = 0
    End If
    
    ' --- マスタ読み込み：SKU初期化（現在庫=0）
    lastM = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastM
        sku = Trim(CStr(wsM.Cells(i, 1).Value))
        If Len(sku) > 0 Then
            dictCur(sku) = 0
        End If
    Next i
    
    ' --- 入出庫集計：入庫+ / 出庫-
    lastIO = wsIO.Cells(wsIO.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastIO
        sku = Trim(CStr(wsIO.Cells(i, 3).Value))
        If dictCur.Exists(sku) Then
            qty = CDbl(Val(wsIO.Cells(i, 4).Value))
            Select Case Replace(Trim(CStr(wsIO.Cells(i, 2).Value)), " ", "")
                Case "入", "入庫": dictCur(sku) = dictCur(sku) + qty
                Case "出", "出庫": dictCur(sku) = dictCur(sku) - qty
                Case Else
                    ' 種別が想定外はスキップ（必要ならログ化）
            End Select
        Else
            ' 未登録SKUはスキップ（必要ならログ化）
        End If
    Next i
    
    ' --- 在庫出力
    wsS.Cells.ClearContents
    wsS.Cells.Interior.ColorIndex = xlNone
    wsS.Range("A1:E1").Value = Array("SKU", "品名", "現在庫", "安全在庫", "不足数")
    
    r = 2
    For i = 2 To lastM
        sku = Trim(CStr(wsM.Cells(i, 1).Value))
        If Len(sku) > 0 Then
            cur = 0
            If dictCur.Exists(sku) Then cur = CDbl(dictCur(sku))
            
            ' 安全在庫：品目マスタC列が空/0なら設定既定を使用
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
    
    MsgBox "在庫更新が完了しました（行数: " & (r - 2) & "）", vbInformation
End Sub

Sub CreateOrderList()
    Dim wsS As Worksheet, wsO As Worksheet
    Dim lastS As Long, i As Long, r As Long
    Dim sku As String, pname As String
    Dim lack As Double
    
    Set wsS = ThisWorkbook.Sheets("在庫")
    Set wsO = ThisWorkbook.Sheets("発注リスト")
    
    ' 発注リストを初期化
    wsO.Cells.ClearContents
    wsO.Range("A1:C1").Value = Array("SKU", "品名", "発注数")
    
    ' 在庫シートの最終行を取得
    lastS = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    
    r = 2
    For i = 2 To lastS
        sku = wsS.Cells(i, 1).Value
        pname = wsS.Cells(i, 2).Value
        lack = wsS.Cells(i, 5).Value ' 不足数列
        
        If lack > 0 Then
            wsO.Cells(r, 1).Value = sku
            wsO.Cells(r, 2).Value = pname
            wsO.Cells(r, 3).Value = lack
            r = r + 1
        End If
    Next i
    
    MsgBox "発注リストを作成しました（" & (r - 2) & "件）", vbInformation
End Sub


Sub SaveOrderListCsv()
    Dim wsSet As Worksheet, wsO As Worksheet
    Dim outDir As String, outFile As String
    Dim wbTmp As Workbook
    Dim usedR As Range
    
    Set wsSet = ThisWorkbook.Sheets("設定")
    Set wsO = ThisWorkbook.Sheets("発注リスト")
    
    ' 出力フォルダ（例：C:\temp\inventory\）
outDir = Trim(CStr(wsSet.Range("B2").Value))

' ① B2未入力のときの案内を具体化
If Len(outDir) = 0 Then
    MsgBox "設定シートのB2セルに保存先フォルダを入力してください。" & vbCrLf & _
           "（例：C:\temp\inventory）", vbExclamation
    Exit Sub
End If

' ② パスの正規化：/ を \ に変換 ＋ 末尾に \ を付ける
outDir = Replace(outDir, "/", "\")
If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"

' ③ フォルダ作成（失敗時のメッセージを具体化）
If Not EnsureFolder(outDir) Then
    MsgBox "保存先フォルダを作成または利用できませんでした。" & vbCrLf & _
           "パスやアクセス権限、セキュリティソフトの設定を確認してください。" & vbCrLf & _
           "対象: " & outDir, vbExclamation
    Exit Sub
End If

' ④ 発注リスト空チェック：Cells→UsedRangeに変更（高速かつ意図に合う）
Dim dataR As Range
Set dataR = GetOrderDataRange(wsO)
If dataR Is Nothing Then
    MsgBox "発注リストが空です。先に［発注リスト作成］を実行してください。", vbExclamation
    Exit Sub
End If

Set usedR = dataR

    
    ' 一時ブックに貼り付けてCSV保存（UTF-8）
    Set wbTmp = Application.Workbooks.Add
    usedR.Copy wbTmp.Sheets(1).Range("A1")
    
    outFile = outDir & "発注_" & Format(Now, "yyyymmdd_hhnn") & ".csv"
    On Error GoTo ErrH
    wbTmp.SaveAs Filename:=outFile, FileFormat:=62 ' xlCSVUTF8=62
    wbTmp.Close SaveChanges:=False
    MsgBox "CSVを保存しました: " & outFile, vbInformation
    Exit Sub
ErrH:
    On Error Resume Next
    wbTmp.Close SaveChanges:=False
    MsgBox "CSV保存でエラー: " & Err.Description, vbExclamation
End Sub

Sub SanityCheck()
    On Error Resume Next
    Dim msg As String
    msg = ""
    msg = msg & "シート存在: " & _
        CStr(Not ThisWorkbook.Sheets("品目マスタ") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("入出庫") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("在庫") Is Nothing) & ", " & _
        CStr(Not ThisWorkbook.Sheets("設定") Is Nothing) & vbCrLf
    
    Dim lastM As Long, lastIO As Long
    lastM = ThisWorkbook.Sheets("品目マスタ").Cells(Rows.Count, 1).End(xlUp).Row
    lastIO = ThisWorkbook.Sheets("入出庫").Cells(Rows.Count, 1).End(xlUp).Row
    msg = msg & "最終行: マスタ=" & lastM & " / 入出庫=" & lastIO & vbCrLf
    
    Dim outDir As String
    outDir = CStr(ThisWorkbook.Sheets("設定").Range("B2").Value)
    msg = msg & "出力フォルダ: [" & outDir & "]" & vbCrLf
    
    MsgBox msg, vbInformation, "SanityCheck"
End Sub


Option Explicit

' 共通：1枚のシートに基本フォーマットを適用
Private Sub FS_FormatCommon(ByVal ws As Worksheet)
    With ws.Cells
        .Font.name = "Meiryo"
        .Font.Size = 10
        .RowHeight = 20
    End With
    ws.Rows(1).RowHeight = 24

    ' 先頭行固定（安全に）
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

' 在庫 A:E（SKU, 品名, 現在庫, 安全在庫, 不足数）
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

    ' 不足(E列>0)を薄赤で強調
    Dim lastRow As Long, rng As Range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1000
    Set rng = ws.Range("A2:E" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$E2>0")
        .Interior.Color = RGB(255, 200, 200)
    End With
End Sub

' 発注リスト A:C（SKU, 品名, 発注数）
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

' 設定 A:B（A1=安全在庫既定値, A2=出力フォルダ）
Private Sub FS_FormatSetting(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:A2")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 48
End Sub

' 品目マスタ A:C（SKU, 品名, 安全在庫）
Private Sub FS_FormatMaster(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:C1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12   ' SKU
    ws.Columns("B").ColumnWidth = 28   ' 品名
    ws.Columns("C").ColumnWidth = 10   ' 安全在庫
    ws.Columns("C").NumberFormatLocal = "#,##0"
End Sub

' 入出庫（想定列：A=日付, B=種別(入/出), C=SKU, D=数量, E=備考）
Private Sub FS_FormatIO(ByVal ws As Worksheet)
    Call FS_FormatCommon(ws)
    With ws.Range("A1:E1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Columns("A").ColumnWidth = 12   ' 日付
    ws.Columns("B").ColumnWidth = 8    ' 種別
    ws.Columns("C").ColumnWidth = 12   ' SKU
    ws.Columns("D").ColumnWidth = 10   ' 数量
    ws.Columns("E").ColumnWidth = 24   ' 備考

    ' 表示形式
    ws.Columns("A").NumberFormatLocal = "yyyy/mm/dd"
    ws.Columns("D").NumberFormatLocal = "#,##0"

    ' （任意）種別の入力規則（入/出のみ）
    On Error Resume Next
    With ws.Range("B2:B1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="入,出"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "選択"
        .InputMessage = "入 or 出 を選択"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "「入」または「出」以外は入力できません。"
        .ShowInput = True
        .ShowError = True
    End With
    On Error GoTo 0
End Sub

Public Sub FormatScreensWithReport_NoHelper()
    Dim ws As Worksheet
    Dim done As String, missing As String

    ' 在庫
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("在庫"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatStock(ws): done = done & "・在庫" & vbCrLf
    Else
        missing = missing & "・在庫" & vbCrLf
    End If

    ' 発注リスト
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("発注リスト"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatOrder(ws): done = done & "・発注リスト" & vbCrLf
    Else
        missing = missing & "・発注リスト" & vbCrLf
    End If

    ' 設定
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("設定"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatSetting(ws): done = done & "・設定" & vbCrLf
    Else
        missing = missing & "・設定" & vbCrLf
    End If

    ' 品目マスタ
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("品目マスタ"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatMaster(ws): done = done & "・品目マスタ" & vbCrLf
    Else
        missing = missing & "・品目マスタ" & vbCrLf
    End If

    ' 入出庫
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("入出庫"): On Error GoTo 0
    If Not ws Is Nothing Then
        Call FS_FormatIO(ws): done = done & "・入出庫" & vbCrLf
    Else
        missing = missing & "・入出庫" & vbCrLf
    End If

    ' レポート
    Dim msg As String: msg = "整形が完了しました。" & vbCrLf & vbCrLf
    If Len(done) > 0 Then msg = msg & "整形したシート：" & vbCrLf & done & vbCrLf
    If Len(missing) > 0 Then msg = msg & "見つからなかったシート：" & vbCrLf & missing & _
        "(タブ名の表記揺れやスペース・全角/半角を確認してください)"
    MsgBox msg, vbInformation, "FormatScreens"
End Sub


