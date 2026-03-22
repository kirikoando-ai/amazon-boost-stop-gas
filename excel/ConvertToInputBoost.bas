Attribute VB_Name = "ConvertToInputBoost"
Option Explicit

' Amazon SP bulk sheet -> input_boost CSV
' Usage:
' 1) Open bulk xlsx in Excel
' 2) Import this module to VBA project
' 3) Run ExportInputBoostCsv

Public Sub ExportInputBoostCsv()
    Dim ws As Worksheet
    Dim outPath As Variant
    Dim lastRow As Long, r As Long
    Dim headerMap As Object
    Dim fnum As Integer
    Dim rowOut(1 To 17) As String
    Dim targetType As String, targetText As String
    Dim productKey As String

    On Error GoTo EH

    Set ws = ResolveSpSheet()
    If ws Is Nothing Then
        MsgBox "SPシートが見つかりません。シート名を確認してください。", vbExclamation
        Exit Sub
    End If

    Set headerMap = BuildHeaderMap(ws)
    If headerMap.Count = 0 Then
        MsgBox "ヘッダー行の取得に失敗しました。", vbExclamation
        Exit Sub
    End If

    outPath = Application.GetSaveAsFilename( _
        InitialFileName:="input_boost_from_bulk.csv", _
        FileFilter:="CSV UTF-8 (*.csv), *.csv")
    If VarType(outPath) = vbBoolean Then
        Exit Sub
    End If

    fnum = FreeFile
    Open CStr(outPath) For Output As #fnum

    Print #fnum, JoinCsv(Array( _
        "boost_campaign_id", "boost_campaign_name", "boost_ad_group_id", "boost_ad_group_name", _
        "product_key", "product_label", "target_type", "target_text", "match_type", _
        "impressions", "clicks", "orders", "sales", "spend", "cpc", "target_key", "existing_exact_in_normal"))

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Nz(GetByHeader(ws, headerMap, r, "プロダクト")) <> "スポンサープロダクト広告" Then GoTo ContinueRow

        targetType = ""
        targetText = ""

        Select Case Nz(GetByHeader(ws, headerMap, r, "エンティティ"))
            Case "キーワード"
                targetType = "keyword"
                targetText = Nz(GetByHeader(ws, headerMap, r, "キーワードテキスト"))
            Case "商品ターゲティング"
                targetType = "product_targeting"
                targetText = Nz(GetByHeader(ws, headerMap, r, "商品ターゲティング式"))
                If targetText = "" Then
                    targetText = Nz(GetByHeader(ws, headerMap, r, "解決済みの商品ターゲティング式（情報提供のみ）"))
                End If
            Case Else
                GoTo ContinueRow
        End Select

        If targetText = "" Then GoTo ContinueRow

        productKey = Nz(GetByHeader(ws, headerMap, r, "ASIN（情報提供のみ）"))
        If productKey = "" Then productKey = Nz(GetByHeader(ws, headerMap, r, "SKU"))
        If productKey = "" Then productKey = Nz(GetByHeader(ws, headerMap, r, "広告グループ名"))
        If productKey = "" Then productKey = Nz(GetByHeader(ws, headerMap, r, "キャンペーンID"))

        rowOut(1) = Nz(GetByHeader(ws, headerMap, r, "キャンペーンID"))
        rowOut(2) = FirstNonEmpty(Array( _
            Nz(GetByHeader(ws, headerMap, r, "キャンペーン名")), _
            Nz(GetByHeader(ws, headerMap, r, "キャンペーン名（情報提供のみ）"))))
        rowOut(3) = Nz(GetByHeader(ws, headerMap, r, "広告グループID"))
        rowOut(4) = FirstNonEmpty(Array( _
            Nz(GetByHeader(ws, headerMap, r, "広告グループ名")), _
            Nz(GetByHeader(ws, headerMap, r, "広告グループ名（情報提供のみ）"))))
        rowOut(5) = productKey
        rowOut(6) = Nz(GetByHeader(ws, headerMap, r, "SKU"))
        rowOut(7) = targetType
        rowOut(8) = targetText
        rowOut(9) = Nz(GetByHeader(ws, headerMap, r, "マッチタイプ"))
        rowOut(10) = Nz(GetByHeader(ws, headerMap, r, "インプレッション数"))
        rowOut(11) = Nz(GetByHeader(ws, headerMap, r, "クリック数"))
        rowOut(12) = Nz(GetByHeader(ws, headerMap, r, "注文"))
        rowOut(13) = Nz(GetByHeader(ws, headerMap, r, "売上"))
        rowOut(14) = Nz(GetByHeader(ws, headerMap, r, "支出"))
        rowOut(15) = Nz(GetByHeader(ws, headerMap, r, "CPC"))
        rowOut(16) = ""
        rowOut(17) = ""

        Print #fnum, JoinCsv(rowOut)

ContinueRow:
    Next r

    Close #fnum
    MsgBox "CSV出力が完了しました: " & CStr(outPath), vbInformation
    Exit Sub

EH:
    On Error Resume Next
    If fnum > 0 Then Close #fnum
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Function ResolveSpSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "スポンサープロダクト広告キャンペーン" Then
            Set ResolveSpSheet = ws
            Exit Function
        End If
    Next ws
    Set ResolveSpSheet = ThisWorkbook.Worksheets(2)
End Function

Private Function BuildHeaderMap(ws As Worksheet) As Object
    Dim d As Object
    Dim lastCol As Long, c As Long
    Dim h As String
    Set d = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim(CStr(ws.Cells(1, c).Value))
        If h <> "" Then d(h) = c
    Next c
    Set BuildHeaderMap = d
End Function

Private Function GetByHeader(ws As Worksheet, d As Object, rowNum As Long, headerName As String) As Variant
    If d.Exists(headerName) Then
        GetByHeader = ws.Cells(rowNum, CLng(d(headerName))).Value
    Else
        GetByHeader = ""
    End If
End Function

Private Function Nz(v As Variant) As String
    If IsError(v) Then
        Nz = ""
    ElseIf IsNull(v) Then
        Nz = ""
    Else
        Nz = Trim(CStr(v))
    End If
End Function

Private Function FirstNonEmpty(values As Variant) As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Trim(CStr(values(i))) <> "" Then
            FirstNonEmpty = Trim(CStr(values(i)))
            Exit Function
        End If
    Next i
    FirstNonEmpty = ""
End Function

Private Function EscapeCsv(value As String) As String
    Dim s As String
    s = Replace(value, """", """""")
    EscapeCsv = """" & s & """"
End Function

Private Function JoinCsv(values As Variant) As String
    Dim i As Long
    Dim parts() As String
    ReDim parts(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        parts(i) = EscapeCsv(CStr(values(i)))
    Next i
    JoinCsv = Join(parts, ",")
End Function
